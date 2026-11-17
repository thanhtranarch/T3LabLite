# -*- coding: utf-8 -*-
"""T3Lab Data Hub, store side: ingest the snapshots Revit wrote
(core/hub_snapshot.py) into a local SQLite database and answer summary /
query / diff from it without Revit (docs/data-hub.md).

Runs in the bridge process (CPython 3, outside Revit); deployed next to
bridge.py as a companion. Stdlib only.

Layout under hub_dir():
  inbox/<name>.ndjson                  written by Revit, ingested then removed
  snapshots/<model_key>/<name>.ndjson.gz   the source of truth
  hub.db                               cache: rebuild() recreates it from snapshots/

hub.db never lives on a synced or network folder (SQLite locks are not
reliable there); share the snapshot files instead.

Versioned rows: an element has one row per distinct content, valid for
snapshots first_seq..last_seq, so an unchanged element costs one row however
many snapshots hold it. Alive at S = first_seq <= S <= last_seq.

Test: lite_guard/test_data_hub.py.
"""
import calendar
import fnmatch
import gzip
import hashlib
import json
import os
import shutil
import sqlite3
import threading
import time

SCHEMA_VERSION = 1
FORMAT = 'hub/1'
DEFAULT_LIMIT, MAX_LIMIT = 50, 500
MAX_GROUPS = 200
MAX_PARAM_CHANGES = 8           # per element in a diff sample
BATCH = 5000
FIELDS = ('category', 'family', 'type', 'level', 'workset', 'phase', 'name')
_SHORT = {'c': 'category', 'f': 'family', 'y': 'type', 'l': 'level', 'ws': 'workset',
          'ph': 'phase', 'n': 'name'}
_KINDS = {'instance': 'i', 'type': 't', 'i': 'i', 't': 't'}
OPS = ('=', '!=', 'contains', 'not_contains', 'empty', 'not_empty', '>', '>=', '<', '<=')

_SCHEMA = """
CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS models (
    model_key TEXT PRIMARY KEY, title TEXT, identity_kind TEXT, last_seq INTEGER);
CREATE TABLE IF NOT EXISTS snapshots (
    model_key TEXT, seq INTEGER, name TEXT UNIQUE, taken_at TEXT, title TEXT, revit TEXT,
    workshared INTEGER, modified INTEGER, version_guid TEXT, saves INTEGER,
    elements INTEGER, types INTEGER, warnings INTEGER, seconds REAL,
    scope TEXT, units TEXT, PRIMARY KEY (model_key, seq));
CREATE TABLE IF NOT EXISTS elements (
    model_key TEXT, uid TEXT, first_seq INTEGER, last_seq INTEGER, hash TEXT, eid INTEGER,
    kind TEXT, category TEXT, family TEXT, type TEXT, level TEXT, workset TEXT,
    phase TEXT, name TEXT, params TEXT, PRIMARY KEY (model_key, uid, first_seq));
CREATE INDEX IF NOT EXISTS elements_alive ON elements (model_key, last_seq, category);
CREATE TABLE IF NOT EXISTS warnings (
    model_key TEXT, seq INTEGER, description TEXT, severity TEXT, uids TEXT, key TEXT);
CREATE INDEX IF NOT EXISTS warnings_seq ON warnings (model_key, seq);
"""


class HubError(Exception):
    """A refusal the agent can act on: message + stable code (+ hint)."""

    def __init__(self, message, code, **extra):
        Exception.__init__(self, message)
        self.payload = dict({'error': message, 'error_code': code, 'retryable': False}, **extra)


def hub_dir():
    """Same rule as core/hub_snapshot.py: %LOCALAPPDATA%\\T3Lab\\hub or T3LAB_HUB_DIR."""
    override = os.environ.get('T3LAB_HUB_DIR')
    if override:
        return override
    base = os.environ.get('LOCALAPPDATA') or os.path.join(os.path.expanduser('~'), '.t3lab')
    return os.path.join(base, 'T3Lab', 'hub')


def element_hash(rec):
    """Content hash of one element record (id and UniqueId excluded)."""
    body = [rec.get('k') or 'i'] + [rec.get(k) for k in ('c', 'f', 'y', 'l', 'ws', 'ph', 'n')]
    body.append(rec.get('p') or {})
    text = json.dumps(body, sort_keys=True, ensure_ascii=True, separators=(',', ':'))
    return hashlib.sha1(text.encode('ascii')).hexdigest()[:16]


def age_hours(taken_at, now=None):
    try:
        then = calendar.timegm(time.strptime(taken_at, '%Y-%m-%dT%H:%M:%SZ'))
    except (TypeError, ValueError):
        return None
    return round(max(0.0, (now if now is not None else time.time()) - then) / 3600.0, 1)


def _num(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def match(value, op, wanted):
    """One `where` test on a parameter value (None = the element has none)."""
    if op == 'empty':
        return value is None or value == ''
    if op == 'not_empty':
        return value is not None and value != ''
    if op in ('>', '>=', '<', '<='):
        a, b = _num(value), _num(wanted)
        if a is None or b is None:
            return False
        return {'>': a > b, '>=': a >= b, '<': a < b, '<=': a <= b}[op]
    if op in ('contains', 'not_contains'):
        hit = value is not None and str(wanted).lower() in str(value).lower()
        return hit if op == 'contains' else not hit
    a, b = _num(value), _num(wanted)
    if a is not None and b is not None:
        same = abs(a - b) < 1e-9
    else:
        same = value is not None and str(value).strip().lower() == str(wanted).strip().lower()
    return same if op == '=' else not same


class Hub(object):
    """One hub folder. `allowed(title)` hides models an agent may not read."""

    def __init__(self, root=None):
        self.root = root or hub_dir()
        self.inbox = os.path.join(self.root, 'inbox')
        self.archive = os.path.join(self.root, 'snapshots')
        self.db_path = os.path.join(self.root, 'hub.db')
        self._lock = threading.Lock()       # one ingest at a time per process

    # ── connection ──
    def _connect(self):
        os.makedirs(self.root, exist_ok=True)
        con = sqlite3.connect(self.db_path, timeout=30)
        con.row_factory = sqlite3.Row
        con.executescript(_SCHEMA)
        row = con.execute("SELECT value FROM meta WHERE key='schema'").fetchone()
        if row is None:
            con.execute("INSERT INTO meta VALUES ('schema', ?)", (str(SCHEMA_VERSION),))
            con.commit()
        elif row['value'] != str(SCHEMA_VERSION):
            con.close()
            raise HubError('hub.db has schema {} (this bridge: {}).'.format(
                row['value'], SCHEMA_VERSION), 'hub_schema',
                hint='Run `bridge.py hub rebuild` to recreate it from the snapshots.')
        return con

    # ── ingest ──
    def pending(self):
        try:
            return sorted(n for n in os.listdir(self.inbox) if n.endswith('.ndjson'))
        except OSError:
            return []

    def ingest_pending(self):
        """Ingest every complete inbox file, oldest first. Returns
        {'ingested': [...], 'errors': [...]}; never raises for one bad file."""
        out = {'ingested': [], 'errors': []}
        names = self.pending()
        if not names:
            return out
        with self._lock:
            # File names start with the model key, so sort by the UTC stamp.
            for name in sorted(names, key=lambda n: n.rsplit('-', 1)[-1]):
                path = os.path.join(self.inbox, name)
                try:
                    out['ingested'].append(self._ingest(path, move=True))
                except HubError as e:
                    out['errors'].append(dict(e.payload, file=name))
                    self._set_aside(path, e.payload.get('error_code') == 'hub_out_of_order')
                except Exception as e:      # unreadable / half-copied file
                    out['errors'].append({'file': name, 'error': str(e)})
                    self._set_aside(path, False)
        return out

    def _set_aside(self, path, keep):
        """A file that cannot be ingested now leaves the inbox, so it is
        reported once: an older snapshot goes to snapshots/ (rebuild() takes
        it in time order), anything else to rejected/."""
        try:
            name = os.path.basename(path).split('.ndjson')[0]
            if keep:
                self._archive(path, name.rsplit('-', 1)[0], name)
                return
            folder = os.path.join(self.root, 'rejected')
            os.makedirs(folder, exist_ok=True)
            os.replace(path, os.path.join(folder, os.path.basename(path)))
        except Exception:
            pass

    def _open(self, path):
        if path.endswith('.gz'):
            return gzip.open(path, 'rt', encoding='utf-8')
        return open(path, encoding='utf-8')

    def _read_frame(self, path):
        """(header, end) of a snapshot file; raises HubError when incomplete."""
        header, last = None, ''
        with self._open(path) as f:
            for line in f:
                if header is None:
                    header = json.loads(line)
                elif line.strip():
                    last = line         # key order is not fixed (IronPython dicts)
        try:
            end = json.loads(last) if last else None
        except ValueError:
            end = None
        if not isinstance(end, dict) or end.get('t') != 'end':
            end = None
        if not header or header.get('t') != 'header' or header.get('format') != FORMAT:
            raise HubError('Not a {} snapshot.'.format(FORMAT), 'hub_bad_snapshot')
        if not end or not end.get('complete'):
            raise HubError('Snapshot is incomplete (no end line).', 'hub_bad_snapshot')
        return header, end

    def _ingest(self, path, move):
        name = os.path.basename(path).split('.ndjson')[0]
        header, end = self._read_frame(path)
        key = header['model_key']
        con = self._connect()
        try:
            con.execute('BEGIN IMMEDIATE')
            if con.execute('SELECT 1 FROM snapshots WHERE name=?', (name,)).fetchone():
                con.rollback()
                if move:        # another bridge process ingested it, or died before archiving
                    if os.path.exists(os.path.join(self.archive, key, name + '.ndjson.gz')):
                        os.remove(path)
                    else:
                        self._archive(path, key, name)
                return {'snapshot': name, 'model_key': key, 'already_ingested': True}
            model = con.execute('SELECT last_seq FROM models WHERE model_key=?', (key,)).fetchone()
            prev_seq = model['last_seq'] if model else 0
            if prev_seq:
                last = con.execute('SELECT taken_at FROM snapshots WHERE model_key=? AND seq=?',
                                   (key, prev_seq)).fetchone()
                if last and header.get('taken_at', '') <= last['taken_at']:
                    raise HubError(
                        'Snapshot {} is older than the latest one of this model.'.format(name),
                        'hub_out_of_order',
                        hint='Run `bridge.py hub rebuild` to re-ingest everything in time order.')
            seq = prev_seq + 1
            previous = dict((r['uid'], (r['hash'], r['first_seq'])) for r in con.execute(
                'SELECT uid, hash, first_seq FROM elements WHERE model_key=? AND last_seq=?',
                (key, prev_seq)))
            inserts, extends, warns = [], [], []
            counts = {'new': 0, 'changed': 0, 'unchanged': 0}
            seen = set()

            def flush():
                con.executemany('INSERT INTO elements VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                                inserts)
                con.executemany('UPDATE elements SET last_seq=? WHERE model_key=? AND uid=? '
                                'AND first_seq=?', extends)
                del inserts[:], extends[:]

            with self._open(path) as f:
                for line in f:
                    rec = json.loads(line)
                    kind = rec.get('t')
                    if kind == 'e':
                        uid = rec.get('u')
                        if not uid or uid in seen:
                            continue
                        seen.add(uid)
                        digest = element_hash(rec)
                        before = previous.get(uid)
                        if before and before[0] == digest:
                            extends.append((seq, key, uid, before[1]))
                            counts['unchanged'] += 1
                        else:
                            counts['changed' if before else 'new'] += 1
                            inserts.append((
                                key, uid, seq, seq, digest, rec.get('i'), rec.get('k') or 'i',
                                rec.get('c'), rec.get('f'), rec.get('y'), rec.get('l'),
                                rec.get('ws'), rec.get('ph'), rec.get('n'),
                                json.dumps(rec.get('p') or {}, ensure_ascii=False,
                                           sort_keys=True, separators=(',', ':'))))
                        if len(inserts) + len(extends) >= BATCH:
                            flush()
                    elif kind == 'w':
                        uids = sorted(rec.get('u') or [])
                        warns.append((key, seq, rec.get('d'), rec.get('s'), json.dumps(uids),
                                      hashlib.sha1(json.dumps([rec.get('d'), uids]).encode(
                                          'utf-8')).hexdigest()[:16]))
            flush()
            con.executemany('INSERT INTO warnings VALUES (?,?,?,?,?,?)', warns)
            con.execute(
                'INSERT INTO snapshots VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                (key, seq, name, header.get('taken_at'), header.get('title'),
                 header.get('revit'), int(bool(header.get('workshared'))),
                 int(bool(header.get('modified'))), header.get('version_guid'),
                 header.get('saves'), end.get('elements'), end.get('types'), len(warns),
                 end.get('seconds'), json.dumps(header.get('scope') or {}),
                 json.dumps(end.get('units') or {})))
            con.execute('INSERT OR REPLACE INTO models VALUES (?,?,?,?)',
                        (key, header.get('title'), header.get('identity_kind'), seq))
            con.commit()
        except Exception:
            con.rollback()
            raise
        finally:
            con.close()
        if move:
            self._archive(path, key, name)
        removed = len(previous) - counts['unchanged'] - counts['changed']
        return {'snapshot': name, 'model_key': key, 'seq': seq, 'document': header.get('title'),
                'elements': end.get('elements'), 'types': end.get('types'),
                'warnings': len(warns),
                'since_previous': (dict(counts, removed=removed) if prev_seq else None)}

    def _archive(self, path, key, name):
        folder = os.path.join(self.archive, key)
        os.makedirs(folder, exist_ok=True)
        target = os.path.join(folder, name + '.ndjson.gz')
        tmp = target + '.tmp'
        with open(path, 'rb') as src, gzip.open(tmp, 'wb') as dst:
            shutil.copyfileobj(src, dst)
        os.replace(tmp, target)
        os.remove(path)

    def rebuild(self):
        """Recreate hub.db from snapshots/ (+ inbox), in time order."""
        with self._lock:
            for suffix in ('', '-wal', '-shm', '-journal'):
                try:
                    os.remove(self.db_path + suffix)
                except OSError:
                    pass
            files = []
            for folder, _dirs, names in os.walk(self.archive):
                files += [os.path.join(folder, n) for n in names if n.endswith('.ndjson.gz')]
            out = {'ingested': [], 'errors': []}
            for path in sorted(files, key=lambda p: os.path.basename(p).split('.')[0]
                               .rsplit('-', 1)[-1]):
                try:
                    out['ingested'].append(self._ingest(path, move=False)['snapshot'])
                except Exception as e:
                    out['errors'].append({'file': os.path.basename(path), 'error': str(e)})
        later = self.ingest_pending()
        out['ingested'] += [i['snapshot'] for i in later['ingested']]
        out['errors'] += later['errors']
        return out

    # ── resolve ──
    def _models(self, con, allowed=None):
        rows = [dict(r) for r in con.execute('SELECT * FROM models ORDER BY title')]
        return [m for m in rows if allowed is None or allowed(m['title'] or '')]

    def _model(self, con, wanted, allowed=None):
        models = self._models(con, allowed)
        if not models:
            raise HubError('The hub has no snapshot yet.', 'hub_empty',
                           hint='Publish one: t3lab_hub action=publish (Revit must be open).')
        if not wanted:
            if len(models) == 1:
                return models[0]
            raise HubError('Several models are in the hub; pass model.', 'hub_model_required',
                           models=[m['title'] for m in models])
        text = str(wanted).strip().lower()
        hits = [m for m in models if m['model_key'] == text or (m['title'] or '').lower() == text]
        if not hits:
            hits = [m for m in models if fnmatch.fnmatch((m['title'] or '').lower(), text)
                    or text in (m['title'] or '').lower()]
        if len(hits) == 1:
            return hits[0]
        raise HubError('{} model(s) match "{}".'.format(len(hits), wanted),
                       'hub_model_not_found' if not hits else 'hub_model_ambiguous',
                       models=[m['title'] for m in (hits or models)])

    def _snapshot(self, con, model, wanted=None):
        """Row of one snapshot: None/'latest', 'previous', a seq, -1 (= previous) or a name."""
        key, last = model['model_key'], model['last_seq']
        seq = None
        if wanted in (None, '', 'latest'):
            seq = last
        elif wanted == 'previous':
            seq = last - 1
        else:
            try:
                seq = int(wanted)
                if seq <= 0:
                    seq = last + seq
            except (TypeError, ValueError):
                row = con.execute('SELECT * FROM snapshots WHERE model_key=? AND name=?',
                                  (key, str(wanted))).fetchone()
                if row:
                    return row
        row = con.execute('SELECT * FROM snapshots WHERE model_key=? AND seq=?',
                          (key, seq)).fetchone() if seq else None
        if row is None:
            raise HubError('No snapshot "{}" for {} (it has {}).'.format(
                wanted, model['title'], last), 'hub_snapshot_not_found',
                hint='Snapshots are numbered 1..{}; "latest" and "previous" also work.'.format(last))
        return row

    @staticmethod
    def _snap_view(row, now=None):
        return {'seq': row['seq'], 'name': row['name'], 'as_of': row['taken_at'],
                'age_hours': age_hours(row['taken_at'], now),
                'unsaved_changes': bool(row['modified'])}

    # ── read ──
    def summary(self, model=None, allowed=None, now=None):
        con = self._connect()
        try:
            if not model:
                models = self._models(con, allowed)
                out = []
                for m in models:
                    snap = self._snapshot(con, m)
                    out.append({'model': m['title'], 'model_key': m['model_key'],
                                'snapshots': m['last_seq'], 'as_of': snap['taken_at'],
                                'age_hours': age_hours(snap['taken_at'], now),
                                'elements': snap['elements'], 'warnings': snap['warnings'],
                                'revit': snap['revit']})
                return {'models': out, 'pending_inbox': len(self.pending())}
            m = self._model(con, model, allowed)
            snap = self._snapshot(con, m)
            key, seq = m['model_key'], snap['seq']
            alive = 'model_key=? AND first_seq<=? AND last_seq>=?'

            def top(field, limit=25, kind='i'):
                rows = con.execute(
                    'SELECT {f} AS v, COUNT(*) AS n FROM elements WHERE {a} AND kind=? '
                    'GROUP BY {f} ORDER BY n DESC LIMIT ?'.format(f=field, a=alive),
                    (key, seq, seq, kind, limit))
                return [{'value': r['v'], 'count': r['n']} for r in rows]

            warns = con.execute(
                'SELECT description AS d, COUNT(*) AS n FROM warnings WHERE model_key=? AND seq=? '
                'GROUP BY description ORDER BY n DESC LIMIT 10', (key, seq))
            history = con.execute(
                'SELECT seq, taken_at, elements, warnings FROM snapshots WHERE model_key=? '
                'ORDER BY seq DESC LIMIT 10', (key,))
            return {
                'model': m['title'], 'model_key': key, 'identity': m['identity_kind'],
                'snapshot': self._snap_view(snap, now), 'revit': snap['revit'],
                'workshared': bool(snap['workshared']), 'scope': json.loads(snap['scope'] or '{}'),
                'elements': snap['elements'], 'types': snap['types'],
                'warnings': snap['warnings'],
                'by_category': top('category'), 'by_level': top('level', 15),
                'top_warnings': [{'description': r['d'], 'count': r['n']} for r in warns],
                'history': [dict(r) for r in history]}
        finally:
            con.close()

    def query(self, model=None, snapshot=None, kind='instance', where=None, group_by=None,
              sum_param=None, select=None, limit=None, offset=0, allowed=None, now=None, **filters):
        """Elements alive at one snapshot. filters: category/family/type/level/
        workset/phase (exact, case-insensitive) + name_contains. where: [{param,
        op, value}]. group_by: fields or 'param:<Name>' -> counts (+ `sum_param`, a
        numeric parameter summed per group). Else rows with `select` parameters."""
        where = self._clean_where(where)
        group_by = [group_by] if isinstance(group_by, str) else list(group_by or [])
        select = [select] if isinstance(select, str) else list(select or [])
        for g in group_by:
            if g not in FIELDS and not g.startswith('param:'):
                raise HubError('group_by "{}" is not a field ({}) or "param:<Name>".'.format(
                    g, ', '.join(FIELDS)), 'bad_arguments')
        try:
            limit = max(1, min(MAX_LIMIT, int(limit or DEFAULT_LIMIT)))
            offset = max(0, int(offset or 0))
        except (TypeError, ValueError):
            raise HubError('limit and offset must be integers.', 'bad_arguments')
        con = self._connect()
        try:
            m = self._model(con, model, allowed)
            snap = self._snapshot(con, m, snapshot)
            sql = ['SELECT * FROM elements WHERE model_key=? AND first_seq<=? AND last_seq>=?']
            args = [m['model_key'], snap['seq'], snap['seq']]
            if kind not in (None, '', 'any'):
                if kind not in _KINDS:
                    raise HubError('kind is instance, type or any.', 'bad_arguments')
                sql.append('AND kind=?')
                args.append(_KINDS[kind])
            for field in FIELDS[:-1]:
                value = filters.get(field)
                if value not in (None, ''):
                    sql.append('AND LOWER({})=?'.format(field))
                    args.append(str(value).strip().lower())
            if filters.get('name_contains'):
                sql.append("AND LOWER(COALESCE(name,'')) LIKE ?")
                args.append('%{}%'.format(str(filters['name_contains']).lower()))
            sql.append('ORDER BY category, eid')
            need_params = bool(where or select or sum_param
                               or any(g.startswith('param:') for g in group_by))
            matched, groups, rows = 0, {}, []
            for r in con.execute(' '.join(sql), args):
                params = json.loads(r['params']) if need_params else {}
                if where and not all(match(params.get(w['param']), w['op'], w.get('value'))
                                     for w in where):
                    continue
                matched += 1
                if group_by:
                    gkey = tuple(params.get(g[6:]) if g.startswith('param:') else r[g]
                                 for g in group_by)
                    slot = groups.setdefault(gkey, [0, 0.0])
                    slot[0] += 1
                    if sum_param:
                        slot[1] += _num(params.get(sum_param)) or 0.0
                elif offset < matched <= offset + limit:
                    row = {'uid': r['uid'], 'id': r['eid']}
                    row.update((f, r[f]) for f in FIELDS if r[f] is not None)
                    for name in select:
                        if name in params:
                            row.setdefault('params', {})[name] = params[name]
                    rows.append(row)
            out = {'model': m['title'], 'snapshot': self._snap_view(snap, now), 'total': matched}
            if group_by:
                ordered = sorted(groups.items(), key=lambda kv: (-kv[1][0], str(kv[0])))
                out['group_by'] = group_by
                out['groups'] = [dict(zip(group_by, k), count=v[0],
                                      **({'sum': round(v[1], 6)} if sum_param else {}))
                                 for k, v in ordered[:MAX_GROUPS]]
                out['groups_total'] = len(ordered)
                if sum_param:
                    out['sum_of'] = sum_param
                    unit = json.loads(snap['units'] or '{}').get(sum_param)
                    if unit:
                        out['unit'] = unit
            else:
                out['rows'] = rows
                end = offset + len(rows)
                out['next_offset'] = end if end < matched else None
            return out
        finally:
            con.close()

    @staticmethod
    def _clean_where(where):
        if not where:
            return []
        if isinstance(where, dict):
            where = [where]
        out = []
        for w in where:
            if not isinstance(w, dict) or not w.get('param'):
                raise HubError('where is a list of {param, op, value}.', 'bad_arguments')
            op = w.get('op') or '='
            if op not in OPS:
                raise HubError('op "{}" is not one of: {}.'.format(op, ' '.join(OPS)),
                               'bad_arguments')
            out.append({'param': w['param'], 'op': op, 'value': w.get('value')})
        return out

    def diff(self, model=None, since=None, to=None, category=None, limit=None,
             allowed=None, now=None):
        """What changed between two snapshots (default previous -> latest)."""
        try:
            limit = max(1, min(100, int(limit or 20)))
        except (TypeError, ValueError):
            raise HubError('limit must be an integer.', 'bad_arguments')
        con = self._connect()
        try:
            m = self._model(con, model, allowed)
            if m['last_seq'] < 2:
                raise HubError('{} has one snapshot; a diff needs two.'.format(m['title']),
                               'hub_one_snapshot',
                               hint='Publish again after the model changed, then diff.')
            b = self._snapshot(con, m, to)
            a = self._snapshot(con, m, since if since not in (None, '') else b['seq'] - 1)
            if a['seq'] == b['seq']:
                raise HubError('since and to are the same snapshot.', 'bad_arguments')
            if a['seq'] > b['seq']:
                a, b = b, a
            key = m['model_key']
            scope_a, scope_b = json.loads(a['scope'] or '{}'), json.loads(b['scope'] or '{}')
            cats = self._common(scope_a.get('categories'), scope_b.get('categories'))
            names = self._common(scope_a.get('parameters'), scope_b.get('parameters'),
                                 lower=False)
            wanted = str(category).strip().lower() if category else None

            def alive(seq):
                out = {}
                for r in con.execute(
                        'SELECT uid, first_seq, hash, kind, category FROM elements WHERE '
                        'model_key=? AND first_seq<=? AND last_seq>=?', (key, seq, seq)):
                    cat = (r['category'] or '').lower()
                    if (cats is not None and cat not in cats) or (wanted and cat != wanted):
                        continue
                    out[r['uid']] = (r['hash'], r['first_seq'], r['category'], r['kind'])
                return out

            old, new = alive(a['seq']), alive(b['seq'])
            added = [u for u in new if u not in old]
            removed = [u for u in old if u not in new]
            changed = [u for u in new if u in old and old[u][0] != new[u][0]]
            by_cat = {}
            for label, uids, source in (('added', added, new), ('removed', removed, old),
                                        ('changed', changed, new)):
                for u in uids:
                    by_cat.setdefault(source[u][2], {'added': 0, 'removed': 0, 'changed': 0})[
                        label] += 1

            def load(uid, first_seq):
                return con.execute('SELECT * FROM elements WHERE model_key=? AND uid=? AND '
                                   'first_seq=?', (key, uid, first_seq)).fetchone()

            def brief(r):
                out = {'uid': r['uid'], 'id': r['eid']}
                out.update((f, r[f]) for f in ('category', 'family', 'type', 'level', 'name')
                           if r[f] is not None)
                if r['kind'] == 't':
                    out['kind'] = 'type'
                return out

            param_counts, samples, real_changed = {}, [], 0
            for u in sorted(changed, key=lambda x: (new[x][2] or '', x)):
                ra, rb = load(u, old[u][1]), load(u, new[u][1])
                pa, pb = json.loads(ra['params']), json.loads(rb['params'])
                changes = {}
                for f in FIELDS[1:]:
                    if ra[f] != rb[f]:
                        changes[f] = [ra[f], rb[f]]
                for name in sorted(set(pa) | set(pb)):
                    if names is not None and name not in names:
                        continue
                    if pa.get(name) != pb.get(name):
                        changes[name] = [pa.get(name), pb.get(name)]
                if not changes:             # differed only outside the common scope
                    by_cat[new[u][2]]['changed'] -= 1
                    continue
                real_changed += 1
                for name in changes:
                    param_counts[name] = param_counts.get(name, 0) + 1
                if len(samples) < limit:
                    item = brief(rb)
                    shown = sorted(changes)[:MAX_PARAM_CHANGES]
                    item['changes'] = dict((n, changes[n]) for n in shown)
                    if len(changes) > len(shown):
                        item['more_changes'] = len(changes) - len(shown)
                    samples.append(item)
            out = {
                'model': m['title'],
                'since': self._snap_view(a, now), 'to': self._snap_view(b, now),
                'added': len(added), 'removed': len(removed), 'changed': real_changed,
                'by_category': sorted(
                    (dict(v, category=c) for c, v in by_cat.items() if any(v.values())),
                    key=lambda d: -(d['added'] + d['removed'] + d['changed'])),
                'changed_parameters': [{'name': n, 'elements': c} for n, c in sorted(
                    param_counts.items(), key=lambda kv: (-kv[1], kv[0]))[:25]],
                'changed_samples': samples,
                'added_samples': [brief(load(u, new[u][1])) for u in sorted(
                    added, key=lambda x: (new[x][2] or '', x))[:limit]],
                'removed_samples': [brief(load(u, old[u][1])) for u in sorted(
                    removed, key=lambda x: (old[x][2] or '', x))[:limit]],
                'warnings': self._warning_diff(con, key, a['seq'], b['seq'], limit)}
            notes = []
            if scope_a != scope_b:
                notes.append('The two snapshots were published with different scopes; only '
                             'the categories and parameters in both are compared.')
            units_a, units_b = json.loads(a['units'] or '{}'), json.loads(b['units'] or '{}')
            moved = sorted(n for n in units_a if n in units_b and units_a[n] != units_b[n])
            if moved:
                out['units_changed'] = moved[:25]
                notes.append('Project units changed for {} parameter(s): their value changes '
                             'may be unit changes.'.format(len(moved)))
            if wanted:
                out['category'] = category
            if notes:
                out['notes'] = notes
            return out
        finally:
            con.close()

    @staticmethod
    def _common(a, b, lower=True):
        """Names allowed by both scopes; None = no restriction."""
        if not a and not b:
            return None
        norm = (lambda s: str(s).lower()) if lower else str
        sets = [set(norm(x) for x in s) for s in (a, b) if s]
        return sets[0] & sets[-1]

    @staticmethod
    def _warning_diff(con, key, a, b, limit):
        def load(seq):
            return dict((r['key'], r['description']) for r in con.execute(
                'SELECT key, description FROM warnings WHERE model_key=? AND seq=?', (key, seq)))
        old, new = load(a), load(b)

        def top(keys, source):
            counts = {}
            for k in keys:
                counts[source[k]] = counts.get(source[k], 0) + 1
            return [{'description': d, 'count': n} for d, n in sorted(
                counts.items(), key=lambda kv: (-kv[1], str(kv[0])))[:min(limit, 10)]]

        fresh = [k for k in new if k not in old]
        gone = [k for k in old if k not in new]
        return {'before': len(old), 'after': len(new), 'new': len(fresh),
                'resolved': len(gone), 'new_top': top(fresh, new), 'resolved_top': top(gone, old)}
