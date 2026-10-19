# -*- coding: utf-8 -*-
"""Find the right tool for a request, in Vietnamese or English.

Offline: BM25 (knowledge.bm25_index, diacritics folded, bigrams) over one
chunk per tool — slug words, name, title_vi and keywords twice, description —
plus a phrase bonus for curated keywords found in the query. The bonus
compares text WITH diacritics when the query has them: folded, "tường" (wall)
and "đối tượng" (object) are the same token "tuong", and only the marks tell
"tô đỏ tường" apart from "chọn đối tượng".

Online (signed in): the toolhub_search RPC (pgroonga + trigram + keywords +
team popularity). Any failure falls back to offline without a word.

search(query, k, kinds, tools, client) -> [{slug, kind, name, score, why}]
"""
from __future__ import unicode_literals

import re

from Intelligence.knowledge import vi_text
from Intelligence.knowledge.bm25_index import BM25Index, doc_id_of
from Intelligence.toolhub import catalog

KEYWORD_BONUS = 2.5

_index_cache = {'sig': None, 'index': None, 'rows': None}


def _doc_text(row):
    words = (row.get('slug') or '').replace('_', ' ')
    title = row.get('title_vi') or ''
    kws = ' '.join(row.get('keywords') or [])
    return ' . '.join([words, row.get('name') or '', title, title, kws, kws,
                       row.get('description') or ''])


def build_index(rows):
    idx = BM25Index()
    for row in rows:
        idx.add_document(row['slug'], [{'page': 0, 'seq': 0, 'text': _doc_text(row)}])
    return idx


def _cached(rows):
    sig = tuple(sorted((r['slug'], r.get('updated_at') or '') for r in rows))
    if _index_cache['sig'] != sig:
        _index_cache.update(sig=sig, index=build_index(rows),
                            rows={r['slug']: r for r in rows})
    return _index_cache['index'], _index_cache['rows']


def _has_marks(text):
    return vi_text.fold_diacritics(text) != text


def _phrase_in(phrase, text):
    return re.search(r'(^|\W)' + re.escape(phrase) + r'(\W|$)', text, re.UNICODE) is not None


def keyword_bonus(query, row):
    q = (query or '').lower()
    marked = _has_marks(q)
    qf = vi_text.fold_diacritics(q)
    bonus = 0.0
    for kw in list(row.get('keywords') or []) + [row.get('title_vi') or '']:
        kw = (kw or '').lower().strip()
        if len(kw) < 2:
            continue
        hit = _phrase_in(kw, q) if marked else _phrase_in(vi_text.fold_diacritics(kw), qf)
        if hit:
            bonus += KEYWORD_BONUS * min(len(kw.split()), 3)
    return bonus


def search_offline(query, k=5, kinds=None, tools=None, rows=None):
    rows = rows if rows is not None else catalog.all_rows(tools)
    if kinds:
        rows = [r for r in rows if r.get('kind') in kinds]
    if not rows or not (query or '').strip():
        return []
    idx, by_slug = _cached(rows)
    scores = {}
    for key, score in idx.search(query, top_k=max(k * 6, 30)):
        scores[doc_id_of(key)] = score
    for slug, row in by_slug.items():
        b = keyword_bonus(query, row)
        if b:
            scores[slug] = scores.get(slug, 0.0) + b
    ranked = sorted(scores.items(), key=lambda kv: (-kv[1], kv[0]))[:k]
    out = []
    for slug, score in ranked:
        row = by_slug[slug]
        out.append({'slug': slug, 'kind': row.get('kind', 'mcp_tool'),
                    'name': row.get('name', slug), 'score': round(score, 4),
                    'why': 'bm25' + ('+kw' if keyword_bonus(query, row) else '')})
    return out


def search(query, k=5, kinds=None, tools=None, client=None):
    if client is not None:
        try:
            if client.signed_in():
                res = client.rpc('toolhub_search', {'q': query, 'kinds': kinds, 'k': k})
                if isinstance(res, list) and res:
                    return res
        except Exception:
            pass
    return search_offline(query, k=k, kinds=kinds, tools=tools)
