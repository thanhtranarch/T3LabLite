# -*- coding: utf-8 -*-
"""`/toolhub …` commands of the T3Lab Assistant — deterministic, no LLM.

    /toolhub                      status + help
    /toolhub login <email>        e-mail a 6-digit sign-in code
    /toolhub code <123456|link>   finish signing in (the digits, or the e-mail's link)
    /toolhub sync                 refresh the catalog cache, upload usage
    /toolhub search <words>       try the tool search
    /toolhub claim-admin          first member only: become admin
    /toolhub submit <slug>        send your draft recipe for review (the owner
                                  reviews it on t3lab.space, Revit tracking)
    /toolhub mine                 your drafts / submissions and the owner's notes
    /toolhub logout
    /save-tool <slug> [title]     save the last turn's tool calls as a recipe draft

handle(text, client=None) -> reply text, or None when `text` is not a
/toolhub command. Network calls happen here, so the dialog calls it from its
worker thread, like the other slash commands.
"""
from __future__ import unicode_literals

import re

from Intelligence.toolhub import auth, catalog, events, recipes, search
from Intelligence.toolhub.client import ToolHubClient, ToolHubError

HELP = (u'Easiest: type `/toolhub` (or pick Tool Hub in the `/` list) to open the Tool Hub '
        u'tab — sign in, sync, your tools and search in one window.\n'
        u'Commands: `/toolhub status` · `/toolhub login you@company.com` · `/toolhub code 123456` · '
        u'`/toolhub sync` · `/toolhub search <words>` · `/toolhub submit <slug>` · '
        u'`/toolhub mine` · '
        u'`/toolhub logout` · `/toolhub claim-admin` (first member only) · '
        u'`/save-tool <slug> [title]` (last turn -> recipe draft)')

_pending_email = [None]
# What the last /toolhub reply asked the user to type next: 'email' (after a
# bare `/toolhub login`), 'code' (after the sign-in e-mail went out) or None.
# Users answer that prompt with the bare value ("tran@wohhup.com.vn") rather
# than retyping the command — see handle_followup().
_awaiting = [None]

_BARE_EMAIL = re.compile(r'^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$')
_BARE_CODE = re.compile(r'^\d{6,10}$')


def parse(text):
    t = (text or u'').strip()
    if not t.lower().startswith(u'/toolhub'):
        return None
    rest = t[len(u'/toolhub'):]
    if rest and not rest[0].isspace():
        return None                       # "/toolhubx" is not ours
    parts = rest.strip().split(None, 1)
    action = parts[0].lower() if parts else u'status'
    arg = parts[1].strip() if len(parts) > 1 else u''
    return action, arg


def _status_text(client):
    st = auth.status()
    cache = catalog.load()
    n = len(cache['rows'])
    if not st:
        return (u'Tool Hub: signed out. {} tools in the local cache.\n'.format(n) + HELP)
    line = u'Tool Hub: signed in as {}. {} tools in the local cache'.format(st['email'], n)
    try:
        me = client.rpc('toolhub_me', {})
        line += u', role {}.'.format(me.get('role')) if isinstance(me, dict) else u'.'
    except ToolHubError as ex:
        line += u'. ({})'.format(ex)
    return line + u'\n' + HELP


def handle_followup(text, client=None):
    """Answer to the prompt the last /toolhub reply showed, typed bare.

    `/toolhub login` without an address asks for one; the user then types
    just the address. That reply used to fall through to the chat pipeline,
    which answered an e-mail address with "the LLM is unreachable". Returns the
    reply, or None when `text` is not such an answer — any other message
    drops the pending prompt, so a later e-mail address in a normal chat is
    never taken for a sign-in.
    """
    want = _awaiting[0]
    if want is None:
        return None
    t = (text or u'').strip()
    if want == u'email' and _BARE_EMAIL.match(t):
        return handle(u'/toolhub login ' + t, client)
    if want == u'code' and (_BARE_CODE.match(t) or auth.link_token(t)):
        return handle(u'/toolhub code ' + t, client)
    _awaiting[0] = None
    return None


def handle(text, client=None):
    parsed = parse(text)
    if parsed is None:
        return None
    action, arg = parsed
    _awaiting[0] = None
    client = client or ToolHubClient()
    try:
        if action in (u'status', u'help'):
            return _status_text(client)

        if action == u'login':
            if not arg.strip():
                _awaiting[0] = u'email'
                return (u'Tool Hub: type the e-mail address the sign-in code should go to '
                        u'(e.g. you@company.com).')
            email = auth.send_code(arg, client._http)
            _pending_email[0] = email
            _awaiting[0] = u'code'
            return (u'A sign-in e-mail was sent to {}. Type the 6-digit code from it — '
                    u'or, if it only has a link, right-click the link, Copy link (do not '
                    u'open it), and paste it here.'.format(email))

        if action == u'code':
            email = _pending_email[0] or (auth.status() or {}).get('email')
            if not email:
                return u'Ask for a code first: `/toolhub login you@company.com`.'
            st = auth.verify(email, arg, client._http)
            _pending_email[0] = None
            me = client.rpc('toolhub_me', {})
            changed, total = catalog.sync(client)
            role = me.get('role') if isinstance(me, dict) else u'viewer'
            return (u'Signed in to the Tool Hub as {} (role {}). Catalog synced: {} tools.'
                    .format(st['email'], role, total))

        if action == u'sync':
            changed, total = catalog.sync(client)
            sent = events.upload(client, force=True)
            return (u'Catalog synced: {} changed, {} tools in the cache. {} usage events uploaded.'
                    .format(changed, total, sent))

        if action == u'search':
            if not arg:
                return u'Type what you want to do: `/toolhub search tô đỏ tường`.'
            hits = search.search(arg, k=5, client=client)
            if not hits:
                return u'No tool matches "{}".'.format(arg)
            lines = [u'{}. `{}` ({}) — {}'.format(i + 1, h['slug'], h.get('kind', ''),
                                                 h.get('why', ''))
                     for i, h in enumerate(hits)]
            return u'Best matches for "{}":\n'.format(arg) + u'\n'.join(lines)

        if action == u'claim-admin':
            me = client.rpc('toolhub_claim_admin', {})
            return u'You are now the Tool Hub admin ({}).'.format(
                me.get('display_name') or auth.status()['email'])

        if action == u'submit':
            if not arg:
                return u'Which draft? `/toolhub submit <slug>`'
            res = client.rpc('toolhub_submit', {'p_slug': arg})
            return (u'Recipe {} sent for review. The extension owner reviews it on '
                    u't3lab.space (Revit tracking); until it is published only you see it. '
                    u'`/toolhub mine` shows the result.'.format(res.get('slug', arg)))

        if action == u'mine':
            try:
                catalog.sync(client)
            except ToolHubError:
                pass                      # offline: show the cache
            mine = sorted((r for r in catalog.load()['rows'].values() if r.get('mine')),
                          key=lambda r: r.get('slug', ''))
            if not mine:
                return u'You have no tools on the Tool Hub yet. Save one with `/save-tool <slug>`.'
            lines = [u'Your Tool Hub tools:']
            for r in mine:
                line = u'- `{}` — {}'.format(r['slug'], r.get('status', '?'))
                if r.get('review_note'):
                    line += u' · owner\'s note: {}'.format(r['review_note'])
                lines.append(line)
            return u'\n'.join(lines)

        if action in (u'logout', u'signout', u'sign-out'):
            auth.sign_out(client._http)
            return u'Signed out of the Tool Hub. The local catalog cache stays for offline search.'

        return u'Unknown Tool Hub command "{}".\n{}'.format(action, HELP)
    except (auth.AuthError, ToolHubError) as ex:
        # A mistyped address or code: keep waiting for the retyped value.
        if action == u'login':
            _awaiting[0] = u'email'
        elif action == u'code' and _pending_email[0]:
            _awaiting[0] = u'code'
        return u'Tool Hub: {}'.format(ex)


_SLUG = re.compile(r'^[a-z0-9_]{2,48}$')


def handle_save(text, calls, user_text, client=None):
    """`/save-tool <slug> [title]` -> reply, or None when not that command."""
    t = (text or u'').strip()
    if not t.lower().startswith(u'/save-tool'):
        return None
    parts = t[len(u'/save-tool'):].strip().split(None, 1)
    if not parts:
        return u'Name the recipe: `/save-tool <slug> [title]`, e.g. `/save-tool color_walls_red Tô đỏ tường`.'
    slug = parts[0].lower()
    title = parts[1].strip() if len(parts) > 1 else u''
    if not _SLUG.match(slug):
        return u'The slug may use a-z, 0-9 and _ (2-48 characters), e.g. color_walls_red.'
    built = recipes.build_from_calls(calls, user_text)
    if not built['steps']:
        return (u'The last turn has no successful tool calls to save. Do the task once, '
                u'then run /save-tool.')
    leaks = recipes.scan_secrets([s['args'] for s in built['steps']], 'steps')
    if leaks:
        return (u'Not saved: {} looks like {}. Recipes are shared with the team; '
                u'redo the task without it (or make it an input) and save again.'
                .format(leaks[0][0], leaks[0][1]))
    client = client or ToolHubClient()
    is_write = any(not _is_read(s['tool']) for s in built['steps'])
    try:
        res = client.rpc('toolhub_save_tool', {
            'p_slug': slug, 'p_kind': 'recipe',
            'p_name': title or slug.replace('_', ' '),
            'p_description': u'Recipe from: {}'.format((user_text or u'')[:200]),
            'p_title_vi': title, 'p_keywords': [],
            'p_input_schema': built['inputs_schema'],
            'p_is_write': is_write, 'p_is_destructive': False,
            'p_plan': {'steps': built['steps']},
        })
    except ToolHubError as ex:
        return u'Tool Hub: {}'.format(ex)
    lines = [u'Saved draft recipe `{}` ({} steps, status {}):'.format(
        slug, len(built['steps']), res.get('status', 'draft') if isinstance(res, dict) else 'draft')]
    for i, s in enumerate(built['steps']):
        lines.append(u'{}. {} {}'.format(i + 1, s['tool'], _short(s['args'])))
    ins = list(built['inputs_schema']['properties'])
    if ins:
        lines.append(u'Inputs: {}'.format(u', '.join(ins)))
    lines.append(u'Send it for review with `/toolhub submit {}`.'.format(slug))
    return u'\n'.join(lines)


def _is_read(tool):
    try:
        from Intelligence.tool_schema import READ_ONLY_TOOL_NAMES
        return tool in READ_ONLY_TOOL_NAMES
    except Exception:
        return False


def _short(args):
    import json
    text = json.dumps(args, ensure_ascii=False, sort_keys=True)
    return text if len(text) <= 160 else text[:157] + u'...'
