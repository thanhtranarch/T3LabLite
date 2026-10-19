# -*- coding: utf-8 -*-
"""What the LLMs Setting ▸ Tool Hub tab shows and does — no WPF in here.

The tab used to not exist: signing in meant knowing to type
`/toolhub login you@company.com` in the chat. This module is the plain-data
layer behind the tab (and testable without Revit); the dialog only renders
the dicts and runs the network calls on a worker thread.

    local_state()            instant, no network: session + local cache
    refresh(client)          role + catalog sync (network)
    send_code(email)         e-mail a sign-in code
    verify(email, code)      code or pasted link -> signed in, then refresh
    sign_out()
    search(query)            team search, offline cache as fallback
"""
from __future__ import unicode_literals

from Intelligence.toolhub import auth, catalog, search as _search
from Intelligence.toolhub.client import NotSignedIn, ToolHubClient, ToolHubError

# Errors the tab shows as they are (they are already user-facing sentences).
UserError = (auth.AuthError, ToolHubError)


def _mine(rows):
    out = []
    for r in sorted((r for r in rows.values() if r.get('mine')),
                    key=lambda r: r.get('slug', '')):
        out.append({'slug': r.get('slug', ''),
                    'title': r.get('title_vi') or r.get('name') or u'',
                    'kind': r.get('kind', u''),
                    'status': r.get('status', u'?'),
                    'note': r.get('review_note') or u''})
    return out


def local_state():
    """Session + local catalog cache. Never touches the network."""
    st = auth.status() or {}
    cache = catalog.load()
    rows = cache.get('rows') or {}
    return {'signed_in': bool(st),
            'email': st.get('email') or u'',
            'role': None,
            'tools': len(rows),
            'last_updated': cache.get('last_updated') or u'',
            'mine': _mine(rows),
            'offline': False,
            'note': u''}


def refresh(client=None):
    """local_state() plus the role and a catalog sync. NETWORK.

    Offline (or the hub is down) is not an error here: the tab still shows
    the cache, flagged `offline`.
    """
    client = client or ToolHubClient()
    role, offline, note = None, False, u''
    if auth.status():
        try:
            me = client.rpc('toolhub_me', {})
            role = me.get('role') if isinstance(me, dict) else None
        except NotSignedIn as ex:
            note = u'{}'.format(ex)                 # session expired / revoked
        except Exception:
            offline = True
    try:
        catalog.sync(client)
    except Exception:
        offline = True
    state = local_state()
    state['role'] = role
    state['offline'] = offline and state['signed_in']
    state['note'] = note
    return state


def send_code(email, client=None):
    client = client or ToolHubClient()
    return auth.send_code(email, client._http)


def verify(email, code, client=None):
    """Finish signing in with the 6-digit code or the pasted link."""
    client = client or ToolHubClient()
    if not (code or u'').strip():
        raise auth.AuthError(u'Type the 6-digit code from the e-mail, or paste '
                             u'its sign-in link.')
    auth.verify(email, code, client._http)
    return refresh(client)


def sign_out(client=None):
    client = client or ToolHubClient()
    auth.sign_out(client._http)
    return local_state()


def search(query, client=None, k=8):
    """[{slug, kind, title, why}] — the team search when signed in, the
    offline cache otherwise."""
    q = (query or u'').strip()
    if not q:
        return []
    client = client or ToolHubClient()
    out = []
    for h in _search.search(q, k=k, client=client) or []:
        out.append({'slug': h.get('slug', u''), 'kind': h.get('kind', u''),
                    'title': h.get('title_vi') or h.get('name') or u'',
                    'why': h.get('why') or h.get('description') or u''})
    return out


def status_line(state):
    """One sentence for the top of the tab."""
    n = state.get('tools', 0)
    tools = u'{} tool{} in the catalog'.format(n, u'' if n == 1 else u's')
    if not state.get('signed_in'):
        return u'Signed out · {} (local cache).'.format(tools)
    who = u'Signed in as {}'.format(state.get('email') or u'?')
    if state.get('role'):
        who += u' · role {}'.format(state['role'])
    line = u'{} · {}'.format(who, tools)
    if state.get('offline'):
        line += u' · offline, showing the local cache'
    return line + u'.'
