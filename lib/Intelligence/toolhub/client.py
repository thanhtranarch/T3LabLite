# -*- coding: utf-8 -*-
"""ToolHubClient.rpc(name, params) — POST /rest/v1/rpc/<name> as the signed-in user.

Every Tool Hub call goes through a public.toolhub_* RPC (PostgREST exposes
only `public`; the tables in `toolhub` have no grants). Errors come back as
ToolHubError with a message the Assistant can show as is.
"""
from __future__ import unicode_literals

import json

from Intelligence.toolhub import auth, config


class ToolHubError(Exception):
    pass


class NotSignedIn(ToolHubError):
    pass


_FRIENDLY = {
    '28000': 'Sign in to the Tool Hub first: LLMs Setting > Tool Hub (or /toolhub login you@company.com)',
    '42501': None,     # the RPC's own message says which role is needed
    'P0002': None,     # "No tool named x."
}


class ToolHubClient(object):

    def __init__(self, http=None):
        self._http = http or auth._default_http

    def signed_in(self):
        return auth.status() is not None

    def rpc(self, name, params=None):
        try:
            token = auth.access_token(self._http)
        except auth.AuthError as ex:
            raise NotSignedIn(str(ex))
        if not token:
            raise NotSignedIn(_FRIENDLY['28000'])
        url, key = config.endpoint()
        headers = {'apikey': key, 'Authorization': 'Bearer ' + token,
                   'Content-Type': 'application/json', 'Accept': 'application/json'}
        try:
            status, text = self._http('POST', '{}/rest/v1/rpc/{}'.format(url, name),
                                      json.dumps(params or {}), headers, config.HTTP_TIMEOUT)
        except Exception as ex:
            raise ToolHubError('The Tool Hub is not reachable ({}). Working offline.'.format(ex))
        try:
            data = json.loads(text) if text else None
        except ValueError:
            data = text
        if status >= 400:
            err = data if isinstance(data, dict) else {}
            code = err.get('code')
            msg = _FRIENDLY.get(code) or err.get('message') or 'HTTP {}'.format(status)
            if status == 401 or code == '28000':
                raise NotSignedIn(msg)
            if status == 404 and not err.get('message'):
                msg = 'The Tool Hub has no function {} (is the database up to date?).'.format(name)
            raise ToolHubError(msg)
        return data
