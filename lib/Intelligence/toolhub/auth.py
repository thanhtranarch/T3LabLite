# -*- coding: utf-8 -*-
"""Tool Hub sign-in: Supabase e-mail OTP, token refresh, session on disk.

    send_code(email)          POST /auth/v1/otp          (a 6-digit code is e-mailed)
    verify(email, code)       POST /auth/v1/verify       -> session
                              (code = the 6 digits, or the pasted sign-in link)
    access_token()            refreshes when < 60 s left -> str or None
    sign_out()

The e-mail must carry the code: in the Supabase dashboard, Authentication →
Email Templates → "Magic Link" has to include {{ .Token }} (the default
template only has a link). See dev/plan/assistant-roadmap-toolhub.md.

Session file %APPDATA%\\T3LabAI\\toolhub\\session.json keeps both tokens as
DPAPI blobs (config.secret_store), so a copied file is useless elsewhere.
`http` is injectable: http(method, url, body, headers, timeout) -> (status, text).
"""
from __future__ import unicode_literals

import json
import time

from Intelligence.toolhub import config

SESSION_FILE = 'session.json'
REFRESH_MARGIN_S = 60


class AuthError(Exception):
    """Plain-English failure the Assistant shows as is."""


def _default_http(method, url, body=None, headers=None, timeout=config.HTTP_TIMEOUT):
    from Services.model_meet_presence import _http
    return _http(method, url, body, headers, timeout)


def _secrets():
    from config import secret_store
    return secret_store


def _post(path, payload, http=None, bearer=None):
    url, key = config.endpoint()
    headers = {'apikey': key, 'Content-Type': 'application/json',
               'Accept': 'application/json'}
    if bearer:
        headers['Authorization'] = 'Bearer ' + bearer
    try:
        status, text = (http or _default_http)(
            'POST', url + path, json.dumps(payload), headers, config.HTTP_TIMEOUT)
    except Exception as ex:
        raise AuthError('The Tool Hub is not reachable ({}). Check the network and try again.'
                        .format(ex))
    try:
        data = json.loads(text) if text else {}
    except ValueError:
        data = {}
    if status >= 400:
        msg = (data.get('msg') or data.get('error_description') or data.get('message')
               or data.get('error') or 'HTTP {}'.format(status))
        raise AuthError(msg)
    return data


def _save(data, email=None):
    """Store a token response. Returns the public part of the session."""
    s = _secrets()
    user = data.get('user') or {}
    expires_at = data.get('expires_at') or (time.time() + int(data.get('expires_in') or 3600))
    old = config.read_json(SESSION_FILE, {}) or {}
    record = {
        'email': email or user.get('email') or old.get('email') or '',
        'user_id': user.get('id') or old.get('user_id') or '',
        'expires_at': int(expires_at),
        'access_token': s.protect(data.get('access_token') or ''),
        'refresh_token': s.protect(data.get('refresh_token') or ''),
    }
    config.write_json(SESSION_FILE, record)
    return {'email': record['email'], 'user_id': record['user_id'],
            'expires_at': record['expires_at']}


def send_code(email, http=None):
    email = (email or '').strip()
    if '@' not in email:
        raise AuthError('Type the e-mail address the code should go to (LLMs Setting > Tool Hub > Send code, or /toolhub login you@company.com).')
    _post('/auth/v1/otp', {'email': email, 'create_user': True}, http)
    return email


def link_token(text):
    """(token_hash, type) from a pasted sign-in link, else None.

    Supabase's default e-mail (no custom SMTP, so the template cannot carry
    {{ .Token }}) holds only a link:
        https://<ref>.supabase.co/auth/v1/verify?token=<hash>&type=magiclink&redirect_to=...
    Copying that link (without opening it: opening uses it up) and pasting it
    into /toolhub code signs in just like the 6-digit code.
    """
    text = (text or '').strip().strip('<>"\'')
    if '/auth/v1/verify' not in text or '?' not in text:
        return None
    try:
        from urllib.parse import parse_qs, urlsplit, unquote
    except ImportError:                          # IronPython 2.7
        from urlparse import parse_qs, urlsplit  # noqa: F401
        from urllib import unquote
    query = parse_qs(urlsplit(text).query)
    token = (query.get('token') or [''])[0]
    kind = (query.get('type') or ['magiclink'])[0]
    if not token:
        return None
    return unquote(token), kind


def verify(email, code, http=None):
    """Finish sign-in with the 6-digit code OR the link from the e-mail."""
    raw = (code or '').strip()
    link = link_token(raw)
    if link:
        token_hash, kind = link
        data = _post('/auth/v1/verify', {'type': kind, 'token_hash': token_hash}, http)
    else:
        digits = raw.replace(' ', '')
        if not digits.isdigit():
            raise AuthError('Type the code from the e-mail (/toolhub code 123456), or paste the '
                            'sign-in link from it (right-click the link > Copy link; do not open it).')
        data = _post('/auth/v1/verify', {'type': 'email', 'email': email, 'token': digits}, http)
    if not data.get('access_token'):
        raise AuthError('That code or link was not accepted. Ask for a new one (LLMs Setting > Tool Hub > Send code, or /toolhub login).')
    return _save(data, email)


def status():
    """Public part of the stored session, or None when signed out."""
    rec = config.read_json(SESSION_FILE)
    if not rec or not rec.get('refresh_token'):
        return None
    return {'email': rec.get('email', ''), 'user_id': rec.get('user_id', ''),
            'expires_at': rec.get('expires_at', 0)}


def access_token(http=None, now=None):
    """A valid access token, refreshed when needed; None when signed out.

    Raises AuthError when the refresh itself is refused (session revoked):
    the caller tells the user to sign in again.
    """
    rec = config.read_json(SESSION_FILE)
    if not rec or not rec.get('refresh_token'):
        return None
    s = _secrets()
    now = time.time() if now is None else now
    if int(rec.get('expires_at') or 0) - REFRESH_MARGIN_S > now:
        token = s.unprotect(rec.get('access_token'))
        if token:
            return token
    refresh = s.unprotect(rec.get('refresh_token'))
    if not refresh:
        raise AuthError('The saved Tool Hub session cannot be read on this account. '
                        'Sign in again in LLMs Setting > Tool Hub (or /toolhub login).')
    try:
        data = _post('/auth/v1/token?grant_type=refresh_token',
                     {'refresh_token': refresh}, http)
    except AuthError as ex:
        sign_out(http, remote=False)
        raise AuthError('The Tool Hub session expired ({}). Sign in again in LLMs Setting > Tool Hub (or /toolhub login).'
                        .format(ex))
    _save(data)
    return data.get('access_token')


def sign_out(http=None, remote=True):
    rec = config.read_json(SESSION_FILE)
    if remote and rec:
        try:
            token = _secrets().unprotect(rec.get('access_token'))
            if token:
                _post('/auth/v1/logout', {}, http, bearer=token)
        except Exception:
            pass                     # signing out locally is what matters
    try:
        import os
        p = config.path(SESSION_FILE)
        if os.path.exists(p):
            os.remove(p)
    except OSError:
        pass
