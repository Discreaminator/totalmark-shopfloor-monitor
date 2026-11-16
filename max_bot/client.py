# -*- coding: utf-8 -*-
"""HTTP-клиент MAX Bot API — только исходящие сообщения."""
import json
import logging
import os
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request

from max_bot import config

log = logging.getLogger(__name__)

_ssl_ctx = None
_ssl_insecure = False


class MaxApiError(Exception):
    def __init__(self, message, status=None, body=None):
        super().__init__(message)
        self.status = status
        self.body = body


def _make_insecure_ctx():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def _ssl_context():
    """
    SSL для Windows/корп. сетей.
    MAX_SSL_VERIFY=0 — сразу без проверки.
    Иначе certifi; при CERTIFICATE_VERIFY_FAILED — автопереход на insecure.
    """
    global _ssl_ctx, _ssl_insecure
    if _ssl_ctx is not None:
        return _ssl_ctx
    raw = (os.environ.get('MAX_SSL_VERIFY', '1') or '1').strip().lower()
    if raw in ('0', 'false', 'no', 'off'):
        _ssl_insecure = True
        _ssl_ctx = _make_insecure_ctx()
        log.warning('MAX: SSL verify выключен (MAX_SSL_VERIFY=0)')
        return _ssl_ctx
    try:
        import certifi
        _ssl_ctx = ssl.create_default_context(cafile=certifi.where())
    except Exception:
        _ssl_ctx = ssl.create_default_context()
    return _ssl_ctx


def _use_insecure_ssl(reason=''):
    global _ssl_ctx, _ssl_insecure
    _ssl_insecure = True
    _ssl_ctx = _make_insecure_ctx()
    log.warning(
        f'MAX: SSL verify отключён автоматически ({reason or "CERTIFICATE_VERIFY_FAILED"}). '
        f'Зафиксируйте MAX_SSL_VERIFY=0 в max_bot.local.env'
    )
    return _ssl_ctx


def _is_ssl_verify_error(err):
    s = str(err).lower()
    return (
        'certificate_verify_failed' in s
        or 'certificate verify failed' in s
        or 'ssl:' in s and 'certificate' in s
    )


def _headers():
    tok = config.token()
    if not tok:
        raise MaxApiError('MAX_BOT_TOKEN не задан')
    return {
        'Authorization': tok,
        'Content-Type': 'application/json; charset=utf-8',
        'Accept': 'application/json',
    }


def request(method, path, query=None, body=None, retries=3):
    url = config.BASE_URL + path
    if query:
        url += '?' + urllib.parse.urlencode({k: v for k, v in query.items() if v is not None})
    data = None if body is None else json.dumps(body, ensure_ascii=False).encode('utf-8')
    last_err = None
    ctx = _ssl_context()
    ssl_fallback_done = _ssl_insecure
    for attempt in range(max(1, retries)):
        req = urllib.request.Request(url, data=data, headers=_headers(), method=method.upper())
        try:
            with urllib.request.urlopen(req, timeout=config.HTTP_TIMEOUT, context=ctx) as resp:
                raw = resp.read().decode('utf-8', errors='replace')
                if not raw:
                    return {}
                try:
                    return json.loads(raw)
                except json.JSONDecodeError:
                    return {'_raw': raw}
        except urllib.error.HTTPError as e:
            raw = e.read().decode('utf-8', errors='replace') if e.fp else ''
            last_err = MaxApiError(
                f'MAX {method} {path} HTTP {e.code}: {raw[:300]}',
                status=e.code, body=raw,
            )
            if e.code in (429, 503) and attempt + 1 < retries:
                time.sleep(0.5 * (2 ** attempt))
                continue
            raise last_err
        except Exception as e:
            if not ssl_fallback_done and _is_ssl_verify_error(e):
                ctx = _use_insecure_ssl(str(e)[:120])
                ssl_fallback_done = True
                # сразу ещё одна попытка без verify
                continue
            last_err = MaxApiError(f'MAX {method} {path}: {e}')
            if attempt + 1 < retries:
                time.sleep(0.5 * (2 ** attempt))
                continue
            raise last_err
    raise last_err or MaxApiError('MAX request failed')


def get_me():
    return request('GET', '/me')


def send_message(text, chat_id=None, user_id=None, notify=True):
    query = {}
    if chat_id is not None:
        query['chat_id'] = int(chat_id)
    if user_id is not None:
        query['user_id'] = int(user_id)
    if not query:
        raise MaxApiError('Нужен chat_id или user_id')
    body = {'text': text, 'notify': bool(notify)}
    try:
        return request('POST', '/messages', query=query, body=body)
    except MaxApiError as e:
        if chat_id is not None and user_id is None and e.status in (400, 404):
            log.info(f'MAX: chat_id={chat_id} не принят, пробую user_id')
            return request('POST', '/messages', query={'user_id': int(chat_id)}, body=body)
        raise
