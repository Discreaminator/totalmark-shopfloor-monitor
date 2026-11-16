# -*- coding: utf-8 -*-
"""Фоновая очередь исходящих вызовов MAX — не блокирует poller."""
import logging
import queue
import threading

log = logging.getLogger(__name__)

_q = queue.Queue(maxsize=200)
_started = False
_lock = threading.Lock()


def start():
    global _started
    with _lock:
        if _started:
            return
        t = threading.Thread(target=_loop, daemon=True, name='max_bot')
        t.start()
        _started = True
        log.info('MAX bot: worker запущен')


def enqueue(fn, *args, **kwargs):
    try:
        _q.put_nowait((fn, args, kwargs))
        return True
    except queue.Full:
        log.warning('MAX bot: очередь исходящих переполнена — drop')
        return False


def _loop():
    while True:
        item = _q.get()
        if item is None:
            break
        fn, args, kwargs = item
        try:
            fn(*args, **kwargs)
        except Exception as e:
            log.error(f'MAX bot worker: {e}')
