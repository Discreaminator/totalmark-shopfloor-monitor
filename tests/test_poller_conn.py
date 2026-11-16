#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Poller должен держать своё соединение и не блокироваться на его закрытии."""
import os
import sys
import threading
import time

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
os.environ['RABBIT_ENABLED'] = '0'
os.environ['LOG_LEVEL'] = 'ERROR'
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import app  # noqa: E402

results = []


def check(name, ok, detail=''):
    results.append(ok)
    print(f"  {'OK  ' if ok else 'FAIL'} {name}{('  — ' + detail) if detail else ''}")


class FakeConn:
    instances = 0

    def __init__(self, hang_close=False):
        FakeConn.instances += 1
        self.n = FakeConn.instances
        self.closed = 0
        self.autocommit = False
        self.close_called = threading.Event()
        self._hang_close = hang_close

    def close(self):
        self.close_called.set()
        if self._hang_close:
            time.sleep(30)  # мёртвый TCP
        self.closed = 1


created = []


def fake_connect(**kwargs):
    c = FakeConn(hang_close=bool(created and created[-1]._hang_close_next))
    c._hang_close_next = False
    created.append(c)
    return c


FakeConn._hang_close_next = False
app.psycopg2.connect = fake_connect

print("\n1. Соединение создаётся один раз и переиспользуется")
app._poller_conn = None
c1 = app._poller_db_conn()
c2 = app._poller_db_conn()
check('то же соединение во втором цикле', c1 is c2, f"создано объектов: {len(created)}")
check('autocommit включён', c1.autocommit is True)
check('соединение видно watchdog для cancel()', app._poller_active_conn is c1)

print("\n2. Закрытое соединение заменяется новым")
c1.closed = 1
c3 = app._poller_db_conn()
check('открыто новое', c3 is not c1)

print("\n3. drop(close=False) не закрывает — поток мог ещё работать")
app._poller_drop_conn(reason='test', close=False)
time.sleep(0.3)
check('close не вызывался', not c3.close_called.is_set())
check('соединение отвязано', app._poller_conn is None)

print("\n4. drop(close=True) не блокирует вызывающий поток на мёртвом закрытии")
app._poller_conn = None
c4 = app._poller_db_conn()
c4._hang_close = True
t0 = time.time()
app._poller_drop_conn(reason='test-hang', close=True)
elapsed = time.time() - t0
closed_started = c4.close_called.wait(2)
check('вызов вернулся мгновенно', elapsed < 1.0, f"{elapsed:.2f} с")
check('закрытие ушло в фоновый поток', closed_started)

print("\n5. Следующий цикл получает рабочее соединение, пока старое висит в close")
c5 = app._poller_db_conn()
check('новое соединение выдано', c5 is not c4 and not c5.closed)

print(f"\nитог: {sum(results)}/{len(results)} пройдено")
sys.exit(0 if all(results) else 1)
