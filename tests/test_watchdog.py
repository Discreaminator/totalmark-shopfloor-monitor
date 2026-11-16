#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Watchdog должен перезапускать только реально застрявший poller, а не долгий догон истории."""
import os
import sys
import threading
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
os.environ['RABBIT_ENABLED'] = '0'
os.environ['LOG_LEVEL'] = 'ERROR'
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import app  # noqa: E402

MAX_SILENCE = 120
results = []


def check(name, ok, detail=''):
    results.append(ok)
    print(f"  {'OK  ' if ok else 'FAIL'} {name}{('  — ' + detail) if detail else ''}")


def iso(delta_sec):
    if delta_sec is None:
        return None
    return (datetime.now() - timedelta(seconds=delta_sec)).isoformat(timespec='seconds')


class FakeThread:
    def is_alive(self):
        return True


def state(ok_age=None, inflight=None, progress=None, stage=None):
    app._poller_thread = FakeThread()
    with app._health_lock:
        app._health['started_at'] = iso(3600)
        app._health['last_poll_ok_at'] = iso(ok_age)
        app._health['poll_started_at'] = iso(inflight)
        app._health['poll_progress_at'] = iso(progress)
        app._health['poll_stage'] = stage
    return app._poller_stale_info(MAX_SILENCE)


print("\n1. Обычная работа: цикла нет, недавно был успешный poll")
info = state(ok_age=5)
check('здоров', not info['stale'], f"reason={info['reason']}")

print("\n2. Догон истории: цикл идёт 400 с, но шаги каждые 3 с")
info = state(ok_age=400, inflight=400, progress=3, stage='fetched:dm')
check('НЕ считается зависанием', not info['stale'], f"reason={info['reason']}")
check('в reason видно стадию и длительность',
      'in_progress' in (info['reason'] or '') and 'dm' in (info['reason'] or ''),
      f"reason={info['reason']}")

print("\n3. Настоящее зависание: цикл идёт 400 с, последний шаг 400 с назад")
info = state(ok_age=400, inflight=400, progress=400, stage='process')
check('зависание найдено', info['stale'], f"reason={info['reason']}")
check('reason говорит об отсутствии прогресса', 'no_progress' in (info['reason'] or ''),
      f"reason={info['reason']}")

print("\n4. Зависание в SQL: шагов внутри цикла не было вовсе")
info = state(ok_age=None, inflight=300, progress=None, stage='fetch')
check('зависание найдено', info['stale'], f"reason={info['reason']}")

print("\n5. Poller молчит между циклами (цикл не начат)")
info = state(ok_age=300, inflight=None, progress=None)
check('зависание найдено по ok_age', info['stale'], f"reason={info['reason']}")

print("\n6. Поток poller мёртв")
app._poller_thread = None
with app._health_lock:
    app._health['started_at'] = iso(3600)
    app._health['last_poll_ok_at'] = iso(5)
    app._health['poll_started_at'] = None
    app._health['poll_progress_at'] = None
info = app._poller_stale_info(MAX_SILENCE)
check('зависание найдено', info['stale'] and info['reason'] == 'thread_dead', f"reason={info['reason']}")

print("\n7. Отметка прогресса реально обновляет health")
with app._health_lock:
    app._health['poll_progress_at'] = iso(500)
app._poller_progress(stage='downtime')
with app._health_lock:
    fresh = app._health['poll_progress_at']
    st = app._health['poll_stage']
age = (datetime.now() - datetime.fromisoformat(fresh)).total_seconds()
check('метка свежая и стадия записана', age < 5 and st == 'downtime', f"age={age:.1f} с, stage={st}")

print(f"\nитог: {sum(results)}/{len(results)} пройдено")
sys.exit(0 if all(results) else 1)
