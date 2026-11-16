#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Старт poller'а должен перечитывать историю за INITIAL_LOAD_HOURS, а не с момента запуска."""
import json
import os
import sys
import tempfile
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

STATE = os.path.join(tempfile.gettempdir(), 'test_poller_state.json')
os.environ['POLLER_STATE_FILE'] = STATE
os.environ['INITIAL_LOAD_HOURS'] = '24'
os.environ['FETCH_WINDOW_MIN'] = '30'
os.environ['FETCH_WINDOW_CATCHUP_MIN'] = '360'
os.environ['RABBIT_ENABLED'] = '0'
os.environ['LOG_LEVEL'] = 'ERROR'

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import app  # noqa: E402  (импорт не подключается к БД: пул поднимается только в __main__)

results = []


def check(name, ok, detail=''):
    results.append((name, ok, detail))
    print(f"  {'OK  ' if ok else 'FAIL'} {name}{('  — ' + detail) if detail else ''}")


def write_state(agg, dm):
    with open(STATE, 'w', encoding='utf-8') as f:
        json.dump({'agg': agg.isoformat() if agg else None,
                   'dm': dm.isoformat() if dm else None}, f)


def reset():
    with app._watermark_lock:
        app._watermarks['agg'] = None
        app._watermarks['dm'] = None


now = datetime.now()
horizon = now - timedelta(hours=24)

print("\n1. Watermark в файле = «сейчас» (обычный рестарт)")
reset()
write_state(now, now)
app.load_poller_state()
agg = app._watermarks['agg']
check('история отматывается на 24 ч назад', abs((agg - horizon).total_seconds()) < 120,
      f"agg={agg:%d.%m %H:%M}, ожидалось ~{horizon:%d.%m %H:%M}")

print("\n2. Watermark в файле старше окна истории (простой 30 дней)")
reset()
old = now - timedelta(days=30)
write_state(old, old)
app.load_poller_state()
agg = app._watermarks['agg']
check('старое значение сохраняется, события не теряются', abs((agg - old).total_seconds()) < 5,
      f"agg={agg:%d.%m %H:%M}")

print("\n3. Файла состояния нет (первый запуск)")
reset()
if os.path.exists(STATE):
    os.remove(STATE)
app.load_poller_state()
agg = app._watermarks['agg']
check('старт с горизонта истории', abs((agg - horizon).total_seconds()) < 120,
      f"agg={agg:%d.%m %H:%M}")

print("\n4. Размер окна запроса")
reset()
with app._watermark_lock:
    app._watermarks['agg'] = now - timedelta(hours=24)
frm, to = app._stream_window('agg', now)
check('далеко позади -> окно догона 360 мин',
      abs((to - frm).total_seconds() / 60 - 360) < 1,
      f"{(to - frm).total_seconds() / 60:.0f} мин")

with app._watermark_lock:
    app._watermarks['agg'] = now - timedelta(minutes=5)
frm, to = app._stream_window('agg', now)
window_min = (to - frm).total_seconds() / 60
check('в реальном времени -> окно не больше 5 мин', window_min <= 5.1,
      f"{window_min:.1f} мин")

print("\n5. Сколько запросов на догон 7 суток")
reset()
with app._watermark_lock:
    app._watermarks['agg'] = now - timedelta(days=7)
cursor, steps = app._watermarks['agg'], 0
while cursor < now - timedelta(seconds=app.FETCH_SAFETY_LAG_SEC) and steps < 10000:
    with app._watermark_lock:
        app._watermarks['agg'] = cursor
    _frm, cursor = app._stream_window('agg', now)
    steps += 1
check('догон укладывается в разумное число запросов', steps <= 40, f"{steps} запросов")

if os.path.exists(STATE):
    os.remove(STATE)

failed = [n for n, ok, _ in results if not ok]
print(f"\nитог: {len(results) - len(failed)}/{len(results)} пройдено")
sys.exit(1 if failed else 0)
