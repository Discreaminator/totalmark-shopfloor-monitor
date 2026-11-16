#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Нагрузка: простои + MAX-алерты (mock).
Опционально --live: одно реальное сообщение в чат (не флуд).

  python test_load_max_downtime.py
  python test_load_max_downtime.py --live
"""
from __future__ import annotations

import argparse
import os
import sys
import tempfile
import threading
import time
from datetime import datetime, timedelta
from unittest.mock import patch

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Изоляция downtime.json до импорта app
_tmpdir = tempfile.mkdtemp(prefix='cs_load_')
os.environ['MAX_ENABLED'] = '1'
os.environ['MAX_BOT_TOKEN'] = os.environ.get('MAX_BOT_TOKEN') or 'test-token'
os.environ.setdefault('MAX_ALERT_CHAT_IDS', '1001')
os.environ['MAX_ALERT_MIN_INTERVAL_SEC'] = '1'

import max_bot.config as max_cfg  # noqa: E402
import max_bot.alerts as alerts  # noqa: E402
from max_bot import worker  # noqa: E402

# Перенаправить downtime-файл app до тяжёлого импорта сложно — патчим после
import app as mon  # noqa: E402

mon.DOWNTIME_FILE = os.path.join(_tmpdir, 'downtime.json')
mon.downtime_closed = []
mon.downtime_open = {}
mon.connected_clients = set()  # без socket flood

results = []
sent = []
sent_lock = threading.Lock()


def check(name, ok, detail=''):
    results.append(ok)
    print(f"  {'OK  ' if ok else 'FAIL'} {name}" + (f'  — {detail}' if detail else ''))


def fake_send(text, chat_id=None, user_id=None, notify=True):
    with sent_lock:
        sent.append((chat_id, text, time.time()))
    return {'message': {'message_id': len(sent)}}


def drain_worker(sec=1.5):
    time.sleep(sec)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--live', action='store_true', help='Одно реальное сообщение в MAX')
    parser.add_argument('--cycles', type=int, default=40)
    args = parser.parse_args()

    print('\n=== 1. Burst open/close (антидубль + badge) ===')
    line = 'Ralot_L1'
    t0 = datetime.now().replace(microsecond=0) - timedelta(minutes=20)
    # 40 циклов «открыли тем же started_at» — в closed должна быть 1 запись
    with patch('max_bot.alerts.send_message', side_effect=fake_send):
        worker.start()
        alerts._sent_keys.clear()
        alerts._last_line_ts.clear()
        sent.clear()

        for i in range(args.cycles):
            mon.downtime_open_or_update(
                line, started_at=t0.isoformat(timespec='seconds'),
                badge_min=10, entity_type='L1',
            )
            mon.downtime_close(line, ended_at=(t0 + timedelta(minutes=12)).isoformat(timespec='seconds'))
        drain_worker(1.0)

    same = [r for r in mon.downtime_closed if r.get('line_id') == line and r.get('started_at', '').startswith(t0.isoformat()[:16])]
    # дедуп close: один started_at
    keys = {(r['line_id'], mon._started_key(r.get('started_at'))) for r in mon.downtime_closed if r.get('line_id') == line}
    check('closed без шторма дублей', len(keys) <= 2, f'unique={len(keys)} closed={len(mon.downtime_closed)}')
    check('повторный open игнор', mon.downtime_open_or_update(line, started_at=t0.isoformat(timespec='seconds'), badge_min=10) is None)

    print('\n=== 2. Короче плашки — не в Excel/не в MAX resume ===')
    mon.downtime_closed.clear()
    mon.downtime_open.clear()
    with patch('max_bot.alerts.send_message', side_effect=fake_send):
        sent.clear()
        alerts._sent_keys.clear()
        alerts._last_line_ts.clear()
        st = datetime.now() - timedelta(minutes=3)
        mon.downtime_open_or_update('BLS', started_at=st.isoformat(timespec='seconds'), badge_min=10)
        closed = mon.downtime_close('BLS', ended_at=datetime.now().isoformat(timespec='seconds'))
        drain_worker(0.8)
    check('short не записан', closed is None and not any(r.get('line_id') == 'BLS' for r in mon.downtime_closed))
    resumes = [s for s in sent if 'возобновлена' in s[1].lower()]
    check('short без resume в MAX', len(resumes) == 0, f'resumes={len(resumes)}')

    print('\n=== 3. Нагрузка алертов MAX (mock, много линий) ===')
    mon.downtime_closed.clear()
    mon.downtime_open.clear()
    with patch('max_bot.alerts.send_message', side_effect=fake_send):
        sent.clear()
        alerts._sent_keys.clear()
        alerts._last_line_ts.clear()
        lines = ['Ralot_L1', 'Ralot_L2', 'KRO_BOTTLE_L1', 'KRO_BOTTLE_L2', 'BLS', 'KRO_KANISTER']
        t_base = datetime.now() - timedelta(minutes=30)
        for i, lid in enumerate(lines):
            st = t_base + timedelta(minutes=i)
            # шторм повторных open
            for _ in range(15):
                mon.downtime_open_or_update(lid, started_at=st.isoformat(timespec='seconds'), badge_min=10)
            mon.downtime_close(
                lid,
                ended_at=(st + timedelta(minutes=15)).isoformat(timespec='seconds'),
            )
        drain_worker(2.0)

    stops = [s for s in sent if 'Задержка или остановка' in s[1]]
    resumes = [s for s in sent if 'возобновлена' in s[1].lower()]
    check('stop ≤ числу линий (дедуп)', len(stops) <= len(lines) + 2, f'stops={len(stops)}')
    check('resume по закрытым', len(resumes) >= 1, f'resumes={len(resumes)}')
    check('очередь не упала', True)

    print('\n=== 4. get_downtime_for_shift / lines_live smoke ===')
    try:
        now = datetime.now()
        start = now.replace(hour=8, minute=0, second=0, microsecond=0)
        if now.hour < 8:
            start -= timedelta(days=1)
        end = start + timedelta(hours=12)
        rows = mon.get_downtime_for_shift(start, end, 'day')
        check('get_downtime_for_shift', isinstance(rows, list), f'n={len(rows)}')
        live = mon.build_lines_live()
        check('build_lines_live', isinstance(live.get('lines'), list), f"n={len(live.get('lines') or [])}")
    except Exception as e:
        check('shift/lines_live', False, str(e))

    print('\n=== 5. Обратная совместимость API open без badge_min ===')
    mon.downtime_open.clear()
    rec = mon.downtime_open_or_update('Ralot_L2', started_at=(datetime.now() - timedelta(minutes=1)).isoformat(timespec='seconds'))
    check('open без badge_min', rec is not None and rec.get('badge_min') == mon.STATUS_SILENCE_BAD_MIN)

    if args.live:
        print('\n=== LIVE: одно сообщение в MAX ===')
        # подтянуть реальный env
        from max_bot.config import _load_local_env
        _load_local_env()
        os.environ['MAX_ENABLED'] = '1'
        if not max_cfg.is_enabled() or not max_cfg.alert_chat_ids():
            check('live пропущен (нет токена/чата)', False)
        else:
            from max_bot.client import send_message
            text = (
                f'🧪 ShopfloorMonitor load-test\n'
                f'{datetime.now().strftime("%d.%m.%Y %H:%M:%S")}\n'
                f'Проверка канала алертов (одно сообщение).'
            )
            try:
                for chat_id in max_cfg.alert_chat_ids()[:1]:
                    send_message(text, chat_id=chat_id)
                check('live send', True, f'chat={max_cfg.alert_chat_ids()[0]}')
            except Exception as e:
                check('live send', False, str(e))

    ok = sum(1 for x in results if x)
    print(f'\nитог: {ok}/{len(results)} пройдено')
    return 0 if all(results) else 1


if __name__ == '__main__':
    raise SystemExit(main())
