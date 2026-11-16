#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""MAX: N=10/20/50/80/+30; тексты задержки/возобновления."""
import os
import sys
import time
from datetime import datetime, timedelta
from unittest.mock import patch

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
os.environ['MAX_ENABLED'] = '1'
os.environ['MAX_BOT_TOKEN'] = 'test-token'
os.environ['MAX_ALERT_CHAT_IDS'] = '1001'
os.environ['MAX_ALERT_FIRST_MIN'] = '10'
os.environ['MAX_ALERT_SECOND_MIN'] = '20'
os.environ['MAX_ALERT_LATE_EVERY_MIN'] = '30'

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from max_bot.alerts import (
    check_open_reminders,
    forget_open,
    milestone_for_duration,
    reset_all_alerts,
    resume_text,
    stop_text,
)

results = []
sent = []


def check(name, ok, detail=''):
    results.append(ok)
    print(f"  {'OK  ' if ok else 'FAIL'} {name}{('  — ' + detail) if detail else ''}")


def fake_send(text, chat_id=None, user_id=None, notify=True):
    sent.append(text)
    return {'ok': True}


print('\n1. milestone_for_duration (10 → 20 → 50 → 80 → 110)')
check('9 → 0', milestone_for_duration(9) == 0)
check('10 → 10', milestone_for_duration(10) == 10)
check('19 → 10', milestone_for_duration(19) == 10)
check('20 → 20', milestone_for_duration(20) == 20)
check('49 → 20', milestone_for_duration(49) == 20)
check('50 → 50', milestone_for_duration(50) == 50)
check('79 → 50', milestone_for_duration(79) == 50)
check('80 → 80', milestone_for_duration(80) == 80)
check('110 → 110', milestone_for_duration(110) == 110)

print('\n2. тексты')
check('stop', stop_text('KRO КАНИСТРЫ 9 КГ', duration_min=20) == 'Задержка на линии KRO КАНИСТРЫ 9 КГ 20 минут.')
check(
    'resume',
    resume_text('BLS', duration_min=50) == 'Работа на линии BLS возобновлена, общее время паузы 50 минут.',
)

print('\n3. расписание алертов')
forget_open('L1')
reset_all_alerts('test')
sent.clear()
start = datetime.now() - timedelta(minutes=10, seconds=5)
open_map = {'L1': {'started_at': start.isoformat(timespec='seconds')}}
with patch('max_bot.alerts.send_message', side_effect=fake_send):
    n = check_open_reminders(open_map, {'L1': 'KRO КАНИСТРЫ 9 КГ'})
    time.sleep(0.35)
check('first N=10', n == 1 and len(sent) == 1 and '10 минут' in sent[0], sent[0] if sent else '')

open_map['L1']['started_at'] = (datetime.now() - timedelta(minutes=20, seconds=5)).isoformat(timespec='seconds')
forget_open('L1')
sent.clear()
start20 = datetime.now() - timedelta(minutes=20, seconds=5)
open_map = {'L1': {'started_at': start20.isoformat(timespec='seconds')}}
with patch('max_bot.alerts.send_message', side_effect=fake_send):
    n = check_open_reminders(open_map, {'L1': 'Line'})
    time.sleep(0.35)
    n2 = check_open_reminders(open_map, {'L1': 'Line'})
    time.sleep(0.2)
check('N=20 once', n == 1 and '20 минут' in sent[0] and n2 == 0, f'n={n} sent={sent}')

forget_open('L1')
sent.clear()
start80 = datetime.now() - timedelta(minutes=80, seconds=5)
open_map = {'L1': {'started_at': start80.isoformat(timespec='seconds')}}
with patch('max_bot.alerts.send_message', side_effect=fake_send):
    n = check_open_reminders(open_map, {'L1': 'Line'})
    time.sleep(0.4)
check('catch-up 80', n >= 1 and any('80 минут' in s for s in sent), f'sent={sent}')

print('\n4. сброс смены')
sent.clear()
with patch('max_bot.alerts.send_message', side_effect=fake_send):
    reset_all_alerts('shift_boundary_08:00')
    n = check_open_reminders({}, {'L1': 'Line'})
    time.sleep(0.15)
check('empty after reset', n == 0 and len(sent) == 0)

print('\n5. ACK причины — стоп задержек + один алерт')
from max_bot.alerts import notify_line_reason, is_reason_acked
forget_open('BLS')
reset_all_alerts('test')
sent.clear()
start = datetime.now() - timedelta(minutes=25)
open_map = {'BLS': {'started_at': start.isoformat(timespec='seconds')}}
with patch('max_bot.alerts.send_message', side_effect=fake_send):
    n = check_open_reminders(open_map, {'BLS': 'BLS'})
    time.sleep(0.3)
    check('delay before ack', n >= 1 and any('Задержка' in s for s in sent), f'sent={sent}')
    sent.clear()
    ok = notify_line_reason('BLS', 'BLS', 'Обед')
    time.sleep(0.3)
    check('reason once', ok and len(sent) == 1 and sent[0] == 'BLS в простое по причине: Обед')
    check('acked flag', is_reason_acked('BLS'))
    sent.clear()
    n2 = check_open_reminders(open_map, {'BLS': 'BLS'})
    time.sleep(0.2)
    check('no more delay', n2 == 0 and len(sent) == 0)
    ok2 = notify_line_reason('BLS', 'BLS', 'Обед')
    time.sleep(0.2)
    check('reason no dup', ok2 is False and len(sent) == 0)

print(f'\nитог: {sum(1 for x in results if x)}/{len(results)} пройдено')
sys.exit(0 if all(results) else 1)
