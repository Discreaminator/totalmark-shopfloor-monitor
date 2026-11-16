# -*- coding: utf-8 -*-
"""Комментарии к причинам простоя (Ремонт/Проблема) → downtime.json / отчёт."""
import os
import sys
import unittest
from datetime import datetime
from unittest import mock

os.environ.setdefault('TSD_POLL_ENABLED', '0')
sys.path.insert(0, r'C:\Users\alexh\Downloads')

import app as mon


class DowntimeCommentTests(unittest.TestCase):
    def setUp(self):
        self._prev_open = dict(mon.downtime_open)
        self._prev_closed = list(mon.downtime_closed)
        with mon.downtime_lock:
            mon.downtime_open.clear()
            mon.downtime_closed[:] = [
                c for c in mon.downtime_closed
                if not str(c.get('line_id', '')).startswith('UT_DT_')
            ]

    def tearDown(self):
        with mon.downtime_lock:
            mon.downtime_open.clear()
            mon.downtime_open.update(self._prev_open)
            mon.downtime_closed[:] = self._prev_closed

    def test_problem_requires_comment(self):
        with mock.patch.object(mon, 'save_downtime'), \
             mock.patch.object(mon, 'connected_clients', 0), \
             mock.patch.object(mon, '_max_notify_reason'):
            rec, err = mon.downtime_ack(
                'UT_DT_A', 'Проблема', started_at='2026-09-18T10:00:00'
            )
            self.assertIsNone(rec)
            self.assertIn('комментарий', (err or '').lower())

    def test_repair_saves_comment(self):
        with mock.patch.object(mon, 'save_downtime'), \
             mock.patch.object(mon, 'connected_clients', 0), \
             mock.patch.object(mon, '_max_notify_reason'):
            rec, err = mon.downtime_ack(
                'UT_DT_B', 'Ремонт',
                started_at='2026-09-18T10:00:00',
                comment='замена подшипника',
            )
            self.assertIsNone(err)
            self.assertEqual(rec.get('reason'), 'Ремонт')
            self.assertEqual(rec.get('comment'), 'замена подшипника')
            self.assertTrue(rec.get('acknowledged'))

    def test_lunch_drops_comment(self):
        with mock.patch.object(mon, 'save_downtime'), \
             mock.patch.object(mon, 'connected_clients', 0), \
             mock.patch.object(mon, '_max_notify_reason'):
            rec, err = mon.downtime_ack(
                'UT_DT_C', 'Обед',
                started_at='2026-09-18T12:00:00',
                comment='не должен сохраниться',
            )
            self.assertIsNone(err)
            self.assertEqual(rec.get('reason'), 'Обед')
            self.assertFalse(rec.get('comment'))

    def test_meta_returns_comment(self):
        closed = [{
            'line_id': 'UT_DT_D',
            'reason': 'Проблема',
            'comment': 'датчик',
            'acknowledged': True,
            'started_at': '2026-09-18T09:00:00',
            'ended_at': '2026-09-18T09:25:00',
        }]
        r, a, c = mon._master_downtime_meta(
            'UT_DT_D',
            datetime(2026, 9, 18, 9, 0, 0),
            datetime(2026, 9, 18, 9, 30, 0),
            closed,
        )
        self.assertEqual(r, 'Проблема')
        self.assertTrue(a)
        self.assertEqual(c, 'датчик')

    def test_snapshot_lists_comment_reasons(self):
        snap = mon._downtime_snapshot_locked()
        self.assertIn('Ремонт', snap.get('comment_reasons') or [])
        self.assertIn('Проблема', snap.get('comment_reasons') or [])


if __name__ == '__main__':
    unittest.main()
