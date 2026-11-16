# -*- coding: utf-8 -*-
"""Бочки: в смену идут только сканы заданий «П» с датой в Comment = дата смены."""
import os
import sys
import unittest
from datetime import datetime
from unittest import mock

os.environ.setdefault('TSD_POLL_ENABLED', '0')
sys.path.insert(0, r'C:\Users\alexh\Downloads')

import app as mon

GOOD = 'Антифриз "ShopfloorMonitor Standard 40" бочка 220 кг, зеленый'
GTIN = '4607075781436'


class TsdStrictPlanDateTests(unittest.TestCase):
    def setUp(self):
        os.environ['TSD_STRICT_PLAN_DATE'] = '1'
        mon._tsd_skipped.clear()
        with mon.data_lock:
            mon.production_data[mon.TSD_LINE_ID][mon.TSD_ENTITY_TYPE]['events'].clear()
        self._p = [
            mock.patch.object(mon, 'connected_clients', 0),
            mock.patch.object(mon, '_tsd_save_state'),
        ]
        for p in self._p:
            p.start()

    def tearDown(self):
        for p in self._p:
            p.stop()
        with mon.data_lock:
            mon.production_data[mon.TSD_LINE_ID][mon.TSD_ENTITY_TYPE]['events'].clear()
        mon._tsd_skipped.clear()

    def _events(self):
        with mon.data_lock:
            return list(mon.production_data[mon.TSD_LINE_ID][mon.TSD_ENTITY_TYPE]['events'])

    def _reconcile(self, comment, target, have, now):
        key = ('job-1', GTIN)
        mon._tsd_reconcile_to_api(
            {key: target}, {key: comment}, {key: GOOD},
            now=now, prev_snapshot={key: have},
        )

    def test_today_job_counted(self):
        now = datetime(2026, 9, 30, 8, 25, 0)
        self._reconcile('П Std 40 зел (220) - 8 30/09', 8, 0, now)
        self.assertEqual(sum(e['qty'] for e in self._events()), 8)
        self.assertEqual(len(mon._tsd_skipped), 0)

    def test_other_date_job_skipped_and_logged(self):
        now = datetime(2026, 9, 30, 8, 25, 0)
        self._reconcile('П Std 40 зел (220) - 8 29/09', 8, 0, now)
        self.assertEqual(self._events(), [])
        self.assertEqual(len(mon._tsd_skipped), 1)
        self.assertEqual(mon._tsd_skipped[0]['qty'], 8)

    def test_comment_without_date_skipped(self):
        now = datetime(2026, 9, 30, 8, 25, 0)
        self._reconcile('П Std 40 зел (220) - 8', 8, 0, now)
        self.assertEqual(self._events(), [])

    def test_night_shift_accepts_both_calendar_days(self):
        now = datetime(2026, 10, 1, 2, 0, 0)  # ночная смена 30.09 20:00 → 01.10 08:00
        self._reconcile('П Komatsu (220) - 10 30/09', 3, 0, now)
        self._reconcile('П Komatsu (220) - 10 01/10', 5, 0, now)
        self.assertEqual(sum(e['qty'] for e in self._events()), 8)

    def test_off_date_events_already_in_shift_are_removed(self):
        now = datetime(2026, 9, 30, 15, 0, 0)
        with mon.data_lock:
            mon._append_event(
                mon.TSD_LINE_ID, mon.TSD_ENTITY_TYPE, 8, GTIN,
                ts=datetime(2026, 9, 30, 8, 25, 0), source='TSD',
                meta={'job_id': 'old', 'comment': 'П Std 40 (220) - 8 29/09'},
            )
            mon._append_event(
                mon.TSD_LINE_ID, mon.TSD_ENTITY_TYPE, 80, '4607075784697',
                ts=datetime(2026, 9, 30, 10, 0, 0), source='TSD',
                meta={'job_id': 'k', 'comment': 'П Komatsu 60/40 (220) - 80 30/09'},
            )
        removed = mon._tsd_drop_off_date_shift_events(now)
        self.assertEqual(removed, 1)
        self.assertEqual([e['qty'] for e in self._events()], [80])
        self.assertEqual(mon._tsd_skipped[-1]['qty'], 8)

    def test_strict_off_restores_old_behavior(self):
        os.environ['TSD_STRICT_PLAN_DATE'] = '0'
        now = datetime(2026, 9, 30, 8, 25, 0)
        self._reconcile('П Std 40 зел (220) - 8 29/09', 8, 0, now)
        self.assertEqual(sum(e['qty'] for e in self._events()), 8)


if __name__ == '__main__':
    unittest.main()
