# -*- coding: utf-8 -*-
"""Режим мастера: причина для прошлых простоев/пауз, «без причины», журнал смены."""
import os
import sys
import unittest
from datetime import datetime, timedelta
from unittest import mock

os.environ.setdefault('TSD_POLL_ENABLED', '0')
sys.path.insert(0, r'C:\Users\alexh\Downloads')

import app as mon

LINE = 'Ralot_L1'


def iso(dt):
    return mon._dt_to_iso(dt)


class MasterJournalTests(unittest.TestCase):
    def setUp(self):
        self.now = datetime.now().replace(microsecond=0)
        self.shift_start = mon._current_shift_start(self.now)
        self._saved_closed = list(mon.downtime_closed)
        self._saved_open = dict(mon.downtime_open)
        mon.downtime_closed.clear()
        mon.downtime_open.clear()
        self._p = [
            mock.patch.object(mon, 'save_downtime'),
            mock.patch.object(mon, 'connected_clients', 0),
            mock.patch.object(mon, '_max_notify_reason'),
        ]
        for p in self._p:
            p.start()

    def tearDown(self):
        for p in self._p:
            p.stop()
        mon.downtime_closed[:] = self._saved_closed
        mon.downtime_open.clear()
        mon.downtime_open.update(self._saved_open)

    def _closed(self, st, en, **kw):
        rec = {
            'id': f'x-{iso(st)}', 'line_id': LINE, 'reason': mon.DEFAULT_DOWNTIME_REASON,
            'comment': None, 'started_at': iso(st), 'ended_at': iso(en),
            'acknowledged': False, 'duration_min': round((en - st).total_seconds() / 60, 1),
            'badge_min': 10,
        }
        rec.update(kw)
        mon.downtime_closed.append(rec)
        return rec

    def test_pending_lists_unacked_closed_in_shift(self):
        st = self.shift_start + timedelta(minutes=1)
        if st + timedelta(minutes=15) > self.now:
            self.skipTest('смена только началась')
        self._closed(st, st + timedelta(minutes=15))
        self._closed(st - timedelta(hours=13), st - timedelta(hours=12, minutes=40))
        pending = mon.downtime_snapshot()['pending']
        self.assertEqual([p['started_at'] for p in pending], [iso(st)])

    def test_ack_closed_does_not_touch_current_open(self):
        old_st = self.now - timedelta(minutes=50)
        self._closed(old_st, old_st + timedelta(minutes=20))
        mon.downtime_open[LINE] = {
            'id': 'cur', 'line_id': LINE, 'reason': None,
            'started_at': iso(self.now - timedelta(minutes=12)), 'acknowledged': False,
        }
        rec, err = mon.downtime_ack(LINE, 'Ремонт', started_at=iso(old_st), comment='насос')
        self.assertIsNone(err)
        self.assertEqual(rec['reason'], 'Ремонт')
        self.assertEqual(rec['comment'], 'насос')
        self.assertTrue(mon.downtime_closed[0]['acknowledged'])
        self.assertFalse(mon.downtime_open[LINE]['acknowledged'])
        self.assertIsNone(mon.downtime_open[LINE]['reason'])
        self.assertEqual(mon.downtime_snapshot()['pending'], [])

    def test_ack_pause_creates_record_visible_to_report(self):
        st = self.now - timedelta(minutes=40)
        en = st + timedelta(minutes=6)
        rec, err = mon.downtime_ack(LINE, 'Переход', started_at=iso(st), ended_at=iso(en))
        self.assertIsNone(err)
        self.assertEqual(rec['kind'], 'pause')
        self.assertNotIn(LINE, mon.downtime_open)
        shift = mon.get_downtime_for_shift(st - timedelta(hours=1), en + timedelta(hours=1))
        reason, acked, _ = mon._master_downtime_meta(LINE, st, en, shift)
        self.assertEqual(reason, 'Переход')
        self.assertTrue(acked)
        # повторная правка — та же запись, без дублей
        mon.downtime_ack(LINE, 'Обед', started_at=iso(st), ended_at=iso(en))
        self.assertEqual(len(mon.downtime_closed), 1)
        self.assertEqual(mon.downtime_closed[0]['reason'], 'Обед')

    def test_ack_pause_rejects_bad_interval(self):
        st = self.now - timedelta(minutes=10)
        _, err = mon.downtime_ack(LINE, 'Обед', started_at=iso(st), ended_at=iso(st - timedelta(minutes=1)))
        self.assertTrue(err)
        self.assertEqual(mon.downtime_closed, [])

    def test_journal_rows_threshold_and_reason(self):
        # вчерашняя дневная смена — уже завершена, результат не зависит от времени запуска
        ref_now = (self.now - timedelta(days=1)).replace(hour=19, minute=0, second=0)
        start, _end, _, _ = mon._shift_window(ref_now)
        t0 = start + timedelta(minutes=5)
        # DM каждые 30 с, с разрывами 4 мин (скрыт), 7 мин (пауза) и 20 мин (простой)
        gaps = {10: 4, 20: 7, 30: 20}
        events, ts = [], t0
        for i in range(40):
            events.append({'ts': ts, 'qty': 1})
            ts += timedelta(minutes=gaps.get(i, 0.5))
        down_st = events[30]['ts']
        self._closed(down_st, down_st + timedelta(minutes=20), reason='Ремонт',
                     comment='лента', acknowledged=True)
        dm = mon.production_data.setdefault(LINE, {}).setdefault('DM', {'events': []})
        saved = list(dm['events'])
        dm['events'].clear()
        dm['events'].extend(events)
        try:
            with mock.patch.object(mon, 'get_plan_for_shift', return_value=0):
                j = mon.build_master_journal(ref_now)
        finally:
            dm['events'].clear()
            dm['events'].extend(saved)
        line = next(ln for ln in j['lines'] if ln['line_id'] == LINE)
        kinds = [(it['kind'], round(it['minutes'])) for it in line['items']]
        self.assertEqual(kinds, [('Пауза', 7), ('Простой', 20)])
        pause, down = line['items']
        self.assertFalse(pause['acknowledged'])
        self.assertEqual(pause['ack_kind'], 'pause')
        self.assertTrue(down['acknowledged'])
        self.assertEqual(down['reason'], 'Ремонт')
        self.assertEqual(down['ack_started'], iso(down_st))
        self.assertEqual(j['unmarked'], 1)

    def test_pending_resets_at_shift_boundary(self):
        day = self.now.date()
        st = datetime(day.year, day.month, day.day, 18, 30)
        self._closed(st, st + timedelta(minutes=25))
        before = datetime(day.year, day.month, day.day, 19, 59)
        after = datetime(day.year, day.month, day.day, 20, 1)
        with mon.downtime_lock:
            self.assertEqual(len(mon._downtime_pending_locked(before)), 1)
            self.assertEqual(mon._downtime_pending_locked(after), [])

    def test_boundary_closed_downtime_not_pending_in_new_shift(self):
        day = self.now.date()
        boundary = datetime(day.year, day.month, day.day, 20, 0)
        self._closed(boundary - timedelta(minutes=40), boundary)
        with mon.downtime_lock:
            self.assertEqual(mon._downtime_pending_locked(boundary + timedelta(minutes=1)), [])

    def test_pause_ack_twice_no_duplicate_even_if_first_lookup_missed(self):
        st = self.now - timedelta(minutes=30)
        en = st + timedelta(minutes=6)
        mon._downtime_ack_pause(LINE, mon._dt_to_iso(st), mon._dt_to_iso(en), 'Обед', None)
        mon._downtime_ack_pause(LINE, mon._dt_to_iso(st), mon._dt_to_iso(en), 'Переход', None)
        self.assertEqual(len(mon.downtime_closed), 1)
        self.assertEqual(mon.downtime_closed[0]['reason'], 'Переход')

    def test_journal_endpoint(self):
        r = mon.app.test_client().get('/api/master_journal')
        self.assertEqual(r.status_code, 200)
        self.assertIn('lines', r.get_json())


if __name__ == '__main__':
    unittest.main()
