# -*- coding: utf-8 -*-
"""Регрессии фильтра TSD jobs / даты в Comment / baseline после потери snapshot."""
import os
import sys
import unittest
from datetime import datetime, date, timedelta

os.environ.setdefault('TSD_POLL_ENABLED', '0')
# не поднимать poller/socketio побочки сильнее необходимого
sys.path.insert(0, r'C:\Users\alexh\Downloads')

import app as mon


class TsdCommentDateTests(unittest.TestCase):
    def test_parse_trailing_plan_date(self):
        d = mon._tsd_parse_comment_plan_date(
            'П Pr 50/50 оранж (220) - 24 07/09', year=2026
        )
        self.assertEqual(d, date(2026, 9, 7))

    def test_parse_ignores_mid_slash_pairs(self):
        d = mon._tsd_parse_comment_plan_date(
            'П Komatsu 60/40 (220) - 80 08/09', year=2026
        )
        self.assertEqual(d, date(2026, 9, 8))

    def test_matches_active_day_shift(self):
        now = datetime(2026, 9, 7, 11, 0, 0)
        self.assertTrue(
            mon._tsd_comment_matches_active_shift(
                'П HS ЭкоПро 30 (1038) - 4 07/09', now
            )
        )
        self.assertFalse(
            mon._tsd_comment_matches_active_shift(
                'П CHN кр (220) - 24 11/09', now
            )
        )


class TsdJobRelevantTests(unittest.TestCase):
    def setUp(self):
        with mon._tsd_fact_lock:
            mon._tsd_fact_snapshot.clear()
        os.environ['TSD_JOB_MAX_AGE_DAYS'] = '30'

    def test_old_create_date_kept_by_comment_plan(self):
        now = datetime(2026, 9, 7, 11, 30, 0)
        job = {
            'Id': 'job-1',
            'Date': '2026-09-04T12:38:53',
            'Comment': 'П Pr 50/50 оранж (220) - 24 07/09',
        }
        self.assertTrue(mon._tsd_job_relevant(job, now=now))

    def test_old_36h_window_would_have_failed(self):
        """Контроль: Date=04.09 при смене 07.09 вне старых 36ч, но Comment держит."""
        now = datetime(2026, 9, 7, 11, 0, 0)
        shift_start = mon._current_shift_start(now)
        window_start = shift_start - timedelta(hours=36)
        job_dt = mon._tsd_parse_job_date('2026-09-04T12:38:53')
        self.assertIsNotNone(job_dt)
        self.assertLess(job_dt, window_start)
        job = {
            'Id': 'job-2',
            'Date': '2026-09-04T12:38:53',
            'Comment': 'П Pr 50/50 оранж (220) - 24 07/09',
        }
        self.assertTrue(mon._tsd_job_relevant(job, now=now))

    def test_ancient_job_without_plan_dropped(self):
        now = datetime(2026, 9, 7, 11, 0, 0)
        job = {
            'Id': 'job-old',
            'Date': '2026-01-01T10:00:00',
            'Comment': 'П древнее без даты плана',
        }
        self.assertFalse(mon._tsd_job_relevant(job, now=now))

    def test_tracked_ancient_still_relevant(self):
        now = datetime(2026, 9, 7, 11, 0, 0)
        with mon._tsd_fact_lock:
            mon._tsd_fact_snapshot[('job-old', '04607075784864')] = 10
        job = {
            'Id': 'job-old',
            'Date': '2026-01-01T10:00:00',
            'Comment': 'П древнее',
        }
        self.assertTrue(mon._tsd_job_relevant(job, now=now))


class TsdBaselineLogicTests(unittest.TestCase):
    def test_baseline_today_vs_future(self):
        now = datetime(2026, 9, 7, 11, 0, 0)
        targets = {
            ('a', 'g1'): 24,
            ('b', 'g2'): 80,
        }
        comments = {
            ('a', 'g1'): 'П Pr 50/50 оранж (220) - 24 07/09',
            ('b', 'g2'): 'П Komatsu 60/40 (220) - 80 08/09',
        }
        reconcile_prev = {}
        for key, fact in targets.items():
            comment = comments.get(key, '')
            if fact > 0 and mon._tsd_comment_matches_active_shift(comment, now):
                reconcile_prev[key] = 0
            else:
                reconcile_prev[key] = int(fact or 0)
        self.assertEqual(reconcile_prev[('a', 'g1')], 0)
        self.assertEqual(reconcile_prev[('b', 'g2')], 80)


if __name__ == '__main__':
    unittest.main()
