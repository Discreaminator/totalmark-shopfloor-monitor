# -*- coding: utf-8 -*-
"""Паузы/простои в сменном отчёте: один источник, без двойного счёта, DM, порог 10 мин."""
import os
import sys
import unittest
from datetime import datetime, timedelta

os.environ.setdefault('TSD_POLL_ENABLED', '0')
sys.path.insert(0, r'C:\Users\alexh\Downloads')
import app as mon


class SilenceClassifyTests(unittest.TestCase):
    def test_pause_vs_downtime_at_10(self):
        rows = mon.classify_report_silences([
            {'start': datetime(2026, 9, 8, 9, 0), 'end': datetime(2026, 9, 8, 9, 7), 'minutes': 7},
            {'start': datetime(2026, 9, 8, 10, 0), 'end': datetime(2026, 9, 8, 10, 10), 'minutes': 10},
            {'start': datetime(2026, 9, 8, 11, 0), 'end': datetime(2026, 9, 8, 11, 15), 'minutes': 15},
        ])
        kinds = [r['kind'] for r in rows]
        self.assertEqual(kinds, ['Пауза', 'Пауза', 'Простой'])

    def test_no_double_sum(self):
        rows = mon.classify_report_silences([
            {'minutes': 5}, {'minutes': 12}, {'minutes': 8},
        ])
        pause = sum(r['minutes'] for r in rows if r['kind'] == 'Пауза')
        down = sum(r['minutes'] for r in rows if r['kind'] == 'Простой')
        self.assertEqual(pause, 13)
        self.assertEqual(down, 12)
        self.assertEqual(pause + down, 25)

    def test_pallet_lines_skipped(self):
        self.assertTrue(mon.is_pallet_line('Ralot_L2'))
        self.assertTrue(mon.is_pallet_line('KRO_BOTTLE_L2'))
        self.assertFalse(mon.is_pallet_line('Ralot_L1'))
        self.assertFalse(mon.is_pallet_line('BLS'))
        self.assertTrue(mon.skip_report_silence_line('TSD_PRODUCTION'))
        self.assertTrue(mon.skip_report_silence_line('Ralot_L2'))
        self.assertFalse(mon.skip_report_silence_line('Ralot_L1'))


class DmGapSilenceTests(unittest.TestCase):
    def test_gaps_from_dm_only(self):
        start = datetime(2026, 9, 8, 8, 0)
        end = datetime(2026, 9, 8, 20, 0)
        events = [
            {'ts': start + timedelta(minutes=10)},
            {'ts': start + timedelta(minutes=14)},   # 4 мин — пауза
            {'ts': start + timedelta(minutes=30)},   # 16 мин — простой
        ]
        info = mon.calc_pauses_from_events(
            events, start, end,
            threshold_min=1,
            include_tail=False,
            breaks=[],
        )
        classified = mon.classify_report_silences(info['intervals'])
        self.assertEqual(len(classified), 2)
        self.assertEqual(classified[0]['kind'], 'Пауза')
        self.assertEqual(classified[1]['kind'], 'Простой')
        # интервалы не пересекаются → суммы независимы
        self.assertAlmostEqual(
            classified[0]['minutes'] + classified[1]['minutes'],
            sum(x['minutes'] for x in classified),
        )


if __name__ == '__main__':
    unittest.main()
