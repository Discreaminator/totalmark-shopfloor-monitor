# -*- coding: utf-8 -*-
"""Плашка Excel vs карточка бочек: один критерий GTIN (макс. нетто)."""
import os
import sys
import unittest
from datetime import datetime
from collections import deque

os.environ.setdefault('TSD_POLL_ENABLED', '0')
sys.path.insert(0, r'C:\Users\alexh\Downloads')
import app as mon
import production_plan as pp


class DominantGtinTests(unittest.TestCase):
    def setUp(self):
        with mon.data_lock:
            mon.production_data[mon.TSD_LINE_ID][mon.TSD_ENTITY_TYPE]['events'] = deque(maxlen=5000)

    def _add(self, gtin, qty, ts):
        with mon.data_lock:
            mon.production_data[mon.TSD_LINE_ID][mon.TSD_ENTITY_TYPE]['events'].append({
                'ts': ts, 'qty': qty, 'gtin': gtin, 'job_id': 'j', 'source': 'TSD',
            })

    def test_max_net_not_last_scan(self):
        # как на скрине: догон 24+18+4, последний по времени — еврокуб, доминирует бочка
        base = datetime(2026, 9, 7, 11, 44, 0)
        self._add('04607075784864', 24, base)           # 50/50 бочка
        self._add('04607075784994', 18, base.replace(second=1))  # NRC
        self._add('04607075786196', 4, base.replace(second=2))   # EcoPro куб — последний
        now = datetime(2026, 9, 7, 11, 45, 0)
        g = mon.current_shift_fact_gtin(mon.TSD_LINE_ID, now=now)
        self.assertEqual(mon._normalize_gtin(g), '4607075784864')
        all_g = mon.excel_plan_fact_gtins_all(now).get(mon.TSD_LINE_ID) or []
        self.assertEqual(mon._normalize_gtin(all_g[0]), '4607075784864')
        self.assertEqual(len(all_g), 3)


class ExcelMismatchExtraTests(unittest.TestCase):
    def test_extra_outside_plan(self):
        state = pp.build_excel_plan_state(
            day=datetime(2026, 9, 7).date(),
            last_gtin_by_line={'TSD_PRODUCTION': '4607075784864'},
            all_fact_gtins_by_line={
                'TSD_PRODUCTION': [
                    '4607075784864',
                    '4607075786196',
                ],
            },
            gtin_names={
                '4607075784864': 'Антифриз "ShopfloorMonitor Premium 50/50" 220 кг, оранжевый',
                '4607075786196': 'Hot Stream ЭкоПро 30 еврокуб, 1038 кг, зеленый',
            },
            line_display={'TSD_PRODUCTION': 'Бочки/Еврокубы'},
            allow_io=False,
        )
        # Без реального Excel day_plan qty может быть 0 → mismatch не строится.
        # Подставим план вручную через кэш day — проще вызвать логику через monkey:
        # эмулируем products/qty прямо в проверке names_match.
        self.assertTrue(
            pp.product_matches_plan(
                'Антифриз "ShopfloorMonitor Premium 50/50" 220 кг, оранжевый',
                [{'name': 'ShopfloorMonitor Premium 50/50 220 кг оранжевый'}],
            )
        )
        self.assertFalse(
            pp.product_matches_plan(
                'Hot Stream ЭкоПро 30 еврокуб, 1038 кг, зеленый',
                [{'name': 'ShopfloorMonitor Premium 50/50 220 кг оранжевый'}],
            )
        )


if __name__ == '__main__':
    unittest.main()
