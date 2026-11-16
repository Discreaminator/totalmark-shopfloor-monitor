# -*- coding: utf-8 -*-
"""Строки переходов в Excel-плане вычитаются из плана линии, а не прибавляются."""
import sys
import unittest
from datetime import date

sys.path.insert(0, r'C:\Users\alexh\Downloads')

import production_plan as pp


def _row(name, line_id, qty, per_pallet=None, changeover=False):
    return {
        'name': name,
        'fasovka': None,
        'per_pallet': per_pallet,
        'excel_line': '',
        'line_id': line_id,
        'changeover': changeover,
        'values': [name, qty],
    }


class PlanChangeoverTests(unittest.TestCase):
    DAY = date(2026, 9, 29)

    def setUp(self):
        self._saved = dict(pp._cache)

    def tearDown(self):
        pp._cache.clear()
        pp._cache.update(self._saved)

    def _load(self, rows):
        pp._cache['rows'] = rows
        pp._cache['date_cols'] = {self.DAY: 1}
        pp._cache['day_plans'] = {}

    def test_all_listed_rows_detected(self):
        for name in pp.CHANGEOVER_ROW_NAMES:
            self.assertTrue(pp.is_changeover_row(name), name)
        self.assertTrue(pp.is_changeover_row('  Переход с жидкости на жидкость RALOT (1) '))
        self.assertFalse(pp.is_changeover_row('Антифриз ShopfloorMonitor Standard 40 5 кг'))

    def test_ralot_changeover_subtracted(self):
        self._load([
            _row('Антифриз A', 'Ralot_L1', 4050, per_pallet=450),
            _row('переход с жидкости на жидкость Ralot (1)', 'Ralot_L1', 450, changeover=True),
        ])
        plan = pp.get_day_plan(self.DAY, allow_io=False)
        ralot = plan['Ralot_L1']
        self.assertEqual(ralot['gross_qty'], 4050)
        self.assertEqual(ralot['changeover_qty'], 450)
        self.assertEqual(ralot['qty'], 3600)
        self.assertEqual(ralot['pallets'], 8)
        self.assertEqual(plan['Ralot_L2']['qty'], 8)
        self.assertEqual([p['name'] for p in ralot['products']], ['Антифриз A'])

    def test_never_negative(self):
        self._load([
            _row('Антифриз B', 'BLS', 100),
            _row('переход с жидкости на жидкость BLS (10)', 'BLS', 300, changeover=True),
        ])
        plan = pp.get_day_plan(self.DAY, allow_io=False)
        self.assertEqual(plan['BLS']['qty'], 0)
        self.assertEqual(plan['BLS']['pallets'], 0)

    def test_kro_17_9_format_change_goes_to_kanister(self):
        self.assertEqual(
            pp.CHANGEOVER_LINE_OVERRIDES.get(pp._norm_row_name('Смена формата на (1,7)/(9)  KRO')),
            'KRO_KANISTER',
        )
        self.assertIsNone(
            pp.CHANGEOVER_LINE_OVERRIDES.get(pp._norm_row_name('смена формата на (0,9)/(3,78) KRO')),
        )


if __name__ == '__main__':
    unittest.main()
