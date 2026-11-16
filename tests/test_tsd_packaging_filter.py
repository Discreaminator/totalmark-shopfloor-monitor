# -*- coding: utf-8 -*-
"""Фильтр TSD: в TSD_PRODUCTION только бочки/еврокубы (не канистры)."""
import os
import sys
import unittest
from datetime import datetime
from unittest import mock

os.environ.setdefault('TSD_POLL_ENABLED', '0')
sys.path.insert(0, r'C:\Users\alexh\Downloads')

import app as mon


# Имена из report_2026-09-09_day_2001.xlsx, лист «Бочки_Еврокубы»
REPORT_09_09 = [
    (4, '4607075785441', 'Антифриз "ShopfloorMonitor Standard 40" оборотный куб 1000 кг, зеленый', True),
    (108, '4607075783911', 'Антифриз "ShopfloorMonitor JPN" канистра 5 кг брутто, красный', False),
    (432, '4607075781511', 'Антифриз "ShopfloorMonitor Premium 40" канистра 5 кг брутто, оранжевый', False),
    (138, '4607075784741', 'Антифриз "ShopfloorMonitor CHN" канистра 5 кг брутто, красный', False),
    (108, '4607075780644', 'Антифриз "ShopfloorMonitor Optima" канистра 5 кг брутто, зеленый', False),
    (108, '4607075785557', 'Антифриз "Coolstream DE" канистра 5 кг брутто, фиолетовый', False),
    (108, '4607075783324', 'Антифриз "ShopfloorMonitor Standard 40" канистра 5 кг брутто, красный', False),
    (540, '4607075784802', 'Антифриз "ShopfloorMonitor Premium 40" (NRC) канистра 5 кг брутто, желтый', False),
    (48, '4607075784918', 'Антифриз "ShopfloorMonitor Standard HD 50/50" канистра 20 кг брутто, красный', False),
    (40, '4607075784796', 'Концентрат антифриза "ShopfloorMonitor CHN C" канистра 1,7 кг брутто, красный', False),
    (45, '4607075784734', 'Антифриз "ShopfloorMonitor CHN" канистра 1 кг брутто, красный', False),
    (90, '4607075785090', 'Тормозная жидкость гидравлическая «Coolstream DOT 4 LV class 7» 1 кг брутто', False),
    (6, '4607075783362', 'Антифриз "ShopfloorMonitor Standard 40" канистра 20 кг брутто, красный', False),
    (80, '4607075785298', 'Антифриз "Coolstream CHN 45" (желтый) бочка 220 кг', True),
    (60, '4607075781528', 'Антифриз "ShopfloorMonitor Premium 40" канистра 10 кг брутто, оранжевый', False),
    (1404, '4607075784208', 'Антифриз "ShopfloorMonitor A-110" канистра 5 кг брутто', False),
    (10, '4607075785106', 'Антифриз "ShopfloorMonitor CHN" бочка 50 кг, красный', True),
]


class TsdPackagingFilterTests(unittest.TestCase):
    def test_report_09_09_names(self):
        kept = rejected = 0
        for qty, gtin, name, expect in REPORT_09_09:
            ok = mon._tsd_is_barrel_or_ibc(good_name=name, gtin=gtin)
            self.assertEqual(
                ok, expect,
                msg=f'{gtin}: {name!r} → {ok}, expected {expect}',
            )
            if ok:
                kept += qty
            else:
                rejected += qty
        self.assertEqual(kept, 94)
        self.assertEqual(rejected, 3235)
        self.assertEqual(kept + rejected, 3329)

    def test_empty_name_denied(self):
        prev = dict(mon.gtin_names)
        try:
            mon.gtin_names = {}
            self.assertFalse(mon._tsd_is_barrel_or_ibc(good_name='', gtin='4607075785298'))
        finally:
            mon.gtin_names = prev

    def test_tm_good_without_word_barrel_uses_catalog(self):
        """Реальный кейс 10.09: TM Good без «бочка», в справочнике — с «бочка»."""
        prev = dict(mon.gtin_names)
        try:
            mon.gtin_names = {
                '4607075785274': 'Концентрат антифриза "ShopfloorMonitor Premium С" бочка 220 кг, желтый',
            }
            tm_good = 'Концентрат антифриза "ShopfloorMonitor Premium С" 220 кг, желтый CS-010105-C-YE'
            self.assertTrue(
                mon._tsd_is_barrel_or_ibc(
                    good_name=tm_good,
                    gtin='04607075785274',
                    comment='П Pr C желт (220) - 24 10/09',
                )
            )
            # даже без справочника — вес >45 кг или (220) в Comment
            mon.gtin_names = {}
            self.assertTrue(mon._tsd_is_barrel_or_ibc(good_name=tm_good, gtin='4607075785274'))
            self.assertTrue(
                mon._tsd_is_barrel_or_ibc(
                    good_name='',
                    gtin='unknown',
                    comment='П Pr C желт (220) - 24 10/09',
                )
            )
        finally:
            mon.gtin_names = prev

    def test_size_threshold_gt_45(self):
        self.assertTrue(mon._tsd_is_barrel_or_ibc(good_name='ОЖ без слова тары 46 кг'))
        self.assertTrue(mon._tsd_is_barrel_or_ibc(good_name='ОЖ 200 л зеленый'))
        self.assertFalse(mon._tsd_is_barrel_or_ibc(good_name='ОЖ 45 кг без канистры'))
        self.assertFalse(mon._tsd_is_barrel_or_ibc(good_name='ОЖ 20 кг брутто'))
        self.assertTrue(
            mon._tsd_is_barrel_or_ibc(good_name='', comment='П тест (46) - 1 10/09')
        )
        self.assertFalse(
            mon._tsd_is_barrel_or_ibc(good_name='', comment='П тест (45) - 1 10/09')
        )
        # «CHN 45» в марке — не вес; без кг/л и без крупного (N) → отказ
        self.assertFalse(mon._tsd_is_barrel_or_ibc(good_name='ShopfloorMonitor CHN 45 желтый'))

    def test_fallback_gtin_names(self):
        prev = dict(mon.gtin_names)
        try:
            mon.gtin_names = {
                '4607075785298': 'ShopfloorMonitor CHN 45 бочка 220 кг',
                '4607075784208': 'ShopfloorMonitor A-110 канистра 5 кг',
            }
            self.assertTrue(mon._tsd_is_barrel_or_ibc(gtin='4607075785298'))
            self.assertFalse(mon._tsd_is_barrel_or_ibc(gtin='4607075784208'))
        finally:
            mon.gtin_names = prev

    def test_canister_not_saved_by_comment_size(self):
        self.assertFalse(
            mon._tsd_is_barrel_or_ibc(
                good_name='Антифриз канистра 5 кг',
                gtin='4607075784208',
                comment='П что-то (220) - 10 10/09',
            )
        )

    def test_emit_skips_canister(self):
        with mon.data_lock:
            mon.production_data[mon.TSD_LINE_ID][mon.TSD_ENTITY_TYPE]['events'].clear()
        with mock.patch.object(mon, 'connected_clients', 0):
            mon._tsd_emit_delta(
                'job-x', '4607075784208', 100,
                'П fake 09/09',
                good_name='Антифриз A-110 канистра 5 кг',
                ts=datetime.now(),
                persist=False,
            )
            mon._tsd_emit_delta(
                'job-y', '4607075785298', 80,
                'П CHN 45 09/09',
                good_name='Coolstream CHN 45 бочка 220 кг',
                ts=datetime.now(),
                persist=False,
            )
        with mon.data_lock:
            evs = list(mon.production_data[mon.TSD_LINE_ID][mon.TSD_ENTITY_TYPE]['events'])
        self.assertEqual(len(evs), 1)
        self.assertEqual(evs[0].get('gtin'), '4607075785298')
        self.assertEqual(int(evs[0].get('qty') or 0), 80)


if __name__ == '__main__':
    unittest.main()

