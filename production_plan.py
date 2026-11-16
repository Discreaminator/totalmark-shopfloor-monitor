# -*- coding: utf-8 -*-
"""
Суточный план производства из Excel (имя файла задаётся PRODUCTION_PLAN_FILE
или шаблонами «План произ-ва*.xlsx» / «*plan*.xlsx» в data/).

Лист «план»: номенклатура + линия + колонки дат («27 авг», …).
Пример маппинга линий (замените под свой цех): RALOT→Ralot_L1,
KRO(+9кг)→KRO_KANISTER, KRO→KRO_BOTTLE_L1, BLS→BLS, Бочки/Еврокубы→TSD_PRODUCTION.
Только дневная смена (plans.day).
"""
from __future__ import annotations

import logging
import math
import os
import re
import threading
import time
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

log = logging.getLogger(__name__)

_ROOT = Path(__file__).resolve().parent

_MONTHS_RU = {
    'янв': 1, 'фев': 2, 'мар': 3, 'апр': 4,
    'май': 5, 'мая': 5, 'июн': 6, 'июл': 7,
    'авг': 8, 'сен': 9, 'окт': 10, 'ноя': 11, 'дек': 12,
}

# штучные линии из Excel (бочки = шт. бочек → ТСД)
EXCEL_PIECE_LINES = ('Ralot_L1', 'KRO_BOTTLE_L1', 'KRO_KANISTER', 'BLS', 'TSD_PRODUCTION')
# паллетные линии = qty_шт / кол-во_на_паллете
EXCEL_PALLET_MAP = {
    'Ralot_L1': 'Ralot_L2',
    'KRO_BOTTLE_L1': 'KRO_BOTTLE_L2',
}
EXCEL_MANAGED_LINES = EXCEL_PIECE_LINES + tuple(EXCEL_PALLET_MAP.values())

_SKIP_LINES = {
    'выдувная', 'налив', 'слив', '',
}

# Строки переходов/смены формата: их количество вычитается из плана линии
CHANGEOVER_ROW_NAMES = (
    'переход с жидкости на жидкость Ralot (5)',
    'переход с жидкости на жидкость Ralot (1)',
    'смена формата (1)→(5)/(5)→(1) Ralot + контрактное',
    'переход с жидкости на жидкость KRO (9)',
    'переход с жидкости на жидкость KRO (3,78)',
    'переход с жидкости на жидкость KRO (0,9)',
    'переход с жидкости на жидкость KRO (1,7)',
    'смена формата на (0,9)/(3,78) KRO',
    'смена формата на (1,7)/(9) KRO',
    'переход с жидкости на жидкость BLS (10)',
    'переход с жидкости на жидкость BLS (20)',
    'смена формата (10)/(20) BLS',
    'смена формата (1)→(5)/(5)→(1) выдувная',
)


def _norm_row_name(s) -> str:
    t = str(s or '').lower().replace('ё', 'е').replace('->', '→')
    return re.sub(r'\s+', '', t)


_CHANGEOVER_KEYS = frozenset(_norm_row_name(n) for n in CHANGEOVER_ROW_NAMES)


def is_changeover_row(name) -> bool:
    return _norm_row_name(name) in _CHANGEOVER_KEYS


CHANGEOVER_LINE_OVERRIDES = {
    _norm_row_name('смена формата на (1,7)/(9) KRO'): 'KRO_KANISTER',
}

_lock = threading.RLock()
_cache = {
    'path': None,
    'mtime': None,
    'date_cols': {},       # date -> col_index (0-based)
    'rows': [],            # list of dicts after last full parse
    'day_plans': {},       # date_iso -> {line_id: {qty, products: [...]}}
    'state': None,         # последний excel_plan snapshot для UI
    'state_at': 0.0,
    'loading': False,
}

_STATE_TTL = float(os.environ.get('PRODUCTION_PLAN_STATE_TTL', '30'))


def _data_dirs():
    dirs = []
    env = (os.environ.get('PRODUCTION_PLAN_DIR') or '').strip()
    if env:
        dirs.append(Path(env))
    dirs.extend([_ROOT / 'data', Path(r'C:\monitor\data')])
    seen = set()
    out = []
    for d in dirs:
        key = str(d)
        if key in seen:
            continue
        seen.add(key)
        out.append(d)
    return out


def resolve_plan_path() -> Path | None:
    explicit = (os.environ.get('PRODUCTION_PLAN_FILE') or '').strip()
    if explicit:
        p = Path(explicit)
        if p.is_file():
            return p
    candidates = []
    for d in _data_dirs():
        if not d.is_dir():
            continue
        for pat in ('План произ-ва*.xlsx', '*plan*.xlsx', '*Plan*.xlsx'):
            for p in d.glob(pat):
                if p.is_file() and not p.name.startswith('~$'):
                    candidates.append(p)
    if not candidates:
        return None
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]


def _parse_header_date(text, year_hint: int) -> date | None:
    if text is None:
        return None
    if isinstance(text, datetime):
        return text.date()
    if isinstance(text, date):
        return text
    s = str(text).strip().lower()
    if not s:
        return None
    # «27 авг», «1 сен», «26 дек2»
    m = re.match(r'^(\d{1,2})\s*([а-яё]{3,})', s)
    if not m:
        return None
    day = int(m.group(1))
    mon_key = m.group(2)[:3]
    month = _MONTHS_RU.get(mon_key)
    if not month:
        # «мая» already in map as мая; «май»
        for k, v in _MONTHS_RU.items():
            if m.group(2).startswith(k):
                month = v
                break
    if not month or not (1 <= day <= 31):
        return None
    try:
        return date(year_hint, month, day)
    except ValueError:
        return None


def _build_date_cols(header_row, base_year: int) -> dict[date, int]:
    """
    Колонки дат идут хронологически. Старт часто с конца декабря (base_year-1).
    """
    date_cols: dict[date, int] = {}
    year = base_year - 1
    prev_month = None
    started = False
    for idx, cell in enumerate(header_row):
        if idx < 14:
            continue
        # пробуем base_year-1 затем base_year при первом попадании
        d = _parse_header_date(cell, year if started else base_year - 1)
        if d is None:
            d = _parse_header_date(cell, base_year)
        if d is None:
            continue
        if not started:
            # первая дата: если декабрь — это конец предыдущего года
            if d.month == 12:
                year = base_year - 1
            else:
                year = base_year
            d = date(year, d.month, d.day)
            started = True
            prev_month = d.month
            date_cols[d] = idx
            continue
        # переход года: декабрь → январь
        if prev_month == 12 and d.month == 1:
            year += 1
        d = date(year, d.month, d.day)
        prev_month = d.month
        # если колонка уже есть — оставляем первую (дубликаты заголовков)
        if d not in date_cols:
            date_cols[d] = idx
    return date_cols


def _is_nine_kg(name: str, fasovka) -> bool:
    try:
        if fasovka is not None and float(fasovka) == 9.0:
            return True
    except (TypeError, ValueError):
        pass
    s = str(name or '').lower().replace(' ', '')
    # 9кг / 9кгбрутто, но не 0,9 / 0.9
    if re.search(r'(?<![0-9.,])9кг', s):
        return True
    s2 = str(name or '').lower()
    if re.search(r'(?<![0-9.,])9\s*кг', s2):
        return True
    return False


def _is_bulk_fill(name: str, fasovka) -> bool:
    """Налив/слив в листе часто стоит в колонке «Бочки/Еврокубы» — в ТСД не тащим."""
    f = str(fasovka or '').lower().replace('ё', 'е').strip()
    n = str(name or '').lower().replace('ё', 'е')
    if f in ('налив', 'слив') or 'налив' in f or 'слив' in f:
        return True
    if 'налив' in n or 'слив' in n:
        return True
    return False


def map_excel_line(excel_line: str, name: str, fasovka) -> str | None:
    raw = str(excel_line or '').strip()
    if not raw:
        return None
    # унификация: кириллическая Р/О → латиница для сравнения
    trans = raw.upper().replace('К', 'K').replace('Р', 'R').replace('О', 'O').replace('А', 'A')
    trans = re.sub(r'\s+', '', trans)
    key = raw.lower().replace('ё', 'е').strip()
    if key in _SKIP_LINES:
        return None
    if 'ВЫДУВ' in trans or trans in ('НАЛИВ', 'СЛИВ'):
        return None
    # Бочки / еврокубы → линия ТСД (отгрузка с производства), без налива/слива
    if (
        'БОЧК' in trans
        or 'ЕВРОКУБ' in trans
        or key in ('бочки/еврокубы', 'бочки', 'еврокубы')
        or key.startswith('бочк')
    ):
        if _is_bulk_fill(name, fasovka):
            return None
        return 'TSD_PRODUCTION'
    if trans == 'RALOT' or key == 'ralot':
        return 'Ralot_L1'
    if trans == 'BLS' or key == 'bls':
        return 'BLS'
    if trans == 'KRO' or key in ('kro', 'кро'):
        if _is_nine_kg(name, fasovka):
            return 'KRO_KANISTER'
        return 'KRO_BOTTLE_L1'
    return None


def _base_year_from_path(path: Path) -> int:
    m = re.search(r'(20\d{2})', path.name)
    if m:
        return int(m.group(1))
    return datetime.now().year


def _load_workbook_rows(path: Path) -> tuple[dict[date, int], list[dict]]:
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        if 'план' in wb.sheetnames:
            ws = wb['план']
        else:
            ws = wb[wb.sheetnames[0]]
        header = None
        rows_out = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            vals = list(row)
            if i == 0:
                header = vals
                continue
            name = vals[0] if len(vals) > 0 else None
            if not name:
                continue
            fasovka = vals[3] if len(vals) > 3 else None
            per_pallet = vals[4] if len(vals) > 4 else None
            excel_line = vals[5] if len(vals) > 5 else None
            changeover = is_changeover_row(name)
            line_id = (
                CHANGEOVER_LINE_OVERRIDES.get(_norm_row_name(name)) if changeover else None
            ) or map_excel_line(str(excel_line or ''), str(name), fasovka)
            if not line_id:
                if changeover:
                    log.info('production plan: переход без линии монитора, пропуск: %s', name)
                continue
            try:
                pp = float(per_pallet) if per_pallet not in (None, '') else 0.0
            except (TypeError, ValueError):
                pp = 0.0
            rows_out.append({
                'name': str(name).strip(),
                'fasovka': fasovka,
                'per_pallet': pp if pp > 0 else None,
                'excel_line': str(excel_line or '').strip(),
                'line_id': line_id,
                'changeover': changeover,
                'values': vals,
            })
        base_year = _base_year_from_path(path)
        date_cols = _build_date_cols(header or [], base_year)
        return date_cols, rows_out
    finally:
        try:
            wb.close()
        except Exception:
            pass


def reload_if_needed(force: bool = False) -> bool:
    path = resolve_plan_path()
    if path is None:
        with _lock:
            _cache['path'] = None
            _cache['mtime'] = None
            _cache['date_cols'] = {}
            _cache['rows'] = []
        return False
    mtime = path.stat().st_mtime
    with _lock:
        if (
            not force
            and _cache['path'] == str(path)
            and _cache['mtime'] == mtime
            and _cache['rows']
        ):
            return True
        if _cache.get('loading'):
            # другой поток уже читает Excel — не дублируем
            return bool(_cache['rows'])
        _cache['loading'] = True
    try:
        date_cols, rows = _load_workbook_rows(path)
    except Exception as e:
        log.error('production plan load %s: %s', path, e)
        with _lock:
            _cache['loading'] = False
        return False
    with _lock:
        _cache['path'] = str(path)
        _cache['mtime'] = mtime
        _cache['date_cols'] = date_cols
        _cache['rows'] = rows
        _cache['day_plans'] = {}
        _cache['loading'] = False
        log.info(
            'production plan loaded: file=%s rows=%s date_cols=%s',
            path.name, len(rows), len(date_cols),
        )
    return True


def get_day_plan(day: date | None = None, allow_io: bool = True) -> dict:
    """
    { line_id: { 'qty': int, 'products': [{'name','qty','fasovka'}] } }

    allow_io=False — только память (snapshot / Socket.IO), без открытия Excel.
    """
    day = day or date.today()
    key = day.isoformat()
    with _lock:
        if key in _cache['day_plans']:
            return {k: dict(v, products=list(v['products'])) for k, v in _cache['day_plans'][key].items()}
        rows_ready = bool(_cache['rows'])

    if not rows_ready:
        if not allow_io:
            return {}
        if not reload_if_needed(force=False):
            return {}

    with _lock:
        date_cols = _cache['date_cols']
        rows = _cache['rows']
        if key in _cache['day_plans']:
            return {k: dict(v, products=list(v['products'])) for k, v in _cache['day_plans'][key].items()}
        col = date_cols.get(day)
        if col is None:
            _cache['day_plans'][key] = {}
            return {}

        by_line: dict[str, dict] = {}
        for rec in rows:
            vals = rec['values']
            if col >= len(vals):
                continue
            raw_q = vals[col]
            try:
                q = float(raw_q)
            except (TypeError, ValueError):
                continue
            if q <= 0:
                continue
            q_int = int(round(q))
            lid = rec['line_id']
            slot = by_line.setdefault(
                lid, {'qty': 0, 'pallets': 0, 'products': [], 'changeover_qty': 0},
            )
            if rec.get('changeover'):
                slot['changeover_qty'] += q_int
                continue
            per_pal = rec.get('per_pallet')
            pallets = None
            if per_pal and per_pal > 0:
                pallets = int(math.ceil(q_int / float(per_pal)))
            slot['qty'] += q_int
            if pallets:
                slot['pallets'] += pallets
            slot['products'].append({
                'name': rec['name'],
                'qty': q_int,
                'fasovka': rec['fasovka'],
                'per_pallet': per_pal,
                'pallets': pallets,
            })
        # План линии = продукция − переходы; паллеты уменьшаем пропорционально
        for slot in by_line.values():
            gross = int(slot['qty'])
            co = int(slot.get('changeover_qty') or 0)
            net = max(0, gross - co)
            slot['gross_qty'] = gross
            slot['qty'] = net
            if co and gross > 0 and slot.get('pallets'):
                slot['pallets'] = int(math.ceil(slot['pallets'] * net / float(gross)))
            elif net == 0:
                slot['pallets'] = 0
        # паллетные линии-спутники (Ralot_L2 / KRO_BOTTLE_L2)
        for piece_line, pallet_line in EXCEL_PALLET_MAP.items():
            src = by_line.get(piece_line)
            if not src:
                by_line[pallet_line] = {'qty': 0, 'pallets': 0, 'products': []}
                continue
            pal_qty = int(src.get('pallets') or 0)
            by_line[pallet_line] = {
                'qty': pal_qty,
                'pallets': pal_qty,
                'products': list(src.get('products') or []),
            }
        _cache['day_plans'][key] = by_line
        return {k: dict(v, products=list(v['products'])) for k, v in by_line.items()}


def normalize_product_name(s: str) -> str:
    if not s:
        return ''
    t = str(s).lower().replace('ё', 'е')
    t = t.replace('«', '"').replace('»', '"').replace('“', '"').replace('”', '"')
    t = re.sub(r'\s+', ' ', t)
    t = re.sub(r'[\"\'`]', '', t)
    return t.strip()


def _tokens(s: str) -> set[str]:
    t = normalize_product_name(s)
    parts = re.findall(r'[a-zа-я0-9]+(?:[./-][a-zа-я0-9]+)?', t)
    stop = {'шт', 'кг', 'л', 'брутто', 'канистра', 'антифриз', 'жидкость', 'охлаждающая'}
    return {p for p in parts if len(p) > 1 and p not in stop}


def names_match(fact_name: str, plan_name: str) -> bool:
    a = normalize_product_name(fact_name)
    b = normalize_product_name(plan_name)
    if not a or not b:
        return False
    if a in b or b in a:
        return True
    ta, tb = _tokens(a), _tokens(b)
    if not ta or not tb:
        return False
    inter = ta & tb
    # достаточно общего ядра бренда/линейки + фасовки
    if len(inter) >= 3:
        return True
    if len(inter) >= 2 and (len(inter) / max(len(ta), len(tb))) >= 0.4:
        return True
    return False


def product_matches_plan(fact_name: str, products: list) -> bool:
    if not fact_name or not products:
        return False
    for p in products:
        if names_match(fact_name, p.get('name') or ''):
            return True
    return False


def build_excel_plan_state(
    day: date | None = None,
    last_gtin_by_line: dict | None = None,
    gtin_names: dict | None = None,
    line_display: dict | None = None,
    allow_io: bool = True,
    all_fact_gtins_by_line: dict | None = None,
) -> dict:
    """Состояние для snapshot/API. allow_io=False — без чтения Excel.

    last_gtin_by_line — основной GTIN линии (макс. нетто смены).
    all_fact_gtins_by_line — все GTIN с нетто > 0 (бочки: бочка + еврокуб в одной смене).
    """
    day = day or date.today()
    day_plan = get_day_plan(day, allow_io=allow_io)
    last_gtin_by_line = last_gtin_by_line or {}
    all_fact_gtins_by_line = all_fact_gtins_by_line or {}
    gtin_names = gtin_names or {}
    line_display = line_display or {}

    def gtin_name(g):
        if not g:
            return None
        s = str(g).strip()
        return gtin_names.get(s) or gtin_names.get(s.lstrip('0') or '0') or s

    lines_out = {}
    mismatches = []
    for line_id in EXCEL_MANAGED_LINES:
        slot = day_plan.get(line_id) or {'qty': 0, 'pallets': 0, 'products': []}
        products = slot.get('products') or []
        qty = int(slot.get('qty') or 0)
        gtin = last_gtin_by_line.get(line_id)
        fact = gtin_name(gtin) if gtin else None
        fact_gtins = list(all_fact_gtins_by_line.get(line_id) or ())
        if gtin and gtin not in fact_gtins:
            fact_gtins = [gtin] + fact_gtins
        mismatch = False
        fact_extra = None
        # сверка продукта — только по штучным линиям и только если есть факт смены
        if line_id in EXCEL_PIECE_LINES and qty > 0 and fact_gtins and products:
            unmatched_names = []
            for fg in fact_gtins:
                fname = gtin_name(fg)
                if fname and not product_matches_plan(fname, products):
                    unmatched_names.append(fname)
            primary_ok = bool(fact) and product_matches_plan(fact, products)
            # вне плана — любой GTIN смены; основной тоже может не совпасть
            if unmatched_names:
                mismatch = True
                # не дублировать основной продукт в «также вне плана»
                extras = [n for n in unmatched_names if n != fact]
                if extras:
                    fact_extra = ' / '.join(extras[:4])
                elif not primary_ok:
                    fact_extra = None
            elif not primary_ok:
                mismatch = True
        lines_out[line_id] = {
            'qty': qty,
            'gross_qty': int(slot.get('gross_qty') or qty),
            'changeover_qty': int(slot.get('changeover_qty') or 0),
            'pallets': int(slot.get('pallets') or 0),
            'products': products,
            'fact_gtin': gtin,
            'fact_name': fact,
            'fact_extra': fact_extra,
            'mismatch': mismatch,
            'display': line_display.get(line_id, line_id),
        }
        if mismatch:
            plan_names = ' / '.join(p['name'] for p in products[:4])
            mismatches.append({
                'line_id': line_id,
                'display': line_display.get(line_id, line_id),
                'fact_name': fact,
                'fact_extra': fact_extra,
                'plan_names': plan_names,
            })

    with _lock:
        source = Path(_cache['path']).name if _cache.get('path') else None
        rows_ready = bool(_cache['rows'])

    state = {
        'ok': bool(source) and rows_ready,
        'date': day.isoformat(),
        'source': source,
        'lines': lines_out,
        'mismatches': mismatches,
        'as_of': datetime.now().isoformat(timespec='seconds'),
        'message': (
            None if (source and rows_ready)
            else ('загрузка плана…' if not rows_ready else 'файл плана не найден')
        ),
    }
    with _lock:
        _cache['state'] = state
        _cache['state_at'] = time.monotonic()
    return state


def get_cached_state() -> dict | None:
    with _lock:
        st = _cache.get('state')
        if not st:
            return None
        if (time.monotonic() - float(_cache.get('state_at') or 0)) > _STATE_TTL * 3:
            return st  # всё равно отдаём last known
        return st


def sync_plans_day(
    plans: dict,
    day: date | None = None,
    last_gtin_by_line: dict | None = None,
    gtin_names: dict | None = None,
    line_display: dict | None = None,
    skip_lines: set | None = None,
    all_fact_gtins_by_line: dict | None = None,
) -> tuple[dict, dict]:
    """
    Обновить plans[line]['day'] для EXCEL_MANAGED_LINES.
    skip_lines — линии с ручным вводом (не трогать до конца смены).
    Возвращает (excel_plan_state, changed_plans_snapshot subset).
    """
    day = day or date.today()
    reload_if_needed(force=False)
    day_plan = get_day_plan(day)
    skip_lines = set(skip_lines or ())
    changed = {}
    for line_id in EXCEL_MANAGED_LINES:
        if line_id in skip_lines:
            continue
        qty = int((day_plan.get(line_id) or {}).get('qty') or 0)
        if line_id not in plans:
            plans[line_id] = {'day': 0, 'night': 0}
        old = int(plans[line_id].get('day', 0) or 0)
        plans[line_id]['day'] = qty
        if old != qty:
            changed[line_id] = qty
    state = build_excel_plan_state(
        day,
        last_gtin_by_line,
        gtin_names,
        line_display,
        all_fact_gtins_by_line=all_fact_gtins_by_line,
    )
    return state, changed


def warm_async(callback=None):
    def _run():
        try:
            reload_if_needed(force=True)
            get_day_plan(date.today())
            if callback:
                callback()
        except Exception as e:
            log.warning('production plan warm: %s', e)
    threading.Thread(target=_run, name='prod_plan_warm', daemon=True).start()
