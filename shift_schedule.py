# -*- coding: utf-8 -*-
"""
Квартальный график смен (Excel) → бригада/мастер на текущую смену.

Файл: data/График*.xlsx (или SHIFT_SCHEDULE_FILE).
Мастера: data/brigade_masters.json  {"4": "Иванов И.И.", ...}
"""
from __future__ import annotations

import json
import logging
import os
import re
import threading
from datetime import date, datetime, timedelta
from pathlib import Path

log = logging.getLogger(__name__)

_ROOT = Path(__file__).resolve().parent
_MONTHS = {
    'январ': 1, 'феврал': 2, 'март': 3, 'апрел': 4,
    'ма': 5, 'июн': 6, 'июл': 7, 'август': 8,
    'сентябр': 9, 'октябр': 10, 'ноябр': 11, 'декабр': 12,
}

_lock = threading.RLock()
_cache = {
    'path': None,
    'mtime': None,
    'masters_mtime': None,
    'by_ym': {},      # (year, month) -> {day: {'д': [brigades], 'н': [brigades]}}
    'missing_ym': set(),  # месяцы, которых нет в файле (чтобы не искать повторно)
    'loading_ym': set(),
    'masters': {},    # "4" -> "ФИО"
    'source': None,
    # готовый ответ для UI/snapshot — без Excel на горячем пути
    'crew_key': None,       # (date_iso, shift)
    'crew_value': None,
    'crew_at': 0.0,
}
_CREW_TTL_SEC = float(os.environ.get('SHIFT_CREW_TTL_SEC', '60'))



def _data_dir() -> Path:
    env = (os.environ.get('SHIFT_SCHEDULE_DIR') or '').strip()
    if env:
        return Path(env)
    local = _ROOT / 'data'
    if local.is_dir():
        return local
    monitor = Path(r'C:\monitor\data')
    if monitor.is_dir():
        return monitor
    return local


def _resolve_schedule_path() -> Path | None:
    explicit = (os.environ.get('SHIFT_SCHEDULE_FILE') or '').strip()
    if explicit:
        p = Path(explicit)
        if p.is_file():
            return p
    candidates = []
    for d in (_ROOT / 'data', Path(r'C:\monitor\data')):
        if d.is_dir():
            candidates.append(d)
    env_dir = (os.environ.get('SHIFT_SCHEDULE_DIR') or '').strip()
    if env_dir:
        candidates.insert(0, Path(env_dir))
    seen = set()
    for d in candidates:
        key = str(d.resolve()) if d.exists() else str(d)
        if key in seen:
            continue
        seen.add(key)
        preferred = list(d.glob('График 3 квартал.xlsx')) + list(d.glob('График*.xlsx'))
        for p in preferred:
            if p.is_file() and not p.name.startswith('~$'):
                return p
    return None


def _masters_path() -> Path:
    env = (os.environ.get('BRIGADE_MASTERS_FILE') or '').strip()
    if env:
        return Path(env)
    for d in (_ROOT / 'data', Path(r'C:\monitor\data')):
        p = d / 'brigade_masters.json'
        if p.is_file():
            return p
    return _data_dir() / 'brigade_masters.json'


def _parse_month_year(text: str, fallback_year: int | None = None) -> tuple[int, int] | None:
    if not text:
        return None
    s = str(text).strip().lower().replace('  ', ' ')
    year = None
    m = re.search(r'(20\d{2})', s)
    if m:
        year = int(m.group(1))
    elif fallback_year:
        year = fallback_year
    month = None
    for key, num in _MONTHS.items():
        if key in s:
            # «май» ловится через «ма» — не путать с «март» (март раньше в словаре длиннее… март has март)
            if key == 'ма' and 'март' in s:
                continue
            if key == 'ма' and not re.search(r'\bмай\b', s) and 'мая' not in s and 'май' not in s.replace(' ', ''):
                # allow "Май 2026"
                if 'май' not in s and 'мая' not in s:
                    continue
            month = num
            break
    if month and year:
        return year, month
    return None


def _sheet_year_month(ws, sheet_name: str) -> tuple[int, int] | None:
    # 1) имя листа
    ym = _parse_month_year(sheet_name)
    if ym:
        return ym
    # 2) ячейки A9–A12 часто «август 2026»
    for r in range(1, 15):
        v = ws.cell(r, 1).value
        if v is None:
            continue
        if isinstance(v, datetime):
            return v.year, v.month
        ym = _parse_month_year(str(v))
        if ym:
            return ym
    return None


def _normalize_mark(val) -> str | None:
    if val is None:
        return None
    if isinstance(val, (int, float)) and val == 0:
        return None
    s = str(val).strip().lower().replace(' ', '')
    if not s:
        return None
    # первая буква: д / н / о
    ch = s[0]
    if ch in ('д', 'd'):
        return 'д'
    if ch in ('н', 'n', 'h'):  # лат. h иногда
        return 'н'
    if ch in ('о', 'o'):
        return 'о'
    return None


def _brigade_num(label: str) -> str | None:
    m = re.search(r'бригад[аы]?\s*(\d+)', str(label), re.I)
    if m:
        return m.group(1)
    return None


def _parse_sheet(ws) -> dict[int, dict[str, list[str]]] | None:
    """
    day -> {'д': ['Бригада 4', ...], 'н': [...]}
    """
    # найти строку с днями 1..28+
    day_row = None
    day_cols: dict[int, int] = {}  # day -> col
    for r in range(1, min(20, (ws.max_row or 1) + 1)):
        nums = {}
        for c in range(1, min(40, (ws.max_column or 1) + 1)):
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)) and 1 <= int(v) <= 31:
                nums[int(v)] = c
            elif isinstance(v, str) and v.strip().isdigit():
                n = int(v.strip())
                if 1 <= n <= 31:
                    nums[n] = c
        if len(nums) >= 20:  # почти полный месяц
            day_row = r
            day_cols = nums
            break
        if len(nums) >= 28:
            day_row = r
            day_cols = nums
            break
    if not day_row or not day_cols:
        return None

    out: dict[int, dict[str, list[str]]] = {}
    for r in range(day_row + 1, min(day_row + 12, (ws.max_row or 1) + 1)):
        label = ws.cell(r, 1).value
        if label is None:
            continue
        label_s = str(label).strip()
        if not re.search(r'бригад', label_s, re.I):
            continue
        brigade = label_s
        for day, col in day_cols.items():
            mark = _normalize_mark(ws.cell(r, col).value)
            if mark not in ('д', 'н'):
                continue
            slot = out.setdefault(day, {'д': [], 'н': []})
            if brigade not in slot[mark]:
                slot[mark].append(brigade)
    return out or None


def _load_masters(force: bool = False) -> dict[str, str]:
    path = _masters_path()
    mtime = path.stat().st_mtime if path.is_file() else None
    if not force and _cache['masters_mtime'] == mtime and _cache['masters'] is not None:
        return _cache['masters']
    masters = {}
    if path.is_file():
        try:
            data = json.loads(path.read_text(encoding='utf-8'))
            if isinstance(data, dict):
                for k, v in data.items():
                    if v is None:
                        continue
                    masters[str(k).strip()] = str(v).strip()
        except Exception as e:
            log.warning('brigade_masters.json: %s', e)
    _cache['masters'] = masters
    _cache['masters_mtime'] = mtime
    return masters


def _invalidate_if_stale(force: bool = False) -> Path | None:
    """Сбросить кэш при смене файла. Возвращает путь или None."""
    path = _resolve_schedule_path()
    if path is None:
        _cache['by_ym'] = {}
        _cache['missing_ym'] = set()
        _cache['path'] = None
        _cache['mtime'] = None
        _cache['source'] = None
        return None
    mtime = path.stat().st_mtime
    if force or _cache['path'] != str(path) or _cache['mtime'] != mtime:
        _cache['by_ym'] = {}
        _cache['missing_ym'] = set()
        _cache['path'] = str(path)
        _cache['mtime'] = mtime
        _cache['source'] = path.name
        _load_masters(force=True)
    else:
        _load_masters(force=False)
    return path


def _sheet_name_hints_ym(sheet_name: str) -> tuple[int, int] | None:
    """Быстрый разбор только по имени листа (без чтения ячеек)."""
    return _parse_month_year(sheet_name)


def _ensure_month(year: int, month: int) -> bool:
    """
    Лениво прочитать один месяц из Excel.
    I/O выполняется вне lock, чтобы не стопорить Socket.IO / snapshot.
    """
    with _lock:
        path = _invalidate_if_stale(force=False)
        if path is None:
            return False
        key = (year, month)
        if key in _cache['by_ym']:
            return True
        if key in _cache['missing_ym']:
            return False
        if key in _cache['loading_ym']:
            return False
        _cache['loading_ym'].add(key)
        path_str = str(path)

    parsed = None
    err = None
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path_str, data_only=True, read_only=True)
        try:
            candidates = []
            fallback = []
            for sheet_name in wb.sheetnames:
                hint = _sheet_name_hints_ym(sheet_name)
                if hint == key:
                    candidates.append(sheet_name)
                elif hint is None:
                    fallback.append(sheet_name)
                elif hint[0] < year - 1:
                    continue

            for sheet_name in candidates:
                ws = wb[sheet_name]
                ym = _sheet_year_month(ws, sheet_name)
                if ym != key:
                    continue
                got = _parse_sheet(ws)
                if got:
                    parsed = got
            if parsed is None:
                for sheet_name in fallback:
                    ws = wb[sheet_name]
                    ym = _sheet_year_month(ws, sheet_name)
                    if ym != key:
                        continue
                    got = _parse_sheet(ws)
                    if got:
                        parsed = got
                        break
        finally:
            try:
                wb.close()
            except Exception:
                pass
    except Exception as e:
        err = e
        log.error('shift schedule open %s: %s', path_str, e)

    with _lock:
        _cache['loading_ym'].discard(key)
        if parsed:
            _cache['by_ym'][key] = parsed
            log.info(
                'shift schedule month %s-%02d ok (%s days) file=%s',
                year, month, len(parsed), Path(path_str).name,
            )
            return True
        if err is None:
            _cache['missing_ym'].add(key)
            log.warning('shift schedule: нет листа %s-%02d в %s', year, month, Path(path_str).name)
        return False


def reload_schedule(force: bool = False) -> bool:
    """Подготовить кэш / прогреть текущий месяц (без обхода всех листов)."""
    with _lock:
        path = _invalidate_if_stale(force=force)
        if path is None:
            return False
        if force:
            _cache['crew_key'] = None
            _cache['crew_value'] = None
            _cache['crew_at'] = 0.0
    now = datetime.now()
    ok = _ensure_month(now.year, now.month)
    prev = datetime(now.year, now.month, 1) - timedelta(days=1)
    _ensure_month(prev.year, prev.month)
    # прогреть ответ для UI
    try:
        get_crew_for_datetime(now, allow_io=True)
    except Exception:
        pass
    return ok


_warm_started = False


def warm_current_async() -> None:
    """Фоновый прогрев, чтобы первый запрос UI не ждал Excel."""
    global _warm_started
    with _lock:
        if _warm_started:
            return
        _warm_started = True

    def _run():
        global _warm_started
        try:
            reload_schedule(force=False)
        except Exception as e:
            log.warning('shift schedule warm: %s', e)
        finally:
            # разрешить повторный warm после смены файла / ошибки
            with _lock:
                # если кэш crew уже есть — больше не спамим потоками
                if not _cache.get('crew_value'):
                    _warm_started = False

    threading.Thread(target=_run, name='shift_sched_warm', daemon=True).start()


def shift_calendar_context(now: datetime | None = None) -> tuple[date, str, str]:
    """
    (calendar_date, shift_type day|night, mark д|н).
    Для ночи 00:00–08:00 дата = вчера (день начала ночной смены в 20:00).
    """
    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    if 8 <= now.hour < 20:
        return now.date(), 'day', 'д'
    if now.hour >= 20:
        return now.date(), 'night', 'н'
    return (now - timedelta(days=1)).date(), 'night', 'н'


def _master_for_brigade(brigade_label: str, masters: dict[str, str]) -> str | None:
    num = _brigade_num(brigade_label)
    if not num:
        return None
    return masters.get(num) or masters.get(str(int(num)))


def _empty_crew(cal_date, shift_type, mark, now, message):
    return {
        'ok': False if message == 'график не загружен' else True,
        'date': cal_date.isoformat(),
        'shift': shift_type,
        'shift_label': 'дневная' if shift_type == 'day' else 'ночная',
        'mark': mark,
        'brigades': [],
        'brigade': None,
        'masters': [],
        'master': None,
        'source': _cache.get('source'),
        'as_of': now.isoformat(timespec='seconds'),
        'message': message,
    }


def get_crew_for_datetime(now: datetime | None = None, allow_io: bool = True) -> dict:
    """
    Текущая (или на момент now) бригада/мастер по графику.

    allow_io=False — только кэш (для snapshot/VPS/connect), без чтения Excel.
    """
    import time as _time

    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)

    cal_date, shift_type, mark = shift_calendar_context(now)
    crew_key = (cal_date.isoformat(), shift_type)

    with _lock:
        cached = _cache.get('crew_value')
        if (
            cached
            and _cache.get('crew_key') == crew_key
            and (_time.monotonic() - float(_cache.get('crew_at') or 0)) < _CREW_TTL_SEC
        ):
            out = dict(cached)
            out['as_of'] = now.isoformat(timespec='seconds')
            return out

    if not allow_io:
        with _lock:
            stale = _cache.get('crew_value')
            if stale and _cache.get('crew_key') == crew_key:
                out = dict(stale)
                out['as_of'] = now.isoformat(timespec='seconds')
                return out
        warm_current_async()
        return _empty_crew(cal_date, shift_type, mark, now, 'загрузка графика…')

    _ensure_month(cal_date.year, cal_date.month)
    masters = _load_masters(force=False)

    by_ym = _cache.get('by_ym') or {}
    month_map = by_ym.get((cal_date.year, cal_date.month))
    if not month_map:
        result = _empty_crew(
            cal_date, shift_type, mark, now,
            f'нет листа {cal_date.year}-{cal_date.month:02d} в графике',
        )
    else:
        day_map = month_map.get(cal_date.day) or {'д': [], 'н': []}
        brigades = list(day_map.get(mark) or [])
        if not brigades:
            result = _empty_crew(
                cal_date, shift_type, mark, now,
                f'на {cal_date.isoformat()} нет бригады с меткой «{mark}»',
            )
            result['ok'] = True
        else:
            masters_list = []
            for b in brigades:
                m = _master_for_brigade(b, masters)
                if m and m not in masters_list:
                    masters_list.append(m)
            result = {
                'ok': True,
                'date': cal_date.isoformat(),
                'shift': shift_type,
                'shift_label': 'дневная' if shift_type == 'day' else 'ночная',
                'mark': mark,
                'brigades': brigades,
                'brigade': ' / '.join(brigades),
                'masters': masters_list,
                'master': ' / '.join(masters_list) if masters_list else None,
                'source': _cache.get('source'),
                'as_of': now.isoformat(timespec='seconds'),
                'message': None,
            }

    with _lock:
        _cache['crew_key'] = crew_key
        _cache['crew_value'] = dict(result)
        _cache['crew_at'] = _time.monotonic()
    return result


def get_crew_for_shift(shift_date, shift_type: str) -> dict:
    """
    Для Excel-отчёта: дата смены (date) + day|night.
    Конвенция как в generate_shift_report: для night shift_date — утро окончания
    (ночь началась вчера в 20:00).
    """
    if hasattr(shift_date, 'year'):
        d = shift_date
    else:
        d = datetime.strptime(str(shift_date), '%Y-%m-%d').date()
    if shift_type == 'day':
        now = datetime(d.year, d.month, d.day, 12, 0, 0)
    else:
        start = datetime(d.year, d.month, d.day, 20, 0, 0) - timedelta(days=1)
        now = start + timedelta(hours=2)  # 22:00 дня начала ночи
    return get_crew_for_datetime(now, allow_io=True)
