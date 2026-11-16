#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import gzip
import base64
import hashlib
import logging
import logging.handlers
import threading
import time
import re
from io import BytesIO
from datetime import datetime, timedelta
import queue
from collections import defaultdict, deque
from urllib import request as urlrequest
from urllib.error import HTTPError, URLError

from flask import Flask, request, jsonify, render_template, send_file
from flask_socketio import SocketIO, emit
from flask_cors import CORS
import psycopg2
from psycopg2 import sql
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import DataBarRule
import pika

import shift_schedule
import production_plan

# Matplotlib
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
import matplotlib.dates as mdates

# PIL для работы с изображениями (опционально)
try:
    from PIL import Image as PILImage
except ImportError:
    PILImage = None
    logging.warning("PIL (Pillow) не установлен, фотографии продуктов не будут вставляться")

# ============================
# НАСТРОЙКИ ЛОГГИРОВАНИЯ
# ============================
LOG_DIR = 'logs'
os.makedirs(LOG_DIR, exist_ok=True)

log_level = os.environ.get('LOG_LEVEL', 'INFO').upper()
numeric_level = getattr(logging, log_level, logging.INFO)

logger = logging.getLogger()
logger.setLevel(numeric_level)

formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

file_handler = logging.handlers.RotatingFileHandler(
    os.path.join(LOG_DIR, 'app.log'),
    maxBytes=10*1024*1024,
    backupCount=5,
    encoding='utf-8'
)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

# Консоль Windows в cp1251 падает на эмодзи в сообщениях: UnicodeEncodeError на каждой записи
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

logging.getLogger('pika').setLevel(logging.WARNING)
logging.getLogger('urllib3').setLevel(logging.WARNING)
logging.getLogger('socketio').setLevel(logging.WARNING)
logging.getLogger('engineio').setLevel(logging.WARNING)

# ============================
# НАСТРОЙКИ ПОДКЛЮЧЕНИЯ К БАЗЕ ДАННЫХ
# (env перекрывает дефолты; без env поведение как раньше)
# ============================
DB_HOST = os.environ.get('DB_HOST', 'localhost')
DB_PORT = int(os.environ.get('DB_PORT', '5432'))
DB_NAME = os.environ.get('DB_NAME', '')
DB_USER = os.environ.get('DB_USER', '')
DB_PASSWORD = os.environ.get('DB_PASSWORD', '')

# ============================
# НАСТРОЙКИ ПОДКЛЮЧЕНИЯ К RABBITMQ
# ============================
RABBIT_HOST = os.environ.get('RABBIT_HOST', '')
RABBIT_PORT = int(os.environ.get('RABBIT_PORT', '5672'))
RABBIT_USER = os.environ.get('RABBIT_USER', '')
RABBIT_PASSWORD = os.environ.get('RABBIT_PASSWORD', '')
RABBIT_VHOST = os.environ.get('RABBIT_VHOST', '/')

RABBIT_QUEUES = [
    q.strip() for q in os.environ.get('RABBIT_QUEUES', 'aggregation').split(',')
    if q.strip()
] or ['aggregation']

RABBIT_EXCHANGES = [
    'to-total-mark-aggregation',
    'from-controllers',
    'to-total-mark-disaggregation',
    'to-total-mark-aggregation-renamed',
    'totalmark_event',
    'printed-codes-exchange',
    'print-device-update-exchange',
]

# ============================
# ОСТАЛЬНЫЕ НАСТРОЙКИ
# ============================
COMPLETED_STATES = [
    'PRINTED', 'COMPLETED', 'ЗАВЕРШЕНО', 'Done',
    'CREATED', 'ACTIVE', 'ВЫПОЛНЕНО', 'CLOSED', 'FINISHED'
]

INITIAL_LOAD_HOURS = int(os.environ.get('INITIAL_LOAD_HOURS', str(24 * 7)))
POLL_INTERVAL = int(os.environ.get('POLL_INTERVAL', '5'))
MAX_EVENTS_PER_POLL = int(os.environ.get('MAX_EVENTS_PER_POLL', '5000'))

# Светофор тишины (как на дашборде):
#   warn — жёлтая точка; bad — плашка «задержка или остановка».
# В лист «Простои» попадают только интервалы ≥ bad (до первого события на линии).
STATUS_SILENCE_WARN_MIN = int(os.environ.get('STATUS_SILENCE_WARN_MIN', '5'))
STATUS_SILENCE_BAD_MIN = int(os.environ.get('STATUS_SILENCE_BAD_MIN', '10'))
DM_SILENCE_WARN_MIN = int(os.environ.get('DM_SILENCE_WARN_MIN', '3'))
DM_SILENCE_BAD_MIN = int(os.environ.get('DM_SILENCE_BAD_MIN', '5'))
# MAX: только индивидуальные коды (DM). Паллетные L2 — без алертов.
MAX_ALERT_EXCLUDE_LINES = frozenset(
    x.strip() for x in os.environ.get(
        'MAX_ALERT_EXCLUDE_LINES', 'Ralot_L2,KRO_BOTTLE_L2,TSD_PRODUCTION'
    ).split(',') if x.strip()
)
# Рабочие часы смены для часового плана (12 ч − 1 ч перерыва)
PLAN_WORK_HOURS = int(os.environ.get('PLAN_WORK_HOURS', '11'))
PAUSE_THRESHOLD_DEFAULT_MIN = int(os.environ.get('PAUSE_THRESHOLD_DEFAULT_MIN', '5'))


def silence_bad_min(entity_type=None):
    """Порог плашки «задержка или остановка», мин."""
    if entity_type == 'DM':
        return float(DM_SILENCE_BAD_MIN)
    return float(STATUS_SILENCE_BAD_MIN)


def max_alert_allowed(line_id, entity_type=None, rec=None):
    """
    MAX только по индивидуальным кодам (DM):
      — Ralot_L2 / KRO_BOTTLE_L2 — никогда;
      — BLS / KRO_KANISTER / L1-линии — только если entity_type == DM;
      — простои по агрегатам (L1/ТУ) — не шлём.
    """
    if not line_id:
        return False
    if line_id in MAX_ALERT_EXCLUDE_LINES:
        return False
    et = entity_type
    if et is None and isinstance(rec, dict):
        et = rec.get('entity_type')
    return et == 'DM'

# Регламентированные перерывы (не считаются паузой/простоем в Excel)
# start/end — (час, минута) относительно календарных суток старта смены
SHIFT_BREAKS = {
    'day': [
        {'start': (12, 0), 'end': (13, 0), 'name': 'Обед'},
    ],
    'night': [
        {'start': (0, 0), 'end': (1, 0), 'name': 'Перерыв'},
    ],
}


def get_shift_break_intervals(shift_start, shift_end, shift_type):
    """Абсолютные интервалы перерывов внутри смены [from, to)."""
    defs = SHIFT_BREAKS.get(shift_type) or []
    intervals = []
    for b in defs:
        sh, sm = b['start']
        eh, em = b['end']
        bs = shift_start.replace(hour=sh, minute=sm, second=0, microsecond=0)
        be = shift_start.replace(hour=eh, minute=em, second=0, microsecond=0)
        # Ночь: перерыв после полуночи — следующие сутки от даты старта 20:00
        if bs < shift_start:
            bs = bs + timedelta(days=1)
            be = be + timedelta(days=1)
        if be <= bs:
            be = be + timedelta(days=1)
        fr = max(bs, shift_start)
        to = min(be, shift_end)
        if to > fr:
            intervals.append({'from': fr, 'to': to, 'name': b.get('name') or 'Перерыв'})
    return intervals


def minutes_excluding_breaks(start, end, breaks):
    """Длительность [start, end) в минутах минус пересечения с перерывами."""
    if not start or not end or end <= start:
        return 0.0
    total = (end - start).total_seconds()
    if not breaks:
        return round(total / 60.0, 1)
    overlap = 0.0
    for b in breaks:
        bf, bt = b['from'], b['to']
        lo = max(start, bf)
        hi = min(end, bt)
        if hi > lo:
            overlap += (hi - lo).total_seconds()
    return round(max(0.0, total - overlap) / 60.0, 1)


PAUSE_THRESHOLD_BY_LINE = {
    'TSD_PRODUCTION': 999,  # сканы ТСД редкие — не считать «простоем линии»
    'Ralot_L1': 5,
    'Ralot_L2': 15,
    'KRO_BOTTLE_L1': 5,
    'KRO_BOTTLE_L2': 15,
    'BLS': 25,
    'KRO_KANISTER': 25,
}


def pause_threshold_for_line(line_id, entity_type=None):
    """Порог паузы (мин): max(порог линии, STATUS_SILENCE_WARN_MIN)."""
    base = PAUSE_THRESHOLD_DEFAULT_MIN
    if line_id and line_id in PAUSE_THRESHOLD_BY_LINE:
        base = PAUSE_THRESHOLD_BY_LINE[line_id]
    else:
        lid = (line_id or '').upper()
        if 'BLS' in lid or 'KANISTER' in lid:
            base = 25
        elif lid.endswith('_L2') or 'ПАЛЛЕТ' in lid or 'PALLET' in lid:
            base = 15
    return max(int(base), STATUS_SILENCE_WARN_MIN)


# Сменный Excel: тишина по индивидуальным кодам (DM).
# ≤ REPORT_SILENCE_PAUSE_MAX_MIN мин → «Пауза», свыше → «Простой».
# Паллетные L2 в этот учёт не входят.
REPORT_SILENCE_PAUSE_MAX_MIN = int(os.environ.get('REPORT_SILENCE_PAUSE_MAX_MIN', '10'))
# Минимальный gap по DM для попадания в отчёт (классификация пауза/простой — отдельно, по 10 мин)
REPORT_SILENCE_DETECT_MIN = float(os.environ.get('REPORT_SILENCE_DETECT_MIN', '3'))
PALLET_LINE_IDS = frozenset({'Ralot_L2', 'KRO_BOTTLE_L2'})
# Линии без учёта тишины в сменном Excel (нет потока DM / редкие сканы)
REPORT_SILENCE_SKIP_LINES = frozenset({'TSD_PRODUCTION'}) | PALLET_LINE_IDS


def is_pallet_line(line_id):
    if not line_id:
        return False
    if line_id in PALLET_LINE_IDS:
        return True
    u = (line_id or '').upper()
    return u.endswith('_L2') or 'ПАЛЛЕТ' in u or 'PALLET' in u


def skip_report_silence_line(line_id):
    """Паллеты L2 и бочки ТСД не участвуют в листе пауз/простоев по DM."""
    if not line_id:
        return True
    if line_id in REPORT_SILENCE_SKIP_LINES:
        return True
    return is_pallet_line(line_id)


def classify_report_silences(intervals, pause_max_min=None):
    """Разметить gap-интервалы: Пауза (≤N мин) / Простой (>N мин)."""
    try:
        pause_max = float(
            pause_max_min if pause_max_min is not None else REPORT_SILENCE_PAUSE_MAX_MIN
        )
    except (TypeError, ValueError):
        pause_max = 10.0
    out = []
    for iv in intervals or []:
        if not isinstance(iv, dict):
            continue
        try:
            mins = float(iv.get('minutes') or 0)
        except (TypeError, ValueError):
            mins = 0.0
        if mins <= 0:
            continue
        kind = 'Пауза' if mins <= pause_max else 'Простой'
        row = dict(iv)
        row['minutes'] = mins
        row['kind'] = kind
        out.append(row)
    return out


def _normalize_downtime_comment(reason, comment):
    """Комментарий только для Ремонт/Проблема; иначе None."""
    reason = (reason or '').strip()
    text = (comment or '').strip()
    if reason not in DOWNTIME_REASONS_NEED_COMMENT:
        return None
    if not text:
        return None
    if len(text) > DOWNTIME_COMMENT_MAX_LEN:
        text = text[:DOWNTIME_COMMENT_MAX_LEN]
    return text


def _master_downtime_meta(line_id, start, end, downtime_recs):
    """Причина/ACK/комментарий мастера, если интервал пересекается с downtime.json."""
    best = _master_downtime_rec(line_id, start, end, downtime_recs)
    if not best:
        return None, False, None
    return (
        (best.get('reason') or None),
        bool(best.get('acknowledged')),
        (best.get('comment') or None),
    )


def _master_downtime_rec(line_id, start, end, downtime_recs):
    """Запись downtime.json с максимальным пересечением интервала (или None)."""
    if not line_id or not start or not end or end <= start:
        return None
    best_ov = 0.0
    best = None
    for rec in downtime_recs or []:
        if rec.get('line_id') != line_id:
            continue
        st = _iso_to_dt(rec.get('started_at'))
        en = _iso_to_dt(rec.get('ended_at')) or st
        if st is None or en is None:
            continue
        lo = max(start, st)
        hi = min(end, en)
        if hi <= lo:
            continue
        ov = (hi - lo).total_seconds()
        if ov > best_ov:
            best_ov = ov
            best = rec
    return best


# SocketIO async mode.
# По умолчанию threading — стабильнее при Excel/Matplotlib/psycopg2 (eventlet часто «глохнет»).
# Вернуть eventlet: export SOCKETIO_ASYNC_MODE=eventlet
SOCKETIO_ASYNC_MODE = os.environ.get('SOCKETIO_ASYNC_MODE', 'threading').strip().lower()
if SOCKETIO_ASYNC_MODE not in ('eventlet', 'threading', 'gevent'):
    SOCKETIO_ASYNC_MODE = 'threading'

# Каталог приложения (не cwd сервиса — иначе reports/static «теряются»)
APP_DIR = os.path.dirname(os.path.abspath(__file__))
REPORTS_DIR = os.path.join(APP_DIR, 'reports')
os.makedirs(REPORTS_DIR, exist_ok=True)

GTIN_NAMES_FILE = os.path.join(APP_DIR, 'data', 'Краткий отчет.xlsx')
gtin_names = {}
# Последний ненулевой GTIN по линии (пока не придёт новый).
# Нужен: утренние DM без aggregate_id → SQL gtin NULL, карточка не должна «забывать» продукт.
last_known_gtin = {}
_last_gtin_lock = threading.Lock()


def _normalize_gtin(gtin):
    if gtin is None:
        return None
    s = str(gtin).strip().lstrip('0')
    return s or None


# GS1 AI (01) = GTIN-14 внутри Data Matrix full_code
_GS1_AI01_RE = re.compile(r'(?:\x1d|^)01(\d{14})')


def gtin_from_full_code(full_code):
    """
    Достать GTIN из dm.full_code (даже без aggregate_id).
    Формат: 01 + 14 цифр + 21... (часто с GS1 FNC1 = \\x1d).
    """
    if not full_code:
        return None
    s = str(full_code)
    m = _GS1_AI01_RE.search(s)
    if m:
        return _normalize_gtin(m.group(1))
    if s.startswith('01') and len(s) >= 16 and s[2:16].isdigit():
        return _normalize_gtin(s[2:16])
    return None


def remember_line_gtin(line_id, gtin):
    """Запомнить GTIN линии, если он ненулевой. Старый не затираем null'ом."""
    g = _normalize_gtin(gtin)
    if not line_id or not g:
        return g
    with _last_gtin_lock:
        prev = last_known_gtin.get(line_id)
        if prev != g:
            last_known_gtin[line_id] = g
            if prev:
                logging.info(f'GTIN {line_id}: {prev} → {g}')
    ensure_product_image_async(g)
    return g


def sticky_line_gtin(line_id, gtin=None, remember=True):
    """Текущий gtin события или последний известный по линии.

    remember=False — только вернуть gtin в payload, не трогая sticky
    (нужно для отрицательных delta ТСД / отмен).
    """
    g = _normalize_gtin(gtin)
    if g:
        if remember:
            remember_line_gtin(line_id, g)
        return g
    with _last_gtin_lock:
        return last_known_gtin.get(line_id)


def current_shift_product_nets(line_id, now=None):
    """
    Нетто qty по GTIN за текущую смену (L1/ТУ/DM/ДЕЗАГРЕГАЦИЯ).

    Возвращает {gtin: {'net': int, 'last_pos': datetime|None}}.
    Один критерий с карточкой бочек: не «последний скан по времени»
    (иначе плашка Excel показывает еврокуб, а карточка — бочку 50/50).
    """
    if not line_id:
        return {}
    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    shift_start = _current_shift_start(now)
    out = {}
    with data_lock:
        types = production_data.get(line_id) or {}
        for etype in ('L1', 'ТУ', 'DM', 'ДЕЗАГРЕГАЦИЯ'):
            bucket = types.get(etype) or {}
            events = bucket.get('events') if isinstance(bucket, dict) else None
            if not events:
                continue
            for ev in events:
                if not isinstance(ev, dict):
                    continue
                try:
                    qty = int(ev.get('qty') or 0)
                except (TypeError, ValueError):
                    try:
                        qty = int(float(ev.get('qty') or 0))
                    except (TypeError, ValueError):
                        qty = 0
                if qty == 0:
                    continue
                g = _normalize_gtin(ev.get('gtin'))
                if not g:
                    continue
                ts = ev.get('ts')
                if not isinstance(ts, datetime):
                    continue
                if ts.tzinfo is not None:
                    ts = ts.replace(tzinfo=None)
                if ts < shift_start or ts > now:
                    continue
                row = out.get(g)
                if not row:
                    row = {'net': 0, 'last_pos': None}
                    out[g] = row
                row['net'] += qty
                if qty > 0 and (row['last_pos'] is None or ts >= row['last_pos']):
                    row['last_pos'] = ts
    return out


def _pick_dominant_gtin(nets_map):
    """GTIN с макс. нетто > 0; при равенстве — более поздний положительный скан."""
    best = None
    best_net = 0
    best_ts = None
    for g, row in (nets_map or {}).items():
        try:
            net = int(row.get('net') or 0)
        except (TypeError, ValueError, AttributeError):
            continue
        if net <= 0:
            continue
        ts = row.get('last_pos') if isinstance(row, dict) else None
        if (
            best is None
            or net > best_net
            or (net == best_net and ts and (best_ts is None or ts > best_ts))
        ):
            best = g
            best_net = net
            best_ts = ts
    return best


def current_shift_fact_gtin(line_id, now=None):
    """
    Основной GTIN выпуска за текущую смену (макс. нетто > 0).

    Нужен для алерта «продукт ≠ план Excel» и должен совпадать с карточкой бочек.
    Sticky с прошлой смены не используется.
    """
    return _pick_dominant_gtin(current_shift_product_nets(line_id, now=now))


def excel_plan_fact_gtins(now=None):
    """Основной GTIN по линиям из выпуска текущей смены — для сверки с Excel."""
    now = now or datetime.now()
    out = {}
    for lid in production_plan.EXCEL_MANAGED_LINES:
        g = current_shift_fact_gtin(lid, now=now)
        if g:
            out[lid] = g
    return out


def excel_plan_fact_gtins_all(now=None):
    """Все GTIN с нетто > 0 за смену (для мульти-SKU бочек/еврокубов)."""
    now = now or datetime.now()
    out = {}
    for lid in production_plan.EXCEL_MANAGED_LINES:
        nets = current_shift_product_nets(lid, now=now)
        ranked = [
            g for g, row in sorted(
                nets.items(),
                key=lambda kv: (
                    -int(kv[1].get('net') or 0),
                    kv[1].get('last_pos') or datetime.min,
                ),
            )
            if int(row.get('net') or 0) > 0
        ]
        if ranked:
            out[lid] = ranked
    return out

# ============================
# СТИЛИ ОТЧЁТОВ (в стиле дашборда)
# ============================
LINE_DISPLAY_NAMES = {
    'Ralot_L1': 'RALOT УПАКОВКИ',
    'Ralot_L2': 'RALOT ПАЛЛЕТЫ',
    'KRO_BOTTLE_L1': 'KRO УПАКОВКИ',
    'KRO_BOTTLE_L2': 'KRO ПАЛЛЕТЫ',
    'BLS': 'BLS',
    'KRO_KANISTER': 'KRO КАНИСТРЫ 9 КГ',
    'TSD_PRODUCTION': 'Бочки/Еврокубы',
}

LINE_COLORS = {
    'Ralot_L1': 'E63946',
    'Ralot_L2': '1D3557',
    'KRO_BOTTLE_L1': '2A9D8F',
    'KRO_BOTTLE_L2': 'E9C46A',
    'BLS': '9B5DE5',
    'KRO_KANISTER': '00BBF9',
    'TSD_PRODUCTION': 'E76F51',
}

BRAND_COLOR = 'F7572E'
SUCCESS_COLOR = '2A9D8F'
WARN_COLOR = 'E9C46A'
DANGER_COLOR = 'E63946'

PRODUCTS_IMG_DIR = os.path.join(APP_DIR, 'static', 'products')


def _atomic_write_text(path, text, encoding='utf-8'):
    """
    Атомарная запись файла на Windows: уникальный .tmp + os.replace + retry.
    Без этого два потока на один path.tmp дают WinError 32 (файл занят).
    """
    directory = os.path.dirname(path) or '.'
    os.makedirs(directory, exist_ok=True)
    tmp = os.path.join(
        directory,
        f'.{os.path.basename(path)}.{os.getpid()}.{threading.get_ident()}.{time.time_ns()}.tmp',
    )
    last_err = None
    try:
        with open(tmp, 'w', encoding=encoding) as f:
            f.write(text)
            f.flush()
            try:
                os.fsync(f.fileno())
            except OSError:
                pass
        for attempt in range(6):
            try:
                os.replace(tmp, path)
                return
            except OSError as e:
                last_err = e
                # WinError 32 / 5 — антивирус или параллельный читатель
                time.sleep(0.05 * (attempt + 1))
        raise last_err
    finally:
        try:
            if os.path.isfile(tmp):
                os.remove(tmp)
        except OSError:
            pass


def _atomic_write_json(path, payload, **json_kwargs):
    text = json.dumps(payload, ensure_ascii=False, **json_kwargs)
    _atomic_write_text(path, text)


def _atomic_save_workbook(wb, filepath):
    """Сохранить xlsx атомарно: tmp → os.replace (защита от двух процессов на один файл)."""
    directory = os.path.dirname(filepath) or '.'
    os.makedirs(directory, exist_ok=True)
    tmp = os.path.join(
        directory,
        f'.{os.path.basename(filepath)}.{os.getpid()}.{time.time_ns()}.tmp.xlsx',
    )
    last_err = None
    try:
        wb.save(tmp)
        for attempt in range(6):
            try:
                os.replace(tmp, filepath)
                return
            except OSError as e:
                last_err = e
                time.sleep(0.05 * (attempt + 1))
        raise last_err
    finally:
        try:
            if os.path.isfile(tmp):
                os.remove(tmp)
        except OSError:
            pass


_instance_lock_fh = None


def _ensure_single_instance():
    """
    Один процесс монитора на каталог APP_DIR.
    Два экземпляра одновременно пишут один report_*.xlsx → битый ZIP (Excel «с ошибкой»).
    """
    global _instance_lock_fh
    if os.environ.get('ALLOW_MULTI_INSTANCE', '0').strip().lower() in ('1', 'true', 'yes'):
        logging.warning('ALLOW_MULTI_INSTANCE=1 — второй экземпляр разрешён (риск порчи отчётов)')
        return
    try:
        import msvcrt
    except ImportError:
        return
    lock_path = os.path.join(APP_DIR, 'monitor.instance.lock')
    _instance_lock_fh = open(lock_path, 'a+b')
    try:
        _instance_lock_fh.seek(0)
        msvcrt.locking(_instance_lock_fh.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError:
        logging.error(
            'Уже запущен другой экземпляр монитора (%s). '
            'Второй процесс портит Excel-отчёты — выход. '
            'Оставьте один процесс на порту 5000. '
            'Принудительно: ALLOW_MULTI_INSTANCE=1',
            lock_path,
        )
        try:
            _instance_lock_fh.close()
        except Exception:
            pass
        sys.exit(1)
    try:
        _instance_lock_fh.seek(0)
        _instance_lock_fh.truncate()
        _instance_lock_fh.write(
            f'pid={os.getpid()} started={datetime.now().isoformat(timespec="seconds")}\n'.encode()
        )
        _instance_lock_fh.flush()
    except Exception:
        pass
    logging.info('Single-instance lock: %s (pid=%s)', lock_path, os.getpid())


class _CrossProcessLock:
    """Неблокирующая файловая блокировка Windows (msvcrt) между процессами."""

    def __init__(self, path):
        self.path = path
        self._fh = None

    def acquire(self):
        try:
            import msvcrt
        except ImportError:
            return True
        os.makedirs(os.path.dirname(self.path) or '.', exist_ok=True)
        self._fh = open(self.path, 'a+b')
        try:
            self._fh.seek(0)
            msvcrt.locking(self._fh.fileno(), msvcrt.LK_NBLCK, 1)
            return True
        except OSError:
            try:
                self._fh.close()
            except Exception:
                pass
            self._fh = None
            return False

    def release(self):
        if self._fh is None:
            return
        try:
            import msvcrt
            self._fh.seek(0)
            msvcrt.locking(self._fh.fileno(), msvcrt.LK_UNLCK, 1)
        except Exception:
            pass
        try:
            self._fh.close()
        except Exception:
            pass
        self._fh = None


def _excel_image_from_bytes(data, width, height, keep):
    """
    openpyxl.Image._data() делает fp.close() на переданном потоке.
    Поэтому каждому рисунку — свой BytesIO; keep удерживает его до wb.save().
    """
    buf = BytesIO(data)
    xl_img = Image(buf)
    # Конструктор Image читает поток и оставляет указатель у EOF —
    # перед save нужен seek(0), иначе картинка пустая / пропадает.
    buf.seek(0)
    xl_img.width = width
    xl_img.height = height
    keep.append(buf)
    return xl_img


def _excel_image_from_stream(stream, width, height, keep):
    """BytesIO/файл → копия байт → Image (исходный поток можно закрыть)."""
    if stream is None:
        return None
    if hasattr(stream, 'getvalue'):
        data = stream.getvalue()
    else:
        stream.seek(0)
        data = stream.read()
    if not data:
        return None
    return _excel_image_from_bytes(data, width, height, keep)


def _excel_product_thumb(img_path, keep, box=(70, 70), display=(55, 55)):
    """Превью продукта для листа Excel или None."""
    if PILImage is None or not img_path or not os.path.isfile(img_path):
        return None
    with PILImage.open(img_path) as pil_img:
        img = pil_img.convert('RGB') if pil_img.mode not in ('RGB', 'L') else pil_img.copy()
        img.thumbnail(box)
        raw = BytesIO()
        img.save(raw, format='JPEG', quality=80)
    return _excel_image_from_bytes(raw.getvalue(), display[0], display[1], keep)

# Стили openpyxl
thin_border = Border(
    left=Side(style='thin', color='DDDDDD'),
    right=Side(style='thin', color='DDDDDD'),
    top=Side(style='thin', color='DDDDDD'),
    bottom=Side(style='thin', color='DDDDDD')
)
header_fill = PatternFill('solid', fgColor=BRAND_COLOR)
header_font = Font(bold=True, color='FFFFFF', size=11)
title_font = Font(bold=True, size=16, color='1A1A1A')
subtitle_font = Font(size=11, color='666666')
kpi_font = Font(bold=True, size=14, color='1A1A1A')

# ============================
# ЗАГРУЗКА НАЗВАНИЙ ПРОДУКТОВ
# ============================
def load_gtin_names():
    """
    Прочитать data/Краткий отчет.xlsx → gtin_names.
    Возвращает dict: ok, count, path, error (для API reload без рестарта).
    """
    global gtin_names
    path = GTIN_NAMES_FILE
    if not os.path.exists(path):
        logging.warning(f"Файл с названиями продуктов не найден: {path}")
        gtin_names = {}
        return {'ok': False, 'count': 0, 'path': path, 'error': 'file_not_found'}
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet = wb.active
        gtin_col = None
        name_col = None
        header_row_index = None
        for row_idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=10, values_only=False), start=1):
            if row:
                for cell in row:
                    if cell.value and "GTIN" in str(cell.value).strip():
                        gtin_col = cell.column - 1
                    if cell.value and "Полное наименование товара" in str(cell.value).strip():
                        name_col = cell.column - 1
                if gtin_col is not None and name_col is not None:
                    header_row_index = row_idx
                    break
        if gtin_col is None or name_col is None:
            logging.error("Не найдены колонки GTIN или 'Полное наименование товара'")
            gtin_names = {}
            return {
                'ok': False, 'count': 0, 'path': path,
                'error': 'columns_not_found',
            }

        new_dict = {}
        for row in sheet.iter_rows(min_row=header_row_index + 1, values_only=True):
            if row and row[gtin_col] is not None:
                gtin_raw = str(row[gtin_col]).strip()
                if not gtin_raw:
                    continue
                gtin = ''.join(filter(str.isdigit, gtin_raw)).lstrip('0')
                if not gtin:
                    continue
                name = str(row[name_col]).strip() if row[name_col] else None
                new_dict[gtin] = name

        gtin_names = new_dict
        logging.info(f"Загружено {len(gtin_names)} названий продуктов из {path}")
        for i, (g, n) in enumerate(list(gtin_names.items())[:5]):
            logging.info(f"Пример: {g} -> {n}")
        try:
            mtime = os.path.getmtime(path)
            mtime_iso = datetime.fromtimestamp(mtime).isoformat(timespec='seconds')
        except OSError:
            mtime_iso = None
        return {
            'ok': True,
            'count': len(gtin_names),
            'path': path,
            'mtime': mtime_iso,
            'error': None,
        }
    except Exception as e:
        logging.error(f"Ошибка загрузки файла {path}: {e}")
        gtin_names = {}
        return {'ok': False, 'count': 0, 'path': path, 'error': str(e)}

# ============================
# ПРИЛОЖЕНИЕ FLASK
# ============================
app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'your-secret-key-here')
CORS(app)

socketio = SocketIO(
    app,
    cors_allowed_origins="*",
    async_mode=SOCKETIO_ASYNC_MODE,
    ping_timeout=int(os.environ.get('SOCKETIO_PING_TIMEOUT', '120')),
    ping_interval=int(os.environ.get('SOCKETIO_PING_INTERVAL', '25')),
    logger=False,
    engineio_logger=False,
)
logging.info(f"SocketIO async_mode={SOCKETIO_ASYNC_MODE}")

# ---- runtime health (антизависание) ----
_health = {
    'started_at': datetime.now().isoformat(timespec='seconds'),
    'last_poll_at': None,
    'last_poll_ok_at': None,
    'poll_started_at': None,  # цикл начался, ещё нет ok — hung fetch/process
    'poll_progress_at': None,  # последний шаг внутри цикла: долгий догон ≠ зависание
    'last_poll_error': None,
    'last_reconcile_at': None,
    'last_report_at': None,
    'last_report_duration_sec': None,
    'poll_count': 0,
    'poller_generation': 0,
}
_health_lock = threading.Lock()

# Broadcast Socket.IO вне poller: emit из фонового потока иначе deadlock/залипание на 5k событий
_emit_queue = queue.Queue(maxsize=int(os.environ.get('EMIT_QUEUE_MAX', '1500')))
MAX_SOCKET_EMITS_PER_POLL = int(os.environ.get('MAX_SOCKET_EMITS_PER_POLL', '200'))


def _socket_emit_worker():
    while True:
        try:
            item = _emit_queue.get()
            if item is None:
                break
            event, payload = item
            socketio.emit(event, payload)
        except Exception as e:
            logging.debug(f"socket emit worker: {e}")


def socket_emit(event, payload):
    """Неблокирующая отправка. При переполнении очереди событие отбрасывается."""
    try:
        _emit_queue.put_nowait((event, payload))
    except queue.Full:
        logging.warning(f"Socket emit queue full — drop {event}")


_emit_thread = threading.Thread(target=_socket_emit_worker, daemon=True, name='socket_emit')
_emit_thread.start()


def _health_set(**kwargs):
    with _health_lock:
        _health.update(kwargs)


def _poller_progress(stage=None):
    """
    Отметка «цикл продвинулся». Долгий догон истории даёт длинный цикл, но живой,
    а зависший poller не двигается вовсе — watchdog различает их по этой метке.
    """
    now_iso = datetime.now().isoformat(timespec='seconds')
    with _health_lock:
        _health['poll_progress_at'] = now_iso
        if stage:
            _health['poll_stage'] = stage


HANG_DUMP_FILE = os.path.join(LOG_DIR, 'hang_traceback.log')


def dump_thread_stacks(tag='manual'):
    """Стек всех потоков в logs/hang_traceback.log — чем именно занят зависший poller."""
    try:
        import faulthandler
        import traceback as _tb
        names = {}
        for th in threading.enumerate():
            if th.ident is not None:
                names[th.ident] = th.name
        with open(HANG_DUMP_FILE, 'a', encoding='utf-8') as f:
            f.write(
                f"\n===== {datetime.now().isoformat(timespec='seconds')} | {tag} | "
                f"threads={threading.active_count()} =====\n"
            )
            for th in threading.enumerate():
                f.write(f"  thread: name={th.name} alive={th.is_alive()} daemon={th.daemon}\n")
            f.flush()
            # Питоновский обход кадров: не обрывается на "<freed thread state>",
            # в отличие от faulthandler, поэтому идёт первым.
            f.write("--- python frames ---\n")
            for ident, frame in sys._current_frames().items():
                f.write(f"\nThread {ident:#x} [{names.get(ident, '?')}]:\n")
                try:
                    f.write(''.join(_tb.format_stack(frame)))
                except Exception as fe:
                    f.write(f"  <format_stack failed: {fe}>\n")
            f.write("--- faulthandler ---\n")
            f.flush()
            try:
                faulthandler.dump_traceback(file=f, all_threads=True)
            except Exception as fe:
                f.write(f"<faulthandler failed: {fe}>\n")
        logging.critical(f"Стеки потоков записаны: {HANG_DUMP_FILE} (tag={tag})")
        return HANG_DUMP_FILE
    except Exception as e:
        logging.error(f"dump_thread_stacks: {e}")
        return None

# Текущие планы смены (редактируются в UI). После окончания смены → history + обнуление.
_DEFAULT_LINE_PLANS = {
    'Ralot_L1': {'day': 0, 'night': 0},
    'Ralot_L2': {'day': 0, 'night': 0},
    'KRO_BOTTLE_L1': {'day': 0, 'night': 0},
    'KRO_BOTTLE_L2': {'day': 0, 'night': 0},
    'BLS': {'day': 0, 'night': 0},
    'KRO_KANISTER': {'day': 0, 'night': 0},
    'TSD_PRODUCTION': {'day': 0, 'night': 0},
}
plans = {k: dict(v) for k, v in _DEFAULT_LINE_PLANS.items()}

# history[key] = {date, shift, saved_at, lines: {line_id: {plan, fact, percent}}}
# key = "YYYY-MM-DD_day" | "YYYY-MM-DD_night"
plans_history = {}

# Ручной ввод плана: Excel не перезаписывает до конца смены (или нового ручного ввода).
# plans_manual_lock[line_id][shift_type] = epoch текущей смены на момент фиксации
plans_manual_lock = {}

PLAN_FILE = os.path.join(APP_DIR, 'plans.json')
plans_lock = threading.RLock()


def _manual_lock_epoch(now=None):
    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    st = 'day' if 8 <= now.hour < 20 else 'night'
    if st == 'day':
        start = now.replace(hour=8, minute=0, second=0, microsecond=0)
    else:
        start = now.replace(hour=20, minute=0, second=0, microsecond=0)
        if now.hour < 8:
            start = start - timedelta(days=1)
    return f"{start.isoformat(timespec='seconds')}_{st}"


def mark_plans_manual(line_ids, shift_types, now=None):
    """Пометить линии/смены как введённые вручную (приоритет над Excel до конца смены)."""
    epoch = _manual_lock_epoch(now)
    with plans_lock:
        for line_id in line_ids:
            if line_id not in plans:
                continue
            slot = plans_manual_lock.setdefault(line_id, {})
            for st in shift_types:
                if st in ('day', 'night'):
                    slot[st] = epoch
    logging.info(
        'plan manual lock: lines=%s shifts=%s epoch=%s',
        list(line_ids), list(shift_types), epoch,
    )


def is_plan_manual_locked(line_id, shift_type, now=None):
    epoch = _manual_lock_epoch(now)
    with plans_lock:
        locked_epoch = (plans_manual_lock.get(line_id) or {}).get(shift_type)
    return bool(locked_epoch) and locked_epoch == epoch


def clear_manual_locks_for_shift_type(shift_type):
    with plans_lock:
        for line_id in list(plans_manual_lock.keys()):
            slot = plans_manual_lock.get(line_id) or {}
            slot.pop(shift_type, None)
            if not slot:
                plans_manual_lock.pop(line_id, None)
    logging.info('plan manual lock cleared for shift_type=%s', shift_type)


def _shift_history_key(shift_date, shift_type):
    if hasattr(shift_date, 'isoformat'):
        d = shift_date.isoformat()
    else:
        d = str(shift_date)
    return f"{d}_{shift_type}"


def load_plans_file():
    """Загрузка plans.json: current + history. Старый формат (только линии) совместим."""
    global plans, plans_history, plans_manual_lock
    if not os.path.exists(PLAN_FILE):
        return
    try:
        with open(PLAN_FILE, 'r', encoding='utf-8') as f:
            loaded = json.load(f)
    except Exception as e:
        logging.error(f"Ошибка чтения {PLAN_FILE}: {e}")
        return

    with plans_lock:
        if isinstance(loaded, dict) and 'current' in loaded:
            cur = loaded.get('current') or {}
            for line in plans:
                if line in cur and isinstance(cur[line], dict):
                    plans[line]['day'] = int(cur[line].get('day', 0) or 0)
                    plans[line]['night'] = int(cur[line].get('night', 0) or 0)
            hist = loaded.get('history') or {}
            if isinstance(hist, dict):
                plans_history = hist
            ml = loaded.get('manual_lock') or {}
            if isinstance(ml, dict):
                plans_manual_lock = {
                    str(lid): {
                        st: str(ep)
                        for st, ep in (vals or {}).items()
                        if st in ('day', 'night') and ep
                    }
                    for lid, vals in ml.items()
                    if isinstance(vals, dict)
                }
        else:
            # Старый формат: {line: {day, night}}
            for line in plans:
                if line in loaded and isinstance(loaded[line], dict):
                    plans[line]['day'] = int(loaded[line].get('day', 0) or 0)
                    plans[line]['night'] = int(loaded[line].get('night', 0) or 0)
            if isinstance(loaded.get('history'), dict):
                plans_history = loaded['history']


def save_plans():
    """Сохранить current + history в plans.json (атомарно через tmp)."""
    with plans_lock:
        payload = {
            'current': {line: dict(vals) for line, vals in plans.items()},
            'history': plans_history,
            'manual_lock': {
                lid: dict(vals) for lid, vals in plans_manual_lock.items() if vals
            },
        }
    try:
        _atomic_write_json(PLAN_FILE, payload, indent=2)
    except Exception as e:
        logging.error(f"Ошибка записи {PLAN_FILE}: {e}")


def get_plan_for_shift(line_id, shift_type, shift_date=None):
    """
    План линии на смену: сначала history (зафиксированный), иначе current.
    shift_date: date | str YYYY-MM-DD | None (только current).
    """
    if shift_date is not None:
        key = _shift_history_key(shift_date, shift_type)
        with plans_lock:
            entry = plans_history.get(key)
            if entry and isinstance(entry.get('lines'), dict):
                line_rec = entry['lines'].get(line_id)
                if line_rec is not None:
                    return int(line_rec.get('plan', 0) or 0)
    with plans_lock:
        return int(plans.get(line_id, {}).get(shift_type, 0) or 0)


def snapshot_shift_to_history(shift_date, shift_type, facts=None, overwrite_plans=False):
    """
    Зафиксировать планы смены в history.
    facts: optional {line_id: fact_qty} — факт исполнения.
    overwrite_plans=False: если запись уже есть, plan не затираем (только fact обновляем).
    """
    key = _shift_history_key(shift_date, shift_type)
    d_str = shift_date.isoformat() if hasattr(shift_date, 'isoformat') else str(shift_date)
    facts = facts or {}
    with plans_lock:
        existing = plans_history.get(key) or {}
        lines_out = dict(existing.get('lines') or {})
        for line_id in plans:
            prev = lines_out.get(line_id) or {}
            if overwrite_plans or line_id not in lines_out:
                plan_val = int(plans.get(line_id, {}).get(shift_type, 0) or 0)
            else:
                plan_val = int(prev.get('plan', 0) or 0)
            fact_val = facts.get(line_id)
            if fact_val is None:
                fact_val = prev.get('fact')
            try:
                fact_val = int(fact_val) if fact_val is not None else None
            except (TypeError, ValueError):
                fact_val = None
            percent = None
            if fact_val is not None and plan_val > 0:
                percent = round(fact_val / plan_val * 100.0, 1)
            lines_out[line_id] = {
                'plan': plan_val,
                'fact': fact_val,
                'percent': percent,
            }
        plans_history[key] = {
            'date': d_str,
            'shift': shift_type,
            'saved_at': datetime.now().isoformat(timespec='seconds'),
            'lines': lines_out,
        }
    save_plans()
    logging.info(f"[PLAN_RESET] history зафиксирован: key={key}")
    return key


def reset_current_plans_for_shift_type(shift_type):
    """Обнулить current.day или current.night после окончания смены."""
    if shift_type not in ('day', 'night'):
        return
    with plans_lock:
        for line_id in plans:
            plans[line_id][shift_type] = 0
        line_ids = list(plans.keys())
    clear_manual_locks_for_shift_type(shift_type)
    save_plans()
    with plans_lock:
        snapshot = {line: dict(vals) for line, vals in plans.items()}
    try:
        socket_emit('plan_update', snapshot)
        logging.info(f"[PLAN_RESET] plan_update отправлен клиентам (обнулён {shift_type})")
    except Exception as e:
        logging.warning(f"[PLAN_RESET] plan_update emit failed: {e}")
    logging.info(
        f"[PLAN_RESET] current.{shift_type} → 0 | lines={line_ids}"
    )


def finalize_shift_plans(shift_date, shift_type, facts=None):
    """
    Конец смены: записать history (plan из current + fact) и обнулить current для этого типа.
    """
    snapshot_shift_to_history(shift_date, shift_type, facts=facts, overwrite_plans=False)
    reset_current_plans_for_shift_type(shift_type)


def sync_excel_production_plans(force_reload=False):
    """
    Подтянуть суточный план из Excel → plans[line].day для Ralot_L1 / KRO_* / BLS / TSD_PRODUCTION.
    """
    try:
        if force_reload:
            production_plan.reload_if_needed(force=True)
        last = excel_plan_fact_gtins()
        with plans_lock:
            skip = {
                lid for lid in production_plan.EXCEL_MANAGED_LINES
                if is_plan_manual_locked(lid, 'day')
            }
            state, changed = production_plan.sync_plans_day(
                plans,
                day=datetime.now().date(),
                last_gtin_by_line=last,
                gtin_names=gtin_names,
                line_display=LINE_DISPLAY_NAMES,
                skip_lines=skip,
                all_fact_gtins_by_line=excel_plan_fact_gtins_all(),
            )
            snapshot = {line: dict(vals) for line, vals in plans.items()}
        save_plans()
        if skip:
            logging.info(f'excel plan sync: skip manual-locked={sorted(skip)}')
        if changed:
            logging.info(f'excel plan sync: changed={changed} source={state.get("source")}')
            socket_emit('plan_update', snapshot)
        else:
            logging.debug(
                f'excel plan sync: no plan changes date={state.get("date")} '
                f'source={state.get("source")}'
            )
        if state is not None:
            state = dict(state)
            state['manual_locked'] = sorted(skip)
        socket_emit('excel_plan', state)
        return state
    except Exception as e:
        logging.error(f'excel plan sync: {e}', exc_info=True)
        return None


def excel_plan_sync_loop():
    interval = float(os.environ.get('PRODUCTION_PLAN_SYNC_SEC', '1800'))  # 30 мин
    interval = max(60.0, interval)
    logging.info(f'Excel production plan: sync каждые {interval:.0f} с')
    # первый проход сразу
    try:
        sync_excel_production_plans(force_reload=True)
    except Exception as e:
        logging.warning(f'excel plan first sync: {e}')
    while True:
        time.sleep(interval)
        try:
            sync_excel_production_plans(force_reload=False)
        except Exception as e:
            logging.error(f'excel plan sync loop: {e}', exc_info=True)


load_plans_file()

# Типы событий для % плана (как карточки в режиме агрегатов / Excel «Выпуск»).
PLAN_FACT_TYPES = ('L1', 'ТУ', 'ДЕЗАГРЕГАЦИЯ')


def _current_shift_type(now=None):
    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    return 'day' if 8 <= now.hour < 20 else 'night'


def plan_completion_moment(events, plan, start_dt, end_dt=None):
    """
    Момент, когда накопленный выпуск впервые достиг плана.
    events: iterable of dict с ts/qty (qty может быть отрицательным у ДЕЗАГРЕГАЦИЯ).
    """
    try:
        plan_n = int(plan or 0)
    except (TypeError, ValueError):
        plan_n = 0
    if plan_n <= 0 or not events:
        return None
    end_dt = end_dt or datetime.now()
    if getattr(start_dt, 'tzinfo', None) is not None:
        start_dt = start_dt.replace(tzinfo=None)
    if getattr(end_dt, 'tzinfo', None) is not None:
        end_dt = end_dt.replace(tzinfo=None)
    rows = []
    for ev in events:
        if not isinstance(ev, dict):
            continue
        ts = ev.get('ts')
        if isinstance(ts, str):
            ts = _iso_to_dt(ts)
        if not isinstance(ts, datetime):
            continue
        if ts.tzinfo is not None:
            ts = ts.replace(tzinfo=None)
        if ts < start_dt or ts > end_dt:
            continue
        try:
            qty = float(ev.get('qty') or 0)
        except (TypeError, ValueError):
            qty = 0.0
        rows.append((ts, qty))
    if not rows:
        return None
    rows.sort(key=lambda x: x[0])
    cum = 0.0
    for ts, qty in rows:
        cum += qty
        if cum >= plan_n:
            return ts
    return None


def _line_shift_plan_events_unlocked(line_id, shift_start, end_dt):
    """Как line_shift_plan_events, но вызывающий уже держит data_lock."""
    out = []
    types = production_data.get(line_id) or {}
    for etype in PLAN_FACT_TYPES:
        bucket = types.get(etype) or {}
        events = bucket.get('events') if isinstance(bucket, dict) else bucket
        for ev in events or []:
            ts = ev.get('ts')
            if not isinstance(ts, datetime):
                continue
            if ts.tzinfo is not None:
                ts = ts.replace(tzinfo=None)
            if shift_start <= ts <= end_dt:
                out.append({'ts': ts, 'qty': ev.get('qty') or 0, 'type': etype})
    return out


def line_shift_plan_events(line_id, shift_start, end_dt=None):
    """События L1/ТУ/ДЕЗАГРЕГАЦИЯ за смену (копия под data_lock)."""
    end_dt = end_dt or datetime.now()
    if getattr(shift_start, 'tzinfo', None) is not None:
        shift_start = shift_start.replace(tzinfo=None)
    if getattr(end_dt, 'tzinfo', None) is not None:
        end_dt = end_dt.replace(tzinfo=None)
    with data_lock:
        return _line_shift_plan_events_unlocked(line_id, shift_start, end_dt)


def line_shift_plan_fact(line_id, shift_start=None, end_dt=None, now=None):
    """Нетто-выпуск для % плана за текущую (или заданную) смену."""
    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    shift_start = shift_start or _current_shift_start(now)
    end_dt = end_dt or now
    return int(sum(ev.get('qty') or 0 for ev in line_shift_plan_events(line_id, shift_start, end_dt)))


def is_line_plan_complete(line_id, now=None, plan=None, fact=None):
    """
    План смены выполнен (≥100%): дальше тишина — не простой и не MAX-алерт.
    План=0 → False (нечего «выполнять»).
    """
    if not line_id:
        return False
    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    if plan is None:
        plan = get_plan_for_shift(line_id, _current_shift_type(now))
    try:
        plan_n = int(plan or 0)
    except (TypeError, ValueError):
        plan_n = 0
    if plan_n <= 0:
        return False
    if fact is None:
        fact = line_shift_plan_fact(line_id, now=now)
    try:
        fact_n = int(fact or 0)
    except (TypeError, ValueError):
        fact_n = 0
    return fact_n >= plan_n


def line_plan_done_at(line_id, now=None, plan=None):
    """Когда впервые достигли плана в текущей смене (или None)."""
    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    if plan is None:
        plan = get_plan_for_shift(line_id, _current_shift_type(now))
    try:
        plan_n = int(plan or 0)
    except (TypeError, ValueError):
        plan_n = 0
    if plan_n <= 0:
        return None
    shift_start = _current_shift_start(now)
    return plan_completion_moment(
        line_shift_plan_events(line_id, shift_start, now),
        plan_n, shift_start, now,
    )


# ============================
# ПРОСТОИ / ПРИЧИНЫ (мастер смены)
# ============================
DOWNTIME_REASONS = ('Переход', 'Обед', 'Ремонт', 'Проблема')
# Для этих причин мастер вводит текст — он попадает в отчёт смены
DOWNTIME_REASONS_NEED_COMMENT = frozenset({'Ремонт', 'Проблема'})
DOWNTIME_COMMENT_MAX_LEN = 500
DEFAULT_DOWNTIME_REASON = 'причина не обозначена'
DOWNTIME_FILE = 'downtime.json'
# Журнал мастера: паузы короче порога не показываются (простои — всегда)
MASTER_JOURNAL_PAUSE_MIN = float(os.environ.get('MASTER_JOURNAL_PAUSE_MIN', '5'))

# closed: list of {id, line_id, reason, started_at, ended_at, acknowledged}
# open:   dict line_id -> {id, line_id, reason|None, started_at, acknowledged}
# RLock: вложенный захват (snapshot внутри close) больше не превращается в вечный deadlock
downtime_lock = threading.RLock()
downtime_closed = []
downtime_open = {}


def _dt_to_iso(dt):
    if dt is None:
        return None
    if isinstance(dt, str):
        return dt
    if isinstance(dt, datetime):
        if dt.tzinfo is not None:
            dt = dt.replace(tzinfo=None)
        return dt.isoformat(timespec='seconds')
    return str(dt)


def _iso_to_dt(s):
    """Парсер datetime → naive local. UTC (Z) корректно переводим в локальное время сервера."""
    if s is None:
        return None
    if isinstance(s, datetime):
        ts = s
    else:
        try:
            raw = str(s).strip()
            # fromisoformat: 'Z' → UTC
            if raw.endswith('Z'):
                raw = raw[:-1] + '+00:00'
            ts = datetime.fromisoformat(raw)
        except Exception:
            return None
    if ts.tzinfo is not None:
        # В локальный timezone сервера, затем naive
        try:
            ts = ts.astimezone().replace(tzinfo=None)
        except Exception:
            ts = ts.replace(tzinfo=None)
    return ts


def _started_key(started_at):
    """Нормализация started_at для дедупа (до секунды)."""
    st = _iso_to_dt(started_at)
    if st is None:
        return str(started_at or '')
    return st.replace(microsecond=0).isoformat(timespec='seconds')


def _dedupe_closed_downtimes(records):
    """
    Один простой = один started_at на линии.
    Конец = первое событие после начала → берём минимальный ended_at.
    """
    best = {}
    for rec in records or []:
        line_id = rec.get('line_id')
        if not line_id:
            continue
        key = (line_id, _started_key(rec.get('started_at')))
        en = _iso_to_dt(rec.get('ended_at'))
        prev = best.get(key)
        if prev is None:
            best[key] = dict(rec)
            continue
        prev_en = _iso_to_dt(prev.get('ended_at'))
        if en is not None and (prev_en is None or en < prev_en):
            merged = dict(rec)
            # сохраняем ACK/причину, если были у любого экземпляра
            if prev.get('acknowledged') and not merged.get('acknowledged'):
                merged['acknowledged'] = True
                merged['reason'] = prev.get('reason') or merged.get('reason')
                if prev.get('comment') and not merged.get('comment'):
                    merged['comment'] = prev.get('comment')
            if prev.get('reason') and (
                not merged.get('reason') or merged.get('reason') == DEFAULT_DOWNTIME_REASON
            ):
                merged['reason'] = prev['reason']
            if prev.get('comment') and not merged.get('comment'):
                merged['comment'] = prev['comment']
            best[key] = merged
        else:
            if rec.get('acknowledged') and not prev.get('acknowledged'):
                prev['acknowledged'] = True
                prev['reason'] = rec.get('reason') or prev.get('reason')
                if rec.get('comment') and not prev.get('comment'):
                    prev['comment'] = rec.get('comment')
            if rec.get('reason') and (
                not prev.get('reason') or prev.get('reason') == DEFAULT_DOWNTIME_REASON
            ):
                prev['reason'] = rec['reason']
            if rec.get('comment') and not prev.get('comment'):
                prev['comment'] = rec['comment']
    out = []
    for rec in best.values():
        st = _iso_to_dt(rec.get('started_at'))
        en = _iso_to_dt(rec.get('ended_at'))
        if st and en and en > st:
            rec['duration_min'] = round((en - st).total_seconds() / 60.0, 1)
        out.append(rec)
    return out


def load_downtime():
    global downtime_closed, downtime_open
    if not os.path.exists(DOWNTIME_FILE):
        downtime_closed = []
        downtime_open = {}
        return
    try:
        with open(DOWNTIME_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        downtime_closed = data.get('closed') or []
        downtime_open = data.get('open') or {}
        # prune older than 14 days
        cutoff = datetime.now() - timedelta(days=14)
        pruned = []
        for rec in downtime_closed:
            end = _iso_to_dt(rec.get('ended_at')) or _iso_to_dt(rec.get('started_at'))
            if end is None or end >= cutoff:
                pruned.append(rec)
        before = len(pruned)
        downtime_closed = _dedupe_closed_downtimes(pruned)
        if len(downtime_closed) != before:
            logging.info(
                f"Простои: дедуп closed {before} → {len(downtime_closed)}"
            )
            save_downtime()
    except Exception as e:
        logging.error(f"Ошибка загрузки {DOWNTIME_FILE}: {e}")
        downtime_closed = []
        downtime_open = {}


def save_downtime():
    try:
        with open(DOWNTIME_FILE, 'w', encoding='utf-8') as f:
            json.dump({'closed': downtime_closed, 'open': downtime_open}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logging.error(f"Ошибка сохранения {DOWNTIME_FILE}: {e}")


def _downtime_pending_locked(now=None):
    """Закрытые простои текущей смены без причины мастера (для «липких» карточек)."""
    now = now or datetime.now()
    shift_start = _current_shift_start(now)
    out = []
    for rec in _dedupe_closed_downtimes(downtime_closed):
        if rec.get('kind') == 'pause' or rec.get('acknowledged'):
            continue
        if skip_report_silence_line(rec.get('line_id')):
            continue
        en = _iso_to_dt(rec.get('ended_at'))
        if en is None or en <= shift_start or en > now + timedelta(minutes=1):
            continue
        out.append(rec)
    out.sort(key=lambda r: (r.get('line_id') or '', r.get('started_at') or ''))
    return out


def _downtime_snapshot_locked():
    """Снимок без взятия лока — вызывать только под downtime_lock."""
    since = datetime.now() - timedelta(hours=24)
    return {
        'open': {k: dict(v) for k, v in downtime_open.items()},
        'closed': [
            r for r in downtime_closed
            if (_iso_to_dt(r.get('ended_at')) or _iso_to_dt(r.get('started_at')) or since) >= since
        ],
        'pending': _downtime_pending_locked(),
        'reasons': list(DOWNTIME_REASONS),
        'comment_reasons': sorted(DOWNTIME_REASONS_NEED_COMMENT),
    }


def downtime_snapshot():
    with downtime_lock:
        return _downtime_snapshot_locked()


def _line_already_closed_silence(line_id, started):
    """Этот started_at уже закрыт первым событием — не открывать снова."""
    if not line_id or started is None:
        return False
    key = _started_key(started)
    for rec in downtime_closed:
        if rec.get('line_id') != line_id:
            continue
        if _started_key(rec.get('started_at')) == key:
            return True
    return False


def _max_notify_open(rec):
    """
    Раньше слал «Задержка N минут» в момент открытия простоя на фронте.
    Карточка открывает простой уже на жёлтом пороге (DM warn = 3 мин), а N
    подставлялся как MAX_ALERT_FIRST_MIN (10) — в чат уходило «10 минут»
    при паузе 3 мин на карточке и в GUI.
    Первый алерт теперь только из max_reminder_loop, когда тишина ≥ FIRST.
    """
    return


def _max_notify_resume(rec):
    if not rec:
        return
    line_id = rec.get('line_id')
    # Resume после ACK причины / DM-тишины — даже если entity на open был не DM
    try:
        from max_bot.alerts import notify_line_resume_if_alerted
        notify_line_resume_if_alerted(
            line_id,
            LINE_DISPLAY_NAMES.get(line_id, line_id),
            ended_at=rec.get('ended_at'),
        )
    except Exception as e:
        logging.error(f'MAX alert resume: {e}')


def _max_notify_reason(line_id, reason):
    """Однократный алерт с причиной; дальше только возобновление."""
    if not line_id or not reason:
        return
    if line_id in MAX_ALERT_EXCLUDE_LINES:
        return
    try:
        from max_bot.alerts import notify_line_reason
        notify_line_reason(
            line_id=line_id,
            display_name=LINE_DISPLAY_NAMES.get(line_id, line_id),
            reason=reason,
        )
    except Exception as e:
        logging.error(f'MAX alert reason: {e}')


def _open_map_for_max_reminders(open_map):
    """Оставить в карте только простои, по которым разрешены MAX-алерты (DM, не L2)."""
    if not open_map:
        return {}
    return {
        lid: rec for lid, rec in open_map.items()
        if isinstance(rec, dict) and max_alert_allowed(lid, rec=rec)
    }


def _max_alert_first_min():
    """Порог тишины для попадания линии в MAX-опрос (из max_bot/config)."""
    try:
        from max_bot.config import MAX_ALERT_FIRST_MIN
        return max(1, int(MAX_ALERT_FIRST_MIN))
    except Exception:
        return 10


def _last_positive_event_ts(events):
    """Последний ts события с qty > 0 (naive)."""
    last_ts = None
    for ev in events or []:
        if (ev.get('qty') or 0) <= 0:
            continue
        ts = ev.get('ts')
        if not isinstance(ts, datetime):
            continue
        if ts.tzinfo is not None:
            ts = ts.replace(tzinfo=None)
        if last_ts is None or ts > last_ts:
            last_ts = ts
    return last_ts


def build_max_dm_silence_map(now=None):
    """
    Тишина ≥ MAX_ALERT_FIRST_MIN (по умолчанию 10 мин) по той же шкале, что
    карточка/GUI: последнее положительное событие L1 / ТУ / DM.
    В смене должны быть DM (иначе не алертим «вчерашний» код).
    ACK причины и план ≥100% — без «Задержка…».
    """
    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    discard_downtimes_for_completed_plans(now)
    shift_start = _current_shift_start(now)
    threshold_sec = float(_max_alert_first_min()) * 60.0
    acked_lines = set()
    try:
        from max_bot.alerts import is_reason_acked
    except Exception:
        is_reason_acked = None
    with downtime_lock:
        for lid, rec in downtime_open.items():
            if rec and rec.get('acknowledged'):
                acked_lines.add(lid)
    out = {}
    with data_lock:
        line_ids = list(production_data.keys())
    for line_id in line_ids:
        if line_id in MAX_ALERT_EXCLUDE_LINES:
            continue
        if line_id in acked_lines:
            continue
        if is_reason_acked and is_reason_acked(line_id):
            continue
        if is_line_plan_complete(line_id, now):
            continue
        with data_lock:
            types = production_data.get(line_id) or {}
            last_dm = _last_positive_event_ts((types.get('DM') or {}).get('events'))
            last_ts = None
            for et in ('L1', 'ТУ', 'DM'):
                t = _last_positive_event_ts((types.get(et) or {}).get('events'))
                if t is not None and (last_ts is None or t > last_ts):
                    last_ts = t
        # Без DM в этой смене — не орём (как раньше).
        if last_dm is None or last_dm < shift_start:
            continue
        if last_ts is None or last_ts < shift_start:
            continue
        age = (now - last_ts).total_seconds()
        if age >= threshold_sec:
            out[line_id] = {
                'started_at': _dt_to_iso(last_ts),
                'entity_type': 'DM',
            }
    return out


def _max_on_dm_activity(line_id, event_time=None):
    """Первый положительный DM после тишины — снять MAX-алерт (если был)."""
    if not max_alert_allowed(line_id, entity_type='DM'):
        return
    try:
        from max_bot.alerts import notify_line_resume_if_alerted
        notify_line_resume_if_alerted(
            line_id,
            LINE_DISPLAY_NAMES.get(line_id, line_id),
            ended_at=event_time or datetime.now(),
        )
    except Exception as e:
        logging.error(f'MAX DM resume {line_id}: {e}')


def max_reminder_loop():
    """
    Каждые ~MAX_REMINDER_POLL_SEC: тишина DM ≥ MAX_ALERT_FIRST_MIN → алерты MAX.
    Отметки N: FIRST → SECOND → +LATE (по умолчанию 10 → 20 → 50 → 80 …).
    """
    interval = float(os.environ.get('MAX_REMINDER_POLL_SEC', '30'))
    while True:
        try:
            from max_bot import config as max_cfg
            if max_cfg.is_enabled():
                from max_bot.alerts import check_open_reminders
                # Тишина как у GUI/карточки (последнее L1/ТУ/DM), порог FIRST мин
                open_map = build_max_dm_silence_map()
                if open_map:
                    n = check_open_reminders(open_map, display_names=LINE_DISPLAY_NAMES)
                    if n:
                        logging.info(f'MAX DM reminders/open sent: {n}')
        except Exception as e:
            logging.error(f'MAX reminder loop: {e}', exc_info=True)
        time.sleep(max(10.0, interval))


def max_shift_alert_reset_loop():
    """
    В 08:00 и 20:00 — закрыть открытые простои + сброс MAX-алертов (новая смена).
    Работает всегда (простоев), MAX-сброс — только если бот включён.
    """
    while True:
        try:
            now = datetime.now()
            today_8 = now.replace(hour=8, minute=0, second=0, microsecond=0)
            today_20 = now.replace(hour=20, minute=0, second=0, microsecond=0)
            if now < today_8:
                next_time = today_8
                label = '08:00'
            elif now < today_20:
                next_time = today_20
                label = '20:00'
            else:
                next_time = today_8 + timedelta(days=1)
                label = '08:00'
            wait = max(1.0, (next_time - now).total_seconds())
            logging.info(
                f'Shift boundary: next at {next_time.strftime("%Y-%m-%d %H:%M")} ({label}) '
                f'in {wait / 60:.1f} мин'
            )
            time.sleep(wait)
            boundary = datetime.now().replace(second=0, microsecond=0)
            # 1) закрыть открытые простои на границе (без «возобновление» в MAX)
            try:
                close_open_downtimes_at_shift_boundary(boundary)
            except Exception as e:
                logging.error(f'Shift boundary downtime close: {e}', exc_info=True)
            # Мастер: «ждут причины» прошлой смены должны пропасть, даже если открытых не было
            try:
                if connected_clients:
                    socket_emit('downtime_update', downtime_snapshot())
            except Exception as e:
                logging.error(f'Shift boundary downtime emit: {e}', exc_info=True)
            # 2) сброс дедупа MAX — только если включён
            try:
                from max_bot import config as max_cfg
                if max_cfg.is_enabled():
                    from max_bot.alerts import reset_all_alerts
                    reset_all_alerts(reason=f'shift_boundary_{label}')
            except Exception as e:
                logging.error(f'MAX alert reset: {e}', exc_info=True)
            # чуть продвинуться, чтобы не сработать дважды в ту же секунду
            time.sleep(2.0)
        except Exception as e:
            logging.error(f'Shift boundary loop: {e}', exc_info=True)
            time.sleep(60)


def _current_shift_start(now=None):
    now = now or datetime.now()
    if 8 <= now.hour < 20:
        return now.replace(hour=8, minute=0, second=0, microsecond=0)
    start = now.replace(hour=20, minute=0, second=0, microsecond=0)
    if now.hour < 8:
        start = start - timedelta(days=1)
    return start


def downtime_open_or_update(line_id, started_at=None, reason=None, acknowledge=False,
                            badge_min=None, entity_type=None, comment=None):
    """Открыть простой или обновить причину. started_at — момент последней активности (начало паузы)."""
    if not line_id:
        return None
    # План выполнен — остановка после выпуска не простой
    if is_line_plan_complete(line_id) and not acknowledge:
        downtime_discard(line_id, reason='plan_done')
        logging.debug(f"Простой {line_id}: open проигнорирован (план ≥100%)")
        return None
    now = datetime.now()
    raw_started = _iso_to_dt(started_at) or now
    # Не раньше начала текущей смены (иначе в отчёте всё «с 08:00»)
    shift_start = _current_shift_start(now)
    from_prev_shift = raw_started < shift_start
    started = shift_start if from_prev_shift else raw_started
    # Не из будущего
    if started > now:
        started = now
    try:
        badge = float(badge_min) if badge_min is not None else silence_bad_min(entity_type)
    except (TypeError, ValueError):
        badge = silence_bad_min(entity_type)
    if badge < 1:
        badge = silence_bad_min(entity_type)
    comment_n = _normalize_downtime_comment(reason, comment) if reason else None

    # Тишина тянется с прошлой смены: не открываем сразу в 08:00/20:00 —
    # ждём порог плашки уже в новой смене (иначе MAX снова орёт «простой с 20:00»).
    if from_prev_shift and (now - shift_start).total_seconds() < badge * 60:
        logging.debug(
            f"Простой {line_id}: отложен до {badge:.0f} мин новой смены "
            f"(тишина с {_dt_to_iso(raw_started)})"
        )
        return None

    is_new = False
    ignored = False
    with downtime_lock:
        existing = downtime_open.get(line_id)
        if existing:
            if reason:
                existing['reason'] = reason
                existing['comment'] = comment_n
            if acknowledge:
                existing['acknowledged'] = True
                existing['acked_at'] = _dt_to_iso(now)
            if badge_min is not None or entity_type:
                # не завышаем порог: берём более строгий (меньший) из уже записанного и нового
                try:
                    prev_b = float(existing.get('badge_min') or badge)
                except (TypeError, ValueError):
                    prev_b = badge
                existing['badge_min'] = min(prev_b, badge)
            save_downtime()
            rec = dict(existing)
        else:
            # Уже закрывали этот же интервал тишины (первое событие) — не плодим дубли в Excel
            if _line_already_closed_silence(line_id, started):
                logging.debug(
                    f"Простой {line_id}: повторный open с started_at={_dt_to_iso(started)} проигнорирован"
                )
                ignored = True
                rec = None
            else:
                is_new = True
                rec = {
                    'id': f"{line_id}_{int(started.timestamp())}_{int(now.timestamp())}",
                    'line_id': line_id,
                    'reason': reason,
                    'comment': comment_n,
                    'started_at': _dt_to_iso(started),
                    'acknowledged': bool(acknowledge and reason),
                    'acked_at': _dt_to_iso(now) if (acknowledge and reason) else None,
                    'badge_min': badge,
                    'entity_type': entity_type,
                }
                if acknowledge and reason:
                    rec['acknowledged'] = True
                downtime_open[line_id] = rec
                save_downtime()
                rec = dict(rec)
        payload = None if ignored else _downtime_snapshot_locked()
    if ignored:
        return None
    if connected_clients and payload is not None:
        socket_emit('downtime_update', payload)
    if is_new and not acknowledge:
        _max_notify_open(rec)
    return rec


def downtime_discard(line_id, reason='plan_done'):
    """
    Снять открытый простой без записи в Excel и без «возобновление» в MAX.
    Используется когда план смены уже выполнен.
    """
    if not line_id:
        return None
    existing = None
    payload = None
    with downtime_lock:
        existing = downtime_open.pop(line_id, None)
        if existing:
            save_downtime()
            payload = _downtime_snapshot_locked()
    if payload is None:
        # всё равно сбросить дедуп MAX (если алертили до достижения плана)
        try:
            from max_bot.alerts import forget_open
            forget_open(line_id)
        except Exception:
            pass
        return None
    if connected_clients:
        socket_emit('downtime_update', payload)
    try:
        from max_bot.alerts import forget_open
        forget_open(line_id)
    except Exception as e:
        logging.debug(f'MAX forget after plan_done {line_id}: {e}')
    logging.info(f'Простой {line_id}: сброшен ({reason}), без записи в отчёт')
    return existing


def discard_downtimes_for_completed_plans(now=None):
    """Закрыть (discard) все открытые простои на линиях с выполненным планом."""
    now = now or datetime.now()
    with downtime_lock:
        lids = list(downtime_open.keys())
    n = 0
    for lid in lids:
        try:
            if is_line_plan_complete(lid, now):
                if downtime_discard(lid, reason='plan_done'):
                    n += 1
        except Exception as e:
            logging.error(f'discard plan_done {lid}: {e}')
    return n


def _downtime_ack_closed(line_id, started_at, reason, comment_n):
    """Проставить причину в closed-запись с тем же started_at. → (rec, snapshot) | (None, None)."""
    key = _started_key(started_at)
    with downtime_lock:
        for c in downtime_closed:
            if c.get('line_id') != line_id:
                continue
            if _started_key(c.get('started_at')) != key:
                continue
            c['reason'] = reason
            c['comment'] = comment_n
            c['acknowledged'] = True
            c['acked_at'] = _dt_to_iso(datetime.now())
            save_downtime()
            return dict(c), _downtime_snapshot_locked()
    return None, None


def _downtime_ack_pause(line_id, started_at, ended_at, reason, comment_n):
    """
    Причина для паузы из журнала (короткая тишина без записи простоя).
    Пишем closed-запись kind='pause' — отчёт подхватит её по пересечению интервалов.
    """
    st = _iso_to_dt(started_at)
    en = _iso_to_dt(ended_at)
    now = datetime.now()
    if st is None or en is None or en <= st:
        return None, None, 'Некорректный интервал паузы'
    if (en - st) > timedelta(hours=12) or st < now - timedelta(days=2) or st > now:
        return None, None, 'Интервал паузы вне допустимого окна'
    rec = {
        'id': f'pause-{line_id}-{_started_key(st)}',
        'line_id': line_id,
        'kind': 'pause',
        'reason': reason,
        'comment': comment_n,
        'started_at': _dt_to_iso(st),
        'ended_at': _dt_to_iso(en),
        'acknowledged': True,
        'acked_at': _dt_to_iso(now),
        'duration_min': round((en - st).total_seconds() / 60.0, 1),
    }
    with downtime_lock:
        # Два мастера одновременно: второй не плодит дубль, а правит ту же запись
        existing, payload = _downtime_ack_closed(line_id, started_at, reason, comment_n)
        if existing is not None:
            return existing, payload, None
        downtime_closed.append(rec)
        save_downtime()
        payload = _downtime_snapshot_locked()
    return dict(rec), payload, None


def downtime_ack(line_id, reason, started_at=None, badge_min=None, entity_type=None,
                 comment=None, ended_at=None):
    """
    Мастер отметил причину — стоп напоминаний MAX, один алерт с причиной, потом только resume.
    ended_at — только из журнала смены: интервал паузы, для которой ещё нет записи.
    """
    reason = (reason or '').strip()
    if reason not in DOWNTIME_REASONS:
        return None, f'Недопустимая причина: {reason}'
    comment_n = _normalize_downtime_comment(reason, comment)
    if reason in DOWNTIME_REASONS_NEED_COMMENT and not comment_n:
        return None, f'Для «{reason}» укажите комментарий'
    notify_label = reason if not comment_n else f'{reason}: {comment_n}'

    # Уже закрытый простой / размеченная пауза с этим started_at — правим его,
    # не трогая текущий открытый простой линии (мастер размечает прошлое).
    if started_at:
        rec, payload = _downtime_ack_closed(line_id, started_at, reason, comment_n)
        if rec is None and ended_at:
            rec, payload, err = _downtime_ack_pause(
                line_id, started_at, ended_at, reason, comment_n,
            )
            if err:
                return None, err
        if rec is not None:
            if connected_clients and payload:
                socket_emit('downtime_update', payload)
            logging.info(f"🛑 Причина (журнал) → {line_id} {rec.get('started_at')}: {notify_label}")
            if rec.get('kind') != 'pause':
                _max_notify_reason(line_id, notify_label)
            return rec, None

    rec = downtime_open_or_update(
        line_id, started_at=started_at, reason=reason, acknowledge=True,
        badge_min=badge_min, entity_type=entity_type, comment=comment_n,
    )
    if rec is None and started_at:
        # Интервал закрылся первым событием между проверкой и open_or_update
        rec, payload = _downtime_ack_closed(line_id, started_at, reason, comment_n)
        if rec is not None:
            if connected_clients and payload:
                socket_emit('downtime_update', payload)
            logging.info(f"🛑 Простой ACK (closed) → {line_id}: {notify_label}")
            _max_notify_reason(line_id, notify_label)
            return rec, None
    logging.info(
        f"🛑 Простой ACK → {line_id}: {notify_label} "
        f"(start={rec.get('started_at') if rec else None})"
    )
    if rec is not None:
        _max_notify_reason(line_id, notify_label)
    return rec, None


def downtime_close(line_id, ended_at=None, default_reason=DEFAULT_DOWNTIME_REASON,
                   notify_max=True):
    """Закрыть открытый простой (при возобновлении событий). Конец = первое событие после начала."""
    if not line_id:
        return None
    end = _iso_to_dt(ended_at) or datetime.now()
    closed = None
    too_short = None
    payload = None
    min_dur_used = None
    with downtime_lock:
        existing = downtime_open.pop(line_id, None)
        if existing:
            reason = existing.get('reason') or default_reason
            comment = existing.get('comment')
            if reason not in DOWNTIME_REASONS_NEED_COMMENT:
                comment = None
            st = _iso_to_dt(existing.get('started_at'))
            # В отчёт — только достигшие плашки «задержка или остановка» (bad)
            try:
                min_dur = float(existing.get('badge_min') or STATUS_SILENCE_BAD_MIN)
            except (TypeError, ValueError):
                min_dur = float(STATUS_SILENCE_BAD_MIN)
            min_dur_used = min_dur
            if st and end > st:
                dur = (end - st).total_seconds() / 60.0
            else:
                dur = 0.0
            if dur < min_dur:
                save_downtime()
                too_short = dur
            else:
                # Не писать второй closed с тем же started_at
                if _line_already_closed_silence(line_id, st):
                    save_downtime()
                    too_short = -1  # маркер «уже записан»
                else:
                    closed = {
                        'id': existing.get('id'),
                        'line_id': line_id,
                        'reason': reason,
                        'comment': comment,
                        'started_at': _dt_to_iso(st) if st else existing.get('started_at'),
                        'ended_at': _dt_to_iso(end),
                        'acknowledged': bool(existing.get('acknowledged')),
                        'duration_min': round(dur, 1),
                        'badge_min': min_dur,
                        'entity_type': existing.get('entity_type'),
                    }
                    downtime_closed.append(closed)
                    save_downtime()
            # Снимок под локом, emit — уже без лока (иначе self-deadlock на downtime_lock)
            payload = _downtime_snapshot_locked()

    if payload is None:
        return None

    if too_short is not None:
        if too_short < 0:
            logging.debug(f"Простой {line_id}: close без записи (уже есть closed с этим started_at)")
        else:
            logging.debug(
                f"Простой {line_id} сброшен без записи "
                f"(длительность {too_short:.1f} < {min_dur_used} мин плашки)"
            )
    else:
        logging.info(
            f"✅ Простой закрыт → {line_id}: {closed.get('reason')} "
            f"{closed.get('started_at')} – {closed.get('ended_at')} ({closed.get('duration_min')} мин)"
        )
    if connected_clients:
        socket_emit('downtime_update', payload)
    if closed and notify_max:
        _max_notify_resume(closed)
    return closed


def close_open_downtimes_at_shift_boundary(boundary_dt=None):
    """
    Конец смены 08:00 / 20:00: закрыть все открытые простои на границе.
    Без MAX «возобновление» — смена закончилась, линии не «поехали».
    """
    boundary = boundary_dt or datetime.now()
    if getattr(boundary, 'tzinfo', None) is not None:
        boundary = boundary.replace(tzinfo=None)
    with downtime_lock:
        line_ids = list(downtime_open.keys())
    if not line_ids:
        logging.info(f'Простои: на границе смены ({boundary.strftime("%H:%M")}) открытых нет')
        return 0
    n = 0
    for line_id in line_ids:
        try:
            # notify_max=False — не слать «работа возобновлена»
            if downtime_close(line_id, ended_at=boundary, notify_max=False) is not None:
                n += 1
            else:
                # даже short/pop без closed — простой снят с open
                n += 1
        except Exception as e:
            logging.error(f'Простой boundary close {line_id}: {e}', exc_info=True)
    logging.info(
        f'Простои: граница смены {boundary.strftime("%Y-%m-%d %H:%M")} — закрыто открытых: {n}'
    )
    return n


def reconcile_stale_open_downtimes():
    """
    Открытый простой с started_at до начала текущей смены — закрыть на shift_start.
    Нужен после рестарта / если граница 08:00/20:00 была пропущена.
    """
    now = datetime.now()
    shift_start = _current_shift_start(now)
    stale = []
    with downtime_lock:
        for line_id, rec in downtime_open.items():
            st = _iso_to_dt(rec.get('started_at'))
            if st is not None and st < shift_start:
                stale.append(line_id)
    if not stale:
        return 0
    n = 0
    for line_id in stale:
        try:
            downtime_close(line_id, ended_at=shift_start, notify_max=False)
            n += 1
        except Exception as e:
            logging.error(f'Простой stale close {line_id}: {e}', exc_info=True)
    if n:
        logging.info(
            f'Простои: reconcile — закрыто {n} устаревших (started до {shift_start.strftime("%Y-%m-%d %H:%M")})'
        )
    return n


def get_downtime_for_shift(start_dt, end_dt, shift_type=None):
    """
    Простои для листа Excel «Простои»:
    - только интервалы ≥ порога плашки (bad);
    - длительность = от начала тишины до первого события на линии
      (дедуп по started_at → минимальный ended_at);
    - пересечение с окном смены, минус регламентированный перерыв.
    """
    now = datetime.now()
    effective_end = end_dt if end_dt <= now else now
    if shift_type is None:
        shift_type = 'day' if start_dt.hour == 8 else 'night'
    breaks = get_shift_break_intervals(start_dt, end_dt, shift_type)
    shift_finished = now >= end_dt + timedelta(minutes=15)
    result = []
    with downtime_lock:
        closed_unique = _dedupe_closed_downtimes(downtime_closed)
        for rec in closed_unique:
            st = _iso_to_dt(rec.get('started_at'))
            en = _iso_to_dt(rec.get('ended_at')) or st
            if st is None:
                continue
            if en <= start_dt or st >= end_dt:
                continue
            try:
                badge = float(rec.get('badge_min') or STATUS_SILENCE_BAD_MIN)
            except (TypeError, ValueError):
                badge = float(STATUS_SILENCE_BAD_MIN)
            # Полная длительность до первого события (не урезанная сменой) — фильтр плашки
            full_dur = (en - st).total_seconds() / 60.0 if en > st else 0.0
            if full_dur < badge and rec.get('kind') != 'pause':
                continue
            clip_st = max(st, start_dt)
            clip_en = min(en, end_dt)
            if clip_en <= clip_st:
                continue
            dur = minutes_excluding_breaks(clip_st, clip_en, breaks)
            if dur <= 0:
                continue
            row = dict(rec)
            row['started_at'] = _dt_to_iso(clip_st)
            row['ended_at'] = _dt_to_iso(clip_en)
            row['duration_min'] = dur
            row['raw_started_at'] = rec.get('started_at')
            row['raw_ended_at'] = rec.get('ended_at')
            result.append(row)

        for line_id, rec in downtime_open.items():
            st = _iso_to_dt(rec.get('started_at'))
            if st is None or st >= end_dt:
                continue
            if st < start_dt - timedelta(hours=12):
                continue
            try:
                badge = float(rec.get('badge_min') or STATUS_SILENCE_BAD_MIN)
            except (TypeError, ValueError):
                badge = float(STATUS_SILENCE_BAD_MIN)
            if shift_finished:
                # Смена закончилась без первого события — в отчёт только если уже была плашка
                if st < start_dt:
                    continue
                clip_st = st
                clip_en = end_dt
                still_open = False
            else:
                if st >= effective_end:
                    continue
                clip_st = max(st, start_dt)
                clip_en = effective_end
                still_open = True
            if clip_en <= clip_st:
                continue
            full_dur = (clip_en - st).total_seconds() / 60.0
            if full_dur < badge:
                continue
            dur = minutes_excluding_breaks(clip_st, clip_en, breaks)
            if dur <= 0:
                continue
            row = dict(rec)
            row['started_at'] = _dt_to_iso(clip_st)
            row['ended_at'] = None if still_open else _dt_to_iso(clip_en)
            row['duration_min'] = dur
            row['raw_started_at'] = rec.get('started_at')
            row['open'] = still_open
            result.append(row)
    result.sort(key=lambda r: r.get('started_at') or '')
    return result


load_downtime()

# RLock: build_lines_live / MAX могут вызывать is_line_plan_complete уже под локом.
# Обычный Lock здесь = вечный deadlock (hang_traceback 24.08: 100+ потоков на data_lock).
data_lock = threading.RLock()
production_data = defaultdict(lambda: defaultdict(lambda: {
    'events': deque(maxlen=int(os.environ.get('EVENTS_DEQUE_MAXLEN', '50000'))),
    'items_per_event': None,
}))
last_update = {}
connected_clients = set()

# Известные агрегаты для SQL-сверки дезагрегаций: line_id -> {agg_id -> meta}
known_aggregates = defaultdict(dict)
# Идемпотентность: уже обработанные отмены (Rabbit + SQL не дают двойной минус)
cancelled_ids = set()
# Состояния агрегата, означающие отмену / разборку
CANCEL_STATES = {
    'CANCEL', 'CANCELLED', 'CANCELED', 'DELETE', 'DELETED', 'UNPACK', 'UNPACKED',
    'DISAGGREGATION', 'DISAGGREGATED', 'REMOVED', 'VOID', 'ОТМЕНЕНО', 'УДАЛЕНО',
    'РАЗОБРАНО', 'РАЗОБРАН'
}
RECONCILE_INTERVAL = 45 * 60  # секунд между сверками (~45 мин)

# ============================
# ОЧИСТКА СТАРЫХ ОТЧЕТОВ
# ============================
def cleanup_old_reports(days=30):
    now = time.time()
    cutoff = now - days * 86400
    deleted = 0
    for fname in os.listdir(REPORTS_DIR):
        filepath = os.path.join(REPORTS_DIR, fname)
        if os.path.isfile(filepath) and fname.endswith('.xlsx'):
            if os.path.getmtime(filepath) < cutoff:
                try:
                    os.remove(filepath)
                    deleted += 1
                except Exception as e:
                    logging.error(f"Ошибка удаления {filepath}: {e}")
    if deleted:
        logging.info(f"Удалено {deleted} старых отчетов (старше {days} дней)")

def cleanup_loop():
    while True:
        cleanup_old_reports(30)
        time.sleep(86400)

# ============================
# ФУНКЦИИ РАБОТЫ С БАЗОЙ ДАННЫХ
# ============================
# Пул соединений Postgres (лимит одновременных клиентов к БД)
_db_pool = None
_db_pool_lock = threading.Lock()
_poller_active_conn = None
_poller_active_conn_lock = threading.Lock()


def _db_connect_kwargs():
    timeout = int(os.environ.get('DB_CONNECT_TIMEOUT', '8'))
    stmt_ms = int(os.environ.get('DB_STATEMENT_TIMEOUT_MS', '30000'))
    lock_ms = int(os.environ.get('DB_LOCK_TIMEOUT_MS', '10000'))
    return dict(
        host=DB_HOST,
        port=DB_PORT,
        dbname=DB_NAME,
        user=DB_USER,
        password=DB_PASSWORD,
        connect_timeout=timeout,
        keepalives=1,
        keepalives_idle=int(os.environ.get('DB_KEEPALIVES_IDLE', '10')),
        keepalives_interval=int(os.environ.get('DB_KEEPALIVES_INTERVAL', '5')),
        keepalives_count=int(os.environ.get('DB_KEEPALIVES_COUNT', '3')),
        options=(
            f'-c statement_timeout={stmt_ms} '
            f'-c lock_timeout={lock_ms} '
            f'-c idle_in_transaction_session_timeout={stmt_ms}'
        ),
    )


def _init_db_pool():
    global _db_pool
    with _db_pool_lock:
        if _db_pool is not None:
            return _db_pool
        minconn = int(os.environ.get('DB_POOL_MIN', '1'))
        maxconn = int(os.environ.get('DB_POOL_MAX', '6'))
        try:
            from psycopg2 import pool as pg_pool
            _db_pool = pg_pool.ThreadedConnectionPool(
                minconn,
                maxconn,
                **_db_connect_kwargs(),
            )
            logging.info(f"DB pool: min={minconn} max={maxconn}")
        except Exception as e:
            logging.warning(f"DB pool недоступен, fallback на connect(): {e}")
            _db_pool = None
        return _db_pool


def get_db_connection():
    """Соединение из пула. Не открываем безлимитный fallback — иначе hung TCP плодит коннекты."""
    pool = _init_db_pool()
    if pool is not None:
        last_err = None
        for attempt in range(3):
            try:
                conn = pool.getconn()
                if conn and getattr(conn, 'closed', 0):
                    pool.putconn(conn, close=True)
                    conn = pool.getconn()
                return conn
            except Exception as e:
                last_err = e
                logging.warning(f"pool.getconn failed ({attempt + 1}/3): {e}")
                time.sleep(0.25 * (attempt + 1))
        logging.error(f"pool.getconn исчерпан: {last_err}")
        raise last_err or RuntimeError('DB pool exhausted')
    return psycopg2.connect(**_db_connect_kwargs())


def _close_db(conn=None, cur=None, discard=False):
    """Вернуть соединение в пул или закрыть (не бросает наружу)."""
    if cur is not None:
        try:
            cur.close()
        except Exception:
            pass
    if conn is None:
        return
    bad = bool(discard) or bool(getattr(conn, 'closed', 0))
    if not bad:
        try:
            from psycopg2.extensions import TRANSACTION_STATUS_IDLE
            if conn.get_transaction_status() != TRANSACTION_STATUS_IDLE:
                try:
                    conn.rollback()
                except Exception:
                    bad = True
        except Exception:
            bad = True
    pool = _db_pool
    if pool is not None:
        try:
            pool.putconn(conn, close=bad)
            return
        except Exception:
            pass
    try:
        conn.close()
    except Exception:
        pass


def _register_poller_conn(conn):
    global _poller_active_conn
    with _poller_active_conn_lock:
        _poller_active_conn = conn


def _unregister_poller_conn(conn):
    global _poller_active_conn
    with _poller_active_conn_lock:
        if _poller_active_conn is conn:
            _poller_active_conn = None


def _cancel_poller_query():
    """Прервать зависший SQL poller из watchdog (другой поток)."""
    with _poller_active_conn_lock:
        conn = _poller_active_conn
    if conn is None:
        return False
    try:
        conn.cancel()
        logging.warning("db_poller: отправлен cancel() зависшему запросу")
        return True
    except Exception as e:
        logging.warning(f"db_poller: cancel() не удался: {e}")
        return False


_poller_conn = None
_poller_conn_lock = threading.Lock()


def _poller_db_conn():
    """
    Отдельное долгоживущее соединение для poller.

    Из пула poller уходит по двум причинам: getconn/putconn идут под общим локом
    с веб-запросами, а закрытие мёртвого TCP там же может встать на минуты
    (таймаут ретрансмиссии Windows), из-за чего цикл висел ~177 с при быстрых SQL.
    autocommit — чтобы SELECT'ы не держали транзакцию открытой.
    """
    global _poller_conn
    with _poller_conn_lock:
        conn = _poller_conn
        if conn is not None and not getattr(conn, 'closed', 0):
            return conn
        conn = psycopg2.connect(**_db_connect_kwargs())
        conn.autocommit = True
        _poller_conn = conn
    _register_poller_conn(conn)
    logging.info("db_poller: открыто отдельное соединение с БД")
    return conn


def _poller_drop_conn(reason='', close=True):
    """
    Выбросить соединение poller: следующий цикл откроет новое.

    close=False, когда соединением ещё может пользоваться зависший поток —
    закрывать его из чужого потока небезопасно, поэтому только отвязываем.
    Само закрытие всегда в отдельном потоке: оно может встать на мёртвом TCP.
    """
    global _poller_conn
    with _poller_conn_lock:
        conn, _poller_conn = _poller_conn, None
    if conn is None:
        return
    _unregister_poller_conn(conn)
    logging.warning(f"db_poller: соединение отброшено ({reason}, close={close})")
    if not close:
        return

    def _close():
        try:
            conn.close()
        except Exception:
            pass

    threading.Thread(target=_close, daemon=True, name='pg_close').start()


def _mpl_close(fig=None):
    """Закрытие фигуры Matplotlib + страховка от накопления backend Agg."""
    try:
        if fig is not None:
            plt.close(fig)
        plt.close('all')
    except Exception:
        pass

def get_last_timestamp():
    """Максимальный ts в памяти. Для watermark не используется (см. POLLER_STATE_FILE)."""
    max_ts = None
    with data_lock:
        for line_id, types in production_data.items():
            for etype, data in types.items():
                if data['events']:
                    last_ts = data['events'][-1]['ts']
                    if max_ts is None or last_ts > max_ts:
                        max_ts = last_ts
    return max_ts


# ============================
# WATERMARK ОПРОСА (по потокам aggregate / datamatrix)
# ============================
# Раньше окно считалось от max(ts) в памяти: при пустом ответе (таймаут SQL)
# watermark не двигался и каждый цикл повторял тот же тяжёлый запрос.
POLLER_STATE_FILE = os.environ.get(
    'POLLER_STATE_FILE',
    os.path.join(APP_DIR, 'poller_state.json'),
)
# Шаг догона: запрос всегда ограничен окном, иначе первый проход сканирует сутки/недели
FETCH_WINDOW_MIN = int(os.environ.get('FETCH_WINDOW_MIN', '30'))
# Окно при чтении истории: шире, чтобы 7 суток не разбивались на сотни запросов.
# Объём каждого запроса всё равно ограничен MAX_EVENTS_PER_POLL.
FETCH_WINDOW_CATCHUP_MIN = int(os.environ.get('FETCH_WINDOW_CATCHUP_MIN', '360'))
# Запас на «поздние» вставки с прошлым created
FETCH_SAFETY_LAG_SEC = int(os.environ.get('FETCH_SAFETY_LAG_SEC', '2'))

_watermarks = {'agg': None, 'dm': None}
_watermark_lock = threading.Lock()
_has_full_code = None  # None = не проверяли, True/False = кэш схемы


def load_poller_state():
    """
    Watermark с диска — только чтобы не терять события при долгом простое.

    production_data живёт в памяти и при рестарте пуста, поэтому watermark нельзя
    брать «как сохранён»: историю за INITIAL_LOAD_HOURS нужно перечитать заново.
    Сохранённое значение используем лишь если оно СТАРШЕ окна истории
    (приложение стояло дольше INITIAL_LOAD_HOURS — иначе события пропадут).
    """
    replay_from = datetime.now() - timedelta(hours=INITIAL_LOAD_HOURS)
    saved = {}
    if os.path.exists(POLLER_STATE_FILE):
        try:
            with open(POLLER_STATE_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f) or {}
        except Exception as e:
            logging.warning(f"Не удалось прочитать {POLLER_STATE_FILE}: {e}")
            saved = {}
    with _watermark_lock:
        for key in ('agg', 'dm'):
            ts = to_naive_dt(saved.get(key))
            _watermarks[key] = min(ts, replay_from) if ts is not None else replay_from
    logging.info(
        f"Watermark на старте: agg={_watermarks['agg']} dm={_watermarks['dm']} "
        f"(история за {INITIAL_LOAD_HOURS} ч, в файле agg={saved.get('agg')} dm={saved.get('dm')})"
    )


def save_poller_state():
    with _watermark_lock:
        payload = {
            key: (val.isoformat() if isinstance(val, datetime) else None)
            for key, val in _watermarks.items()
        }
    try:
        _atomic_write_json(POLLER_STATE_FILE, payload, indent=2)
    except Exception as e:
        logging.warning(f"Не удалось сохранить {POLLER_STATE_FILE}: {e}")


def _stream_window(stream, now):
    """[from, to] для потока: watermark .. min(now-lag, watermark+окно)."""
    with _watermark_lock:
        wm = _watermarks.get(stream)
    if wm is None:
        wm = now - timedelta(hours=INITIAL_LOAD_HOURS)
    upper_limit = now - timedelta(seconds=FETCH_SAFETY_LAG_SEC)
    behind_min = (now - wm).total_seconds() / 60.0
    window = FETCH_WINDOW_CATCHUP_MIN if behind_min > 2 * FETCH_WINDOW_MIN else FETCH_WINDOW_MIN
    win_end = wm + timedelta(minutes=window)
    return wm, min(win_end, upper_limit)


def _commit_watermark(stream, value):
    if value is None:
        return
    with _watermark_lock:
        cur = _watermarks.get(stream)
        if cur is None or value > cur:
            _watermarks[stream] = value


AGG_SQL = """
WITH picked AS (
    SELECT a.id, a.source, a.aggregation_level, a.created, a.gtin{full_code_col}
    FROM aggregate a
    WHERE (a.state IN %s OR a.state IS NULL)
      AND a.aggregation_level IN (1, 2)
      AND a.created > %s
      AND a.created <= %s
    ORDER BY a.created
    LIMIT %s
)
SELECT
    p.source AS line_name,
    GREATEST(
        COALESCE((SELECT COUNT(*) FROM datamatrix dm WHERE dm.aggregate_id = p.id), 0),
        1
    ) AS quantity,
    CASE
        WHEN p.aggregation_level = 1 THEN 'L1'
        WHEN p.aggregation_level = 2 THEN 'ТУ'
        ELSE 'L1'
    END AS type,
    p.created AS event_time,
    p.gtin AS gtin,
    p.id AS aggregate_id,
    {full_code_expr} AS full_code,
    p.aggregation_level AS agg_level
FROM picked p
ORDER BY p.created
"""

DM_SQL = """
SELECT
    dm.source AS line_name,
    1 AS quantity,
    'DM' AS type,
    dm.created AS event_time,
    a.gtin AS gtin,
    dm.aggregate_id AS aggregate_id,
    {full_code_expr} AS full_code,
    NULL::integer AS agg_level
FROM datamatrix dm
LEFT JOIN aggregate a ON dm.aggregate_id = a.id
WHERE (dm.state IN %s OR dm.state IS NULL)
  AND dm.created > %s
  AND dm.created <= %s
ORDER BY dm.created
LIMIT %s
"""


def _agg_sql(with_full_code):
    if with_full_code:
        return AGG_SQL.format(full_code_col=', a.full_code', full_code_expr='p.full_code')
    return AGG_SQL.format(full_code_col='', full_code_expr='NULL::text')


def _dm_sql(with_full_code):
    if with_full_code:
        return DM_SQL.format(full_code_expr='COALESCE(dm.full_code, a.full_code)')
    return DM_SQL.format(full_code_expr='NULL::text')


def fetch_new_events():
    """
    Инкрементальная выборка агрегатов L1/ТУ и DM.

    Два независимых запроса вместо UNION ALL + ORDER BY: LIMIT прижат к каждому
    потоку, поэтому Postgres читает индекс по created, а не сортирует всё окно.
    Возвращает (rows, pending_watermarks) — watermark двигаем только после
    успешной обработки, иначе события потеряются при сбое.
    """
    global _has_full_code
    now = datetime.now()
    states = tuple(COMPLETED_STATES)
    limit = int(MAX_EVENTS_PER_POLL)
    conn = None
    cur = None
    rows = []
    pending = {}
    steps = {}
    t_step = time.time()
    try:
        conn = _poller_db_conn()
        steps['conn'] = round(time.time() - t_step, 2)
        cur = conn.cursor()

        for stream, sql_builder in (('agg', _agg_sql), ('dm', _dm_sql)):
            win_from, win_to = _stream_window(stream, now)
            if win_to <= win_from:
                continue
            params = (states, win_from, win_to, limit)
            t_sql = time.time()
            try:
                use_full_code = _has_full_code is not False
                cur.execute(sql_builder(use_full_code), params)
                if _has_full_code is None:
                    _has_full_code = True
            except Exception as e_code:
                if _has_full_code is None:
                    logging.warning(f"Схема без full_code ({e_code}) — переключаюсь без него")
                    _has_full_code = False
                    conn.rollback()
                    cur.execute(sql_builder(False), params)
                else:
                    raise

            stream_rows = cur.fetchall()
            sql_sec = time.time() - t_sql
            steps[stream] = round(sql_sec, 2)
            _poller_progress(stage=f'fetched:{stream}')
            if sql_sec > 5:
                logging.warning(
                    f"SQL {stream} медленный: {sql_sec:.1f} с "
                    f"({win_from:%d.%m %H:%M:%S} – {win_to:%d.%m %H:%M:%S}, строк={len(stream_rows)})"
                )

            if stream_rows:
                max_created = max(
                    (to_naive_dt(r[3]) for r in stream_rows if to_naive_dt(r[3]) is not None),
                    default=None,
                )
                # Полное окно выбрано — следующий цикл продолжит с последнего created
                pending[stream] = max_created if len(stream_rows) >= limit else win_to
                rows.extend(stream_rows)
            else:
                pending[stream] = win_to

        rows.sort(key=lambda r: to_naive_dt(r[3]) or now)
        if rows:
            logging.info(f"Найдено событий (агрегаты + DM): {len(rows)}")
        else:
            logging.debug("Новых событий не найдено")
        return rows, pending
    except Exception as e:
        logging.error(f"Ошибка запроса к БД: {e}")
        _poller_drop_conn(reason=str(e)[:120])
        return [], {}
    finally:
        t_close = time.time()
        if cur is not None:
            try:
                cur.close()
            except Exception:
                pass
        steps['close'] = round(time.time() - t_close, 2)
        _health_set(last_fetch_steps=steps)

_event_counter = 0


def _append_event(line_id, entity_type, qty, gtin, ts=None, aggregate_id=None,
                  full_code=None, source=None, meta=None):
    """Добавить событие в production_data и (опционально) emit в socket. Вызывать под data_lock."""
    now = datetime.now()
    if ts is None:
        ts = now
    else:
        ts = to_naive_dt(ts) or now

    raw_gtin = _normalize_gtin(gtin)
    # Sticky только от прироста: −qty (ТSD reconcile / дезагрегация) не должен
    # перебивать продукт линии (иначе после −KOMATSU в «факте» остаётся KOMATSU).
    if raw_gtin and (qty or 0) > 0:
        remember_line_gtin(line_id, raw_gtin)
    event = {
        'ts': ts,
        'qty': qty,
        'gtin': raw_gtin,  # фактический из SQL (None у DM без aggregate_id)
    }
    if aggregate_id is not None:
        event['aggregate_id'] = aggregate_id
    if full_code is not None:
        event['full_code'] = full_code
    if source is not None:
        event['source'] = source
    if meta:
        event.update(meta)

    data = production_data[line_id][entity_type]
    data['events'].append(event)
    if (qty or 0) > 0:
        data['items_per_event'] = qty

    cutoff = now - timedelta(days=7)
    while data['events'] and data['events'][0]['ts'] < cutoff:
        data['events'].popleft()

    last_update[line_id] = now
    return event


def apply_disaggregation(line_id, qty, gtin=None, level=1, aggregate_id=None,
                         full_code=None, source='rabbit', event_time=None):
    """
    Единая точка учёта отмены/разборки агрегата.
    - Пишет отрицательное событие в ДЕЗАГРЕГАЦИЯ (для плана / Excel / карточек «Агрегаты»).
    - Если level == 1 (L1) — зеркалит −qty в поток DM (карточки «штучные коды»).
    - Идемпотентность по aggregate_id через cancelled_ids.
    Возвращает True, если событие применено.
    """
    if not line_id or not qty or qty <= 0:
        return False

    with data_lock:
        if aggregate_id is not None:
            if aggregate_id in cancelled_ids:
                logging.debug(f"Дезагрегация {aggregate_id} уже учтена, пропуск")
                return False
            cancelled_ids.add(aggregate_id)
            # Обновляем known: qty живых = 0
            if aggregate_id in known_aggregates.get(line_id, {}):
                known_aggregates[line_id][aggregate_id]['qty'] = 0
                known_aggregates[line_id][aggregate_id]['cancelled'] = True

        ts = event_time or datetime.now()
        _append_event(
            line_id, 'ДЕЗАГРЕГАЦИЯ', -qty, gtin, ts=ts,
            aggregate_id=aggregate_id, full_code=full_code, source=source,
            meta={'level': level}
        )

        # Каскад L1 → DM: минусуем штуки, чтобы фильтр «DM» тоже отражал отмену
        if level == 1:
            _append_event(
                line_id, 'DM', -qty, gtin, ts=ts,
                aggregate_id=aggregate_id, full_code=full_code, source=source,
                meta={'reason': 'parent_cancel', 'level': level}
            )

    logging.info(
        f"📉 Дезагрегация → {line_id}: -{qty} шт. "
        f"(level={level}, id={aggregate_id}, gtin={gtin}, source={source})"
    )

    if connected_clients:
        payload_base = {
            'line_id': line_id,
            'quantity': -qty,
            'timestamp': (to_naive_dt(ts) or datetime.now()).isoformat(),
            'queue': source,
            'gtin': sticky_line_gtin(line_id, gtin),
            'aggregate_id': aggregate_id,
        }
        socket_emit('production_data', {**payload_base, 'entity_type': 'ДЕЗАГРЕГАЦИЯ'})
        if level == 1:
            socket_emit('production_data', {**payload_base, 'entity_type': 'DM'})

    return True


def process_new_events(rows):
    """Обработка строк БД. socketio.emit — ТОЛЬКО вне data_lock (иначе deadlock)."""
    global _event_counter
    now = datetime.now()
    resume_lines = {}
    dm_activity_lines = {}
    emit_payloads = []
    processed_log_n = 0

    with data_lock:
        for row in rows:
            if len(row) >= 8:
                line_name, quantity, entity_type, event_time, gtin, aggregate_id, full_code, agg_level = row[:8]
            else:
                line_name, quantity, entity_type, event_time, gtin = row[:5]
                aggregate_id, full_code, agg_level = None, None, None

            if quantity == 0:
                quantity = 1

            gtin_normalized = _normalize_gtin(gtin)
            if not gtin_normalized and full_code:
                gtin_normalized = gtin_from_full_code(full_code)
            event_time = to_naive_dt(event_time) or event_time

            _append_event(
                line_name, entity_type, quantity, gtin_normalized, ts=event_time,
                aggregate_id=aggregate_id, full_code=full_code, source='DB'
            )

            if entity_type in ('L1', 'ТУ') and aggregate_id is not None:
                level = 1 if entity_type == 'L1' else 2
                if agg_level is not None:
                    try:
                        level = int(agg_level)
                    except (TypeError, ValueError):
                        pass
                prev = known_aggregates[line_name].get(aggregate_id)
                if not prev or not prev.get('cancelled'):
                    known_aggregates[line_name][aggregate_id] = {
                        'qty': quantity,
                        'level': level,
                        'gtin': gtin_normalized,
                        'full_code': full_code,
                        'ts': event_time,
                        'cancelled': False,
                    }

            if entity_type in ('L1', 'ТУ', 'DM') and quantity and quantity > 0:
                prev_ts = resume_lines.get(line_name)
                if prev_ts is None or (event_time and event_time > prev_ts):
                    resume_lines[line_name] = event_time or now
            if entity_type == 'DM' and quantity and quantity > 0:
                prev_dm = dm_activity_lines.get(line_name)
                if prev_dm is None or (event_time and event_time > prev_dm):
                    dm_activity_lines[line_name] = event_time or now

            emit_payloads.append({
                'line_id': line_name,
                'entity_type': entity_type,
                'quantity': quantity,
                'timestamp': event_time.isoformat() if hasattr(event_time, 'isoformat') else str(event_time),
                'queue': 'DB',
                # В сокет — sticky GTIN (для карточек), пока SQL не даст новый
                'gtin': sticky_line_gtin(line_name, gtin_normalized),
                'aggregate_id': aggregate_id,
            })

            _event_counter += 1
            if _event_counter % 100 == 0:
                processed_log_n = _event_counter
            if _event_counter % 500 == 0:
                _poller_progress()

    if processed_log_n:
        logging.info(f"Обработано {processed_log_n} событий из БД")

    # emit вне lock и вне poller-thread wait: очередь, иначе deadlock Socket.IO
    if connected_clients and emit_payloads:
        to_send = emit_payloads
        if len(to_send) > MAX_SOCKET_EMITS_PER_POLL:
            logging.info(
                f"poller: emit последних {MAX_SOCKET_EMITS_PER_POLL} из {len(to_send)} событий"
            )
            to_send = to_send[-MAX_SOCKET_EMITS_PER_POLL:]
        for payload in to_send:
            socket_emit('production_data', payload)

    for line_id, ts in resume_lines.items():
        try:
            downtime_close(line_id, ended_at=ts)
            _poller_progress(stage='downtime')
        except Exception as e:
            logging.error(f"downtime_close {line_id}: {e}")

    for line_id, ts in dm_activity_lines.items():
        try:
            _max_on_dm_activity(line_id, ts)
        except Exception as e:
            logging.error(f"MAX DM activity {line_id}: {e}")


_poller_thread = None
_poller_stop = threading.Event()
_poller_generation = 0
_poller_start_lock = threading.Lock()


def _poller_backlog_sec():
    """Насколько watermark отстал от «сейчас» (для догона без ожидания POLL_INTERVAL)."""
    now = datetime.now()
    with _watermark_lock:
        values = [v for v in _watermarks.values() if isinstance(v, datetime)]
    if not values:
        return None
    return (now - min(values)).total_seconds()


def db_poller(generation=None):
    """Основной цикл опроса БД. Выходит по stop или смене generation (старый поток после restart)."""
    if generation is None:
        generation = _poller_generation
    logging.info(f"db_poller: поток запущен gen={generation}")
    while not _poller_stop.is_set() and generation == _poller_generation:
        t0 = time.time()
        rows = None
        try:
            now_iso = datetime.now().isoformat(timespec='seconds')
            _health_set(last_poll_at=now_iso, poll_started_at=now_iso,
                        poll_progress_at=now_iso, poll_stage='fetch')
            rows, pending = fetch_new_events()
            fetch_sec = time.time() - t0
            if generation != _poller_generation:
                break
            t_process = time.time()
            if rows:
                _health_set(poll_stage='process')
                process_new_events(rows)
            process_sec = time.time() - t_process
            # Watermark двигаем только после успешной обработки
            for stream, value in (pending or {}).items():
                _commit_watermark(stream, value)
            if pending:
                save_poller_state()
            dur = time.time() - t0
            backlog = _poller_backlog_sec()
            with _health_lock:
                _health['poll_count'] = int(_health.get('poll_count') or 0) + 1
                _health['last_poll_ok_at'] = datetime.now().isoformat(timespec='seconds')
                _health['last_poll_error'] = None
                _health['last_poll_duration_sec'] = round(dur, 3)
                _health['last_fetch_sec'] = round(fetch_sec, 3)
                _health['last_process_sec'] = round(process_sec, 3)
                _health['last_poll_rows'] = len(rows) if rows else 0
                _health['backlog_sec'] = round(backlog, 1) if backlog is not None else None
                _health['poll_started_at'] = None
                _health['poll_progress_at'] = None
                _health['poll_stage'] = 'idle'
            if dur > 15:
                with _health_lock:
                    steps = _health.get('last_fetch_steps')
                logging.warning(
                    f"db_poller медленный цикл: {dur:.1f} с "
                    f"(fetch={fetch_sec:.1f} с, process={process_sec:.1f} с, "
                    f"шаги fetch={steps}, событий={len(rows) if rows else 0})"
                )
        except Exception as e:
            logging.error(f"Ошибка в poller: {e}", exc_info=True)
            _health_set(last_poll_error=str(e)[:500], poll_started_at=None,
                        poll_progress_at=None, poll_stage='error')
            backlog = None
        if generation != _poller_generation:
            break
        # Догон истории идёт окнами: не ждём POLL_INTERVAL, пока отставание велико
        catching_up = backlog is not None and backlog > FETCH_WINDOW_MIN * 60
        _poller_stop.wait(timeout=0.2 if catching_up else max(POLL_INTERVAL, 1))
    logging.info(f"db_poller: поток остановлен gen={generation}")


def start_db_poller(reason='manual'):
    """Запуск / перезапуск потока poller (daemon). Старый поток не продолжает цикл."""
    global _poller_thread, _poller_generation
    with _poller_start_lock:
        _poller_generation += 1
        gen = _poller_generation
        _poller_stop.set()
        _cancel_poller_query()
        t = _poller_thread
        if t is not None and t.is_alive() and t is not threading.current_thread():
            t.join(timeout=5)
        # Не плодим потоки: если прежние не умерли, они держат лок/сокет — рестарт не поможет
        alive_pollers = [
            th for th in threading.enumerate()
            if th.name.startswith('db_poller-') and th.is_alive()
        ]
        max_alive = int(os.environ.get('POLLER_MAX_ALIVE_THREADS', '3'))
        if len(alive_pollers) >= max_alive:
            logging.critical(
                f"db_poller: {len(alive_pollers)} потоков не завершились "
                f"({[th.name for th in alive_pollers]}) — новый не запускаю, нужен рестарт процесса"
            )
            dump_thread_stacks(tag='poller-threads-piled-up')
            return _poller_thread

        _poller_stop.clear()
        _health_set(poller_generation=gen, poll_started_at=None)
        _poller_thread = threading.Thread(
            target=db_poller, args=(gen,), daemon=True, name=f'db_poller-{gen}'
        )
        _poller_thread.start()
        logging.info(f"db_poller: start ({reason}) gen={gen}")
        return _poller_thread


def reconcile_disaggregations():
    """
    SQL-сверка известных агрегатов: «тихие» удаления / CANCEL / частичная разборка.
    Сравнивает known_aggregates с текущим состоянием в БД и применяет apply_disaggregation.
    """
    with data_lock:
        # snapshot: list of (line_id, agg_id, meta)
        snapshot = []
        for line_id, aggs in known_aggregates.items():
            for agg_id, meta in aggs.items():
                if meta.get('cancelled'):
                    continue
                if agg_id in cancelled_ids:
                    continue
                if not meta.get('qty') or meta['qty'] <= 0:
                    continue
                snapshot.append((line_id, agg_id, dict(meta)))

    if not snapshot:
        logging.debug("Сверка: нет известных агрегатов для проверки")
        return 0

    ids = [agg_id for _, agg_id, _ in snapshot]
    applied = 0
    chunk_size = 500
    conn = None
    cur = None
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        for i in range(0, len(ids), chunk_size):
            chunk_ids = ids[i:i + chunk_size]
            try:
                cur.execute(
                    """
                    SELECT a.id,
                           a.full_code,
                           a.aggregation_level,
                           a.source,
                           a.state,
                           a.gtin,
                           (SELECT COUNT(*) FROM datamatrix dm WHERE dm.aggregate_id = a.id) AS live_dm
                    FROM aggregate a
                    WHERE a.id = ANY(%s)
                    """,
                    (chunk_ids,)
                )
            except Exception as e_fc:
                logging.warning(f"Reconcile SELECT с full_code: {e_fc}, fallback")
                try:
                    conn.rollback()
                except Exception:
                    pass
                cur.execute(
                    """
                    SELECT a.id,
                           NULL::text AS full_code,
                           a.aggregation_level,
                           a.source,
                           a.state,
                           a.gtin,
                           (SELECT COUNT(*) FROM datamatrix dm WHERE dm.aggregate_id = a.id) AS live_dm
                    FROM aggregate a
                    WHERE a.id = ANY(%s)
                    """,
                    (chunk_ids,)
                )
            rows = cur.fetchall()
            found = {}
            for r in rows:
                found[r[0]] = {
                    'full_code': r[1],
                    'level': r[2],
                    'source': r[3],
                    'state': r[4],
                    'gtin': r[5],
                    'live_dm': r[6] or 0,
                }

            for line_id, agg_id, meta in snapshot:
                if agg_id not in chunk_ids:
                    continue
                saved_qty = meta.get('qty') or 0
                level = meta.get('level') or 1
                gtin = meta.get('gtin')
                full_code = meta.get('full_code')

                info = found.get(agg_id)
                if info is None:
                    delta = saved_qty
                else:
                    state_raw = (info.get('state') or '')
                    state_up = str(state_raw).strip().upper()
                    if state_up in CANCEL_STATES or state_raw in CANCEL_STATES:
                        delta = saved_qty
                    else:
                        live = info.get('live_dm') or 0
                        if level == 2:
                            delta = 0
                        else:
                            delta = max(0, saved_qty - live)
                    if info.get('full_code'):
                        full_code = info['full_code']
                    if info.get('gtin'):
                        gtin = str(info['gtin']).strip().lstrip('0') or gtin
                    if info.get('level') is not None:
                        try:
                            level = int(info['level'])
                        except (TypeError, ValueError):
                            pass

                if delta > 0:
                    ok = apply_disaggregation(
                        line_id=line_id,
                        qty=delta,
                        gtin=gtin,
                        level=level,
                        aggregate_id=agg_id,
                        full_code=full_code,
                        source='sql_reconcile',
                        event_time=datetime.now(),
                    )
                    if ok:
                        applied += 1
                        with data_lock:
                            if agg_id in known_aggregates.get(line_id, {}):
                                remaining = saved_qty - delta
                                if remaining <= 0:
                                    known_aggregates[line_id][agg_id]['qty'] = 0
                                    known_aggregates[line_id][agg_id]['cancelled'] = True
                                else:
                                    known_aggregates[line_id][agg_id]['qty'] = remaining
    except Exception as e:
        logging.error(f"Ошибка reconcile_disaggregations: {e}", exc_info=True)
        return applied
    finally:
        _close_db(conn, cur)

    if applied:
        logging.info(f"🔄 Сверка: применено {applied} дезагрегаций из {len(snapshot)} известных")
    else:
        logging.debug(f"Сверка: изменений нет (проверено {len(snapshot)} агрегатов)")
    try:
        _health_set(last_reconcile_at=datetime.now().isoformat(timespec='seconds'))
    except Exception:
        pass
    return applied


def reconcile_loop():
    while True:
        try:
            time.sleep(RECONCILE_INTERVAL)
            reconcile_disaggregations()
        except Exception as e:
            logging.error(f"Ошибка в reconcile_loop: {e}", exc_info=True)
            time.sleep(60)


def send_alert_email(subject, body):
    """
    Письмо о проблеме сервера.
    Env:
      ALERT_EMAIL_TO=a@x.ru,b@y.ru
      ALERT_EMAIL_FROM=monitor@x.ru
      SMTP_HOST=smtp.example.ru
      SMTP_PORT=587
      SMTP_USER=...
      SMTP_PASSWORD=...
      SMTP_TLS=1          # STARTTLS (587)
      SMTP_SSL=0          # 1 = SSL с ходу (465)
    Без SMTP_HOST / ALERT_EMAIL_TO — только лог, без исключения.
    """
    to_raw = (os.environ.get('ALERT_EMAIL_TO') or '').strip()
    host = (os.environ.get('SMTP_HOST') or '').strip()
    if not to_raw or not host:
        logging.info("Alert email пропущен (не заданы ALERT_EMAIL_TO / SMTP_HOST)")
        return False
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header

    from_addr = (os.environ.get('ALERT_EMAIL_FROM') or to_raw.split(',')[0]).strip()
    to_list = [x.strip() for x in to_raw.split(',') if x.strip()]
    port = int(os.environ.get('SMTP_PORT', '587'))
    user = (os.environ.get('SMTP_USER') or '').strip()
    password = os.environ.get('SMTP_PASSWORD') or ''
    use_ssl = os.environ.get('SMTP_SSL', '0').strip() in ('1', 'true', 'yes')
    use_tls = os.environ.get('SMTP_TLS', '1').strip() in ('1', 'true', 'yes')

    msg = MIMEText(body, 'plain', 'utf-8')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = from_addr
    msg['To'] = ', '.join(to_list)

    try:
        if use_ssl:
            smtp = smtplib.SMTP_SSL(host, port, timeout=20)
        else:
            smtp = smtplib.SMTP(host, port, timeout=20)
        with smtp:
            smtp.ehlo()
            if use_tls and not use_ssl:
                smtp.starttls()
                smtp.ehlo()
            if user:
                smtp.login(user, password)
            smtp.sendmail(from_addr, to_list, msg.as_string())
        logging.info(f"Alert email отправлен → {to_list}: {subject}")
        return True
    except Exception as e:
        logging.error(f"Не удалось отправить alert email: {e}", exc_info=True)
        return False


def _poller_stale_info(max_silence):
    """
    Зависание = нет успешного poll ИЛИ цикл не продвигается дольше порога.
    last_poll_at (старт цикла) НЕ считается здоровьем — иначе hung SQL выглядит как OK.
    Длительность цикла сама по себе тоже не признак: догон истории идёт минутами,
    но с отметками прогресса, а зависший poller не двигается совсем.
    """
    with _health_lock:
        ok_at = _health.get('last_poll_ok_at')
        started_at = _health.get('poll_started_at')
        progress_at = _health.get('poll_progress_at')
        stage = _health.get('poll_stage')
        err = _health.get('last_poll_error')
        busy = _report_busy
        poll_count = _health.get('poll_count')
        boot = _health.get('started_at')
        gen = _health.get('poller_generation')
    now = datetime.now()
    thread_alive = bool(_poller_thread is not None and _poller_thread.is_alive())

    def _age(iso):
        if not iso:
            return None
        try:
            return (now - datetime.fromisoformat(str(iso))).total_seconds()
        except Exception:
            return None

    ok_age = _age(ok_at)
    inflight_age = _age(started_at)
    progress_age = _age(progress_at)
    boot_age = _age(boot)
    # Пока цикл в работе, о жизни судим по последнему шагу внутри него
    stuck_age = progress_age if (inflight_age is not None and progress_age is not None) else inflight_age

    stale = False
    reason = None
    if not thread_alive:
        stale = boot_age is not None and boot_age > 15
        reason = 'thread_dead'
    elif stuck_age is not None and stuck_age > max_silence:
        stale = True
        reason = f'no_progress_{int(stuck_age)}s_stage_{stage or "?"}'
    elif inflight_age is not None and progress_age is not None:
        # Цикл идёт и продвигается: ok_age устарел законно (догон истории)
        reason = f'in_progress_{int(inflight_age)}s_stage_{stage or "?"}'
    elif ok_age is None:
        if boot_age is not None and boot_age > max_silence:
            stale = True
            reason = 'no_ok_poll'
    elif ok_age > max_silence:
        stale = True
        reason = f'ok_age_{int(ok_age)}s'

    age_for_health = inflight_age if inflight_age is not None else ok_age
    return {
        'stale': stale,
        'reason': reason,
        'ok_age': ok_age,
        'inflight_age': inflight_age,
        'progress_age': progress_age,
        'stage': stage,
        'age': age_for_health,
        'err': err,
        'busy': busy,
        'poll_count': poll_count,
        'thread_alive': thread_alive,
        'last_poll_ok_at': ok_at,
        'poll_started_at': started_at,
        'generation': gen,
    }


def health_watchdog_loop():
    """
    CRITICAL + email, если poller молчит.
    Сначала cancel() SQL, затем restart с новым generation (старый поток не оживёт).
    """
    max_silence = int(os.environ.get('POLLER_WATCHDOG_SEC', '120'))
    cooldown = int(os.environ.get('ALERT_EMAIL_COOLDOWN_SEC', '900'))
    last_alert_ts = 0.0
    was_stale = False
    last_restart_ts = 0.0
    restart_cooldown = int(os.environ.get('POLLER_RESTART_COOLDOWN_SEC', '60'))

    while True:
        try:
            time.sleep(30)
            info = _poller_stale_info(max_silence)
            clients = len(connected_clients)
            now_ts = time.time()

            if info['stale']:
                logging.critical(
                    f"WATCHDOG: poller stale reason={info['reason']} "
                    f"ok_age={info['ok_age']} inflight={info['inflight_age']} "
                    f"progress_age={info['progress_age']} stage={info['stage']} "
                    f"alive={info['thread_alive']} err={info['err']} report_busy={info['busy']}"
                )
                dump_thread_stacks(tag=f"watchdog-{info['reason']}")
                if now_ts - last_restart_ts >= restart_cooldown:
                    try:
                        _cancel_poller_query()
                        # Соединение может быть мёртвым на уровне TCP: новый цикл
                        # должен открыть своё, а не наследовать зависшее
                        _poller_drop_conn(reason='watchdog', close=False)
                        start_db_poller(reason=f"watchdog-{info['reason']}")
                        last_restart_ts = now_ts
                    except Exception as e:
                        logging.error(f"WATCHDOG restart failed: {e}")

                if now_ts - last_alert_ts >= cooldown:
                    host_name = os.environ.get('ALERT_HOSTNAME') or os.environ.get('HOSTNAME') or 'coolstream-monitor'
                    age_s = info['age'] if info['age'] is not None else -1
                    subj = f"[ShopfloorMonitor] ЗАВИСАНИЕ: poller {age_s:.0f}с · {host_name}"
                    body = (
                        f"Производственный монитор не опрашивает БД дольше допустимого.\n\n"
                        f"Хост: {host_name}\n"
                        f"Причина: {info['reason']}\n"
                        f"Молчание (ok): {info['ok_age']}\n"
                        f"Цикл in-flight: {info['inflight_age']}\n"
                        f"Поток жив: {info['thread_alive']}\n"
                        f"Последний ok poll: {info['last_poll_ok_at']}\n"
                        f"Ошибка poller: {info['err'] or '—'}\n"
                        f"Отчёт формируется: {info['busy']}\n"
                        f"Клиентов Socket.IO: {clients}\n"
                        f"poll_count: {info['poll_count']}\n"
                        f"Время сервера: {datetime.now().isoformat(timespec='seconds')}\n\n"
                        f"Выполнен cancel SQL + автоперезапуск потока poller.\n"
                        f"Проверьте logs/app.log и GET /api/health\n"
                    )
                    try:
                        send_alert_email(subj, body)
                    except Exception as e:
                        logging.error(f"alert email: {e}")
                    last_alert_ts = now_ts
                    was_stale = True
            else:
                if was_stale:
                    host_name = os.environ.get('ALERT_HOSTNAME') or os.environ.get('HOSTNAME') or 'coolstream-monitor'
                    try:
                        send_alert_email(
                            f"[ShopfloorMonitor] ВОССТАНОВЛЕН · {host_name}",
                            f"Poller снова отвечает.\n"
                            f"Хост: {host_name}\n"
                            f"last_poll_ok: {info['last_poll_ok_at']}\n"
                            f"Время: {datetime.now().isoformat(timespec='seconds')}\n",
                        )
                    except Exception as e:
                        logging.error(f"alert email recovery: {e}")
                    was_stale = False
                    last_alert_ts = 0.0
        except Exception as e:
            logging.error(f"watchdog: {e}", exc_info=True)


# ============================
# ФУНКЦИИ ДЛЯ RABBITMQ
# ============================
def rabbitmq_listener():
    while True:
        try:
            credentials = pika.PlainCredentials(RABBIT_USER, RABBIT_PASSWORD)
            parameters = pika.ConnectionParameters(
                host=RABBIT_HOST,
                port=RABBIT_PORT,
                virtual_host=RABBIT_VHOST,
                credentials=credentials,
                heartbeat=600,
                blocked_connection_timeout=300
            )
            connection = pika.BlockingConnection(parameters)
            channel = connection.channel()

            logging.info(f"✅ Подключен к RabbitMQ: {RABBIT_HOST}:{RABBIT_PORT}")

            for queue_name in RABBIT_QUEUES:
                try:
                    channel.queue_declare(queue=queue_name, durable=True, passive=False)
                    channel.basic_consume(
                        queue=queue_name,
                        on_message_callback=lambda ch, method, properties, body, q=queue_name:
                            callback_queue(ch, method, properties, body, q),
                        auto_ack=True
                    )
                    logging.info(f"  📨 Подписан на очередь: {queue_name}")
                except Exception as e:
                    logging.error(f"  ❌ Ошибка подписки на очередь {queue_name}: {e}")

            for exchange_name in RABBIT_EXCHANGES:
                try:
                    channel.exchange_declare(exchange=exchange_name, exchange_type='direct', durable=True)
                    result = channel.queue_declare(queue='', exclusive=True)
                    queue_name = result.method.queue
                    channel.queue_bind(exchange=exchange_name, queue=queue_name, routing_key='#')
                    channel.basic_consume(
                        queue=queue_name,
                        on_message_callback=lambda ch, method, properties, body, exch=exchange_name:
                            callback_exchange(ch, method, properties, body, exch),
                        auto_ack=True
                    )
                    logging.info(f"  📨 Подписан на обменник {exchange_name} через очередь {queue_name}")
                except Exception as e:
                    logging.error(f"  ❌ Ошибка подписки на обменник {exchange_name}: {e}")

            logging.info("=" * 60)
            logging.info("⏳ Ожидание сообщений...")
            channel.start_consuming()

        except pika.exceptions.AMQPConnectionError as e:
            logging.error(f"❌ Ошибка подключения к RabbitMQ: {e}")
            time.sleep(10)
        except Exception as e:
            logging.error(f"❌ Непредвиденная ошибка: {e}")
            time.sleep(10)

def callback_queue(ch, method, properties, body, queue_name):
    raw = body.decode('utf-8')
    logging.info(f"📨 [QUEUE] {queue_name} | размер: {len(raw)} байт")
    if logging.getLogger().isEnabledFor(logging.DEBUG):
        logging.debug(f"Содержимое: {raw[:200]}...")

    try:
        msg = json.loads(raw)
    except json.JSONDecodeError:
        logging.warning(f"Не JSON в очереди {queue_name}: {raw[:100]}")
        return

    is_disaggregation = False
    quantity = 0
    line_id = None
    gtin = None
    aggregate_id = None
    full_code = None
    level = 1

    if msg.get('operation_type') == 'UNPACK':
        is_disaggregation = True
        quantity = msg.get('quantity', 0) or msg.get('count', 0)
        line_id = msg.get('line_id') or msg.get('lineId') or msg.get('source')
        gtin = msg.get('gtin')
        aggregate_id = msg.get('aggregate_id') or msg.get('id') or msg.get('aggregateId')
        full_code = msg.get('full_code') or msg.get('fullCode') or msg.get('code')
        level = msg.get('aggregation_level') or msg.get('level') or 1
    elif msg.get('action') == 'delete':
        is_disaggregation = True
        quantity = msg.get('count', 1) or msg.get('quantity', 0)
        line_id = msg.get('line_id') or msg.get('lineId') or msg.get('source')
        gtin = msg.get('gtin')
        aggregate_id = msg.get('aggregate_id') or msg.get('id') or msg.get('aggregateId')
        full_code = msg.get('full_code') or msg.get('fullCode') or msg.get('code')
        level = msg.get('aggregation_level') or msg.get('level') or 1
    elif msg.get('type') == 'UNPACK' or msg.get('type') == 'DISAGGREGATION':
        is_disaggregation = True
        quantity = msg.get('quantity', 0) or msg.get('count', 0)
        line_id = msg.get('line_id') or msg.get('lineId') or msg.get('source')
        gtin = msg.get('gtin')
        aggregate_id = msg.get('aggregate_id') or msg.get('id') or msg.get('aggregateId')
        full_code = msg.get('full_code') or msg.get('fullCode') or msg.get('code')
        level = msg.get('aggregation_level') or msg.get('level') or 1

    if not is_disaggregation:
        if any(key in str(msg).lower() for key in ['unpack', 'disagg', 'delete', 'remove']):
            is_disaggregation = True
            quantity = msg.get('quantity', 0) or msg.get('count', 0)
            line_id = msg.get('line_id') or msg.get('lineId') or msg.get('source')
            gtin = msg.get('gtin')
            aggregate_id = msg.get('aggregate_id') or msg.get('id') or msg.get('aggregateId')
            full_code = msg.get('full_code') or msg.get('fullCode') or msg.get('code')
            level = msg.get('aggregation_level') or msg.get('level') or 1
            logging.warning(f"⚠️ Определено как дезагрегация по ключевым словам в {queue_name}")

    try:
        level = int(level) if level is not None else 1
    except (TypeError, ValueError):
        level = 1

    # Если qty не пришёл — пробуем взять из known_aggregates
    if is_disaggregation and line_id and (not quantity or quantity <= 0) and aggregate_id:
        with data_lock:
            meta = known_aggregates.get(line_id, {}).get(aggregate_id)
            if meta and meta.get('qty'):
                quantity = meta['qty']
                if not gtin:
                    gtin = meta.get('gtin')
                if not full_code:
                    full_code = meta.get('full_code')
                level = meta.get('level', level)

    if not gtin and full_code:
        gtin = gtin_from_full_code(full_code)

    if is_disaggregation and quantity and quantity > 0 and line_id:
        apply_disaggregation(
            line_id=line_id,
            qty=int(quantity),
            gtin=gtin,
            level=level,
            aggregate_id=aggregate_id,
            full_code=full_code,
            source=f'rabbit:{queue_name}',
            event_time=datetime.now(),
        )
    else:
        logging.debug(f"ℹ️ Сообщение не обработано: {msg.get('type', msg.get('operation_type', 'unknown'))}")

def callback_exchange(ch, method, properties, body, exchange_name):
    raw = body.decode('utf-8')
    logging.info(f"📨 [EXCHANGE] из {exchange_name} | размер: {len(raw)} байт")
    if logging.getLogger().isEnabledFor(logging.DEBUG):
        logging.debug(f"Содержимое: {raw[:200]}...")

# ============================
# ФУНКЦИИ СТАТИСТИКИ
# ============================

def to_naive_dt(ts):
    """Приводит timestamp к naive datetime (без tz) — иначе ключи интервалов не совпадают."""
    if ts is None:
        return None
    if isinstance(ts, str):
        try:
            ts = datetime.fromisoformat(ts.replace('Z', '+00:00'))
        except Exception:
            return None
    if not isinstance(ts, datetime):
        return None
    if ts.tzinfo is not None:
        ts = ts.replace(tzinfo=None)
    return ts


def calc_pauses_from_events(events, start_dt, end_dt, threshold_min=3, include_tail=True,
                            breaks=None, ignore_after=None):
    """
    Паузы по промежуткам между событиями.

    - Пауза = gap между уникальными метками >= threshold_min минут
      после вычета регламентированных перерывов (обед и т.п.).
    - Ведущий простой (до первого события) не считаем.
    - Хвостовой: include_tail=True — онлайн; False — сменный Excel без хвоста до 20:00.
    - ignore_after: после выполнения плана тишина/gaps не пауза.
    - intervals: {start, end, minutes} для листа «Паузы».
    """
    empty = {
        'count': 0, 'avg': 0.0, 'max': 0.0, 'total': 0.0,
        'threshold_min': threshold_min, 'intervals': [],
    }
    if not events:
        return empty

    now = datetime.now()
    effective_end = end_dt if end_dt <= now else now
    if start_dt >= effective_end:
        return empty

    if ignore_after is not None:
        if getattr(ignore_after, 'tzinfo', None) is not None:
            ignore_after = ignore_after.replace(tzinfo=None)
        if ignore_after <= start_dt:
            return empty
        if ignore_after < effective_end:
            effective_end = ignore_after

    window_minutes = (effective_end - start_dt).total_seconds() / 60.0
    if window_minutes <= 0:
        return empty

    try:
        thr_min = float(threshold_min)
    except (TypeError, ValueError):
        thr_min = 5.0
    if thr_min < 0:
        thr_min = 0.0

    unique_times = set()
    for ev in events:
        ts = ev['ts'] if isinstance(ev, dict) else ev
        if isinstance(ts, str):
            try:
                ts = datetime.fromisoformat(ts.replace('Z', '+00:00'))
                if ts.tzinfo is not None:
                    ts = ts.replace(tzinfo=None)
            except Exception:
                continue
        if not isinstance(ts, datetime):
            continue
        if ts.tzinfo is not None:
            ts = ts.replace(tzinfo=None)
        if start_dt <= ts < effective_end:
            unique_times.add(ts.replace(microsecond=0))

    times = sorted(unique_times)
    if not times:
        return {**empty, 'threshold_min': thr_min}

    breaks = breaks or []
    intervals = []

    def _maybe_pause(g_start, g_end):
        if ignore_after is not None and g_start >= ignore_after:
            return
        if ignore_after is not None and g_end > ignore_after:
            g_end = ignore_after
        if g_end <= g_start:
            return
        minutes = minutes_excluding_breaks(g_start, g_end, breaks)
        if minutes >= thr_min:
            intervals.append({
                'start': g_start,
                'end': g_end,
                'minutes': minutes,
            })

    for i in range(1, len(times)):
        _maybe_pause(times[i - 1], times[i])

    if include_tail:
        _maybe_pause(times[-1], effective_end)

    pauses = [x['minutes'] for x in intervals]
    count = len(pauses)
    total = sum(pauses) if pauses else 0.0
    # Потолок: не больше рабочего времени окна (минус перерывы)
    work_window = minutes_excluding_breaks(start_dt, effective_end, breaks)
    if work_window > 0 and total > work_window:
        total = work_window
    elif total > window_minutes:
        total = window_minutes

    return {
        'count': count,
        'avg': (sum(pauses) / count) if count else 0.0,
        'max': min(max(pauses), window_minutes) if pauses else 0.0,
        'total': total,
        'threshold_min': thr_min,
        'intervals': intervals,
    }


def compute_stats(line_id, entity_type):
    with data_lock:
        data = production_data.get(line_id, {}).get(entity_type)
        if not data or not data['events']:
            return None
        events = list(data['events'])
    interval_minutes = 10
    intervals = {}
    for ev in events:
        key = ev['ts'].replace(minute=(ev['ts'].minute // interval_minutes)*interval_minutes, second=0, microsecond=0)
        intervals[key] = intervals.get(key, 0) + ev['qty']
    if not intervals:
        return None
    sorted_times = sorted(intervals.keys())
    values = [intervals[t] for t in sorted_times]
    total = sum(values)
    active_vals = [v for v in values if v > 0]
    if not active_vals:
        return {'total': total, 'active_intervals': 0, 'total_intervals': len(values)}
    n_active = len(active_vals)
    n_total = len(values)
    avg_active = sum(active_vals) / n_active
    median = sorted(active_vals)[n_active // 2]
    max_val = max(active_vals)
    mean = avg_active
    variance = sum((x - mean) ** 2 for x in active_vals) / n_active
    std = variance ** 0.5
    variation = std / mean if mean else 0
    # Паузы: порог по линии; хвост в API по диапазону событий ≈ 0
    threshold = pause_threshold_for_line(line_id, entity_type)
    if events:
        ev_start = min(ev['ts'] for ev in events)
        ev_end = max(ev['ts'] for ev in events)
        pause_info = calc_pauses_from_events(
            events, ev_start, ev_end,
            threshold_min=threshold,
            include_tail=True,
        )
    else:
        pause_info = {'count': 0, 'avg': 0.0, 'max': 0.0, 'total': 0.0}
    pause_count = pause_info['count']
    avg_pause = pause_info['avg']
    max_pause = pause_info['max']

    items_per_event = data['items_per_event']
    if items_per_event is None and len(events) > 0:
        items_per_event = total // len(events) if len(events) > 0 else None
    return {
        'total': total,
        'events': len(events),
        'items_per_event': items_per_event,
        'total_intervals': n_total,
        'active_intervals': n_active,
        'avg_active': round(avg_active, 2),
        'median': median,
        'max': max_val,
        'std': round(std, 2),
        'variation': round(variation, 4),
        'pauses': pause_count,
        'avg_pause': round(avg_pause, 1),
        'max_pause': round(max_pause, 1)
    }

def compute_hourly(line_id, entity_type):
    with data_lock:
        data = production_data.get(line_id, {}).get(entity_type)
        if not data or not data['events']:
            return {}
        now = datetime.now()
        cutoff = now - timedelta(hours=24)
        hourly = {}
        for ev in data['events']:
            if ev['ts'] < cutoff:
                continue
            hour_key = ev['ts'].strftime('%Y-%m-%d %H:00')
            hourly[hour_key] = hourly.get(hour_key, 0) + ev['qty']
    return dict(sorted(hourly.items()))

# ============================
# КОНТРОЛЬНЫЕ КАРТЫ (ритмичность)
# ============================
def shewhart_individuals(values):
    """
    Карта индивидуальных значений (I-chart, Shewhart).
    Оценка σ через скользящий размах MR: σ̂ = mean(MR) / d2, d2(n=2)=1.128
    Возвращает dict: points, cl, ucl, lcl, sigma, mr_bar, ooc_idx (индексы вне границ).
    """
    xs = [float(v) for v in values if v is not None]
    n = len(xs)
    if n < 2:
        return None
    cl = sum(xs) / n
    mrs = [abs(xs[i] - xs[i - 1]) for i in range(1, n)]
    mr_bar = sum(mrs) / len(mrs) if mrs else 0.0
    d2 = 1.128
    sigma = (mr_bar / d2) if d2 else 0.0
    ucl = cl + 3.0 * sigma
    lcl = cl - 3.0 * sigma
    # Для времени/количества отрицательный LCL не имеет смысла
    if lcl < 0:
        lcl = 0.0
    ooc = [i for i, v in enumerate(xs) if v > ucl or v < lcl]
    return {
        'points': xs,
        'cl': cl,
        'ucl': ucl,
        'lcl': lcl,
        'sigma': sigma,
        'mr_bar': mr_bar,
        'ooc_idx': ooc,
        'ooc_count': len(ooc),
        'n': n,
    }


def _series_relative_span(values):
    """(max-min)/mean — мера «есть ли разброс». 0 = плоская линия."""
    xs = [float(v) for v in values if v is not None]
    if len(xs) < 2:
        return 0.0
    mean = sum(xs) / len(xs)
    if abs(mean) < 1e-12:
        return 1.0 if (max(xs) - min(xs)) > 0 else 0.0
    return (max(xs) - min(xs)) / abs(mean)


def control_chart_series_for_line(prod_events, intervals_10, line_id=None):
    """
    Ряд для карты ритма.

    Важно для BLS/KRO/L2: qty паллеты часто константа (48, 48, 48…) —
    карта по «шт. за 10 мин» становится горизонтальной прямой на CL.
    Поэтому:
      1) промежутки между событиями (мин) — основной сигнал ритма;
      2) число событий за 10 мин (не сумма qty) — если gaps мало;
      3) сумма qty в активных слотах — только если есть реальный разброс.
    """
    events = [e for e in (prod_events or []) if e.get('ts') and (e.get('qty') or 0) > 0]
    if len(events) < 2:
        return None

    # Время каждого события (не схлопываем в одну секунду разные qty — только microsecond)
    times = sorted(e['ts'].replace(microsecond=0) for e in events)

    # --- 1) Промежутки между последовательными событиями (мин) ---
    gaps = []
    gap_labels = []
    for i in range(1, len(times)):
        g = (times[i] - times[i - 1]).total_seconds() / 60.0
        # отбрасываем нулевые (L1+ТУ в одну секунду) — иначе MR=0 и ложный «идеальный» ритм
        if g >= (1.0 / 60.0):
            gaps.append(g)
            gap_labels.append(times[i])

    # --- 2) Число событий в 10-мин слотах на окне [first, last] ---
    count_series = []
    count_labels = []
    if intervals_10 and times:
        t0, t1 = times[0], times[-1]
        for dt, _q in intervals_10:
            if dt < t0.replace(minute=(t0.minute // 10) * 10, second=0, microsecond=0):
                continue
            if dt > t1:
                break
            n_ev = sum(
                1 for t in times
                if dt <= t < dt + timedelta(minutes=10)
            )
            count_series.append(n_ev)
            count_labels.append(dt)

    # --- 3) Сумма qty в активных 10-мин слотах ---
    qty_series = []
    qty_labels = []
    if intervals_10:
        for dt, q in intervals_10:
            if q and q > 0:
                qty_series.append(q)
                qty_labels.append(dt)

    candidates = []
    if len(gaps) >= 2:
        candidates.append({
            'mode': 'gaps',
            'values': gaps,
            'labels': gap_labels,
            'ylabel': 'Интервал между событиями, мин',
            'title_suffix': 'промежутки между событиями',
            'span': _series_relative_span(gaps),
            'priority': 3,  # выше = предпочтительнее при равном span
        })
    if len(count_series) >= 2:
        candidates.append({
            'mode': 'event_count',
            'values': count_series,
            'labels': count_labels,
            'ylabel': 'Число событий за 10 мин',
            'title_suffix': 'число событий / 10 мин',
            'span': _series_relative_span(count_series),
            'priority': 2,
        })
    if len(qty_series) >= 2:
        candidates.append({
            'mode': 'rate',
            'values': qty_series,
            'labels': qty_labels,
            'ylabel': 'Шт. за активный 10-мин интервал',
            'title_suffix': 'выпуск / 10 мин (активные)',
            'span': _series_relative_span(qty_series),
            'priority': 1,
        })

    if not candidates:
        return None

    # Паллетные / редкие линии: всегда предпочитаем gaps, если там есть разброс
    line_l = (line_id or '').upper()
    prefer_gaps = any(k in line_l for k in ('BLS', 'KRO', '_L2')) or (
        len(gaps) >= 2 and (sum(gaps) / len(gaps)) >= 8.0
    )

    # Отбрасываем почти плоские ряды (span < 2%), кроме случая «другого нет»
    varied = [c for c in candidates if c['span'] >= 0.02]
    pool = varied if varied else candidates

    if prefer_gaps:
        gap_c = next((c for c in pool if c['mode'] == 'gaps'), None)
        if gap_c and gap_c['span'] >= 0.02:
            chosen = gap_c
        else:
            # gaps плоские — берём max span
            chosen = max(pool, key=lambda c: (c['span'], c['priority']))
    else:
        chosen = max(pool, key=lambda c: (c['span'], c['priority']))

    return {
        'mode': chosen['mode'],
        'values': chosen['values'],
        'labels': chosen['labels'],
        'ylabel': chosen['ylabel'],
        'title_suffix': chosen['title_suffix'],
    }


def render_control_chart_png(series, spc, display_name, line_color):
    """PNG контрольной карты I-chart."""
    if not series or not spc:
        return None
    values = spc['points']
    labels = series.get('labels') or list(range(1, len(values) + 1))
    color_hex = f'#{line_color}' if not str(line_color).startswith('#') else line_color

    fig, ax = plt.subplots(figsize=(11, 3.8))
    x = list(range(1, len(values) + 1))

    ax.axhline(spc['cl'], color='#1A1A1A', linewidth=1.4, linestyle='-', label=f"CL {spc['cl']:.1f}")
    ax.axhline(spc['ucl'], color='#E63946', linewidth=1.2, linestyle='--', label=f"UCL {spc['ucl']:.1f}")
    ax.axhline(spc['lcl'], color='#E63946', linewidth=1.2, linestyle='--', label=f"LCL {spc['lcl']:.1f}")
    # зоны ±1σ / ±2σ (лёгкие)
    if spc['sigma'] > 0:
        ax.axhline(spc['cl'] + spc['sigma'], color='#E9C46A', linewidth=0.8, linestyle=':', alpha=0.8)
        ax.axhline(spc['cl'] - spc['sigma'], color='#E9C46A', linewidth=0.8, linestyle=':', alpha=0.8)
        ax.axhline(spc['cl'] + 2 * spc['sigma'], color='#E9C46A', linewidth=0.8, linestyle=':', alpha=0.5)
        ax.axhline(max(0, spc['cl'] - 2 * spc['sigma']), color='#E9C46A', linewidth=0.8, linestyle=':', alpha=0.5)

    ax.plot(x, values, color=color_hex, linewidth=2.0, marker='o', markersize=4.5, label='Наблюдения')
    ooc = set(spc.get('ooc_idx') or [])
    if ooc:
        ax.scatter(
            [x[i] for i in ooc],
            [values[i] for i in ooc],
            color='#E63946', s=48, zorder=5, label=f'Вне границ ({len(ooc)})'
        )

    # Масштаб по данным + границам (иначе при σ≈0 ось «схлопывает» картинку в прямую)
    y_min = min([spc['lcl']] + values)
    y_max = max([spc['ucl']] + values)
    if y_max - y_min < 1e-6:
        pad = max(1.0, abs(y_max) * 0.1 + 0.5)
        y_min, y_max = y_min - pad, y_max + pad
    else:
        pad = (y_max - y_min) * 0.12
        y_min = max(0.0, y_min - pad) if y_min >= 0 else y_min - pad
        y_max = y_max + pad
    ax.set_ylim(y_min, y_max)

    ax.set_title(
        f'Контрольная карта ритма · {display_name} · {series["title_suffix"]}',
        fontsize=11, fontweight='bold', color='#1A1A1A'
    )
    ax.set_ylabel(series['ylabel'])
    ax.set_xlabel('Порядковый номер наблюдения')
    ax.grid(True, linestyle='--', alpha=0.35)
    ax.legend(loc='upper right', fontsize=8, framealpha=0.92)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    # подписи оси X — время, если это datetime
    if labels and hasattr(labels[0], 'strftime'):
        step = max(1, len(labels) // 10)
        ticks = list(range(0, len(labels), step))
        ax.set_xticks([x[i] for i in ticks])
        ax.set_xticklabels([labels[i].strftime('%H:%M') for i in ticks], rotation=0, fontsize=8)
    else:
        ax.xaxis.set_major_locator(MaxNLocator(integer=True, nbins=12))

    fig.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=120, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    _mpl_close(fig)
    return buf


# ============================
# ГЕНЕРАЦИЯ ОТЧЕТА
# ============================
# Один отчёт за раз; CPU-bound — в отдельном потоке/tpool
report_lock = threading.Lock()
_report_busy = False
_report_executor = None

try:
    from eventlet import tpool as _eventlet_tpool
except ImportError:
    _eventlet_tpool = None


class ReportBusyError(Exception):
    """Параллельная генерация отчёта уже идёт."""
    pass


def _get_report_executor():
    global _report_executor
    if _report_executor is None:
        from concurrent.futures import ThreadPoolExecutor
        _report_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix='report')
    return _report_executor


def run_generate_shift_report(shift_date=None, shift_type=None, do_reconcile=True):
    """
    Excel/Matplotlib вне основного цикла запросов:
    - eventlet → tpool
    - threading → ThreadPoolExecutor (таймаут REPORT_TIMEOUT_SEC)
    Межпроцессный report.gen.lock: два app.py на одном каталоге не пишут один xlsx.
    """
    global _report_busy
    cross = _CrossProcessLock(os.path.join(APP_DIR, 'report.gen.lock'))
    if not cross.acquire():
        logging.warning(
            "Генерация отчёта уже выполняется в другом процессе — пропуск "
            "(иначе xlsx портится)"
        )
        raise ReportBusyError('report locked by another process')
    if not report_lock.acquire(blocking=False):
        cross.release()
        logging.warning("Генерация отчёта уже выполняется — пропуск параллельного запуска")
        raise ReportBusyError('report already running')
    _report_busy = True
    t0 = time.time()
    try:
        def _job():
            return generate_shift_report(
                shift_date=shift_date,
                shift_type=shift_type,
                do_reconcile=do_reconcile,
            )

        timeout = int(os.environ.get('REPORT_TIMEOUT_SEC', '300'))
        if SOCKETIO_ASYNC_MODE == 'eventlet' and _eventlet_tpool is not None:
            result = _eventlet_tpool.execute(_job)
        else:
            fut = _get_report_executor().submit(_job)
            result = fut.result(timeout=timeout)
        dur = time.time() - t0
        logging.info(f"Отчёт готов за {dur:.1f} с: {result}")
        _health_set(
            last_report_at=datetime.now().isoformat(timespec='seconds'),
            last_report_duration_sec=round(dur, 1),
        )
        return result
    except ReportBusyError:
        raise
    except Exception as e:
        logging.error(f"Ошибка run_generate_shift_report: {e}", exc_info=True)
        raise
    finally:
        _report_busy = False
        try:
            report_lock.release()
        except Exception:
            pass
        cross.release()


def _excel_sheet_title(name, used=None):
    """Имя листа Excel: без \\ / ? * [ ], ≤31 символов, уникально в книге."""
    raw = str(name or 'Линия')
    title = ''.join('_' if ch in '\\/?*[]' else ch for ch in raw).strip() or 'Линия'
    title = title[:31]
    if used is None:
        return title
    base = title
    n = 2
    while title in used:
        suffix = f'_{n}'
        title = (base[: 31 - len(suffix)] + suffix)
        n += 1
    used.add(title)
    return title


def _line_report_silence_rows(line_id, release_events, dm_events, start_dt, end_dt,
                              effective_end, shift_type, shift_date):
    """
    Паузы/простои линии для листа «Простои и паузы» и журнала мастера.
    → (plan, plan_done_at, silence_rows с kind Пауза/Простой).
    """
    breaks = get_shift_break_intervals(start_dt, end_dt, shift_type)
    plan = get_plan_for_shift(line_id, shift_type, shift_date)
    plan_done_at = plan_completion_moment(release_events, plan, start_dt, effective_end)
    if skip_report_silence_line(line_id):
        return plan, plan_done_at, []
    pause_info = calc_pauses_from_events(
        dm_events, start_dt, end_dt,
        threshold_min=REPORT_SILENCE_DETECT_MIN,
        include_tail=False,
        breaks=breaks,
        ignore_after=plan_done_at,
    )
    return plan, plan_done_at, classify_report_silences(pause_info.get('intervals') or [])


def _shift_window(now=None):
    """Текущая смена → (start_dt, end_dt, shift_type, shift_date) как в generate_shift_report."""
    now = now or datetime.now()
    start_dt = _current_shift_start(now)
    shift_type = 'day' if start_dt.hour == 8 else 'night'
    shift_date = start_dt.date() if shift_type == 'day' else (start_dt + timedelta(days=1)).date()
    return start_dt, start_dt + timedelta(hours=12), shift_type, shift_date


def build_master_journal(now=None):
    """
    Журнал смены для мастера: те же интервалы тишины, что в Excel, по линиям хронологически.
    Паузы короче MASTER_JOURNAL_PAUSE_MIN не показываем; причина — из downtime.json по пересечению.
    """
    now = now or datetime.now()
    start_dt, end_dt, shift_type, shift_date = _shift_window(now)
    effective_end = min(end_dt, now)
    with data_lock:
        snapshot = [
            (line_id, etype, list(data['events']))
            for line_id, types in production_data.items()
            for etype, data in types.items()
            if etype in ('L1', 'ТУ', 'ДЕЗАГРЕГАЦИЯ', 'DM')
        ]
    by_line = {}
    for line_id, etype, events in snapshot:
        if skip_report_silence_line(line_id):
            continue
        bucket = by_line.setdefault(line_id, {'release': [], 'dm': []})
        for ev in events:
            ts = to_naive_dt(ev.get('ts'))
            if ts is None or not (start_dt <= ts < end_dt):
                continue
            qty = ev.get('qty') or 0
            if etype == 'DM':
                if qty > 0:
                    bucket['dm'].append({'ts': ts, 'qty': qty})
            else:
                bucket['release'].append({'ts': ts, 'qty': qty, 'type': etype})

    shift_downtimes = get_downtime_for_shift(start_dt, end_dt, shift_type=shift_type)
    lines = []
    for line_id in sorted(by_line):
        b = by_line[line_id]
        if not b['dm']:
            continue
        b['release'].sort(key=lambda e: e['ts'])
        _plan, done_at, rows = _line_report_silence_rows(
            line_id, b['release'], b['dm'], start_dt, end_dt, effective_end,
            shift_type, shift_date,
        )
        items = []
        for iv in rows:
            st_iv, en_iv = iv.get('start'), iv.get('end')
            if not isinstance(st_iv, datetime) or not isinstance(en_iv, datetime):
                continue
            if done_at and st_iv >= done_at:
                continue
            mins = float(iv.get('minutes') or 0)
            kind = iv.get('kind')
            if kind == 'Пауза' and mins < MASTER_JOURNAL_PAUSE_MIN:
                continue
            rec = _master_downtime_rec(line_id, st_iv, en_iv, shift_downtimes)
            reason = rec.get('reason') if rec else None
            acked = bool(rec and rec.get('acknowledged'))
            items.append({
                'kind': kind,
                'start': _dt_to_iso(st_iv),
                'end': _dt_to_iso(en_iv),
                'minutes': round(mins, 1),
                'reason': reason if acked else None,
                'comment': (rec.get('comment') if rec and acked else None),
                'acknowledged': acked,
                # Ключ для ack: started_at записи простоя; без записи — интервал паузы
                'ack_started': (rec.get('raw_started_at') or rec.get('started_at')) if rec else _dt_to_iso(st_iv),
                'ack_kind': None if rec else 'pause',
                'ack_ended': None if rec else _dt_to_iso(en_iv),
            })
        if items:
            lines.append({
                'line_id': line_id,
                'name': LINE_DISPLAY_NAMES.get(line_id, line_id),
                'items': items,
            })
    with downtime_lock:
        pending = _downtime_pending_locked(now)
    return {
        'shift_start': _dt_to_iso(start_dt),
        'shift_end': _dt_to_iso(end_dt),
        'now': _dt_to_iso(now),
        'pending': pending,
        'pause_min': MASTER_JOURNAL_PAUSE_MIN,
        'pause_max': REPORT_SILENCE_PAUSE_MAX_MIN,
        'lines': lines,
        'unmarked': sum(1 for ln in lines for it in ln['items'] if not it['acknowledged']),
    }


def generate_shift_report(shift_date=None, shift_type=None, do_reconcile=True):
    now = datetime.now()
    if shift_date is None:
        shift_date = now.date()
    if shift_type is None:
        shift_type = 'day' if 8 <= now.hour < 20 else 'night'

    if shift_type == 'day':
        start_dt = datetime(shift_date.year, shift_date.month, shift_date.day, 8, 0, 0)
        end_dt = start_dt + timedelta(hours=12)
    else:
        start_dt = datetime(shift_date.year, shift_date.month, shift_date.day, 20, 0, 0) - timedelta(days=1)
        end_dt = start_dt + timedelta(hours=12)

    shift_label = 'Дневная' if shift_type == 'day' else 'Ночная'
    logging.info(f"Генерация отчёта за {shift_type} смену {shift_date} ({start_dt} – {end_dt})")

    # Сверка «тихих» отмен — опционально (тяжелая SQL, не должна валить HTTP)
    if do_reconcile:
        try:
            reconcile_disaggregations()
        except Exception as e:
            logging.error(f"Сверка перед отчётом: {e}", exc_info=True)

    # ---------- Сбор данных: короткий snapshot под lock, разбор вне lock ----------
    # Выпуск = L1 + ТУ + ДЕЗАГРЕГАЦИЯ (qty ≤ 0). DM — для учёта пауз/простоев (не в выпуск).
    lines_data = {}
    cancel_events = []  # для листа «Отмены»
    with data_lock:
        snapshot = []
        for line_id, types in production_data.items():
            for etype, data in types.items():
                if etype not in ('L1', 'ТУ', 'ДЕЗАГРЕГАЦИЯ', 'DM'):
                    continue
                snapshot.append((line_id, etype, list(data['events'])))

    for line_id, etype, events in snapshot:
        line_events = lines_data.get(line_id)
        if line_events is None:
            line_events = []
            lines_data[line_id] = line_events
        for ev in events:
            ts = to_naive_dt(ev.get('ts'))
            if ts is None:
                continue
            if start_dt <= ts < end_dt:
                item = {
                    'ts': ts,
                    'qty': ev['qty'],
                    'gtin': ev.get('gtin') or 'NO_GTIN',
                    'type': etype,
                    'aggregate_id': ev.get('aggregate_id'),
                    'full_code': ev.get('full_code'),
                    'source': ev.get('source'),
                    'level': ev.get('level'),
                }
                line_events.append(item)
                if etype == 'ДЕЗАГРЕГАЦИЯ':
                    cancel_events.append({**item, 'line_id': line_id})
    lines_data = {lid: evs for lid, evs in lines_data.items() if evs}

    if not lines_data:
        logging.warning("Нет данных за смену, отчёт не создан.")
        return None

    wb = openpyxl.Workbook()
    # Держим BytesIO до wb.save: openpyxl.Image._data() закрывает поток при записи
    _excel_img_keep = []
    wb.remove(wb.active)

    # =========================================================
    # ЛИСТ «СВОДКА»
    # =========================================================
    ws = wb.create_sheet('Сводка', 0)

    # Заголовок
    ws.merge_cells('A1:Q1')
    ws['A1'] = f'Производственный отчёт · {shift_label} смена · {start_dt.strftime("%d.%m.%Y")}'
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(vertical='center')
    ws.row_dimensions[1].height = 28

    ws.merge_cells('A2:Q2')
    ws['A2'] = f'{start_dt.strftime("%d.%m.%Y %H:%M")} — {end_dt.strftime("%d.%m.%Y %H:%M")}  |  Сформирован: {now.strftime("%d.%m.%Y %H:%M")}'
    ws['A2'].font = subtitle_font

    try:
        crew = shift_schedule.get_crew_for_shift(shift_date, shift_type)
    except Exception as e:
        logging.warning(f'shift_crew for report: {e}')
        crew = {}
    brigade_txt = crew.get('brigade') or '—'
    master_txt = crew.get('master') or '—'
    ws.merge_cells('A3:Q3')
    ws['A3'] = f'Бригада: {brigade_txt}'
    ws['A3'].font = subtitle_font
    ws.merge_cells('A4:Q4')
    ws['A4'] = f'Мастер смены: {master_txt}'
    ws['A4'].font = subtitle_font

    # Шапка таблицы (паузы/простои по DM без двойного счёта; паллеты L2 не входят)
    headers = [
        'Линия', 'Выпуск, шт.', 'Событий', 'Ср. / 10 мин',
        'Медиана', 'Макс.', 'СКО', 'Var %',
        'Паузы', 'Сумма пауз, мин', 'Простои', 'Сумма простоев, мин',
        'План', 'Выполнение %', 'Доступность %', 'Эффективность %', 'Статус'
    ]
    header_row = 6
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[header_row].height = 32

    # Ширины колонок
    widths = [22, 12, 10, 12, 10, 10, 10, 9, 9, 14, 11, 11, 10, 13, 13, 14, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Эффективное окно смены (для незавершённой — до now)
    effective_end = end_dt if end_dt <= now else now

    line_stats = {}
    row = 7

    for line_id, events in sorted(lines_data.items()):
        # Нормализуем все ts сразу (naive) — иначе 10-мин корзины и минуты «пустые»
        norm_events = []
        for ev in events:
            ts = to_naive_dt(ev.get('ts'))
            if ts is None:
                continue
            if not (start_dt <= ts < end_dt):
                continue
            norm_events.append({**ev, 'ts': ts})

        # Нетто-выпуск: L1/ТУ + отрицательные ДЕЗАГРЕГАЦИЯ (и −qty ТСД-коррекции).
        # Для бочек раньше брали только плюсы (бандаж от phantom −24); после фикса
        # absent-job reconcile снова считаем нетто — как live/дашборд.
        release_events = [
            e for e in norm_events
            if e.get('type') in ('L1', 'ТУ', 'ДЕЗАГРЕГАЦИЯ')
        ]
        total_qty = sum(e['qty'] for e in release_events)
        # События выпуска (без отмен) — для интервалов, «Событий», доступности
        prod_events = [
            e for e in norm_events
            if e.get('type') in ('L1', 'ТУ') and (e.get('qty') or 0) > 0
        ]
        # Паузы/простои отчёта — только движение индивидуальных кодов (DM)
        dm_events = [
            e for e in norm_events
            if e.get('type') == 'DM' and (e.get('qty') or 0) > 0
        ]
        total_events = len(prod_events)
        cancel_qty = sum(e['qty'] for e in norm_events if e.get('type') == 'ДЕЗАГРЕГАЦИЯ')

        def _build_intervals(evts, start, end, minutes):
            """Корзины выпуска по minutes; ключи — naive datetime."""
            buckets = {}
            cur = start.replace(second=0, microsecond=0)
            cur = cur.replace(minute=(cur.minute // minutes) * minutes)
            while cur < end:
                buckets[cur] = 0
                cur += timedelta(minutes=minutes)
            for ev in evts:
                ts = ev['ts']
                rounded = ts.replace(
                    minute=(ts.minute // minutes) * minutes,
                    second=0, microsecond=0
                )
                if rounded in buckets:
                    buckets[rounded] += ev['qty']
                elif buckets:
                    nearest = min(buckets.keys(), key=lambda k: abs((k - rounded).total_seconds()))
                    if abs((nearest - rounded).total_seconds()) < minutes * 60:
                        buckets[nearest] += ev['qty']
            ordered = sorted(buckets.items())
            vals = [q for _, q in ordered]
            act = [v for v in vals if v > 0]
            return ordered, vals, act

        def _std_cv(samples):
            """СКО и коэффициент вариации (%) по списку значений (n>=2)."""
            n = len(samples)
            if n < 2:
                return 0.0, 0.0
            mean = sum(samples) / n
            if mean == 0:
                return 0.0, 0.0
            var = sum((x - mean) ** 2 for x in samples) / (n - 1)
            s = var ** 0.5
            cv = (s / mean) * 100.0
            return s, cv

        def _inter_event_gaps_min(evts):
            """Промежутки между событиями выпуска, минуты (ритм линии)."""
            times = sorted({e['ts'].replace(microsecond=0) for e in evts if e.get('ts')})
            if len(times) < 2:
                return []
            return [(times[i] - times[i - 1]).total_seconds() / 60.0 for i in range(1, len(times))]

        # Для графика всегда 10-мин корзины
        interval_minutes = 10
        sorted_intervals, values, active = _build_intervals(
            prod_events, start_dt, end_dt, interval_minutes
        )
        n_active = len(active)
        avg_active = sum(active) / n_active if n_active else 0.0
        median = sorted(active)[n_active // 2] if n_active else 0
        max_val = max(active) if active else 0

        # ----- СКО / Var % -----
        # Проблема BLS/KRO/L2: qty часто = 1 у каждого события → CV по qty = 0.
        # Порядок:
        #  1) активные 10-мин корзины, если значений >=2 и есть разброс;
        #  2) те же корзины 30 / 60 мин (реже паллеты → несколько штук в слоте);
        #  3) промежутки между событиями (мин) — ритм, даже при qty≡1;
        #  4) qty событий (если размеры агрегатов реально разные).
        std, variation = 0.0, 0.0
        stat_source = None  # для отладки в лог

        def _try_samples(samples, source):
            nonlocal std, variation, stat_source
            if len(samples) < 2:
                return False
            s, cv = _std_cv(samples)
            # принимаем, даже если s≈0 (тогда честный ноль), но ищем разброс дальше только если s==0 и есть альтернатива
            std, variation, stat_source = s, cv, source
            return s > 1e-9

        # 1) 10 мин
        if n_active >= 2:
            _try_samples(active, '10min')

        # 2) укрупнение корзин, если нет разброса или мало активных
        if std <= 1e-9:
            for mins in (30, 60):
                _, _, act_m = _build_intervals(prod_events, start_dt, end_dt, mins)
                if len(act_m) >= 2 and _try_samples(act_m, f'{mins}min'):
                    # для отображения avg/median/max — по той же шкале, если 10-мин бедна
                    if n_active < 2:
                        avg_active = sum(act_m) / len(act_m)
                        median = sorted(act_m)[len(act_m) // 2]
                        max_val = max(act_m)
                    break

        # 3) ритм: промежутки между событиями (особенно qty≡1)
        if std <= 1e-9:
            gaps = _inter_event_gaps_min(prod_events)
            if len(gaps) >= 2 and _try_samples(gaps, 'gaps_min'):
                pass

        # 4) размеры агрегатов
        if std <= 1e-9:
            event_qtys = [e['qty'] for e in prod_events if e.get('qty') and e['qty'] > 0]
            if len(event_qtys) >= 2:
                _try_samples(event_qtys, 'event_qty')
                if n_active < 2:
                    avg_active = sum(event_qtys) / len(event_qtys)
                    median = sorted(event_qtys)[len(event_qtys) // 2]
                    max_val = max(event_qtys)

        if total_events >= 2 and std <= 1e-9:
            logging.debug(
                f"Отчёт {line_id}: СКО/Var остались 0 "
                f"(events={total_events}, active_10={n_active}, source={stat_source})"
            )
        elif stat_source:
            logging.debug(
                f"Отчёт {line_id}: СКО={std:.2f}, Var={variation:.1f}% (source={stat_source})"
            )

        # Паузы/простои: gap по DM. Паллеты L2 и ТСД/бочки — не учитываем.
        # ≤10 мин = Пауза, >10 мин = Простой (один ряд интервалов, без двойного счёта).
        plan, plan_done_at, silence_rows = _line_report_silence_rows(
            line_id, release_events, dm_events, start_dt, end_dt, effective_end,
            shift_type, shift_date,
        )
        pause_rows = [x for x in silence_rows if x.get('kind') == 'Пауза']
        downtime_rows = [x for x in silence_rows if x.get('kind') == 'Простой']
        pause_count = len(pause_rows)
        pause_total = sum(float(x.get('minutes') or 0) for x in pause_rows)
        avg_pause = (pause_total / pause_count) if pause_count else 0.0
        max_pause = max((float(x.get('minutes') or 0) for x in pause_rows), default=0.0)
        downtime_count = len(downtime_rows)
        downtime_total = sum(float(x.get('minutes') or 0) for x in downtime_rows)
        pause_threshold = REPORT_SILENCE_PAUSE_MAX_MIN
        pause_intervals = silence_rows

        percent = (total_qty / plan * 100.0) if plan > 0 else 0.0

        # ----- Доступность и эффективность (как на дашборде) -----
        minute_set = set()
        for ev in prod_events:
            ts = ev['ts']
            if start_dt <= ts <= effective_end:
                minute_set.add(ts.replace(second=0, microsecond=0))

        total_minutes = max(1, int((effective_end - start_dt).total_seconds() // 60) + 1)
        active_minutes = len(minute_set)
        availability = (active_minutes / float(total_minutes)) * 100.0
        plan_completion = percent
        efficiency = (availability * plan_completion) / 100.0 if plan > 0 else None

        availability = round(availability, 1)
        efficiency = round(efficiency, 1) if efficiency is not None else None

        # Статус
        if plan <= 0:
            status = 'Нет плана'
            status_color = '888888'
        elif percent >= 100:
            status = 'Выполнен'
            status_color = SUCCESS_COLOR
        elif percent >= 80:
            status = 'Близко'
            status_color = WARN_COLOR
        else:
            status = 'Ниже плана'
            status_color = DANGER_COLOR

        line_stats[line_id] = {
            'total_qty': total_qty,
            'total_events': total_events,
            'cancel_qty': cancel_qty,
            'avg_active': avg_active,
            'median': median,
            'max_val': max_val,
            'std': std,
            'variation': variation,
            'pause_count': pause_count,
            'pause_total': pause_total,
            'avg_pause': avg_pause,
            'max_pause': max_pause,
            'downtime_count': downtime_count,
            'downtime_total': downtime_total,
            'pause_intervals': pause_intervals,
            'pause_threshold': pause_threshold,
            'plan': plan,
            'plan_done_at': plan_done_at,
            'percent': percent,
            'availability': availability,
            'efficiency': efficiency,
            'status': status,
            'status_color': status_color,
            'intervals': sorted_intervals,
            'events': norm_events,
            'prod_events': prod_events,
        }

        display_name = LINE_DISPLAY_NAMES.get(line_id, line_id)
        line_color = LINE_COLORS.get(line_id, BRAND_COLOR)

        values_row = [
            display_name, total_qty, total_events, round(avg_active, 1),
            median, max_val, round(std, 1), round(variation, 1),
            pause_count, round(pause_total, 1),
            downtime_count, round(downtime_total, 1),
            plan, round(percent, 1),
            availability,
            efficiency if efficiency is not None else '—',
            status
        ]

        for col, val in enumerate(values_row, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')

            if col == 1:  # название линии
                cell.font = Font(bold=True, color=line_color)
                cell.alignment = Alignment(horizontal='left', vertical='center')
            elif col == 14:  # % выполнения
                cell.font = Font(bold=True, color=status_color)
            elif col == 15:  # доступность
                av_color = SUCCESS_COLOR if availability >= 85 else (WARN_COLOR if availability >= 60 else DANGER_COLOR)
                cell.font = Font(bold=True, color=av_color)
            elif col == 16:  # эффективность
                if efficiency is not None:
                    ef_color = SUCCESS_COLOR if efficiency >= 85 else (WARN_COLOR if efficiency >= 60 else DANGER_COLOR)
                    cell.font = Font(bold=True, color=ef_color)
            elif col == 17:  # статус
                cell.font = Font(bold=True, color=status_color)

        # Data bar для % выполнения
        ws.conditional_formatting.add(
            f'N{row}',
            DataBarRule(start_type='num', start_value=0, end_type='num', end_value=100,
                        color=status_color, showValue=True, minLength=None, maxLength=None)
        )
        row += 1

    # Итоговая строка
    if line_stats:
        total_all = sum(s['total_qty'] for s in line_stats.values())
        plan_all = sum(s['plan'] for s in line_stats.values())
        pause_all = sum(s.get('pause_total') or 0 for s in line_stats.values())
        down_all = sum(s.get('downtime_total') or 0 for s in line_stats.values())
        ws.cell(row=row, column=1, value='ИТОГО').font = Font(bold=True)
        ws.cell(row=row, column=2, value=total_all).font = Font(bold=True)
        ws.cell(row=row, column=10, value=round(pause_all, 1)).font = Font(bold=True)
        ws.cell(row=row, column=12, value=round(down_all, 1)).font = Font(bold=True)
        ws.cell(row=row, column=13, value=plan_all).font = Font(bold=True)
        if plan_all > 0:
            pct = total_all / plan_all * 100
            cell = ws.cell(row=row, column=14, value=round(pct, 1))
            cell.font = Font(bold=True, color=SUCCESS_COLOR if pct >= 100 else DANGER_COLOR)
        # Пояснение: минуты по линиям параллельны — сумма ≠ «простой завода»
        note_row = row + 1
        ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=17)
        ws.cell(
            row=note_row, column=1,
            value=(
                'Паузы/простои: gap по индивидуальным кодам (DM); ≤10 мин = пауза, >10 = простой; '
                'L2 и бочки ТСД исключены. ИТОГО мин — сумма по линиям (работают параллельно).'
            ),
        ).font = Font(size=9, italic=True, color='666666')

    # =========================================================
    # ЛИСТЫ ПО ЛИНИЯМ
    # =========================================================
    used_sheet_titles = {'Сводка', 'Ритм', 'По часам', 'Простои и паузы', 'Отмены'}
    for line_id, st in line_stats.items():
        display_name = LINE_DISPLAY_NAMES.get(line_id, line_id)
        line_color = LINE_COLORS.get(line_id, BRAND_COLOR)
        ws = wb.create_sheet(_excel_sheet_title(display_name, used_sheet_titles))

        # --- Шапка ---
        ws.merge_cells('A1:H1')
        ws['A1'] = f'{display_name}'
        ws['A1'].font = Font(bold=True, size=16, color=line_color)
        ws.row_dimensions[1].height = 26

        ws.merge_cells('A2:H2')
        ws['A2'] = f'{shift_label} смена  ·  {start_dt.strftime("%d.%m.%Y %H:%M")} – {end_dt.strftime("%H:%M")}'
        ws['A2'].font = subtitle_font

        # --- KPI-блок ---
        kpi_row = 4
        eff_str = f"{st['efficiency']:.1f}%" if st.get('efficiency') is not None else '—'
        kpis = [
            ('Выпуск', f"{st['total_qty']} шт."),
            ('План', f"{st['plan']} шт." if st['plan'] else '—'),
            ('Выполнение', f"{st['percent']:.1f}%"),
            ('Доступность', f"{st['availability']:.1f}%"),
            ('Эффективность', eff_str),
            ('Статус', st['status']),
        ]
        for i, (label, value) in enumerate(kpis):
            col = 1 + i * 2
            ws.cell(row=kpi_row, column=col, value=label).font = Font(size=9, color='888888')
            cell = ws.cell(row=kpi_row + 1, column=col, value=value)
            cell.font = Font(bold=True, size=13, color=line_color if i == 0 else '1A1A1A')
            if i == 2:  # % выполнения
                cell.font = Font(bold=True, size=13, color=st['status_color'])
            if i == 3:  # доступность
                av = st.get('availability') or 0
                av_color = SUCCESS_COLOR if av >= 85 else (WARN_COLOR if av >= 60 else DANGER_COLOR)
                cell.font = Font(bold=True, size=13, color=av_color)
            if i == 4:  # эффективность
                ef = st.get('efficiency')
                if ef is not None:
                    ef_color = SUCCESS_COLOR if ef >= 85 else (WARN_COLOR if ef >= 60 else DANGER_COLOR)
                    cell.font = Font(bold=True, size=13, color=ef_color)
            if i == 5:  # статус
                cell.font = Font(bold=True, size=13, color=st['status_color'])

        # --- Продукты (GTIN) ---
        # Кол-во: нетто за смену (L1/ТУ + ДЕЗАГРЕГАЦИЯ).
        # Первое/Последнее: только положительные события выпуска (сканы / агрегаты),
        # иначе −qty reconcile у бочек сдвигает «последнее» на момент коррекции.
        gtin_events = {}
        for ev in st['events']:
            # DM — те же штуки, что уже внутри L1/ТУ; иначе количество удваивается
            if ev.get('type') not in ('L1', 'ТУ', 'ДЕЗАГРЕГАЦИЯ'):
                continue
            g = ev.get('gtin') or 'NO_GTIN'
            if g not in gtin_events:
                gtin_events[g] = {'first': None, 'last': None, 'qty': 0, 'count': 0}
            q = ev.get('qty') or 0
            gtin_events[g]['qty'] += q
            if ev.get('type') in ('L1', 'ТУ') and q > 0 and ev.get('ts'):
                gtin_events[g]['count'] += 1
                ts = ev['ts']
                if gtin_events[g]['first'] is None or ts < gtin_events[g]['first']:
                    gtin_events[g]['first'] = ts
                if gtin_events[g]['last'] is None or ts > gtin_events[g]['last']:
                    gtin_events[g]['last'] = ts

        prod_start = 7
        ws.cell(row=prod_start, column=1, value='Продукты за смену').font = Font(bold=True, size=12)
        ws.merge_cells(f'A{prod_start}:E{prod_start}')

        headers_p = ['GTIN', 'Наименование', 'Кол-во', 'Первое', 'Последнее']
        for col, h in enumerate(headers_p, 1):
            cell = ws.cell(row=prod_start + 1, column=col, value=h)
            cell.fill = PatternFill('solid', fgColor=line_color)
            cell.font = Font(bold=True, color='FFFFFF', size=10)
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border

        r = prod_start + 2
        for gtin, info in sorted(
            gtin_events.items(),
            key=lambda x: (x[1]['first'] is None, x[1]['first'] or datetime.max),
        ):
            # Чистый минус / ноль без выпуска — это reconcile/коррекция, не продукт смены
            if (info.get('qty') or 0) <= 0:
                continue
            name = gtin_names.get(str(gtin).lstrip('0'), gtin_names.get(gtin, 'Неизвестно'))
            ws.cell(row=r, column=1, value=gtin).border = thin_border
            ws.cell(row=r, column=2, value=name).border = thin_border
            ws.cell(row=r, column=3, value=info['qty']).border = thin_border
            ws.cell(row=r, column=3).alignment = Alignment(horizontal='center')
            first_s = info['first'].strftime('%H:%M') if info['first'] else '—'
            last_s = info['last'].strftime('%H:%M') if info['last'] else '—'
            ws.cell(row=r, column=4, value=first_s).border = thin_border
            ws.cell(row=r, column=4).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=5, value=last_s).border = thin_border
            ws.cell(row=r, column=5).alignment = Alignment(horizontal='center')

            # Фото (если есть)
            img_path = None
            gtin_s = str(gtin).strip()
            for candidate in (f'{gtin_s}.jpg', f'{gtin_s.lstrip("0")}.jpg', f'{gtin}.jpg'):
                p = os.path.join(PRODUCTS_IMG_DIR, candidate)
                if os.path.isfile(p):
                    img_path = p
                    break
            if img_path:
                try:
                    xl_img = _excel_product_thumb(img_path, _excel_img_keep)
                    if xl_img is not None:
                        ws.add_image(xl_img, f'F{r}')
                        ws.row_dimensions[r].height = 45
                        ws.column_dimensions['F'].width = 10
                except Exception as e:
                    logging.warning(f'Фото {gtin} ({img_path}): {e}')
            r += 1

        # --- График ---
        chart_row = r + 2
        try:
            times = [dt for dt, _ in st['intervals']]
            quantities = [q for _, q in st['intervals']]

            fig, ax = plt.subplots(figsize=(11, 4.2))
            color_hex = f'#{line_color}'
            ax.fill_between(times, 0, quantities, color=color_hex, alpha=0.25)
            ax.plot(times, quantities, color=color_hex, linewidth=2.2, label='Выпуск / 10 мин')
            ax.axhline(y=st['avg_active'], color='#888888', linestyle='--', linewidth=1, alpha=0.7, label=f"Среднее ({st['avg_active']:.1f})")

            ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
            ax.xaxis.set_major_locator(mdates.HourLocator(interval=1))
            ax.yaxis.set_major_locator(MaxNLocator(integer=True))
            ax.set_xlabel('Время')
            ax.set_ylabel('Шт. за 10 мин')
            ax.set_title(f'Производительность · {display_name}', fontsize=12, fontweight='bold', color='#1A1A1A')
            ax.grid(True, linestyle='--', alpha=0.4)
            ax.legend(loc='upper right', framealpha=0.9)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            plt.xticks(rotation=0)
            fig.tight_layout()

            buf = BytesIO()
            plt.savefig(buf, format='png', dpi=120, bbox_inches='tight', facecolor='white')
            buf.seek(0)
            _mpl_close(fig)

            img = _excel_image_from_stream(buf, 720, 280, _excel_img_keep)
            if img is not None:
                ws.add_image(img, f'A{chart_row}')
        except Exception as e:
            logging.error(f'График {line_id}: {e}')
            ws.cell(row=chart_row, column=1, value=f'Ошибка графика: {e}')

        # --- Контрольная карта ритма (Shewhart I-chart) ---
        cc_row = chart_row + 16
        spc_info = None
        try:
            series = control_chart_series_for_line(
                st.get('prod_events') or [],
                st.get('intervals') or [],
                line_id=line_id,
            )
            spc = shewhart_individuals(series['values']) if series else None
            if series and spc:
                spc_info = {
                    'mode': series['mode'],
                    'cl': spc['cl'],
                    'ucl': spc['ucl'],
                    'lcl': spc['lcl'],
                    'sigma': spc['sigma'],
                    'ooc_count': spc['ooc_count'],
                    'n': spc['n'],
                }
                st['spc'] = spc_info
                buf_cc = render_control_chart_png(series, spc, display_name, line_color)
                img_cc = _excel_image_from_stream(buf_cc, 720, 255, _excel_img_keep)
                if img_cc is not None:
                    ws.add_image(img_cc, f'A{cc_row}')
                note = (
                    f"CL={spc['cl']:.1f}  UCL={spc['ucl']:.1f}  LCL={spc['lcl']:.1f}  "
                    f"σ̂={spc['sigma']:.2f}  вне границ: {spc['ooc_count']} из {spc['n']}"
                )
                ws.cell(row=cc_row + 14, column=1, value=note).font = Font(size=9, color='555555')
            else:
                ws.cell(
                    row=cc_row, column=1,
                    value='Контрольная карта: недостаточно точек (нужно ≥2 наблюдений)'
                ).font = Font(size=10, color='888888')
        except Exception as e:
            logging.error(f'Контрольная карта {line_id}: {e}', exc_info=True)
            ws.cell(row=cc_row, column=1, value=f'Ошибка контрольной карты: {e}')

        # --- Детальная статистика ---
        stat_row = cc_row + 16
        ws.cell(row=stat_row, column=1, value='Детальная статистика').font = Font(bold=True, size=12)

        details = [
            ('Событий', st['total_events']),
            ('Среднее за активный интервал (10 мин)', round(st['avg_active'], 2)),
            ('Медиана', st['median']),
            ('Максимум', st['max_val']),
            ('Стандартное отклонение', round(st['std'], 2)),
            ('Коэффициент вариации', f"{st['variation']:.1f}%"),
            ('Учёт тишины', 'DM; ≤10 мин=пауза, >10=простой; L2 не учитывается'),
            ('Количество пауз (≤10 мин)', st.get('pause_count') or 0),
            ('Суммарная длительность пауз, мин', round(st.get('pause_total', 0), 1)),
            ('Количество простоев (>10 мин)', st.get('downtime_count') or 0),
            ('Суммарная длительность простоев, мин', round(st.get('downtime_total', 0), 1)),
            ('Доступность (доля активных минут)', f"{st['availability']:.1f}%"),
            ('Выполнение плана', f"{st['percent']:.1f}%"),
            ('Эффективность (Дост. × План / 100)',
             f"{st['efficiency']:.1f}%" if st.get('efficiency') is not None else '— (план не задан)'),
        ]
        spc_info = st.get('spc')
        if spc_info:
            _mode = spc_info.get('mode')
            mode_ru = (
                'промежутки, мин' if _mode == 'gaps'
                else 'число событий / 10 мин' if _mode == 'event_count'
                else 'выпуск / 10 мин'
            )
            details.extend([
                ('Контрольная карта (метрика)', mode_ru),
                ('CL (центр)', round(spc_info['cl'], 2)),
                ('UCL (верхняя граница)', round(spc_info['ucl'], 2)),
                ('LCL (нижняя граница)', round(spc_info['lcl'], 2)),
                ('σ̂ (по скользящему размаху)', round(spc_info['sigma'], 2)),
                ('Точек вне границ UCL/LCL', f"{spc_info['ooc_count']} из {spc_info['n']}"),
            ])
        for i, (label, val) in enumerate(details):
            ws.cell(row=stat_row + 1 + i, column=1, value=label).font = Font(color='555555')
            ws.cell(row=stat_row + 1 + i, column=2, value=val)

        for col in range(1, 7):
            ws.column_dimensions[get_column_letter(col)].width = 18
        ws.column_dimensions['B'].width = 42

        # Отмены в деталях линии (если есть)
        if st.get('cancel_qty'):
            cancel_row = stat_row + 2 + len(details)
            ws.cell(row=cancel_row, column=1, value='Отмены (дезагрегация), шт.').font = Font(color='555555')
            cell = ws.cell(row=cancel_row, column=2, value=st['cancel_qty'])
            cell.font = Font(bold=True, color=DANGER_COLOR)

    # =========================================================
    # ЛИСТ «РИТМ» — сводка контрольных карт по линиям
    # =========================================================
    ws_r = wb.create_sheet('Ритм', 1)  # сразу после Сводки
    ws_r.merge_cells('A1:H1')
    ws_r['A1'] = f'Ритмичность (контрольные карты Shewhart) · {shift_label} смена · {start_dt.strftime("%d.%m.%Y")}'
    ws_r['A1'].font = title_font
    ws_r.row_dimensions[1].height = 26
    ws_r.merge_cells('A2:H2')
    ws_r['A2'] = (
        'I-chart: CL ± 3σ̂, σ̂ = mean(MR)/1.128. '
        'Для паллетных линий — промежутки между событиями (мин); иначе — активный выпуск / 10 мин. '
        'Красные точки — вне UCL/LCL.'
    )
    ws_r['A2'].font = subtitle_font

    rhythm_headers = [
        'Линия', 'Метрика', 'n', 'CL', 'UCL', 'LCL', 'σ̂', 'Вне границ'
    ]
    for col, h in enumerate(rhythm_headers, 1):
        cell = ws_r.cell(row=4, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws_r.row_dimensions[4].height = 28

    rr = 5
    for line_id, st in sorted(line_stats.items()):
        display = LINE_DISPLAY_NAMES.get(line_id, line_id)
        spc = st.get('spc')
        if not spc:
            row_vals = [display, '—', '—', '—', '—', '—', '—', 'недостаточно данных']
        else:
            _mode = spc.get('mode')
            mode_ru = (
                'промежутки, мин' if _mode == 'gaps'
                else 'число событий / 10 мин' if _mode == 'event_count'
                else 'выпуск / 10 мин'
            )
            ooc = spc.get('ooc_count', 0)
            n = spc.get('n', 0)
            row_vals = [
                display,
                mode_ru,
                n,
                round(spc['cl'], 2),
                round(spc['ucl'], 2),
                round(spc['lcl'], 2),
                round(spc['sigma'], 2),
                f'{ooc} ({(100.0 * ooc / n):.0f}%)' if n else ooc,
            ]
        for col, val in enumerate(row_vals, 1):
            cell = ws_r.cell(row=rr, column=col, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center' if col != 1 else 'left')
            if col == 1:
                cell.font = Font(bold=True, color=LINE_COLORS.get(line_id, BRAND_COLOR))
            if col == 8 and spc and spc.get('ooc_count', 0) > 0:
                cell.font = Font(bold=True, color=DANGER_COLOR)
        rr += 1

    for i, w in enumerate([22, 18, 8, 10, 10, 10, 10, 14], 1):
        ws_r.column_dimensions[get_column_letter(i)].width = w

    # ----- Базовые правила -----
    guide_row = rr + 1
    ws_r.cell(row=guide_row, column=1, value='Как читать карту').font = Font(bold=True, size=12)
    basics = [
        'CL (центральная линия) — средний уровень процесса за смену; UCL/LCL = CL ± 3σ̂ (оценка σ по скользящему размаху).',
        'Точки внутри UCL–LCL — обычная случайная вариация. Реагировать на каждую точку не нужно.',
        'Красная точка вне границ — особая причина: искать сбой, а не «усреднять» смену.',
        'Карта промежутков (мин): рост = медленнее/паузы; карта выпуска /10 мин: рост = быстрее/плотнее.',
        'Полные графики — на листах линий; в таблице выше — только сводные границы и число выбросов.',
    ]
    for i, t in enumerate(basics):
        cell = ws_r.cell(row=guide_row + 1 + i, column=1, value=t)
        cell.font = Font(size=10, color='555555')
        ws_r.merge_cells(
            start_row=guide_row + 1 + i, start_column=1,
            end_row=guide_row + 1 + i, end_column=8
        )

    # ----- Примеры паттернов -----
    pat_row = guide_row + 2 + len(basics)
    ws_r.cell(row=pat_row, column=1, value='Примеры интерпретации паттернов').font = Font(bold=True, size=12)
    ws_r.merge_cells(start_row=pat_row, start_column=1, end_row=pat_row, end_column=8)

    pat_headers = ['Паттерн', 'Что видно на карте', 'Возможный смысл', 'Что проверить на линии']
    for col, h in enumerate(pat_headers, 1):
        cell = ws_r.cell(row=pat_row + 1, column=col, value=h)
        cell.fill = PatternFill('solid', fgColor='1D3557')
        cell.font = Font(bold=True, color='FFFFFF', size=10)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws_r.row_dimensions[pat_row + 1].height = 30

    patterns = [
        (
            'Выброс вверх (UCL)',
            'Одна точка выше верхней границы',
            'Для промежутков — аномально длинная пауза; для выпуска — разовый пик (двойной цикл, догон)',
            'Журнал простоев, затор, смена катушки/рулона, ручное вмешательство, скан-дубль'
        ),
        (
            'Выброс вниз (LCL)',
            'Одна точка ниже нижней границы',
            'Для выпуска — провал такта; для промежутков — аномально короткий цикл (пачка событий)',
            'Кратковременный стоп без фиксации, ошибка учёта, перестройка, нехватка материала'
        ),
        (
            'Сдвиг уровня',
            '7+ точек подряд только выше или только ниже CL',
            'Процесс «переехал» на новый средний ритм (не случайность)',
            'Смена SKU/формата, скорость линии, смена бригады, новый режим упаковки'
        ),
        (
            'Тренд',
            '6+ точек подряд монотонно растут или падают',
            'Постепенное ускорение или замедление',
            'Прогревающий прогрев, износ, загрязнение, усталость персонала, конец партии'
        ),
        (
            'Прилипание к границе',
            'Много точек у UCL или LCL, но без явного выброса',
            'Процесс на пределе устойчивости — высокий риск срыва',
            'Нагрузка на узкое место, нестабильный подающий поток, частые микропросто'
        ),
        (
            'Пила / высокая амплитуда',
            'Частые скачки вверх-вниз, широкий размах',
            'Большая общая вариабельность; план плохо предсказуем',
            'Смешение продуктов, ручные операции, нестабильная маркировка, разные размеры паллет'
        ),
        (
            'Плоский участок',
            'Длинная серия почти одинаковых значений',
            'Стабильный такт — хорошо; на карте промежутков «ровная полка» после паузы — восстановление',
            'Нормальный режим; если полка = нули выпуска — линия стояла (см. лист «Простои»)'
        ),
        (
            'Серия из 2–3 у границы',
            'Две-три точки подряд около UCL/LCL',
            'Начало особой причины; ещё не «правило 7», но уже сигнал смотреть линию',
            'Не ждать конца смены: принтер, транспортёр, агрегат, сканер'
        ),
        (
            'Одиночный шип + возврат',
            'Выброс и сразу возврат к CL',
            'Эпизод устранён; процесс снова под контролем',
            'Зафиксировать причину в «Простои» (Ремонт/Проблема), чтобы не повторилось'
        ),
        (
            'Много красных точек',
            'ooc_count высок относительно n',
            'Процесс нестабилен: границы «ломаются», управление по среднему бессмысленно',
            'Разделить смену на фазы (SKU, скорости); устранить хронические простои; не сравнивать с «идеалом»'
        ),
    ]

    pr = pat_row + 2
    for name, visible, meaning, check in patterns:
        vals = [name, visible, meaning, check]
        for col, val in enumerate(vals, 1):
            cell = ws_r.cell(row=pr, column=col, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(vertical='center', wrap_text=True, horizontal='left')
            if col == 1:
                cell.font = Font(bold=True, size=10)
            else:
                cell.font = Font(size=9, color='333333')
        ws_r.row_dimensions[pr].height = 48
        pr += 1

    # ширины под длинный текст примеров
    for i, w in enumerate([22, 28, 36, 40], 1):
        ws_r.column_dimensions[get_column_letter(i)].width = w

    note_row = pr + 1
    ws_r.cell(row=note_row, column=1, value='Практика для мастера смены').font = Font(bold=True, size=11)
    practice = [
        '1) Сначала смотрите красные точки и длинные паузы на карте промежутков — они совпадают с реальными остановами.',
        '2) Сдвиг/тренд без единичного выброса — не «поломка», а смена режима: зафиксируйте продукт и скорость.',
        '3) Если на BLS/KRO карта в режиме «промежутки» — сравнивайте с листом «Простои» (причина от мастера).',
        '4) Не сужайте UCL/LCL вручную: границы пересчитаются сами при следующей смене по факту процесса.',
    ]
    for i, t in enumerate(practice):
        cell = ws_r.cell(row=note_row + 1 + i, column=1, value=t)
        cell.font = Font(size=10, color='555555')
        ws_r.merge_cells(
            start_row=note_row + 1 + i, start_column=1,
            end_row=note_row + 1 + i, end_column=8
        )

    # =========================================================
    # ЛИСТ «ПО ЧАСАМ» (факт / расчётный план / % исполнения)
    # План часа = план_смены / PLAN_WORK_HOURS (11); час перерыва = 0
    # =========================================================
    if shift_type == 'day':
        hour_list = list(range(8, 20))
        break_hour = 12
    else:
        hour_list = list(range(20, 24)) + list(range(0, 8))
        break_hour = 0

    ws_h = wb.create_sheet('По часам')
    ws_h.merge_cells('A1:N1')
    ws_h['A1'] = (
        f'Выпуск по часам · {shift_label} смена · {start_dt.strftime("%d.%m.%Y")}  |  '
        f'план часа = план смены / {PLAN_WORK_HOURS} ч (перерыв {break_hour:02d}:00 — 0)'
    )
    ws_h['A1'].font = title_font
    ws_h.row_dimensions[1].height = 26
    ws_h.merge_cells('A2:N2')
    ws_h['A2'] = f'{start_dt.strftime("%d.%m.%Y %H:%M")} — {end_dt.strftime("%d.%m.%Y %H:%M")}'
    ws_h['A2'].font = subtitle_font

    # Заголовок: Линия | 08:00 | ... | Итого факт | Итого план | %
    h_headers = ['Линия']
    for h in hour_list:
        h_headers.append(f'{h:02d}:00')
    h_headers.extend(['Итого факт', 'Итого план', '% смены'])
    for col, h in enumerate(h_headers, 1):
        cell = ws_h.cell(row=4, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws_h.row_dimensions[4].height = 28

    # Подзаголовок: в каждой часовой колонке — факт / план (%)
    r_h = 5
    for line_id, st in line_stats.items():
        display = LINE_DISPLAY_NAMES.get(line_id, line_id)
        plan_total = float(st.get('plan') or 0)
        plan_per_hour = (plan_total / PLAN_WORK_HOURS) if plan_total > 0 else 0.0

        # Факт по часам из prod_events
        hour_fact = {h: 0 for h in hour_list}
        for ev in (st.get('prod_events') or []):
            ts = ev.get('ts')
            if not isinstance(ts, datetime):
                continue
            if not (start_dt <= ts < end_dt):
                continue
            h = ts.hour
            if h in hour_fact:
                q = ev.get('qty') or 0
                try:
                    q = int(q)
                except (TypeError, ValueError):
                    q = 0
                if q > 0:
                    hour_fact[h] += q

        row_vals = [display]
        sum_fact = 0
        sum_plan = 0.0
        for h in hour_list:
            fact = hour_fact.get(h, 0)
            plan_h = 0.0 if h == break_hour else plan_per_hour
            sum_fact += fact
            sum_plan += plan_h
            if plan_h > 0:
                pct = fact / plan_h * 100.0
                cell_txt = f'{fact} / {round(plan_h)} ({pct:.0f}%)'
            elif plan_total > 0 and h == break_hour:
                cell_txt = f'{fact} / 0 (перерыв)'
            else:
                cell_txt = str(fact)
            row_vals.append(cell_txt)

        sum_pct = (sum_fact / sum_plan * 100.0) if sum_plan > 0 else 0.0
        row_vals.extend([sum_fact, round(sum_plan) if sum_plan else 0, round(sum_pct, 1)])

        for col, val in enumerate(row_vals, 1):
            cell = ws_h.cell(row=r_h, column=col, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center' if col != 1 else 'left', wrap_text=True)
            if col == len(row_vals) and sum_plan > 0:
                if sum_pct >= 100:
                    cell.font = Font(bold=True, color='2A9D8F')
                elif sum_pct >= 80:
                    cell.font = Font(bold=True, color='B8860B')
                else:
                    cell.font = Font(bold=True, color=DANGER_COLOR)
        r_h += 1

    if not line_stats:
        ws_h.cell(row=r_h, column=1, value='Нет данных').font = Font(color='888888')

    ws_h.column_dimensions['A'].width = 22
    for col in range(2, len(h_headers) + 1):
        ws_h.column_dimensions[get_column_letter(col)].width = 14

    # =========================================================
    # ЛИСТ «ПРОСТОИ И ПАУЗЫ» (один источник: gap по DM, без двойного счёта)
    # ≤10 мин = Пауза, >10 мин = Простой; паллетные L2 не входят.
    # Причина мастера подставляется при пересечении с downtime.json (не отдельной строкой).
    # =========================================================
    shift_downtimes = get_downtime_for_shift(start_dt, end_dt, shift_type=shift_type)
    master_dt = [
        rec for rec in (shift_downtimes or [])
        if not skip_report_silence_line(rec.get('line_id'))
    ]

    ws_d = wb.create_sheet('Простои и паузы')
    ws_d.merge_cells('A1:H1')
    ws_d['A1'] = (
        f'Простои и паузы · {shift_label} смена · {start_dt.strftime("%d.%m.%Y")}'
    )
    ws_d['A1'].font = title_font
    ws_d.row_dimensions[1].height = 26
    ws_d.merge_cells('A2:H2')
    ws_d['A2'] = (
        f'{start_dt.strftime("%d.%m.%Y %H:%M")} — {end_dt.strftime("%d.%m.%Y %H:%M")}  |  '
        f'только DM; gap ≥{REPORT_SILENCE_DETECT_MIN:g} мин; '
        f'≤{REPORT_SILENCE_PAUSE_MAX_MIN} мин = пауза, '
        f'>{REPORT_SILENCE_PAUSE_MAX_MIN} мин = простой; L2 и бочки ТСД исключены; '
        f'обед вычтен; без двойного счёта'
    )
    ws_d['A2'].font = subtitle_font

    dt_headers = [
        'Линия', 'Тип', 'Начало', 'Конец', 'Длительность, мин',
        'Причина (мастер)', 'Комментарий', 'Отмечено мастером',
    ]
    for col, h in enumerate(dt_headers, 1):
        cell = ws_d.cell(row=4, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws_d.row_dimensions[4].height = 28

    r = 5
    total_pause = 0.0
    total_down = 0.0
    any_row = False
    for line_id, st in sorted(line_stats.items(), key=lambda x: x[0]):
        if skip_report_silence_line(line_id):
            continue
        display = LINE_DISPLAY_NAMES.get(line_id, line_id)
        done_at = st.get('plan_done_at')
        for iv in (st.get('pause_intervals') or []):
            st_iv = iv.get('start')
            en_iv = iv.get('end')
            if done_at and isinstance(st_iv, datetime) and st_iv >= done_at:
                continue
            try:
                mins = float(iv.get('minutes') or 0)
            except (TypeError, ValueError):
                mins = 0.0
            if mins <= 0:
                continue
            kind = iv.get('kind') or (
                'Пауза' if mins <= REPORT_SILENCE_PAUSE_MAX_MIN else 'Простой'
            )
            reason, acked, comment = _master_downtime_meta(line_id, st_iv, en_iv, master_dt)
            if kind == 'Пауза':
                total_pause += mins
            else:
                total_down += mins
            any_row = True
            row_vals = [
                display,
                kind,
                st_iv.strftime('%d.%m %H:%M:%S') if isinstance(st_iv, datetime) else '—',
                en_iv.strftime('%d.%m %H:%M:%S') if isinstance(en_iv, datetime) else '—',
                round(mins, 1),
                reason or '—',
                comment or '—',
                'Да' if acked else 'Нет',
            ]
            for col, val in enumerate(row_vals, 1):
                cell = ws_d.cell(row=r, column=col, value=val)
                if col == 7:
                    # текст мастера, а не формула (комментарий «=…» не исполняется)
                    cell.data_type = 's'
                cell.border = thin_border
                cell.alignment = Alignment(
                    horizontal='left' if col in (1, 7) else 'center',
                    wrap_text=(col == 7),
                )
                if col == 2:
                    cell.font = Font(
                        bold=True,
                        color=WARN_COLOR if kind == 'Пауза' else DANGER_COLOR,
                    )
                if col == 6 and reason and reason != DEFAULT_DOWNTIME_REASON:
                    cell.font = Font(bold=True)
            r += 1

    if any_row:
        ws_d.cell(row=r, column=1, value='ИТОГО паузы').font = Font(bold=True)
        ws_d.cell(row=r, column=5, value=round(total_pause, 1)).font = Font(bold=True)
        r += 1
        ws_d.cell(row=r, column=1, value='ИТОГО простои').font = Font(bold=True)
        ws_d.cell(row=r, column=5, value=round(total_down, 1)).font = Font(bold=True)
        r += 1
        ws_d.cell(row=r, column=1, value='ИТОГО (паузы+простои)').font = Font(bold=True)
        ws_d.cell(row=r, column=5, value=round(total_pause + total_down, 1)).font = Font(bold=True)
    else:
        ws_d.cell(
            row=r, column=1,
            value='Пауз и простоев по индивидуальным кодам за смену не обнаружено',
        ).font = Font(color='888888')

    for i, w in enumerate([22, 12, 18, 18, 16, 18, 36, 16], 1):
        ws_d.column_dimensions[get_column_letter(i)].width = w

    # =========================================================
    # ЛИСТ «ОТМЕНЫ»
    # =========================================================
    if cancel_events:
        ws_c = wb.create_sheet('Отмены')
        ws_c.merge_cells('A1:G1')
        ws_c['A1'] = f'Отмены / дезагрегация · {shift_label} смена · {start_dt.strftime("%d.%m.%Y")}'
        ws_c['A1'].font = title_font
        ws_c.row_dimensions[1].height = 26
        ws_c.merge_cells('A2:G2')
        ws_c['A2'] = f'{start_dt.strftime("%d.%m.%Y %H:%M")} — {end_dt.strftime("%d.%m.%Y %H:%M")}'
        ws_c['A2'].font = subtitle_font

        cancel_headers = ['Линия', 'Уровень', 'GTIN', 'Кол-во', 'Время', 'full_code / id', 'Источник']
        for col, h in enumerate(cancel_headers, 1):
            cell = ws_c.cell(row=4, column=col, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border

        r = 5
        for ev in sorted(cancel_events, key=lambda x: (x.get('ts') or datetime.min)):
            line_id = ev.get('line_id', '')
            display = LINE_DISPLAY_NAMES.get(line_id, line_id)
            level = ev.get('level')
            if level is None:
                level = '—'
            elif int(level) == 1:
                level = 'L1'
            elif int(level) == 2:
                level = 'ТУ'
            code_or_id = ev.get('full_code') or ev.get('aggregate_id') or '—'
            row_vals = [
                display,
                level,
                ev.get('gtin') or '—',
                ev.get('qty'),
                ev['ts'].strftime('%H:%M:%S') if ev.get('ts') else '—',
                str(code_or_id),
                ev.get('source') or '—',
            ]
            for col, val in enumerate(row_vals, 1):
                cell = ws_c.cell(row=r, column=col, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center' if col != 1 else 'left')
                if col == 4:
                    cell.font = Font(bold=True, color=DANGER_COLOR)
            r += 1

        for i, w in enumerate([22, 10, 16, 10, 12, 28, 18], 1):
            ws_c.column_dimensions[get_column_letter(i)].width = w

        total_cancel = sum(e.get('qty') or 0 for e in cancel_events)
        ws_c.cell(row=r, column=1, value='ИТОГО').font = Font(bold=True)
        cell = ws_c.cell(row=r, column=4, value=total_cancel)
        cell.font = Font(bold=True, color=DANGER_COLOR)

    # Зафиксировать plan/fact в plans.json (history) — можно править plan вручную и пересобрать отчёт
    try:
        facts = {
            lid: int(st.get('total_qty') or 0)
            for lid, st in line_stats.items()
        }
        # plan в history: если записи ещё нет — из current; если есть — plan не затираем
        snapshot_shift_to_history(shift_date, shift_type, facts=facts, overwrite_plans=False)
    except Exception as e:
        logging.warning(f"Не удалось записать plan/fact в history: {e}")

    # Сохранение (атомарно: два процесса не перемешают ZIP одного файла)
    filename = f"report_{shift_date}_{shift_type}_{now.strftime('%H%M')}.xlsx"
    filepath = os.path.join(REPORTS_DIR, filename)
    n_imgs = sum(len(getattr(ws, '_images', []) or []) for ws in wb.worksheets)
    try:
        _atomic_save_workbook(wb, filepath)
    finally:
        # Потоки уже прочитаны openpyxl при save; отпускаем ссылки
        _excel_img_keep.clear()
    logging.info(f"✅ Отчёт сохранён: {filepath} · изображений: {n_imgs}")
    return filepath

# ============================
# ПЛАНИРОВЩИК (исправлен)
# ============================
_scheduler_timer = None

def schedule_next_report():
    now = datetime.now()
    today_8_01 = now.replace(hour=8, minute=1, second=0, microsecond=0)
    today_20_01 = now.replace(hour=20, minute=1, second=0, microsecond=0)

    if now < today_8_01:
        next_time = today_8_01
        shift_type = 'night'
        shift_date = now.date()
    elif now < today_20_01:
        next_time = today_20_01
        shift_type = 'day'
        shift_date = now.date()
    else:
        next_time = today_8_01 + timedelta(days=1)
        shift_type = 'night'
        shift_date = next_time.date()

    delta = (next_time - now).total_seconds()
    if delta < 0:
        delta = 0

    logging.info(
        f"⏰ Следующая генерация отчёта: {next_time.strftime('%Y-%m-%d %H:%M')} "
        f"(смена {shift_type} {shift_date}) через {delta//60:.0f} мин"
    )

    def report_wrapper():
        try:
            run_generate_shift_report(shift_date, shift_type, do_reconcile=True)
        except Exception as e:
            logging.error(f"Ошибка автогенерации отчёта: {e}", exc_info=True)
        finally:
            schedule_next_report()

    global _scheduler_timer
    if _scheduler_timer:
        _scheduler_timer.cancel()
    _scheduler_timer = threading.Timer(delta, report_wrapper)
    _scheduler_timer.daemon = True
    _scheduler_timer.start()


# ============================
# СБРОС ПЛАНОВ ПОСЛЕ СМЕНЫ (08:05 / 20:05)
# ============================
_plan_reset_timer = None


def schedule_plan_reset():
    """
    В 20:05 — зафиксировать дневную смену в history и обнулить current.day.
    В 08:05 — зафиксировать ночную смену в history и обнулить current.night.
    (Отчёт уже строится ~08:01/20:01 и пишет fact в history.)
    """
    global _plan_reset_timer
    now = datetime.now()
    today_8_05 = now.replace(hour=8, minute=5, second=0, microsecond=0)
    today_20_05 = now.replace(hour=20, minute=5, second=0, microsecond=0)

    if now < today_8_05:
        next_time = today_8_05
        ended_shift = 'night'
        # ночная смена «принадлежит» календарной дате утра (конец 08:00)
        shift_date = now.date()
    elif now < today_20_05:
        next_time = today_20_05
        ended_shift = 'day'
        shift_date = now.date()
    else:
        next_time = today_8_05 + timedelta(days=1)
        ended_shift = 'night'
        shift_date = (now + timedelta(days=1)).date()

    delta = max(1.0, (next_time - now).total_seconds())
    logging.info(
        f"[PLAN_RESET] scheduled shift={ended_shift} date={shift_date} "
        f"at={next_time.strftime('%Y-%m-%d %H:%M:%S')} in_min={delta / 60:.1f}"
    )

    def _reset_wrapper():
        t0 = datetime.now()
        try:
            logging.info(
                f"[PLAN_RESET] START shift={ended_shift} date={shift_date} "
                f"ts={t0.strftime('%Y-%m-%d %H:%M:%S')}"
            )
            # history: plan из current (если ещё не записан), fact уже мог проставить отчёт
            key = snapshot_shift_to_history(
                shift_date, ended_shift, facts=None, overwrite_plans=False
            )
            reset_current_plans_for_shift_type(ended_shift)
            logging.info(
                f"[PLAN_RESET] OK shift={ended_shift} date={shift_date} "
                f"history_key={key} duration_s={(datetime.now() - t0).total_seconds():.2f}"
            )
        except Exception as e:
            logging.error(
                f"[PLAN_RESET] FAIL shift={ended_shift} date={shift_date}: {e}",
                exc_info=True,
            )
        finally:
            schedule_plan_reset()

    if _plan_reset_timer:
        try:
            _plan_reset_timer.cancel()
        except Exception:
            pass
    _plan_reset_timer = threading.Timer(delta, _reset_wrapper)
    _plan_reset_timer.daemon = True
    _plan_reset_timer.start()


# ============================
# ДОГОНЯЮЩАЯ ГЕНЕРАЦИЯ ПРОПУЩЕННЫХ СМЕН
# ============================
def catch_up_missed_reports():
    now = datetime.now()
    candidates = []

    day_end = now.replace(hour=20, minute=0, second=0, microsecond=0)
    if now >= day_end:
        candidates.append((day_end.date(), 'day'))
    else:
        yesterday = (now - timedelta(days=1)).date()
        candidates.append((yesterday, 'day'))

    night_end = now.replace(hour=8, minute=0, second=0, microsecond=0)
    if now >= night_end:
        candidates.append((night_end.date(), 'night'))
    else:
        yesterday = (now - timedelta(days=1)).date()
        candidates.append((yesterday, 'night'))

    for days_ago in range(2, 4):
        d = (now - timedelta(days=days_ago)).date()
        candidates.append((d, 'day'))
        candidates.append((d, 'night'))

    existing = set(os.listdir(REPORTS_DIR))
    generated = 0
    for shift_date, shift_type in candidates:
        pattern = f"report_{shift_date}_{shift_type}_"
        if any(f.startswith(pattern) for f in existing):
            continue
        logging.info(f"🔄 Догоняем пропущенный отчёт: {shift_date} {shift_type}")
        try:
            # без reconcile на старте — быстрее, меньше шансов завесить boot
            run_generate_shift_report(shift_date, shift_type, do_reconcile=False)
            generated += 1
        except Exception as e:
            logging.error(f"Ошибка догоняющей генерации {shift_date} {shift_type}: {e}")

    if generated:
        logging.info(f"✅ Сгенерировано {generated} пропущенных отчётов")
    else:
        logging.info("ℹ️ Пропущенных отчётов не найдено")

# ============================
# HTTP ЭНДПОИНТЫ
# ============================
@app.route('/')
def index():
    return render_template('index.html', master_mode=False)


@app.route('/master')
def master_ui():
    """Локальный UI мастера смены: только линии с остановкой/простоем."""
    return render_template('index.html', master_mode=True)


@app.route('/api/health')
def api_health():
    """Для мониторинга / systemd / nginx: 200 = ок, 503 = poller завис."""
    with _health_lock:
        h = dict(_health)
    h['report_busy'] = bool(_report_busy)
    h['clients'] = len(connected_clients)
    h['async_mode'] = SOCKETIO_ASYNC_MODE
    max_silence = int(os.environ.get('POLLER_WATCHDOG_SEC', '120'))
    info = _poller_stale_info(max_silence)
    h['last_poll_age_sec'] = round(info['age'], 1) if info['age'] is not None else None
    h['last_poll_ok_age_sec'] = round(info['ok_age'], 1) if info['ok_age'] is not None else None
    h['poll_in_flight_age_sec'] = round(info['inflight_age'], 1) if info['inflight_age'] is not None else None
    h['poller_thread_alive'] = info['thread_alive']
    h['stale_reason'] = info['reason']
    h['status'] = 'stale' if info['stale'] else 'ok'
    with _watermark_lock:
        h['watermarks'] = {
            k: (v.isoformat() if isinstance(v, datetime) else None)
            for k, v in _watermarks.items()
        }
    h['emit_queue'] = _emit_queue.qsize()
    h['tsd'] = {
        'enabled': _tsd_poll_enabled(),
        'seeded': _tsd_seeded,
        'polls': _tsd_stats.get('polls'),
        'emits': _tsd_stats.get('emits'),
        'last_ok': _tsd_stats.get('last_ok'),
        'last_error': _tsd_stats.get('last_error'),
        'last_reconcile': _tsd_stats.get('last_reconcile'),
        'snapshot_keys': len(_tsd_fact_snapshot),
        'shift_fact': _tsd_shift_fact_qty(),
        'state_file': _tsd_state_path(),
        'state_exists': os.path.isfile(_tsd_state_path()),
    }
    code = 503 if info['stale'] else 200
    return jsonify(h), code


@app.route('/api/diag', methods=['GET', 'POST'])
def api_diag():
    """Снимок потоков + запись стеков в logs/hang_traceback.log (для разбора зависаний)."""
    dump_path = dump_thread_stacks(tag='api_diag') if request.method == 'POST' else None
    max_silence = int(os.environ.get('POLLER_WATCHDOG_SEC', '120'))
    with _watermark_lock:
        wm = {k: (v.isoformat() if isinstance(v, datetime) else None) for k, v in _watermarks.items()}
    return jsonify({
        'poller': _poller_stale_info(max_silence),
        'watermarks': wm,
        'threads': [
            {'name': t.name, 'alive': t.is_alive(), 'daemon': t.daemon}
            for t in threading.enumerate()
        ],
        'emit_queue': _emit_queue.qsize(),
        'report_busy': bool(_report_busy),
        'clients': len(connected_clients),
        'db_pool': bool(_db_pool is not None),
        'dump_file': dump_path,
    })


@app.route('/api/reports')
def list_reports():
    files = []
    pattern = re.compile(r'report_(\d{4}-\d{2}-\d{2})_(day|night)_(\d{4})\.xlsx')
    for f in os.listdir(REPORTS_DIR):
        if f.endswith('.xlsx'):
            filepath = os.path.join(REPORTS_DIR, f)
            stat = os.stat(filepath)
            size_str = f"{stat.st_size / 1024:.1f} КБ"
            modified_dt = datetime.fromtimestamp(stat.st_mtime)
            match = pattern.match(f)
            if match:
                shift_date = match.group(1)
                shift_type = match.group(2)
                created_time = match.group(3)
                shift_date_display = datetime.strptime(shift_date, '%Y-%m-%d').strftime('%d.%m.%Y')
                created_time_display = f"{created_time[:2]}:{created_time[2:]}"
            else:
                shift_date_display = '—'
                shift_type = 'unknown'
                created_time_display = '—'
            files.append({
                'name': f,
                'size_str': size_str,
                'modified_str': modified_dt.strftime('%d.%m.%Y %H:%M'),
                'shift_date': shift_date_display,
                'shift_type': shift_type,
                'created_time': created_time_display,
                'modified': stat.st_mtime,
            })
    files.sort(key=lambda x: x['modified'], reverse=True)

    if request.headers.get('Accept', '').startswith('application/json'):
        return jsonify([{
            'name': f['name'],
            'size': os.path.getsize(os.path.join(REPORTS_DIR, f['name'])),
            'modified': datetime.fromtimestamp(f['modified']).isoformat()
        } for f in files])
    else:
        return render_template('reports.html', files=files)

@app.route('/api/report/<filename>', methods=['GET'])
def download_report(filename):
    filepath = os.path.join(REPORTS_DIR, filename)
    if not os.path.exists(filepath):
        return jsonify({'error': 'File not found'}), 404
    return send_file(filepath, as_attachment=True)

@app.route('/api/report/<filename>', methods=['DELETE'])
def delete_report(filename):
    filepath = os.path.join(REPORTS_DIR, filename)
    if not os.path.exists(filepath):
        return jsonify({'error': 'File not found'}), 404
    try:
        os.remove(filepath)
        return jsonify({'status': 'ok'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/time')
def server_time():
    return jsonify({'time': datetime.now().isoformat()})

@app.route('/api/gtin_names')
def get_gtin_names():
    return jsonify(gtin_names)


@app.route('/api/gtin_names/reload', methods=['POST', 'GET'])
def reload_gtin_names():
    """
    Принудительно перечитать «Краткий отчет.xlsx» без рестарта монитора.
    Рассылает gtin_names клиентам по Socket.IO.
    """
    result = load_gtin_names()
    if result.get('ok'):
        socket_emit('gtin_names', dict(gtin_names))
        logging.info(
            'gtin_names reload: count=%s mtime=%s',
            result.get('count'), result.get('mtime'),
        )
        return jsonify({
            'status': 'ok',
            'count': result.get('count'),
            'path': result.get('path'),
            'mtime': result.get('mtime'),
        })
    return jsonify({
        'status': 'error',
        'error': result.get('error') or 'reload_failed',
        'path': result.get('path'),
        'count': result.get('count') or 0,
    }), 500

@app.route('/api/plan', methods=['GET'])
def get_plans():
    """Текущие планы (для UI). История — GET /api/plan/history."""
    with plans_lock:
        return jsonify({line: dict(vals) for line, vals in plans.items()})


@app.route('/api/plan', methods=['POST'])
def set_plans():
    """Обновить текущие планы day/night по линиям (ручной ввод — приоритет над Excel до конца смены)."""
    data = request.json
    if not data:
        return jsonify({'error': 'No data'}), 400
    touched_lines = []
    touched_shifts = set()
    with plans_lock:
        for line, values in data.items():
            if line in plans and isinstance(values, dict):
                if 'day' in values:
                    plans[line]['day'] = int(values.get('day', 0) or 0)
                    touched_shifts.add('day')
                if 'night' in values:
                    plans[line]['night'] = int(values.get('night', 0) or 0)
                    touched_shifts.add('night')
                if 'day' in values or 'night' in values:
                    touched_lines.append(line)
        snapshot = {line: dict(vals) for line, vals in plans.items()}
    if touched_lines and touched_shifts:
        mark_plans_manual(touched_lines, touched_shifts)
    save_plans()
    socket_emit('plan_update', snapshot)
    return jsonify({'status': 'ok', 'manual_locked': sorted(touched_lines)})


@app.route('/api/plan/history', methods=['GET'])
def get_plan_history():
    """
    ?date=YYYY-MM-DD&shift=day|night — одна смена;
    без параметров — вся history.
    """
    date_s = request.args.get('date')
    shift_s = request.args.get('shift')
    with plans_lock:
        if date_s and shift_s in ('day', 'night'):
            key = _shift_history_key(date_s, shift_s)
            return jsonify({key: plans_history.get(key)})
        return jsonify(plans_history)


@app.route('/api/plan/history', methods=['POST'])
def set_plan_history():
    """
    Править зафиксированный план смены (перед пересборкой отчёта).
    body: { "date": "YYYY-MM-DD", "shift": "day"|"night",
            "lines": { "Ralot_L1": 3000, ... } }  // plan qty
    """
    data = request.json or {}
    date_s = data.get('date')
    shift_s = data.get('shift')
    lines_in = data.get('lines') or {}
    if not date_s or shift_s not in ('day', 'night'):
        return jsonify({'error': 'date and shift=day|night required'}), 400
    try:
        datetime.strptime(date_s, '%Y-%m-%d')
    except ValueError:
        return jsonify({'error': 'date must be YYYY-MM-DD'}), 400

    key = _shift_history_key(date_s, shift_s)
    with plans_lock:
        entry = plans_history.get(key) or {
            'date': date_s,
            'shift': shift_s,
            'lines': {},
        }
        lines_out = dict(entry.get('lines') or {})
        for line_id, plan_val in lines_in.items():
            if line_id not in plans:
                continue
            prev = lines_out.get(line_id) or {}
            try:
                plan_i = int(plan_val or 0)
            except (TypeError, ValueError):
                plan_i = 0
            fact_val = prev.get('fact')
            percent = None
            if fact_val is not None and plan_i > 0:
                try:
                    percent = round(int(fact_val) / plan_i * 100.0, 1)
                except (TypeError, ValueError):
                    percent = None
            lines_out[line_id] = {
                'plan': plan_i,
                'fact': fact_val,
                'percent': percent,
            }
        entry['lines'] = lines_out
        entry['date'] = date_s
        entry['shift'] = shift_s
        entry['saved_at'] = datetime.now().isoformat(timespec='seconds')
        entry['edited'] = True
        plans_history[key] = entry
    save_plans()
    logging.info(f"History plan updated: {key}")
    return jsonify({'status': 'ok', 'key': key, 'entry': plans_history.get(key)})


@app.route('/api/downtime', methods=['GET'])
def api_downtime_get():
    return jsonify(downtime_snapshot())


@app.route('/api/master_journal')
def api_master_journal():
    """Журнал простоев/пауз текущей смены для режима мастера."""
    try:
        return jsonify(build_master_journal())
    except Exception as e:
        logging.error(f'master_journal: {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': str(e), 'lines': []}), 500


@app.route('/api/downtime', methods=['POST'])
def api_downtime_post():
    """
    body:
      action: 'ack' | 'open' | 'close' | 'discard'
      line_id: str
      reason: str (для ack)
      started_at: ISO (опционально)
      ended_at: ISO (для close)
    """
    data = request.json or {}
    action = (data.get('action') or '').strip().lower()
    line_id = data.get('line_id') or data.get('lineId')
    if not line_id:
        return jsonify({'status': 'error', 'message': 'line_id required'}), 400

    if action == 'ack':
        rec, err = downtime_ack(
            line_id, data.get('reason'), started_at=data.get('started_at'),
            badge_min=data.get('badge_min'), entity_type=data.get('entity_type'),
            comment=data.get('comment') or data.get('note'),
            ended_at=data.get('ended_at') if data.get('kind') == 'pause' else None,
        )
        if err:
            return jsonify({'status': 'error', 'message': err}), 400
        return jsonify({'status': 'ok', 'record': rec, **downtime_snapshot()})

    if action == 'open':
        rec = downtime_open_or_update(
            line_id,
            started_at=data.get('started_at'),
            reason=data.get('reason'),
            acknowledge=bool(data.get('acknowledge')),
            badge_min=data.get('badge_min'),
            entity_type=data.get('entity_type'),
            comment=data.get('comment') or data.get('note'),
        )
        # record=null + ignored=true — уже закрытый интервал (обратная совместимость: status всё ещё ok)
        return jsonify({
            'status': 'ok',
            'record': rec,
            'ignored': rec is None,
            **downtime_snapshot(),
        })

    if action == 'close':
        rec = downtime_close(line_id, ended_at=data.get('ended_at'))
        return jsonify({'status': 'ok', 'record': rec, **downtime_snapshot()})

    if action == 'discard':
        # План выполнен / снять без Excel и без resume в MAX
        rec = downtime_discard(line_id, reason=data.get('reason') or 'plan_done')
        return jsonify({'status': 'ok', 'record': rec, **downtime_snapshot()})

    return jsonify({'status': 'error', 'message': 'action must be ack|open|close|discard'}), 400


@app.route('/api/stats')
def api_stats():
    result = {}
    for line_id, types in production_data.items():
        result[line_id] = {}
        for etype in types:
            stats = compute_stats(line_id, etype)
            if stats:
                result[line_id][etype] = stats
    return jsonify(result)


def _product_name_for_gtin(gtin):
    if not gtin:
        return None
    s = str(gtin).strip()
    if not s:
        return None
    return gtin_names.get(s.lstrip('0')) or gtin_names.get(s) or s


def build_lines_live():
    """
    Снимок линий для app_gui (read-only): статус, продукт, выпуск смены.
    Не меняет production_data / простои.
    """
    now = datetime.now()
    shift_start = _current_shift_start(now)
    open_dt = {}
    with downtime_lock:
        open_dt = {k: dict(v) for k, v in downtime_open.items()}

    rows = []
    with data_lock:
        line_ids = sorted(production_data.keys(), key=lambda x: LINE_DISPLAY_NAMES.get(x, x))
        for line_id in line_ids:
            types = production_data.get(line_id) or {}
            last_ts = None
            last_gtin = None
            last_gtin_ts = None
            last_etype = None
            qty_shift = 0
            qty_dm_shift = 0
            active_types = []
            type_last = {}

            for etype, data in types.items():
                events = (data.get('events') if isinstance(data, dict) else None) or []
                if not events:
                    continue
                et_last = None
                for ev in events:
                    ts = ev.get('ts')
                    if not isinstance(ts, datetime):
                        continue
                    if ts.tzinfo is not None:
                        ts = ts.replace(tzinfo=None)
                    qty = ev.get('qty') or 0
                    if ts >= shift_start:
                        if etype == 'DM':
                            if qty > 0:
                                qty_dm_shift += qty
                        elif etype in PLAN_FACT_TYPES:
                            qty_shift += qty
                    if qty > 0:
                        if et_last is None or ts > et_last:
                            et_last = ts
                        if last_ts is None or ts > last_ts:
                            last_ts = ts
                            last_etype = etype
                        g = _normalize_gtin(ev.get('gtin'))
                        if g and (last_gtin_ts is None or ts >= last_gtin_ts):
                            last_gtin = g
                            last_gtin_ts = ts
                            remember_line_gtin(line_id, g)
                if et_last is not None:
                    active_types.append(etype)
                    type_last[etype] = et_last

            if not last_gtin:
                last_gtin = sticky_line_gtin(line_id)

            if last_ts is None and line_id not in open_dt and qty_shift == 0 and qty_dm_shift == 0:
                # пустая линия без событий и без простоя — пропускаем шум
                continue

            age_sec = None
            if last_ts is not None:
                age_sec = max(0.0, (now - last_ts).total_seconds())

            # порог плашки: если последняя активность DM — DM-пороги, иначе агрегаты
            if last_etype == 'DM':
                warn_sec = DM_SILENCE_WARN_MIN * 60
                bad_sec = DM_SILENCE_BAD_MIN * 60
            else:
                warn_sec = STATUS_SILENCE_WARN_MIN * 60
                bad_sec = STATUS_SILENCE_BAD_MIN * 60

            dt_rec = open_dt.get(line_id)
            # fact уже посчитан в этом же проходе — НЕ вызывать is_line_plan_complete
            # без fact: он снова берёт data_lock (раньше это был deadlock на Lock).
            plan_done = is_line_plan_complete(line_id, now, fact=qty_shift)

            if age_sec is None:
                status = 'idle'
            elif plan_done:
                status = 'ok'  # план выполнен — тишина не простой
            elif age_sec >= bad_sec:
                status = 'bad'
            elif age_sec >= warn_sec:
                status = 'warn'
            else:
                status = 'ok'
            if dt_rec and status in ('warn', 'bad', 'idle') and not plan_done:
                status = 'downtime'

            rows.append({
                'line_id': line_id,
                'display_name': LINE_DISPLAY_NAMES.get(line_id, line_id),
                'status': status,
                'last_ts': last_ts.isoformat(timespec='seconds') if last_ts else None,
                'age_sec': round(age_sec, 1) if age_sec is not None else None,
                'entity_types': sorted(active_types),
                'last_entity_type': last_etype,
                'gtin': last_gtin,
                'product': _product_name_for_gtin(last_gtin),
                'qty_shift': int(qty_shift),
                'qty_dm_shift': int(qty_dm_shift),
                'plan_done': bool(plan_done),
                'downtime_open': bool(dt_rec) and not plan_done,
                'downtime_reason': (dt_rec or {}).get('reason') if not plan_done else None,
                'downtime_started_at': (dt_rec or {}).get('started_at') if not plan_done else None,
            })

    # линии только в open downtime, без событий в памяти
    known = {r['line_id'] for r in rows}
    for line_id, rec in open_dt.items():
        if line_id in known:
            continue
        rows.append({
            'line_id': line_id,
            'display_name': LINE_DISPLAY_NAMES.get(line_id, line_id),
            'status': 'downtime',
            'last_ts': None,
            'age_sec': None,
            'entity_types': [],
            'last_entity_type': None,
            'gtin': None,
            'product': None,
            'qty_shift': 0,
            'qty_dm_shift': 0,
            'downtime_open': True,
            'downtime_reason': rec.get('reason'),
            'downtime_started_at': rec.get('started_at'),
        })

    status_order = {'bad': 0, 'downtime': 1, 'warn': 2, 'idle': 3, 'ok': 4}
    rows.sort(key=lambda r: (status_order.get(r['status'], 9), r['display_name']))
    return {
        'server_now': now.isoformat(timespec='seconds'),
        'shift_start': shift_start.isoformat(timespec='seconds'),
        'lines': rows,
    }


def build_frontend_snapshot(now=None):
    """
    Полный снимок для фронта / VPS-визуализатора
    (события текущей смены + планы + простои + имена + hourly).
    """
    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    MAX_INIT_EVENTS = int(os.environ.get('MAX_INIT_EVENTS', '30000'))
    cutoff = _current_shift_start(now)
    with data_lock:
        copied = {}
        last_gtin = {}
        for line_id, types in production_data.items():
            copied[line_id] = {}
            for etype, data in types.items():
                evs = []
                for ev in data['events']:
                    ts = to_naive_dt(ev.get('ts'))
                    if ts is None or ts >= cutoff:
                        evs.append(ev)
                if len(evs) > MAX_INIT_EVENTS:
                    evs = evs[-MAX_INIT_EVENTS:]
                copied[line_id][etype] = evs
            sticky = sticky_line_gtin(line_id)
            if sticky:
                last_gtin[line_id] = sticky

    initial_data = {}
    for line_id, types in copied.items():
        initial_data[line_id] = {}
        for etype, evs in types.items():
            initial_data[line_id][etype] = [
                {
                    'ts': ev['ts'].isoformat() if hasattr(ev['ts'], 'isoformat') else str(ev['ts']),
                    'qty': ev['qty'],
                    'gtin': ev.get('gtin'),
                }
                for ev in evs
            ]

    hourly = {}
    for line_id in list(initial_data.keys()):
        hourly[line_id] = {}
        for etype in ('L1', 'ТУ'):
            h = compute_hourly(line_id, etype)
            if h:
                hourly[line_id][etype] = h

    with plans_lock:
        plans_snap = {line: dict(vals) for line, vals in plans.items()}

    try:
        shift_crew = shift_schedule.get_crew_for_datetime(now, allow_io=False)
    except Exception as e:
        logging.warning(f'shift_crew snapshot: {e}')
        shift_crew = {'ok': False, 'message': str(e)}

    try:
        # Алерт «продукт ≠ план» — факт текущей смены (макс. нетто = как на карточке)
        last_for_plan = excel_plan_fact_gtins(now)
        excel_plan = production_plan.build_excel_plan_state(
            day=now.date(),
            last_gtin_by_line=last_for_plan,
            gtin_names=gtin_names,
            line_display=LINE_DISPLAY_NAMES,
            allow_io=False,
            all_fact_gtins_by_line=excel_plan_fact_gtins_all(now),
        )
    except Exception as e:
        logging.warning(f'excel_plan snapshot: {e}')
        excel_plan = {'ok': False, 'message': str(e), 'mismatches': [], 'lines': {}}

    snap = {
        'server_now': now.isoformat(timespec='seconds'),
        'shift_start': cutoff.isoformat(timespec='seconds'),
        'data': initial_data,
        'plans': plans_snap,
        'downtime': downtime_snapshot(),
        'gtin_names': dict(gtin_names),
        'last_gtin': last_gtin,
        'hourly': hourly,
        'shift_crew': shift_crew,
        'excel_plan': excel_plan,
    }
    img_map, img_hashes = _vps_product_image_manifest(snap)
    snap['product_image_hashes'] = img_hashes
    # локальный кэш путей для последующего POST недостающих файлов
    snap['_local_image_paths'] = img_map
    return snap


def _load_env_file_if_present(path):
    if not os.path.isfile(path):
        return
    try:
        with open(path, encoding='utf-8') as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith('#') or '=' not in line:
                    continue
                key, _, val = line.partition('=')
                key = key.strip()
                val = val.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = val
    except Exception as e:
        logging.warning(f'env file {path}: {e}')


def _sha256_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def _find_local_product_image(gtin):
    """Путь к jpg/png продукта на заводе (с нулями и без)."""
    if not gtin:
        return None
    s = str(gtin).strip()
    if not s:
        return None
    variants = [s, s.lstrip('0') or '0']
    names = []
    for v in variants:
        names.extend([f'{v}.jpg', f'{v}.jpeg', f'{v}.png', f'{v}.webp'])
    seen = set()
    for name in names:
        if name in seen:
            continue
        seen.add(name)
        p = os.path.join(PRODUCTS_IMG_DIR, name)
        if os.path.isfile(p) and os.path.getsize(p) > 0:
            return p
    return None


_img_dl_lock = threading.Lock()
_img_dl_inflight = set()
_img_dl_failed_until = {}  # normalized gtin -> monotonic
_IMG_DL_FAIL_COOLDOWN_SEC = float(os.environ.get('PRODUCT_IMG_FAIL_COOLDOWN_SEC', '3600'))


def ensure_product_image_async(gtin):
    """
    Если нет файла в static/products — в фоне запустить image_downloader.
    Не блокирует poller / Socket.IO.
    """
    if not gtin:
        return
    raw = str(gtin).strip()
    g = _normalize_gtin(raw) or raw
    if not g:
        return
    if _find_local_product_image(g) or _find_local_product_image(raw):
        return
    now = time.monotonic()
    with _img_dl_lock:
        if g in _img_dl_inflight:
            return
        until = _img_dl_failed_until.get(g)
        if until and now < until:
            return
        _img_dl_inflight.add(g)

    def _run():
        try:
            try:
                import image_downloader as idl
            except ImportError:
                logging.warning('image_downloader.py не найден — картинка %s не скачается', g)
                with _img_dl_lock:
                    _img_dl_failed_until[g] = time.monotonic() + _IMG_DL_FAIL_COOLDOWN_SEC
                return
            save_dir = os.path.abspath(PRODUCTS_IMG_DIR)
            os.makedirs(save_dir, exist_ok=True)
            idl.SAVE_DIR = save_dir
            if idl.image_exists(g) or idl.image_exists(raw):
                return
            logging.info('product image: скачивание GTIN %s …', g)
            path = idl.download_image(g)
            if not path and raw != g:
                path = idl.download_image(raw)
            if path and os.path.isfile(path):
                # UI часто просит /static/products/<gtin>.jpg — продублируем при необходимости
                try:
                    base, ext = os.path.splitext(path)
                    targets = {
                        os.path.join(save_dir, f'{g}.jpg'),
                        os.path.join(save_dir, f'{g}{ext or ".jpg"}'),
                        os.path.join(save_dir, f'{raw}.jpg'),
                    }
                    for dest in targets:
                        if os.path.abspath(dest) == os.path.abspath(path):
                            continue
                        if not os.path.isfile(dest) or os.path.getsize(dest) == 0:
                            import shutil
                            shutil.copy2(path, dest)
                except Exception as e:
                    logging.debug('product image copy aliases: %s', e)
                logging.info('product image: OK %s → %s', g, path)
            else:
                logging.warning('product image: не удалось скачать GTIN %s', g)
                with _img_dl_lock:
                    _img_dl_failed_until[g] = time.monotonic() + _IMG_DL_FAIL_COOLDOWN_SEC
        except Exception as e:
            logging.warning('product image download %s: %s', g, e)
            with _img_dl_lock:
                _img_dl_failed_until[g] = time.monotonic() + _IMG_DL_FAIL_COOLDOWN_SEC
        finally:
            with _img_dl_lock:
                _img_dl_inflight.discard(g)

    threading.Thread(target=_run, daemon=True, name=f'img_dl_{g[-8:]}').start()


def _vps_product_image_manifest(snap):
    """
    dest_name (как в URL /static/...) -> local_path,
    и dest_name -> sha256.
    """
    paths = {}
    hashes = {}
    logo = os.path.join('static', 'logo.png')
    if os.path.isfile(logo):
        paths['logo.png'] = logo
        hashes['logo.png'] = _sha256_file(logo)

    gtins = set()
    for g in (snap.get('last_gtin') or {}).values():
        if g:
            gtins.add(str(g).strip())
    for types in (snap.get('data') or {}).values():
        if not isinstance(types, dict):
            continue
        for evs in types.values():
            if not isinstance(evs, list):
                continue
            for ev in evs:
                if isinstance(ev, dict) and ev.get('gtin'):
                    gtins.add(str(ev.get('gtin')).strip())

    for gtin in gtins:
        src = _find_local_product_image(gtin)
        if not src:
            ensure_product_image_async(gtin)
            continue
        ext = os.path.splitext(src)[1].lower() or '.jpg'
        if ext not in ('.jpg', '.jpeg', '.png', '.webp'):
            ext = '.jpg'
        # фронт всегда просит .jpg с тем GTIN, что в last_gtin
        dests = [f'products/{gtin}.jpg']
        stripped = (gtin.lstrip('0') or '0')
        if stripped != gtin:
            dests.append(f'products/{stripped}.jpg')
        if ext != '.jpg':
            dests.append(f'products/{gtin}{ext}')
        sha = _sha256_file(src)
        for dest in dests:
            paths[dest] = src
            hashes[dest] = sha
    return paths, hashes


def _vps_post_gzip_json(url, token, payload, timeout):
    body = gzip.compress(
        json.dumps(payload, ensure_ascii=False, separators=(',', ':')).encode('utf-8'),
        compresslevel=6,
    )
    req = urlrequest.Request(
        url,
        data=body,
        method='POST',
        headers={
            'Content-Type': 'application/json',
            'Content-Encoding': 'gzip',
            'Authorization': f'Bearer {token}',
            'User-Agent': 'coolstream-monitor-push/1.0',
        },
    )
    with urlrequest.urlopen(req, timeout=timeout) as resp:
        code = int(getattr(resp, 'status', None) or resp.getcode())
        raw = resp.read()
        data = {}
        if raw:
            try:
                data = json.loads(raw.decode('utf-8'))
            except Exception:
                data = {}
        return code, data


def push_product_images_once(need_names, local_paths, token, ingest_url, timeout):
    """Дослать недостающие фото на VPS (base64)."""
    if not need_names or not local_paths:
        return 0
    img_url = ingest_url.rstrip('/')
    if img_url.endswith('/api/ingest'):
        img_url = img_url[: -len('/api/ingest')] + '/api/ingest_images'
    else:
        img_url = img_url + '/api/ingest_images'
    max_files = int(os.environ.get('VPS_PUSH_IMAGES_PER_TICK', '12'))
    max_bytes = int(os.environ.get('VPS_PUSH_IMAGE_MAX_BYTES', str(2 * 1024 * 1024)))
    files = {}
    for name in need_names:
        if len(files) >= max_files:
            break
        src = local_paths.get(name)
        if not src or not os.path.isfile(src):
            continue
        size = os.path.getsize(src)
        if size > max_bytes:
            logging.warning(f'VPS image skip {name}: {size} bytes')
            continue
        with open(src, 'rb') as f:
            files[name] = base64.b64encode(f.read()).decode('ascii')
    if not files:
        return 0
    code, data = _vps_post_gzip_json(img_url, token, {'files': files}, timeout)
    if code >= 400:
        logging.warning(f'VPS images HTTP {code}')
        return 0
    saved = (data or {}).get('saved')
    n = len(saved) if isinstance(saved, list) else len(files)
    logging.info(f'VPS images: отправлено {n} файл(ов)')
    return n


def _vps_push_enabled():
    _load_env_file_if_present(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vps_push.env'))
    _load_env_file_if_present(os.path.join('vps_visualizer', 'push.env'))
    if os.environ.get('VPS_PUSH_ENABLED', '0').strip().lower() in ('0', 'false', 'no', ''):
        return False
    return bool((os.environ.get('VPS_PUSH_URL') or '').strip() and (os.environ.get('VPS_INGEST_TOKEN') or '').strip())


def push_frontend_snapshot_once():
    """Один gzip POST снимка на VPS. Возвращает True при 2xx."""
    url = (os.environ.get('VPS_PUSH_URL') or '').strip()
    token = (os.environ.get('VPS_INGEST_TOKEN') or '').strip()
    if not url or not token:
        return False
    snap = build_frontend_snapshot()
    local_paths = snap.pop('_local_image_paths', None) or {}
    timeout = float(os.environ.get('VPS_PUSH_TIMEOUT_SEC', '20'))
    try:
        code, data = _vps_post_gzip_json(url, token, snap, timeout)
        if code >= 400:
            logging.warning(f'VPS push HTTP {code}')
            return False
        need = (data or {}).get('need_images') or []
        if need:
            try:
                push_product_images_once(need, local_paths, token, url, timeout)
            except Exception as e:
                logging.warning(f'VPS images: {e}')
        return True
    except HTTPError as e:
        logging.warning(f'VPS push HTTPError {e.code}: {e.reason}')
        return False
    except URLError as e:
        logging.warning(f'VPS push URLError: {e.reason}')
        return False
    except Exception as e:
        logging.warning(f'VPS push: {e}')
        return False


def vps_push_loop():
    """Фон: раз в VPS_PUSH_INTERVAL_SEC секунд — снимок на Timeweb."""
    interval = float(os.environ.get('VPS_PUSH_INTERVAL_SEC', '5'))
    interval = max(2.0, interval)
    logging.info(f'VPS push: старт → {os.environ.get("VPS_PUSH_URL")} каждые {interval:.0f} с')
    fail_streak = 0
    while True:
        try:
            ok = push_frontend_snapshot_once()
            if ok:
                if fail_streak:
                    logging.info('VPS push: снова ок')
                fail_streak = 0
            else:
                fail_streak += 1
                if fail_streak in (1, 5, 20) or fail_streak % 60 == 0:
                    logging.warning(f'VPS push: неудача #{fail_streak}')
        except Exception as e:
            fail_streak += 1
            logging.error(f'VPS push loop: {e}', exc_info=True)
        time.sleep(interval)


# ============================
# ТСД: опрос TotalMark (факт отгрузки с производства)
# ============================
# Только задания с Comment, начинающимся на «П». SerialNumber не нужен.
# Diff QuantityFact → production_data['TSD_PRODUCTION']['L1'] (отображение: «Бочки/Еврокубы»).
# Смена 08–20 / 20–08 на фронте; старт/конец скана = первый/последний прирост.

TSD_LINE_ID = 'TSD_PRODUCTION'
TSD_ENTITY_TYPE = 'L1'  # попадает в PLAN_TYPES / план–факт

def _tsd_state_path():
    env = (os.environ.get('TSD_STATE_FILE') or '').strip()
    if env:
        return env
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tsd_state.json')


TSD_STATE_FILE = _tsd_state_path()

_tsd_fact_lock = threading.Lock()
_tsd_state_io_lock = threading.Lock()  # сериализация записи tsd_state.json (WinError 32)
_tsd_fact_snapshot = {}  # (job_id, gtin) -> int QuantityFact (capped)
_tsd_seeded = False
_tsd_stats = {'polls': 0, 'emits': 0, 'last_ok': None, 'last_error': None, 'last_reconcile': None}
# Сканы «П» с датой в Comment не на эту смену: в факт не идут, видны в /api/tsd_events
_tsd_skipped = deque(maxlen=1000)


def _tsd_strict_plan_date():
    return os.environ.get('TSD_STRICT_PLAN_DATE', '1').strip().lower() not in ('0', 'false', 'no')


def _tsd_record_skipped(job_id, gtin, delta, comment, good_name, ts, why):
    _tsd_skipped.append({
        'ts': (ts or datetime.now()).isoformat(timespec='seconds'),
        'job_id': job_id,
        'gtin': _normalize_gtin(gtin) or gtin,
        'qty': int(delta or 0),
        'comment': (comment or '')[:200],
        'good': (good_name or '')[:200] or None,
        'why': why,
    })


def _tsd_load_env():
    base = os.path.dirname(os.path.abspath(__file__))
    _load_env_file_if_present(os.path.join(base, 'tsd.env'))
    _load_env_file_if_present(os.path.join(base, 'tsd_push.env'))
    # путь к state после чтения env (если задали TSD_STATE_FILE)
    global TSD_STATE_FILE
    TSD_STATE_FILE = _tsd_state_path()


def _tsd_poll_enabled():
    _tsd_load_env()
    return os.environ.get('TSD_POLL_ENABLED', '0').strip().lower() not in ('0', 'false', 'no', '')


def _tsd_cap_fact(fact, plan):
    """В ТМ нельзя просканировать больше плана задания — режем факт."""
    try:
        f = int(fact or 0)
    except (TypeError, ValueError):
        f = 0
    try:
        p = int(plan or 0)
    except (TypeError, ValueError):
        p = 0
    if f < 0:
        f = 0
    if p > 0 and f > p:
        return p
    return f


def _tsd_shift_fact_qty(now=None):
    """Сумма qty событий бочек за текущую смену (в памяти)."""
    now = now or datetime.now()
    shift_start = _current_shift_start(now)
    total = 0
    with data_lock:
        events = production_data[TSD_LINE_ID][TSD_ENTITY_TYPE]['events']
        for ev in events:
            ts = ev.get('ts')
            if not isinstance(ts, datetime):
                continue
            if ts.tzinfo is not None:
                ts = ts.replace(tzinfo=None)
            if shift_start <= ts <= now:
                total += int(ev.get('qty') or 0)
    return total


def _tsd_memory_qty_by_key(now=None):
    """
    Сумма qty по (job_id, gtin) за текущую смену — диагностика/UI.
    Сверку с абсолютным QuantityFact API делать через snapshot, не через эту сумму:
    иначе на границе смены старые задания «П» заново падают в «сегодня».
    """
    now = now or datetime.now()
    shift_start = _current_shift_start(now)
    out = {}
    with data_lock:
        events = production_data[TSD_LINE_ID][TSD_ENTITY_TYPE]['events']
        for ev in events:
            ts = ev.get('ts')
            if not isinstance(ts, datetime):
                continue
            if ts.tzinfo is not None:
                ts = ts.replace(tzinfo=None)
            if not (shift_start <= ts <= now):
                continue
            jid = str(ev.get('job_id') or '').strip()
            gtin = _normalize_gtin(ev.get('gtin')) or str(ev.get('gtin') or '').strip()
            if not jid or not gtin:
                key = ('', gtin or '?')
            else:
                key = (jid, gtin)
            out[key] = out.get(key, 0) + int(ev.get('qty') or 0)
    return out


def _tsd_drop_false_shift_catchup(job_dates, targets=None, now=None, comments_by_key=None):
    """
    Убрать из текущей смены ложный догон: события по старому заданию, у которых
    сумма qty за смену == полному QuantityFact API (типичный dump на границе смены).

    Не трогаем:
    - частичные приросты (сканы сегодня по вчерашнему Job.Date);
    - задания с датой плана в Comment на текущую смену (создание Job могло быть раньше).
    """
    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    shift_start = _current_shift_start(now)
    targets = targets or {}
    comments_by_key = comments_by_key or {}

    # ключи, которые выглядят как полный dump абсолютного факта в новую смену
    shift_sums = {}
    comments_from_events = {}
    with data_lock:
        events = production_data[TSD_LINE_ID][TSD_ENTITY_TYPE]['events']
        maxlen = getattr(events, 'maxlen', None)
        for ev in events:
            ts = ev.get('ts')
            if not isinstance(ts, datetime):
                continue
            if ts.tzinfo is not None:
                ts = ts.replace(tzinfo=None)
            if not (shift_start <= ts <= now):
                continue
            jid = str(ev.get('job_id') or '').strip()
            gtin = _normalize_gtin(ev.get('gtin')) or str(ev.get('gtin') or '').strip()
            if not jid or not gtin:
                continue
            key = (jid, gtin)
            try:
                q = int(ev.get('qty') or 0)
            except (TypeError, ValueError):
                q = 0
            shift_sums[key] = shift_sums.get(key, 0) + q
            if ev.get('comment') and key not in comments_from_events:
                comments_from_events[key] = ev.get('comment')

    drop_keys = set()
    for key, sq in shift_sums.items():
        jid, gtin = key
        jdt = job_dates.get(jid)
        target = int(targets.get(key, 0) or 0)
        if jdt is None or jdt >= shift_start:
            continue
        if target <= 0 or sq != target:
            continue
        comment = comments_by_key.get(key) or comments_from_events.get(key) or ''
        if _tsd_comment_matches_active_shift(comment, now):
            continue
        drop_keys.add(key)

    if not drop_keys:
        return 0, 0

    removed = 0
    removed_qty = 0
    with data_lock:
        events = production_data[TSD_LINE_ID][TSD_ENTITY_TYPE]['events']
        maxlen = getattr(events, 'maxlen', None)
        keep = deque(maxlen=maxlen)
        for ev in events:
            ts = ev.get('ts')
            if isinstance(ts, datetime) and ts.tzinfo is not None:
                ts = ts.replace(tzinfo=None)
            jid = str(ev.get('job_id') or '').strip()
            gtin = _normalize_gtin(ev.get('gtin')) or str(ev.get('gtin') or '').strip()
            key = (jid, gtin)
            in_shift = isinstance(ts, datetime) and shift_start <= ts <= now
            if in_shift and key in drop_keys:
                removed += 1
                try:
                    removed_qty += int(ev.get('qty') or 0)
                except (TypeError, ValueError):
                    pass
                continue
            keep.append(ev)
        production_data[TSD_LINE_ID][TSD_ENTITY_TYPE]['events'] = keep
    if removed:
        logging.warning(
            f'TSD: снято ложных догонов в текущей смене: '
            f'keys={len(drop_keys)} events={removed} qty={removed_qty:+d}'
        )
    return removed, removed_qty


def _tsd_drop_off_date_shift_events(now=None):
    """
    Снять из текущей смены события по заданиям «П», у которых дата в Comment
    не совпадает со сменой (уже попавшие до включения строгого режима).
    События без комментария не трогаем — по ним дату не определить.
    """
    if not _tsd_strict_plan_date():
        return 0
    now = now or datetime.now()
    shift_start = _current_shift_start(now)
    removed = []
    with data_lock:
        events = production_data[TSD_LINE_ID][TSD_ENTITY_TYPE]['events']
        keep = deque(maxlen=getattr(events, 'maxlen', None))
        for ev in events:
            ts = ev.get('ts')
            if isinstance(ts, datetime) and ts.tzinfo is not None:
                ts = ts.replace(tzinfo=None)
            comment = ev.get('comment') or ''
            if (
                isinstance(ts, datetime) and shift_start <= ts <= now
                and comment and not _tsd_comment_matches_active_shift(comment, now)
            ):
                removed.append((ev, ts))
                continue
            keep.append(ev)
        if removed:
            production_data[TSD_LINE_ID][TSD_ENTITY_TYPE]['events'] = keep
    for ev, ts in removed:
        _tsd_record_skipped(
            ev.get('job_id'), ev.get('gtin'), ev.get('qty'), ev.get('comment'),
            ev.get('good'), ts, 'снято из смены: дата в комментарии не совпадает',
        )
    if removed:
        logging.warning(
            f'TSD: снято из смены событий с чужой датой плана: {len(removed)} '
            f'qty={sum(int(e.get("qty") or 0) for e, _ in removed):+d}'
        )
    return len(removed)


def _tsd_save_state():
    """Снимок QuantityFact + события линии — чтобы после рестарта факт не обнулялся."""
    path = _tsd_state_path()
    with _tsd_fact_lock:
        snap = [
            {'job_id': j, 'gtin': g, 'fact': int(f)}
            for (j, g), f in _tsd_fact_snapshot.items()
        ]
    events_out = []
    with data_lock:
        for ev in production_data[TSD_LINE_ID][TSD_ENTITY_TYPE]['events']:
            ts = ev.get('ts')
            if isinstance(ts, datetime):
                if ts.tzinfo is not None:
                    ts = ts.replace(tzinfo=None)
                ts_s = ts.isoformat(timespec='seconds')
            else:
                ts_s = str(ts) if ts is not None else None
            events_out.append({
                'ts': ts_s,
                'qty': ev.get('qty'),
                'gtin': ev.get('gtin'),
                'job_id': ev.get('job_id'),
                'comment': ev.get('comment'),
                'good': ev.get('good'),
                'source': ev.get('source') or 'TSD',
            })
    payload = {
        'saved_at': datetime.now().isoformat(timespec='seconds'),
        'snapshot': snap,
        'events': events_out,
    }
    # Один писатель + уникальный tmp: иначе на Windows os.replace → WinError 32
    with _tsd_state_io_lock:
        try:
            _atomic_write_json(path, payload)
        except Exception as e:
            logging.warning(f'TSD state save {path}: {e}')


def _tsd_load_state():
    """Восстановить snapshot и события бочек из tsd_state.json. True если файл прочитан."""
    global _tsd_seeded, TSD_STATE_FILE
    _tsd_load_env()
    TSD_STATE_FILE = _tsd_state_path()
    path = TSD_STATE_FILE
    if not os.path.isfile(path):
        return False
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f) or {}
    except Exception as e:
        logging.warning(f'TSD state load {path}: {e}')
        return False

    snap_in = data.get('snapshot') or []
    restored_snap = {}
    if isinstance(snap_in, list):
        for row in snap_in:
            if not isinstance(row, dict):
                continue
            j = str(row.get('job_id') or '').strip()
            g = _normalize_gtin(row.get('gtin')) or str(row.get('gtin') or '').strip()
            if not j or not g:
                continue
            try:
                restored_snap[(j, g)] = int(row.get('fact') or 0)
            except (TypeError, ValueError):
                continue

    events_in = data.get('events') or []
    n_ev = 0
    rebuilt_from_events = {}
    if isinstance(events_in, list):
        with data_lock:
            existing = production_data[TSD_LINE_ID][TSD_ENTITY_TYPE]['events']
            if not existing:
                for row in events_in:
                    if not isinstance(row, dict):
                        continue
                    try:
                        qty = int(row.get('qty') or 0)
                    except (TypeError, ValueError):
                        continue
                    if qty == 0:
                        continue
                    gtin = _normalize_gtin(row.get('gtin')) or str(row.get('gtin') or '').strip()
                    if not gtin:
                        continue
                    ts = to_naive_dt(row.get('ts')) or _iso_to_dt(row.get('ts'))
                    if ts is None:
                        continue
                    jid = str(row.get('job_id') or '').strip()
                    if jid and gtin:
                        rebuilt_from_events[(jid, gtin)] = (
                            int(rebuilt_from_events.get((jid, gtin), 0)) + qty
                        )
                    _append_event(
                        TSD_LINE_ID, TSD_ENTITY_TYPE, qty, gtin, ts=ts,
                        source=row.get('source') or 'TSD',
                        meta={
                            'job_id': row.get('job_id'),
                            'comment': row.get('comment'),
                            'good': row.get('good'),
                            'restored': True,
                        },
                    )
                    n_ev += 1

    # События без snapshot → иначе первый poll заново эмитит весь QuantityFact
    if n_ev and not restored_snap and rebuilt_from_events:
        for key, net in rebuilt_from_events.items():
            if net > 0:
                restored_snap[key] = net
        logging.warning(
            f'TSD state: snapshot был пуст, восстановлен из events '
            f'({len(restored_snap)} keys, file={path})'
        )

    with _tsd_fact_lock:
        _tsd_fact_snapshot.update(restored_snap)
    if restored_snap or n_ev:
        _tsd_seeded = True
    logging.info(
        f'TSD state loaded: snapshot={len(restored_snap)} events={n_ev} file={path}'
    )
    return True


def _tsd_comment_is_production(comment):
    """Префикс производства: кириллическая «П» в начале Comment (не латинская P — иначе ловит «Pr C…»)."""
    c = (comment or '').lstrip()
    return c.startswith('П')


# Линия TSD_PRODUCTION = крупная тара (>45 кг/л). Comment «П» этого не гарантирует.
TSD_MIN_PACK_KG = float(os.environ.get('TSD_MIN_PACK_KG', '45') or '45')
_TSD_PACK_DENY_RE = re.compile(r'канистр', re.IGNORECASE)
_TSD_PACK_ALLOW_RE = re.compile(
    r'бочк|еврокуб|оборотн\w*\s+куб|\bкуб\b',
    re.IGNORECASE,
)
# «220 кг», «1000 л», «1,7 кг» — число сразу перед ед. измерения
_TSD_PACK_SIZE_RE = re.compile(
    r'(?<![\d./])(\d+(?:[.,]\d+)?)\s*(?:кг|л|kg|l)\b',
    re.IGNORECASE,
)
# В Comment плана: «П Pr C желт (220) - 24 10/09»
_TSD_PACK_COMMENT_SIZE_RE = re.compile(r'\((\d+(?:[.,]\d+)?)\)')


def _tsd_parse_pack_size(raw):
    """Число из «220» / «1,7» → float или None."""
    if raw is None:
        return None
    s = str(raw).strip().replace(',', '.')
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _tsd_text_has_large_pack(text, min_kg=None):
    """True, если в тексте есть размер тары > min_kg (кг или л)."""
    threshold = TSD_MIN_PACK_KG if min_kg is None else float(min_kg)
    for m in _TSD_PACK_SIZE_RE.finditer(text or ''):
        val = _tsd_parse_pack_size(m.group(1))
        if val is not None and val > threshold:
            return True
    return False


def _tsd_comment_has_large_pack(comment, min_kg=None):
    """True, если в Comment есть (N) с N > min_kg (типичный размер бочки/куба)."""
    threshold = TSD_MIN_PACK_KG if min_kg is None else float(min_kg)
    c = comment or ''
    if _TSD_PACK_DENY_RE.search(c):
        return False
    for m in _TSD_PACK_COMMENT_SIZE_RE.finditer(c):
        val = _tsd_parse_pack_size(m.group(1))
        if val is not None and val > threshold:
            return True
    return False


def _tsd_gtin_catalog_name(gtin=None):
    g = _normalize_gtin(gtin) or (str(gtin).strip() if gtin else '')
    if not g:
        return ''
    try:
        names = gtin_names or {}
    except NameError:
        names = {}
    return str(names.get(g) or names.get(g.lstrip('0')) or '').strip()


def _tsd_packaging_names(good_name=None, gtin=None):
    """Все доступные имена: Good из API + справочник (TM часто короче каталога)."""
    out = []
    seen = set()
    for raw in ((good_name or '').strip(), _tsd_gtin_catalog_name(gtin)):
        if not raw:
            continue
        key = raw.casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(raw)
    return out


def _tsd_packaging_name(good_name=None, gtin=None):
    """Первое доступное имя (для логов)."""
    names = _tsd_packaging_names(good_name=good_name, gtin=gtin)
    return names[0] if names else ''


def _tsd_name_is_barrel_or_ibc(name):
    """Вердикт по одной строке: True/False/None (пусто)."""
    name = (name or '').strip()
    if not name:
        return None
    if _TSD_PACK_DENY_RE.search(name):
        return False
    if _TSD_PACK_ALLOW_RE.search(name) or _tsd_text_has_large_pack(name):
        return True
    return False


def _tsd_is_barrel_or_ibc(good_name=None, gtin=None, comment=None):
    """
    True для крупной тары: бочка / еврокуб / любой размер > TSD_MIN_PACK_KG (кг или л).

    Смотрим Good и gtin_names вместе: в TotalMark Good бывает без «бочка»
    («… 220 кг»), а в «Кратком отчёте» — полное «бочка 220 кг».
    Явная «канистра» в любом имени → отказ.
    Запасной сигнал — (N) в Comment при N > порога (напр. (220)).
    """
    for name in _tsd_packaging_names(good_name=good_name, gtin=gtin):
        v = _tsd_name_is_barrel_or_ibc(name)
        if v is False:
            return False
        if v is True:
            return True
    if _tsd_comment_has_large_pack(comment):
        return True
    return False


def _tsd_parse_comment_plan_date(comment, year=None):
    """
    Дата плана из хвоста Comment: «… - 24 07/09» → date(year, 9, 7).
    Не путать с «50/50» / «60/40» в середине строки — якорь на конец.
    """
    c = (comment or '').strip()
    if not c:
        return None
    m = re.search(r'(?<!\d)(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\s*$', c)
    if not m:
        return None
    try:
        day = int(m.group(1))
        month = int(m.group(2))
    except (TypeError, ValueError):
        return None
    y_raw = m.group(3)
    now_y = (year if year is not None else datetime.now().year)
    if y_raw:
        try:
            y = int(y_raw)
        except (TypeError, ValueError):
            y = now_y
        if y < 100:
            y += 2000
    else:
        y = int(now_y)
    try:
        return datetime(y, month, day).date()
    except ValueError:
        return None


def _tsd_comment_matches_active_shift(comment, now=None):
    """Comment DD/MM попадает в календарные дни текущей смены (вкл. ночную через полночь)."""
    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    plan = _tsd_parse_comment_plan_date(comment, year=now.year)
    if plan is None:
        return False
    shift_start = _current_shift_start(now)
    d0 = shift_start.date()
    d1 = now.date()
    if d0 > d1:
        d0, d1 = d1, d0
    return d0 <= plan <= d1


def _tsd_parse_job_date(raw):
    if not raw:
        return None
    if isinstance(raw, datetime):
        return raw.replace(tzinfo=None) if raw.tzinfo else raw
    s = str(raw).strip()
    for fmt in ('%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
        try:
            return datetime.strptime(s[:19] if 'T' in s or ' ' in s else s[:10], fmt)
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(s.replace('Z', ''))
    except ValueError:
        return None


def _tsd_basic_auth_header():
    """Basic из env (готовая base64) или логин/пароль; из env TSD_USER / TSD_PASSWORD или TSD_AUTH_BASIC."""
    ready = (os.environ.get('TSD_AUTH_BASIC') or '').strip()
    if ready:
        if ready.lower().startswith('basic '):
            return ready
        return f'Basic {ready}'
    user = os.environ.get('TSD_USER', '')
    password = os.environ.get('TSD_PASSWORD', '')
    token = base64.b64encode(f'{user}:{password}'.encode('utf-8')).decode('ascii')
    return f'Basic {token}'


def _tsd_http_json(path, body, timeout=20):
    base = (os.environ.get('TSD_BASE_URL') or '').rstrip('/')
    url = f'{base}{path}'
    raw = json.dumps(body, ensure_ascii=False).encode('utf-8')
    req = urlrequest.Request(url, data=raw, method='POST')
    req.add_header('Authorization', _tsd_basic_auth_header())
    req.add_header('Content-Type', 'application/json;charset=utf-8')
    req.add_header('User-Agent', os.environ.get('TSD_USER_AGENT', '1C+Enterprise/8.3'))
    req.add_header('ApplicationVersion', os.environ.get('TSD_APP_VERSION', '0.7.09.1904'))
    req.add_header('Content-Language', 'ru_RU')
    req.add_header('Accept-Encoding', 'identity')
    with urlrequest.urlopen(req, timeout=timeout) as resp:
        data = resp.read()
        if not data:
            return {}
        return json.loads(data.decode('utf-8'))


def _tsd_list_jobs():
    org = (os.environ.get('TSD_ORG') or '').strip()
    wh = (os.environ.get('TSD_WAREHOUSE') or '').strip()
    data = _tsd_http_json('/hs/Honest/jobs/shipment', {
        'Organization': org,
        'Warehouse': wh,
    })
    jobs = data.get('Jobs') if isinstance(data, dict) else None
    return jobs if isinstance(jobs, list) else []


def _tsd_job_info(job_id):
    data = _tsd_http_json('/hs/Honest/job/shipment/info', {'Id': job_id})
    return data if isinstance(data, dict) else {}


def _tsd_job_relevant(job, now=None):
    """
    Открытое задание «П» из списка Jobs.

    Job.Date часто = день создания (за несколько дней до розлива), дата плана — в Comment
    («… 07/09»). Окно «смена−36ч» по Date отбрасывало живые задания после потери snapshot.
    """
    if not isinstance(job, dict):
        return False
    if not _tsd_comment_is_production(job.get('Comment')):
        return False
    job_id = str(job.get('Id') or '').strip()
    if not job_id:
        return False
    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    with _tsd_fact_lock:
        if any(k[0] == job_id for k in _tsd_fact_snapshot):
            return True
    try:
        max_age_days = int(os.environ.get('TSD_JOB_MAX_AGE_DAYS', '30') or '30')
    except (TypeError, ValueError):
        max_age_days = 30
    if max_age_days <= 0:
        return True
    past = (now - timedelta(days=max_age_days)).date()
    future = (now + timedelta(days=max_age_days)).date()
    plan_d = _tsd_parse_comment_plan_date(job.get('Comment') or '', year=now.year)
    if plan_d is not None and past <= plan_d <= future:
        return True
    job_dt = _tsd_parse_job_date(job.get('Date'))
    if job_dt is not None and past <= job_dt.date() <= future:
        return True
    # В списке Jobs без дат — не отбрасываем (API уже отфильтровал открытые)
    if job_dt is None and plan_d is None:
        return True
    return False


def _tsd_emit_delta(job_id, gtin, delta, comment, good_name=None, ts=None, persist=True):
    """Записать delta (+/−) в production_data и отдать в Socket.IO."""
    if not delta or not gtin:
        return
    if not _tsd_is_barrel_or_ibc(good_name=good_name, gtin=gtin, comment=comment):
        logging.info(
            f'TSD skip non-barrel/IBC: {int(delta):+d} gtin={gtin} '
            f'job={job_id} good={(good_name or "")[:80]!r}'
        )
        return
    now = ts or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    gtin_n = _normalize_gtin(gtin) or str(gtin).strip()
    with data_lock:
        _append_event(
            TSD_LINE_ID, TSD_ENTITY_TYPE, int(delta), gtin_n, ts=now,
            source='TSD',
            meta={
                'job_id': job_id,
                'comment': (comment or '')[:200],
                'good': (good_name or '')[:200] if good_name else None,
            },
        )
    payload = {
        'line_id': TSD_LINE_ID,
        'entity_type': TSD_ENTITY_TYPE,
        'quantity': int(delta),
        'timestamp': now.isoformat(),
        'queue': 'TSD',
        'gtin': sticky_line_gtin(TSD_LINE_ID, gtin_n, remember=(int(delta) > 0)),
        'aggregate_id': None,
    }
    if connected_clients:
        socket_emit('production_data', payload)
    _tsd_stats['emits'] = int(_tsd_stats.get('emits') or 0) + 1
    logging.info(
        f'TSD: {delta:+d} {TSD_LINE_ID}/{gtin_n} job={job_id} '
        f'comment={(comment or "")[:60]!r}'
    )
    if persist:
        _tsd_save_state()


def _tsd_refresh_sticky_from_shift(now=None):
    """
    Sticky GTIN бочек = продукт с макс. нетто > 0 за текущую смену.
    Тот же критерий, что у карточки и плашки Excel.
    """
    best = current_shift_fact_gtin(TSD_LINE_ID, now=now)
    if best:
        remember_line_gtin(TSD_LINE_ID, best)
    return best


def _tsd_reconcile_to_api(
    targets,
    comments_by_key,
    goods_by_key,
    now=None,
    prev_snapshot=None,
    fetch_failed_job_ids=None,
    jobs_seen_ids=None,
):
    """
    Сверка с API по абсолютному snapshot: delta = target − prev_snapshot.

    Важно:
    - Задание пропало из списка Jobs — НЕ эмитить −QuantityFact (тихий drop snapshot).
    - job/shipment/info упал — НЕ трогать ключи этого job (freeze), иначе следующий
      успешный poll заново вольёт весь QuantityFact в текущую смену.
    - Отрицательная delta только если задание в targets и QuantityFact реально упал.
    """
    now = now or datetime.now()
    if prev_snapshot is None:
        with _tsd_fact_lock:
            prev_snapshot = dict(_tsd_fact_snapshot)
    fetch_failed_job_ids = set(fetch_failed_job_ids or ())
    jobs_seen_ids = set(jobs_seen_ids or ())
    adjusted = 0
    dropped_absent = 0

    for key in set(targets) | set(prev_snapshot):
        job_id, gtin = key
        if job_id in fetch_failed_job_ids:
            continue
        target = int(targets.get(key, 0) or 0)
        have = int(prev_snapshot.get(key, 0) or 0)
        if key not in targets:
            if have:
                dropped_absent += have
            continue
        delta = target - have
        if delta == 0:
            continue
        comment = comments_by_key.get(key, '')
        if _tsd_strict_plan_date() and not _tsd_comment_matches_active_shift(comment, now):
            _tsd_record_skipped(
                job_id, gtin, delta, comment, goods_by_key.get(key), now,
                'дата в комментарии не совпадает со сменой',
            )
            logging.info(
                f'TSD skip off-date: {delta:+d} gtin={gtin} job={job_id} comment={comment[:60]!r}'
            )
            continue
        _tsd_emit_delta(
            job_id, gtin, delta,
            comments_by_key.get(key, ''),
            good_name=goods_by_key.get(key),
            ts=now,
            persist=False,
        )
        adjusted += delta

    if adjusted or dropped_absent or fetch_failed_job_ids:
        _tsd_stats['last_reconcile'] = (
            f'{datetime.now().isoformat(timespec="seconds")} adj={adjusted:+d} '
            f'absent_drop={dropped_absent} frozen_jobs={len(fetch_failed_job_ids)} '
            f'prev={sum(prev_snapshot.values())} api={sum(targets.values())}'
        )
        logging.info(
            f'TSD reconcile: snapshot→API adj={adjusted:+d} '
            f'absent_jobs_qty={dropped_absent} fetch_failed={len(fetch_failed_job_ids)} '
            f'(было {sum(prev_snapshot.values())}, цель {sum(targets.values())})'
        )
    return adjusted


def _tsd_build_next_snapshot(targets, prev_snapshot, fetch_failed_job_ids):
    """Новый snapshot = успешные targets + замороженные ключи failed info."""
    fetch_failed_job_ids = set(fetch_failed_job_ids or ())
    next_snap = dict(targets)
    for key, fact in (prev_snapshot or {}).items():
        if key[0] in fetch_failed_job_ids and key not in next_snap:
            next_snap[key] = int(fact or 0)
    return next_snap


def tsd_poll_once():
    """Один цикл: список jobs → info по «П» → сверка факта с API (с потолком плана)."""
    global _tsd_seeded
    now = datetime.now()
    with _tsd_fact_lock:
        prev_snapshot = dict(_tsd_fact_snapshot)

    jobs = _tsd_list_jobs()

    # Пустой список Jobs при непустом snapshot — чаще сбой API, чем «все закрыли»
    if prev_snapshot and not jobs:
        logging.warning(
            f'TSD: пустой Jobs при snapshot={len(prev_snapshot)} — пропуск цикла '
            f'(защита от повторного догона QuantityFact)'
        )
        _tsd_stats['last_error'] = 'empty_jobs_skip'
        return 0

    jobs_seen_ids = {
        str(j.get('Id') or '').strip()
        for j in jobs
        if isinstance(j, dict) and str(j.get('Id') or '').strip()
    }
    relevant = [j for j in jobs if _tsd_job_relevant(j, now=now)]
    targets = {}  # (job_id, gtin) -> capped fact
    comments_by_key = {}
    goods_by_key = {}
    job_dates = {}
    fetch_failed_job_ids = set()
    fetched_ok_job_ids = set()

    for job in relevant:
        job_id = str(job.get('Id') or '').strip()
        comment = job.get('Comment') or ''
        job_dates[job_id] = _tsd_parse_job_date(job.get('Date'))
        try:
            info = _tsd_job_info(job_id)
        except Exception as e:
            logging.warning(f'TSD job info {job_id}: {e}')
            fetch_failed_job_ids.add(job_id)
            continue
        fetched_ok_job_ids.add(job_id)
        items = info.get('Items') if isinstance(info, dict) else None
        if not isinstance(items, list):
            items = []
        comment = (info.get('Comment') if isinstance(info, dict) else None) or comment
        if not _tsd_comment_is_production(comment):
            continue
        skipped_pack = 0
        for item in items:
            if not isinstance(item, dict):
                continue
            gtin_raw = item.get('GTIN')
            gtin = _normalize_gtin(gtin_raw) or (str(gtin_raw).strip() if gtin_raw else '')
            if not gtin:
                continue
            good = item.get('Good')
            if not _tsd_is_barrel_or_ibc(good_name=good, gtin=gtin, comment=comment):
                skipped_pack += 1
                continue
            fact = _tsd_cap_fact(item.get('QuantityFact'), item.get('QuantityPlan'))
            key = (job_id, gtin)
            targets[key] = fact
            comments_by_key[key] = comment
            goods_by_key[key] = good
        if skipped_pack:
            logging.info(
                f'TSD: job={job_id} пропуск не-бочка/куб Items={skipped_pack} '
                f'comment={(comment or "")[:50]!r}'
            )

    # Все info упали, а snapshot был — не затираем и не эмитим
    if prev_snapshot and not targets and fetch_failed_job_ids and not fetched_ok_job_ids:
        logging.warning(
            f'TSD: все job info failed ({len(fetch_failed_job_ids)}) — пропуск reconcile/snapshot'
        )
        _tsd_stats['last_error'] = 'all_info_failed_skip'
        return 0

    # Снять уже попавший в смену ложный догон старых заданий
    _tsd_drop_false_shift_catchup(
        job_dates, targets=targets, now=now, comments_by_key=comments_by_key,
    )
    _tsd_drop_off_date_shift_events(now)

    if jobs and not relevant:
        logging.warning(
            f'TSD: Jobs={len(jobs)}, но relevant=0 — проверьте фильтр дат/Comment '
            f'(TSD_JOB_MAX_AGE_DAYS={os.environ.get("TSD_JOB_MAX_AGE_DAYS", "30")})'
        )

    if not _tsd_seeded and not prev_snapshot:
        # Холодный старт без state: только snapshot, без emit в текущую смену
        with _tsd_fact_lock:
            _tsd_fact_snapshot.clear()
            _tsd_fact_snapshot.update(targets)
        _tsd_seeded = True
        logging.info(
            f'TSD: seed snapshot only (no emit), jobs={len(jobs)} relevant={len(relevant)} '
            f'keys={len(targets)} api_fact={sum(targets.values())}'
        )
    else:
        # Snapshot потерян (absent_drop / сбой), а процесс уже seeded:
        # не валить весь QuantityFact в смену. База = текущий API-факт,
        # кроме заданий с датой плана в Comment на эту смену — их догоняем в UI.
        reconcile_prev = prev_snapshot
        if _tsd_seeded and not prev_snapshot and targets:
            reconcile_prev = {}
            today_keys = 0
            for key, fact in targets.items():
                comment = comments_by_key.get(key, '')
                if fact > 0 and _tsd_comment_matches_active_shift(comment, now):
                    reconcile_prev[key] = 0
                    today_keys += 1
                else:
                    reconcile_prev[key] = int(fact or 0)
            logging.warning(
                f'TSD: пустой snapshot при seeded — baseline без истории; '
                f'догон плана смены keys={today_keys}/{len(targets)} '
                f'api_fact={sum(targets.values())}'
            )

        _tsd_reconcile_to_api(
            targets, comments_by_key, goods_by_key,
            now=now, prev_snapshot=reconcile_prev,
            fetch_failed_job_ids=fetch_failed_job_ids,
            jobs_seen_ids=jobs_seen_ids,
        )
        next_snap = _tsd_build_next_snapshot(targets, reconcile_prev, fetch_failed_job_ids)
        with _tsd_fact_lock:
            _tsd_fact_snapshot.clear()
            _tsd_fact_snapshot.update(next_snap)
        if not _tsd_seeded:
            _tsd_seeded = True
            logging.info(
                f'TSD: seed ok, jobs={len(jobs)} relevant={len(relevant)} '
                f'keys={len(targets)} api_fact={sum(targets.values())} '
                f'mem_shift={_tsd_shift_fact_qty(now)}'
            )

    sticky = _tsd_refresh_sticky_from_shift(now)
    if sticky:
        try:
            sync_excel_production_plans(force_reload=False)
        except Exception as e:
            logging.debug(f'TSD sticky→excel_plan: {e}')

    _tsd_save_state()
    _tsd_stats['polls'] = int(_tsd_stats.get('polls') or 0) + 1
    _tsd_stats['last_ok'] = now.isoformat(timespec='seconds')
    _tsd_stats['last_error'] = None
    return len(relevant)



def tsd_poll_loop():
    interval = float(os.environ.get('TSD_POLL_SEC', '5'))
    interval = max(3.0, interval)
    logging.info(
        f'TSD poll: старт → {(os.environ.get("TSD_BASE_URL") or "(TSD_BASE_URL not set)").rstrip("/")} '
        f'каждые {interval:.0f} с, линия={TSD_LINE_ID}, префикс Comment=П'
    )
    try:
        _tsd_load_state()
    except Exception as e:
        logging.warning(f'TSD load state: {e}')
    fail_streak = 0
    while True:
        try:
            tsd_poll_once()
            if fail_streak:
                logging.info('TSD poll: снова ок')
            fail_streak = 0
        except HTTPError as e:
            fail_streak += 1
            _tsd_stats['last_error'] = f'HTTP {e.code}'
            if fail_streak in (1, 5, 20) or fail_streak % 60 == 0:
                logging.warning(f'TSD poll HTTP {e.code} #{fail_streak}')
        except URLError as e:
            fail_streak += 1
            _tsd_stats['last_error'] = str(e.reason)
            if fail_streak in (1, 5, 20) or fail_streak % 60 == 0:
                logging.warning(f'TSD poll URLError: {e.reason} #{fail_streak}')
        except Exception as e:
            fail_streak += 1
            _tsd_stats['last_error'] = str(e)
            logging.error(f'TSD poll loop: {e}', exc_info=True)
        time.sleep(interval)


@app.route('/api/lines_live')
def api_lines_live():
    """Снимок линий/продуктов для внешней панели app_gui.py."""
    return jsonify(build_lines_live())


@app.route('/api/frontend_snapshot')
def api_frontend_snapshot():
    """Полный снимок состояния фронта (для VPS / отладки)."""
    snap = build_frontend_snapshot()
    snap.pop('_local_image_paths', None)
    return jsonify(snap)


@app.route('/api/tsd_events')
def api_tsd_events():
    """Сканы бочек (ТСД) за текущую смену: засчитанные и отброшенные, с заданием и комментарием."""
    now = datetime.now()
    shift_start = _current_shift_start(now)

    def _name(g):
        g = _normalize_gtin(g) or ''
        return gtin_names.get(g) or gtin_names.get(g.lstrip('0')) or None

    counted = []
    with data_lock:
        for ev in production_data[TSD_LINE_ID][TSD_ENTITY_TYPE]['events']:
            ts = ev.get('ts')
            if isinstance(ts, datetime) and ts.tzinfo is not None:
                ts = ts.replace(tzinfo=None)
            if not (isinstance(ts, datetime) and shift_start <= ts <= now):
                continue
            counted.append({
                'ts': ts.isoformat(timespec='seconds'),
                'qty': ev.get('qty'),
                'gtin': ev.get('gtin'),
                'name': ev.get('good') or _name(ev.get('gtin')),
                'job_id': ev.get('job_id'),
                'comment': ev.get('comment'),
            })
    shift_iso = shift_start.isoformat(timespec='seconds')
    skipped = [
        dict(s, name=s.get('good') or _name(s.get('gtin')))
        for s in list(_tsd_skipped)
        if (s.get('ts') or '') >= shift_iso
    ]
    return jsonify({
        'shift_start': shift_iso,
        'now': now.isoformat(timespec='seconds'),
        'strict_plan_date': _tsd_strict_plan_date(),
        'counted_qty': sum(int(e.get('qty') or 0) for e in counted),
        'counted': counted,
        'skipped_qty': sum(int(s.get('qty') or 0) for s in skipped),
        'skipped': skipped,
    })


@app.route('/api/shift_crew')
def api_shift_crew():
    """Текущая бригада/мастер по квартальному графику."""
    try:
        return jsonify(shift_schedule.get_crew_for_datetime())
    except Exception as e:
        logging.error(f'shift_crew: {e}', exc_info=True)
        return jsonify({'ok': False, 'message': str(e)}), 500


@app.route('/api/excel_plan')
def api_excel_plan():
    """Суточный план из Excel + сверка продуктов."""
    try:
        last = excel_plan_fact_gtins()
        state = production_plan.build_excel_plan_state(
            day=datetime.now().date(),
            last_gtin_by_line=last,
            gtin_names=gtin_names,
            line_display=LINE_DISPLAY_NAMES,
            allow_io=False,
            all_fact_gtins_by_line=excel_plan_fact_gtins_all(),
        )
        return jsonify(state)
    except Exception as e:
        logging.error(f'excel_plan: {e}', exc_info=True)
        return jsonify({'ok': False, 'message': str(e)}), 500


@app.route('/api/hourly/<line_id>/<entity_type>')
def api_hourly(line_id, entity_type):
    hourly = compute_hourly(line_id, entity_type)
    return jsonify(hourly)

@app.route('/api/generate_report', methods=['POST'])
def generate_report_endpoint():
    try:
        data = request.json or {}
        shift_date_str = data.get('shift_date')
        shift_type = data.get('shift_type')
        if shift_date_str:
            try:
                shift_date = datetime.strptime(shift_date_str, '%Y-%m-%d').date()
            except ValueError:
                return jsonify({'status': 'error', 'message': 'Invalid date format, use YYYY-MM-DD'}), 400
        else:
            shift_date = None
        logging.info(f"Запрос на генерацию отчета: дата={shift_date}, смена={shift_type}")
        # tpool: не блокируем eventlet/Socket.IO на время openpyxl+matplotlib
        result = run_generate_shift_report(shift_date, shift_type, do_reconcile=True)
        if result:
            return jsonify({'status': 'ok', 'file': os.path.basename(result)})
        return jsonify({'status': 'error', 'message': 'No data for this shift'}), 404
    except ReportBusyError:
        return jsonify({'status': 'error', 'message': 'Отчёт уже формируется, повторите через минуту'}), 409
    except Exception as e:
        logging.error(f"Ошибка генерации отчета: {e}", exc_info=True)
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ============================
# WEBSOCKET СОБЫТИЯ
# ============================
@socketio.on('connect')
def handle_connect():
    client_id = request.sid
    connected_clients.add(client_id)
    logging.info(f"✅ Клиент подключен: {client_id} (Всего: {len(connected_clients)})")
    snap = build_frontend_snapshot()
    n_dm = sum(len(types.get('DM') or []) for types in (snap.get('data') or {}).values())
    logging.info(
        f"init_data: cutoff={snap.get('shift_start')} "
        f"DM_events={n_dm} lines={len(snap.get('data') or {})}"
    )
    emit('init_data', {'data': snap.get('data') or {}})
    emit('plan_update', snap.get('plans') or {})
    emit('downtime_update', snap.get('downtime') or {})
    if snap.get('excel_plan'):
        emit('excel_plan', snap['excel_plan'])

@socketio.on('disconnect')
def handle_disconnect():
    client_id = request.sid
    if client_id in connected_clients:
        connected_clients.remove(client_id)
    logging.info(f"❌ Клиент отключен: {client_id} (Всего: {len(connected_clients)})")

# ============================
# ЗАПУСК
# ============================
if __name__ == '__main__':
    _ensure_single_instance()
    cleanup_old_reports(30)
    cleanup_thread = threading.Thread(target=cleanup_loop, daemon=True)
    cleanup_thread.start()

    load_gtin_names()
    load_poller_state()
    _init_db_pool()

    try:
        shift_schedule.warm_current_async()
    except Exception as e:
        logging.warning(f'shift schedule warm: {e}')

    threading.Thread(
        target=excel_plan_sync_loop, daemon=True, name='excel_plan_sync'
    ).start()

    # Догоняющие отчёты — в фоне, чтобы не блокировать старт HTTP
    if os.environ.get('SKIP_CATCH_UP', '0') != '1':
        threading.Thread(
            target=lambda: catch_up_missed_reports(),
            name='catch_up_reports',
            daemon=True,
        ).start()
    else:
        logging.info('SKIP_CATCH_UP=1 — пропуск догоняющих отчётов')

    os.makedirs('templates', exist_ok=True)

    # Граница смены 08:00/20:00: закрытие открытых простоев (+ сброс MAX, если включён).
    # Поток нужен всегда — не только при MAX_ENABLED.
    try:
        reconcile_stale_open_downtimes()
    except Exception as e:
        logging.warning(f'Простои reconcile при старте: {e}')
    threading.Thread(
        target=max_shift_alert_reset_loop, daemon=True, name='shift_boundary'
    ).start()
    logging.info('Shift boundary: поток 08:00/20:00 (закрытие простоев + сброс MAX-алертов)')

    try:
        from max_bot import is_enabled as max_enabled
        if max_enabled():
            from max_bot.client import get_me
            from max_bot.worker import start as max_worker_start
            max_worker_start()
            threading.Thread(
                target=max_reminder_loop, daemon=True, name='max_reminders'
            ).start()
            try:
                me = get_me()
                from max_bot.config import alert_schedule_summary
                logging.info(
                    f'MAX alerts: включены, бот={me.get("name")} (@{me.get("username")}), '
                    f'только DM текущей смены; N={alert_schedule_summary()}; сброс 08:00/20:00'
                )
            except Exception as e:
                logging.warning(
                    f'MAX alerts: включены, но GET /me не удался ({e}). '
                    f'Алерты всё равно активны. Расписание: см. max_bot/alert_schedule.example.env. '
                    f'При SSL: pip install certifi или MAX_SSL_VERIFY=0 в max_bot.local.env'
                )
        else:
            logging.info('MAX alerts: выключены (MAX_ENABLED=0 или нет токена)')
    except Exception as e:
        logging.warning(f'MAX alerts: проверка при старте не удалась: {e}')

    start_db_poller(reason='boot')

    # RABBIT_ENABLED=0 для тестового запуска рядом с рабочим: consume с auto_ack
    # забирает сообщения у боевого экземпляра, и тот теряет дезагрегации
    if os.environ.get('RABBIT_ENABLED', '1').strip() not in ('0', 'false', 'no'):
        rabbitmq_thread = threading.Thread(target=rabbitmq_listener, daemon=True, name='rabbitmq')
        rabbitmq_thread.start()
    else:
        logging.warning('RABBIT_ENABLED=0 — слушатель RabbitMQ не запущен (дезагрегации только из SQL-сверки)')

    reconcile_thread = threading.Thread(target=reconcile_loop, daemon=True, name='reconcile')
    reconcile_thread.start()

    scheduler_thread = threading.Thread(target=schedule_next_report, daemon=True, name='report_sched')
    scheduler_thread.start()

    plan_reset_thread = threading.Thread(target=schedule_plan_reset, daemon=True, name='plan_reset')
    plan_reset_thread.start()

    watchdog_thread = threading.Thread(target=health_watchdog_loop, daemon=True, name='watchdog')
    watchdog_thread.start()

    if _vps_push_enabled():
        threading.Thread(target=vps_push_loop, daemon=True, name='vps_push').start()
    else:
        logging.info('VPS push: выкл (VPS_PUSH_ENABLED=0 или нет URL/токена)')

    if _tsd_poll_enabled():
        threading.Thread(target=tsd_poll_loop, daemon=True, name='tsd_poll').start()
    else:
        logging.info('TSD poll: выкл (TSD_POLL_ENABLED=0)')

    logging.info(
        f"Старт: async_mode={SOCKETIO_ASYNC_MODE}  "
        f"health=/api/health  pool=on"
    )

    try:
        socketio.run(
            app,
            host='0.0.0.0',
            port=5000,
            debug=False,
            use_reloader=False,
            log_output=True,
            allow_unsafe_werkzeug=True,
        )
    except KeyboardInterrupt:
        print("\n👋 Завершение работы...")
    except Exception as e:
        print(f"❌ Ошибка запуска сервера: {e}")