# -*- coding: utf-8 -*-
"""
Исходящие алерты в MAX:
  — тишина по последнему L1/ТУ/DM (как карточка/GUI) ≥ MAX_ALERT_FIRST_MIN (10 мин);
  — в смене должны быть DM; N в тексте: 10 → 20 → 50 → 80…;
  — не шлём с открытия карточки на жёлтом пороге (3 мин) — только reminder loop;
  — не шлём L2; сброс 08:00 / 20:00; план ≥100% — не алерт.
"""
import logging
import threading
from datetime import datetime

from max_bot import config, worker
from max_bot.client import send_message

log = logging.getLogger(__name__)

_dedup_lock = threading.Lock()
_reminder_milestone = {}
_reminder_started = {}
_open_sent = {}
# После кнопки причины: стоп «Задержка…», ждём только возобновление
_reason_acked = {}  # line_id -> reason str
_reason_sent = {}   # line_id -> reason (дедуп однократного алерта)


def _parse_started(started_at):
    if isinstance(started_at, datetime):
        ts = started_at
    else:
        try:
            raw = str(started_at).strip()
            if raw.endswith('Z'):
                raw = raw[:-1] + '+00:00'
            ts = datetime.fromisoformat(raw)
        except Exception:
            return datetime.now()
    if getattr(ts, 'tzinfo', None) is not None:
        ts = ts.replace(tzinfo=None)
    return ts


def format_started(started_at):
    return _parse_started(started_at).strftime('%d.%m.%Y %H:%M:%S')


def alert_first_min():
    try:
        return max(1, int(getattr(config, 'MAX_ALERT_FIRST_MIN', 10) or 10))
    except (TypeError, ValueError):
        return 10


def alert_second_min():
    try:
        second = int(getattr(config, 'MAX_ALERT_SECOND_MIN', 20) or 20)
    except (TypeError, ValueError):
        second = 20
    return max(alert_first_min(), second)


def alert_late_every_min():
    try:
        return max(1, int(getattr(config, 'MAX_ALERT_LATE_EVERY_MIN', 30) or 30))
    except (TypeError, ValueError):
        return 30


# Совместимость со старыми именами
def reminder_early_every_min():
    return alert_first_min()


def reminder_early_until_min():
    return alert_second_min()


def reminder_late_every_min():
    return alert_late_every_min()


def reminder_every_min():
    return alert_first_min()


def milestone_for_duration(dur_min):
    """
    Отметка N для текста алерта по длительности тишины (мин).
    По умолчанию: 10 → 20 → 50 → 80 → 110 … (см. MAX_ALERT_* в config).
    """
    try:
        dur = float(dur_min)
    except (TypeError, ValueError):
        return 0

    first = alert_first_min()
    second = alert_second_min()
    late = alert_late_every_min()

    if dur < first:
        return 0
    if dur < second:
        return first
    if dur < second + late:
        return second
    extra = int((dur - second) // late)
    return second + extra * late


def _fmt_min(n):
    try:
        return str(int(round(float(n))))
    except (TypeError, ValueError):
        return str(n)


def stop_text(display_name, line_id=None, started_at=None, duration_min=None):
    """Задержка на линии {имя как во фронте} N минут."""
    name = display_name or line_id or '—'
    n = _fmt_min(duration_min if duration_min is not None else alert_first_min())
    return f'Задержка на линии {name} {n} минут.'


def resume_text(display_name, line_id=None, duration_min=None):
    """Работа на линии {имя} возобновлена, общее время паузы N минут."""
    name = display_name or line_id or '—'
    n = _fmt_min(duration_min if duration_min is not None else 0)
    return f'Работа на линии {name} возобновлена, общее время паузы {n} минут.'


alert_text = stop_text


def _send_to_chats(text):
    chats = config.alert_chat_ids()
    if not chats:
        log.warning('MAX: MAX_ALERT_CHAT_IDS пуст — алерт некуда отправить')
        return
    for chat_id in chats:
        try:
            send_message(text, chat_id=chat_id)
            log.info(f'MAX: алерт отправлен chat_id={chat_id}')
        except Exception as e:
            log.error(f'MAX: не удалось отправить в {chat_id}: {e}')


def reason_text(display_name, reason):
    name = display_name or '—'
    return f'{name} в простое по причине: {reason}'


def is_reason_acked(line_id):
    with _dedup_lock:
        return line_id in _reason_acked


def forget_open(line_id, started_at=None):
    with _dedup_lock:
        _reminder_milestone.pop(line_id, None)
        _reminder_started.pop(line_id, None)
        _open_sent.pop(line_id, None)
        _reason_acked.pop(line_id, None)
        _reason_sent.pop(line_id, None)


def reset_all_alerts(reason='shift_boundary'):
    """Сброс дедупа на границе смены (08:00 / 20:00)."""
    with _dedup_lock:
        n = len(_open_sent) + len(_reminder_started) + len(_reason_acked)
        _reminder_milestone.clear()
        _reminder_started.clear()
        _open_sent.clear()
        _reason_acked.clear()
        _reason_sent.clear()
    log.info(f'MAX: сброс всех алертов ({reason}), сброшено ключей≈{n}')
    return n


def notify_line_reason(line_id, display_name, reason):
    """
    Однократно: «{линия} в простое по причине: {причина}».
    Останавливает дальнейшие алерты задержки до возобновления / сброса смены.
    """
    if not config.is_enabled():
        return False
    reason = (reason or '').strip()
    if not line_id or not reason:
        return False
    worker.start()
    with _dedup_lock:
        if _reason_sent.get(line_id) == reason and line_id in _reason_acked:
            log.debug(f'MAX: причина уже отправлена {line_id} {reason}')
            return False
        _reason_sent[line_id] = reason
        _reason_acked[line_id] = reason
        # чтобы resume сработал даже без предшествующей «Задержка»
        if line_id not in _open_sent:
            _open_sent[line_id] = f'reason:{reason}'
            _reminder_started[line_id] = _open_sent[line_id]
    text = reason_text(display_name or line_id, reason)
    worker.enqueue(_send_to_chats, text)
    log.info(f'MAX: причина простоя {line_id}: {reason}')
    return True


def notify_line_stop(line_id, display_name, started_at, duration_min=None):
    """Первый алерт. duration_min — фактическая тишина; иначе подставляется FIRST (не вызывать с карточки на 3 мин)."""
    if not config.is_enabled():
        return
    with _dedup_lock:
        if line_id in _reason_acked:
            return
    worker.start()
    key = format_started(started_at)
    n = int(duration_min) if duration_min is not None else alert_first_min()
    with _dedup_lock:
        if line_id in _reason_acked:
            return
        if _open_sent.get(line_id) == key and int(_reminder_milestone.get(line_id) or 0) >= n:
            return
        _reminder_started[line_id] = key
        _reminder_milestone[line_id] = n
        _open_sent[line_id] = key
    text = stop_text(display_name or line_id, line_id, started_at, duration_min=n)
    worker.enqueue(_send_to_chats, text)
    log.info(f'MAX: задержка {line_id} N={n} start={key}')


def notify_line_still_down(line_id, display_name, started_at, duration_min, milestone):
    """Алерт на отметке milestone минут (N в тексте = milestone)."""
    if not config.is_enabled():
        return False
    if milestone < 1:
        return False
    worker.start()
    key = format_started(started_at)
    with _dedup_lock:
        if line_id in _reason_acked:
            return False
        cur_start = _reminder_started.get(line_id)
        if cur_start and cur_start != key:
            _reminder_started[line_id] = key
            _reminder_milestone[line_id] = 0
            _open_sent.pop(line_id, None)
        last = int(_reminder_milestone.get(line_id) or 0)
        if milestone <= last:
            return False
        _reminder_milestone[line_id] = milestone
        _reminder_started[line_id] = key
        _open_sent[line_id] = key
    text = stop_text(display_name or line_id, line_id, started_at, duration_min=milestone)
    worker.enqueue(_send_to_chats, text)
    log.info(f'MAX: задержка {line_id} N={milestone}')
    return True


def check_open_reminders(open_map, display_names=None, now=None):
    """
    open_map: line_id -> {started_at, ...} только тишина DM текущей смены.
    После ACK причины — пропуск (ждём только возобновление).
    """
    if not config.is_enabled():
        return 0
    if not open_map:
        return 0
    now = now or datetime.now()
    if getattr(now, 'tzinfo', None) is not None:
        now = now.replace(tzinfo=None)
    display_names = display_names or {}
    sent_n = 0
    for line_id, rec in open_map.items():
        if not isinstance(rec, dict):
            continue
        with _dedup_lock:
            if line_id in _reason_acked:
                continue
        started = _parse_started(rec.get('started_at'))
        if started is None:
            continue
        dur_min = (now - started).total_seconds() / 60.0
        ms = milestone_for_duration(dur_min)
        if ms < 1:
            continue
        name = display_names.get(line_id, line_id)
        if notify_line_still_down(line_id, name, started, dur_min, ms):
            sent_n += 1
    return sent_n


def notify_line_resume(line_id, display_name, duration_min, started_at=None):
    if not config.is_enabled():
        return
    worker.start()
    forget_open(line_id, started_at)
    text = resume_text(display_name or line_id, line_id, duration_min)
    worker.enqueue(_send_to_chats, text)


def notify_line_resume_if_alerted(line_id, display_name, ended_at=None):
    if not config.is_enabled():
        return False
    ended = _parse_started(ended_at) if ended_at else datetime.now()
    with _dedup_lock:
        if (
            line_id not in _open_sent
            and line_id not in _reminder_started
            and line_id not in _reason_acked
        ):
            return False
        start_key = _reminder_started.get(line_id) or _open_sent.get(line_id)
        last_n = int(_reminder_milestone.get(line_id) or 0)
    dur = last_n if last_n > 0 else None
    if start_key and dur is None and not str(start_key).startswith('reason:'):
        try:
            st = datetime.strptime(start_key, '%d.%m.%Y %H:%M:%S')
            dur = round((ended - st).total_seconds() / 60.0)
        except Exception:
            dur = 0
    if dur is None:
        dur = 0
    notify_line_resume(line_id, display_name, dur, started_at=start_key)
    return True


def should_send(line_id, started_at):
    key = format_started(started_at)
    with _dedup_lock:
        return _open_sent.get(line_id) != key
