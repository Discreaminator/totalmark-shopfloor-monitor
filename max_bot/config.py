# -*- coding: utf-8 -*-
"""
Конфиг MAX-бота: токен, чаты, расписание алертов задержки.

Файлы (порядок загрузки, позже перекрывает раньше):
  1. max_bot.local.env       — секреты (токен, chat_id), не коммитить
  2. max_bot/alert_schedule.env — только тайминги алертов (опционально)
  3. переменные окружения процесса

Шаблон таймингов: max_bot/alert_schedule.example.env
"""
import os

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_LOCAL_ENV = os.path.join(_ROOT, 'max_bot.local.env')
_SCHEDULE_ENV = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'alert_schedule.env')


def _load_local_env(path):
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
    except Exception:
        pass


_load_local_env(_LOCAL_ENV)
_load_local_env(_SCHEDULE_ENV)

BASE_URL = os.environ.get('MAX_API_BASE', 'https://platform-api2.max.ru').rstrip('/')
HTTP_TIMEOUT = float(os.environ.get('MAX_HTTP_TIMEOUT', '8'))
# Минимальный интервал между любыми исходящими (антиспам HTTP), сек
ALERT_MIN_INTERVAL_SEC = int(os.environ.get('MAX_ALERT_MIN_INTERVAL_SEC', '180'))


def _int_env(name, default, minimum=1):
    try:
        return max(minimum, int(os.environ.get(name, default)))
    except (TypeError, ValueError):
        return max(minimum, int(default))


# =============================================================================
# Расписание алертов «Задержка на линии … N минут»
#
# Тишина должна длиться ≥ MAX_ALERT_FIRST_MIN (по последнему L1/ТУ/DM, как GUI).
# Текст сообщения: N = ближайшая достигнутая отметка:
#   FIRST → SECOND → SECOND+LATE → SECOND+2·LATE → …
#
# По умолчанию: 10 → 20 → 50 → 80 → 110 → 140 …
#   • 10 мин — первое оповещение
#   • 20 мин — второе
#   • далее каждые 30 мин (50, 80, 110…)
#
# Связанные переменные (app.py): MAX_REMINDER_POLL_SEC — период опроса (~30 с)
# =============================================================================
MAX_ALERT_FIRST_MIN = _int_env('MAX_ALERT_FIRST_MIN', 10)
MAX_ALERT_SECOND_MIN = max(MAX_ALERT_FIRST_MIN, _int_env('MAX_ALERT_SECOND_MIN', 20))
MAX_ALERT_LATE_EVERY_MIN = _int_env('MAX_ALERT_LATE_EVERY_MIN', 30)

# Устаревшие имена (оставлены для совместимости .env; не используются в логике)
REMINDER_EVERY_MIN = MAX_ALERT_FIRST_MIN
REMINDER_EARLY_EVERY_MIN = MAX_ALERT_FIRST_MIN
REMINDER_EARLY_UNTIL_MIN = MAX_ALERT_SECOND_MIN
REMINDER_LATE_EVERY_MIN = MAX_ALERT_LATE_EVERY_MIN


def _flag(name, default='0'):
    return os.environ.get(name, default).strip().lower() not in ('0', 'false', 'no', '')


def _id_list(name):
    raw = os.environ.get(name, '') or ''
    out = []
    for part in raw.replace(';', ',').split(','):
        part = part.strip()
        if not part:
            continue
        try:
            out.append(int(part))
        except ValueError:
            continue
    return out


def is_enabled():
    if not _flag('MAX_ENABLED', '0'):
        return False
    return bool((os.environ.get('MAX_BOT_TOKEN') or '').strip())


def token():
    return (os.environ.get('MAX_BOT_TOKEN') or '').strip()


def alert_chat_ids():
    return _id_list('MAX_ALERT_CHAT_IDS')


def alert_schedule_summary():
    """Краткое описание отметок N для логов и документации."""
    f, s, l = MAX_ALERT_FIRST_MIN, MAX_ALERT_SECOND_MIN, MAX_ALERT_LATE_EVERY_MIN
    return f'{f} → {s} → {s + l} → {s + 2 * l} → … (+{l} мин)'
