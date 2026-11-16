# TotalMark / TotalCore Shopfloor Monitor

Открытый reference-проект цехового монитора для производств на стеке **TotalMark / TotalCore** (Honest API, агрегация кодов, опционально RabbitMQ и PostgreSQL conveyor).

Показывает выпуск по линиям в реальном времени, простои, Excel-план смены, учёт крупной тары со **ТСД через Honest API**, Excel-отчёты и режимы «мастер» / «цех/ТВ».

> Это **обезличенная** публичная копия. В репозитории нет заводских логинов, паролей, tenant ID, внутренних IP и брендинга конкретного предприятия. Все подключения — через переменные окружения (см. `.env.example`).

## Для кого

- Интеграторы и IT цеха, у кого уже есть TotalMark / ТСД / агрегаторы.
- Разработчики и AI-агенты, которым нужна рабочая схема «как читать факт с Honest API и показать на мониторе».

## Быстрый старт (без боя)

```powershell
git clone https://github.com/Discreaminator/totalmark-shopfloor-monitor.git
cd totalmark-shopfloor-monitor
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt   # если файла нет — Flask, flask-socketio, openpyxl, psycopg2-binary, pika
copy .env.example .env            # заполните своими значениями

$env:RABBIT_ENABLED=0
$env:VPS_PUSH_ENABLED=0
$env:MAX_ENABLED=0
$env:TSD_POLL_ENABLED=0
python app.py
```

Откройте `http://127.0.0.1:5000`.

**Не** включайте локально `RABBIT_ENABLED` / `TSD_POLL_ENABLED` против боевого контура чужого или своего завода без понимания последствий — опрос ТСД и очереди заберут события у работающей системы.

## Документация

| Документ | Содержание |
|----------|------------|
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | Архитектура монитора, источники данных, линии, UI |
| [docs/CONNECTIONS.md](docs/CONNECTIONS.md) | Что к чему подключается (БД, Rabbit, Excel, ТСД, алерты) |
| [docs/TSD_HONEST_API.md](docs/TSD_HONEST_API.md) | Перехват/чтение факта ТСД через Honest API — правила фильтрации |
| [docs/ERP_BARRELS_FROM_TSD.md](docs/ERP_BARRELS_FROM_TSD.md) | ТЗ-шаблон: факт бочек/еврокубов из TotalMark в 1С:ERP |
| [README_MAX_BOT.md](README_MAX_BOT.md) | Опциональные алерты простоя в мессенджер MAX |
| [.env.example](.env.example) | Полный список переменных без секретов |

## Что обязательно настроить под себя

1. `TSD_BASE_URL`, `TSD_USER` / `TSD_PASSWORD` (или `TSD_AUTH_BASIC`), `TSD_ORG`, `TSD_WAREHOUSE` — ваш tenant Honest.
2. Линии в коде (`Ralot_L1`, `KRO_*`, `BLS`, `TSD_PRODUCTION`) — **пример** топологии цеха; переименуйте/сократите под свои SKU.
3. Excel-план: `PRODUCTION_PLAN_FILE` и маппинг в `production_plan.py`.
4. БД агрегатов / Rabbit — если используете TotalCore conveyor; иначе оставьте выключенными.

## Лицензия и ответственность

Код публикуется как reference «as is». Проверяйте доступы TotalMark у своего вендора/админа. Не коммитьте `.env`, `*.local.env`, `tsd_state.json`, отчёты и фото продукции.
