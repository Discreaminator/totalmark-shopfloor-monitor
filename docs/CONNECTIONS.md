# Что к чему подключается

Краткая карта для человека или агента: какие внешние системы нужны и какие переменные их включают.

```
                    ┌─────────────────────┐
   Excel план ─────►│                     │
   «Краткий отчет»─►│  Shopfloor Monitor  │◄── браузеры цеха / ТВ / мастер
                    │  (Flask+Socket.IO)  │
   PostgreSQL ─────►│                     │──► Excel-отчёты смен
   (агрегаты DM)    │                     │──► опц. MAX-алерты
   RabbitMQ ───────►│                     │
   Honest API ТСД ─►│                     │
                    └─────────────────────┘
```

## 1. Honest API (ТСД / задания отгрузки)

| | |
|--|--|
| Назначение | Факт сканирования крупной тары (бочки, еврокубы) |
| Протокол | HTTPS POST JSON, HTTP Basic |
| База | `TSD_BASE_URL` = `https://tm.totalmark.pro/<TENANT>` |
| Методы | `/hs/Honest/jobs/shipment`, `/hs/Honest/job/shipment/info` |
| Учётка | Выдаёт админ TotalMark. Учётка интеграции `api` часто **не** пускает на `Honest` — нужна отдельная учётка с правом чтения Honest. |
| Env | `TSD_POLL_ENABLED`, `TSD_BASE_URL`, `TSD_USER`, `TSD_PASSWORD` или `TSD_AUTH_BASIC`, `TSD_ORG`, `TSD_WAREHOUSE` |

Подробные правила фильтрации: [TSD_HONEST_API.md](TSD_HONEST_API.md).

## 2. PostgreSQL conveyor (агрегаты / DM)

| | |
|--|--|
| Назначение | Живой выпуск с линий маркировки (индивидуальные коды, агрегаты) |
| Env | `DB_HOST`, `DB_PORT`, `DB_NAME`, `DB_USER`, `DB_PASSWORD` |
| Замечание | Без БД монитор всё равно стартует; линии по DM будут пустыми, ТСД может работать отдельно. |

## 3. RabbitMQ TotalMark

| | |
|--|--|
| Назначение | События агрегации / дезагрегации (дополнение к SQL-поллеру) |
| Env | `RABBIT_ENABLED=1`, `RABBIT_HOST`, `RABBIT_USER`, `RABBIT_PASSWORD`, `RABBIT_QUEUES` |
| Опасность | Два потребителя на одну очередь = потеря сообщений. На бою — один инстанс. |

## 4. Excel

| Файл | Роль |
|------|------|
| План производства (`PRODUCTION_PLAN_FILE` или `data/План произ-ва*.xlsx`) | Суточный план → `plans[line].day` |
| `data/Краткий отчет.xlsx` | GTIN → полное имя (и косвенно фильтр тары ТСД) |
| `reports/report_YYYY-MM-DD.xlsx` | Сменные отчёты (генерирует монитор) |

Перечитать справочник GTIN без рестарта: `POST|GET /api/gtin_names/reload`.

## 5. Браузеры / режимы UI

| URL | Режим |
|-----|-------|
| `/` | Обычный дашборд |
| `/?master=1` | Режим мастера (причины простоев, журнал) |
| параметры kiosk / tv | Цеховой/ТВ экран «одна смена на одном экране» |

`/api/downtime` без отдельной auth — держите порт только в LAN цеха или за reverse-proxy с auth.

## 6. Опционально

- **MAX-бот** — алерты тишины DM: `max_bot.local.env` (не в git).
- **image_downloader.py** — картинки GTIN из национального каталога: `NK_API_KEY`.
- **SMTP / hang_alert_watcher** — внешний watchdog зависания поллера.

## Минимальный контур «только ТСД-бочки»

```
TSD_POLL_ENABLED=1
TSD_BASE_URL=https://tm.totalmark.pro/<TENANT>
TSD_USER=...
TSD_PASSWORD=...
TSD_ORG=<uuid organization>
TSD_WAREHOUSE=<uuid warehouse>
RABBIT_ENABLED=0
```

Линия `TSD_PRODUCTION` начнёт наполняться дельтами `QuantityFact`. Диагностика: `GET /api/tsd_events`.
