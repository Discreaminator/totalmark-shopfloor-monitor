# Архитектура Shopfloor Monitor

Flask + Socket.IO приложение. Один процесс обслуживает UI и фоновые поллеры.

## Назначение

1. Живой выпуск по линиям (агрегаты/DM из БД или Rabbit + крупная тара из ТСД).
2. План смены (ручной + Excel).
3. Простои / паузы, режим мастера (причины).
4. Excel-отчёт смены.
5. Диагностика ТСД и сверка продукта с планом.

## Пример топологии линий

Идентификаторы ниже — **пример** цеха розлива. Замените в `app.py` / UI / `production_plan.py` на свои.

| ID | Смысл |
|----|--------|
| `Ralot_L1` | Штучный выпуск линии A |
| `Ralot_L2` | Паллеты от A (`ceil(шт / на_паллете)`) |
| `KRO_BOTTLE_L1` / `L2` | Бутылка / паллеты |
| `KRO_KANISTER` | Канистры |
| `BLS` | Отдельная линия |
| `TSD_PRODUCTION` | Бочки/еврокубы со сканеров ТСД |

## Потоки данных

1. **SQL-поллер** (watermark в `poller_state.json`) читает новые агрегаты.
2. **Rabbit** (опционально) — события дезагрегации / агрегации.
3. **TSD-поллер** — Honest API → дельты в `TSD_PRODUCTION`.
4. **production_plan sync** — Excel → `plans[line].day` (с учётом `manual_lock`).
5. **Socket.IO** пушит `production_data`, `plan_update`, `gtin_names`, `downtime_update`.

## Важные файлы runtime (не в git)

| Файл | Назначение |
|------|------------|
| `tsd_state.json` | снимок факта ТСД |
| `plans.json` | ручные планы + lock |
| `downtime.json` | простои |
| `poller_state.json` | watermark БД |
| `data/*.xlsx` | план и краткий отчёт |
| `reports/` | сменные Excel |
| `.env` / `max_bot.local.env` | секреты |

Запись JSON/xlsx — атомарная (tmp + replace) + lockfiles, чтобы два процесса не портили ZIP/xlsx.

## API (основные)

| Метод | Назначение |
|-------|------------|
| `GET /api/health` | живость |
| `GET /api/gtin_names` | справочник |
| `POST\|GET /api/gtin_names/reload` | перечитать Excel без рестарта |
| `GET /api/tsd_events` | диагностика ТСД |
| downtime API | открытие/причина/ack простоев |

## Эксплуатация

1. Один инстанс на порт (или `ALLOW_MULTI_INSTANCE` только осознанно).
2. После деплоя кода — перезапуск.
3. Тесты не гонять против боевого Rabbit/TotalMark.
4. Новые GTIN: обновить «Краткий отчет» → `/api/gtin_names/reload`.
