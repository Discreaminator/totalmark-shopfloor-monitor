# Honest API: чтение факта ТСД (крупная тара)

Документ описывает проверенную на практике схему **только чтения** заданий ТСД TotalMark через сервис **Honest** — ту же, что использует линия `TSD_PRODUCTION` в этом мониторе и шаблон ТЗ для 1С:ERP.

Секреты (URL tenant, логин, пароль, UUID организации/склада) **не публикуются** — берите у своего администратора TotalMark.

## 1. Подключение

| Параметр | Как задать |
|----------|------------|
| Base URL | `https://tm.totalmark.pro/<ВАШ_TENANT_ID>` → env `TSD_BASE_URL` |
| Сервис | `Honest` (не путать с публикацией `api`) |
| Auth | HTTP Basic → `TSD_USER` + `TSD_PASSWORD` или готовый `TSD_AUTH_BASIC` |
| Content-Type | `application/json;charset=utf-8` |

Полезный заголовок User-Agent (как у 1С): `1C+Enterprise/8.3` (`TSD_USER_AGENT`).

**Типичная ловушка:** учётка, которая работает на `/hs/api/...`, на `/hs/Honest/...` отвечает **401**. Нужна учётка с правом именно на Honest (часто отдельная «ТСД»-роль). Значения логина/пароля — только в вашем `.env`.

Интеграция должна **только читать**. Ничего не закрывайте и не пишите в задания с монитора.

## 2. Методы

### 2.1. Список открытых заданий

`POST {TSD_BASE_URL}/hs/Honest/jobs/shipment`

```json
{
  "Organization": "<UUID организации>",
  "Warehouse": "<UUID склада>"
}
```

Ответ: массив `Jobs[]` с полями как минимум `Id`, `Date`, `Comment`, `Number`.

Важные свойства API:

- возвращаются **только открытые** задания;
- параметра периода **нет**;
- **закрытое задание пропадает** из списка — факт после закрытия этим методом уже не получить.

Поэтому нужен фоновый опрос (у монитора — минуты; для ERP — см. [ERP_BARRELS_FROM_TSD.md](ERP_BARRELS_FROM_TSD.md)).

### 2.2. Факт по заданию

`POST {TSD_BASE_URL}/hs/Honest/job/shipment/info`  
тело: `{"Id": "<UUID задания>"}`

В `Items[]`:

| Поле | Смысл |
|------|--------|
| `Good` | Наименование из TotalMark |
| `GTIN` | Код товара |
| `QuantityPlan` | План **задания** (не план ERP) |
| `QuantityFact` | Накопительный факт (1 скан крупной тары = +1) |

## 3. Правила учёта (обязательные фильтры)

Без фильтров в факт попадают складские задания, канистры и чужие даты. Ниже — правила, обкатанные на реальных ошибках операторов.

### 3.1. Только производственные задания

`Comment` после `lstrip` начинается с **кириллической «П»**.  
Латинская `P` не подходит (ловятся складские «Pr C…»).

### 3.2. Дата в хвосте Comment = дни текущей смены

- Смены по умолчанию: день 08:00–20:00, ночь 20:00–08:00.
- Дата в **конце** комментария: `DD/MM` или `DD/MM/YY(YY)`.
- Регулярка: `(?<!\d)(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\s*$`  
  (не путать с «50/50» в середине строки).
- Ночная смена: дата начала **или** окончания смены.
- Нет даты / чужая дата → не считать. В мониторе: `TSD_STRICT_PLAN_DATE=1`.

### 3.3. Только крупная тара (порог `TSD_MIN_PACK_KG`, по умолчанию 45)

Проверяются `Good`, имя из справочника GTIN и Comment:

1. Есть «канистр» → **отказ**.
2. Есть «бочк» / «еврокуб» / «куб» → учёт.
3. Есть «N кг» / «N л» с N > порога → учёт.
4. Запасной сигнал: `(N)` в Comment с N > порога.
5. Иначе — не учитывать.

### 3.4. Потолок по плану задания

```
учитываемый_факт = min(QuantityFact, QuantityPlan)  # если Plan > 0
```

Это защита от массового ошибочного сканирования, **не** сравнение с планом ERP/Excel.

### 3.5. Ключ состояния

Ключ строки: `(JobId, GTIN)`.  
Монитор хранит снимок в `tsd_state.json` и эмитит **дельту** в линию `TSD_PRODUCTION` / entity `L1`.

## 4. Псевдокод опроса (как в мониторе)

```python
snapshot = {}  # (job_id, gtin) -> last_capped_fact

def poll():
    jobs = post("/hs/Honest/jobs/shipment", {Organization, Warehouse})["Jobs"]
    for job in jobs:
        comment = (job.get("Comment") or "").lstrip()
        if not comment.startswith("П"):
            continue
        if not comment_date_in_active_shift(comment):
            continue
        info = post("/hs/Honest/job/shipment/info", {"Id": job["Id"]})
        for item in info.get("Items") or []:
            if not is_large_pack(item, comment):
                continue
            plan, fact = item["QuantityPlan"] or 0, item["QuantityFact"] or 0
            capped = min(fact, plan) if plan > 0 else fact
            key = (job["Id"], item["GTIN"])
            delta = capped - snapshot.get(key, 0)
            if delta:
                emit_production(delta, gtin=item["GTIN"], comment=comment)
            snapshot[key] = capped
    save_snapshot(snapshot)
```

Отрицательная дельта возможна (сняли лишний скан) — монитор умеет её отражать; для ERP обычно берут только положительный прирост за смену.

## 5. Диагностика в этом проекте

| Endpoint | Зачем |
|----------|--------|
| `GET /api/tsd_events` | последние принятые / пропущенные события ТСД |
| статус в health / stats | `polls`, `last_ok`, `last_error`, размер snapshot |

## 6. Чеклист перед боем

1. Проверить Basic-auth на Honest отдельно от `api`.
2. Подставить **свои** Organization / Warehouse UUID (одна пара на склад бочек).
3. Убедиться, что Comment заданий производства начинается с «П» и содержит дату `DD/MM`.
4. Один процесс-поллер на контур (или осознанный shared state).
5. Не коммитить `tsd_state.json` и `.env`.
