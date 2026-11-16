#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import requests
from datetime import date, datetime
from urllib.parse import urlparse
import json
import psycopg2

# ============= НАСТРОЙКИ =============
API_KEY = os.environ.get("NK_API_KEY", "")  # ключ национального каталога — только из env
BEARER_TOKEN = os.environ.get("NK_BEARER_TOKEN") or None

SAVE_DIR = "static/products"
os.makedirs(SAVE_DIR, exist_ok=True)

# URL из документации
API_BASE_URLS = [
    "https://апи.национальный-каталог.рф",
    "https://api.national-catalog.ru",
    "https://api.nk.sandbox.crptech.ru",   # тестовый контур
]

API_METHODS = [
    "/v3/feed-product",
    "/v3/product",
]

# ============= ЛОГИРОВАНИЕ =============
def log(msg, level="INFO"):
    print(f"[{level}] {msg}")

# ============= ПОЛУЧЕНИЕ GTIN ИЗ БД =============
def get_gtins_from_aggregates_today() -> list[str]:
    try:
        conn = psycopg2.connect(
            host=os.environ.get("DB_HOST", "localhost"),
            port=int(os.environ.get("DB_PORT", "5432")),
            dbname=os.environ.get("DB_NAME", ""),
            user=os.environ.get("DB_USER", ""),
            password=os.environ.get("DB_PASSWORD", ""),
        )
        cur = conn.cursor()
        cur.execute("""
            SELECT DISTINCT gtin
            FROM aggregate
            WHERE date(created) = CURRENT_DATE
              AND gtin IS NOT NULL
              AND gtin != ''
        """)
        gtins = [row[0] for row in cur.fetchall()]
        cur.close()
        conn.close()
        log(f"Найдено {len(gtins)} уникальных GTIN в агрегатах за сегодня", "INFO")
        return gtins
    except Exception as e:
        log(f"Ошибка подключения к БД: {e}", "ERROR")
        return []

# ============= ПРОВЕРКА НАЛИЧИЯ (с нормализацией) =============
def image_exists(gtin: str) -> bool:
    normalized = gtin.lstrip('0') or '0'
    for ext in ('.jpg', '.jpeg', '.png', '.gif'):
        path = os.path.join(SAVE_DIR, f"{normalized}{ext}")
        if os.path.isfile(path) and os.path.getsize(path) > 0:
            log(f"[{gtin}] Файл уже есть: {path}", "SKIP")
            return True
    return False

# ============= СКАЧИВАНИЕ =============
def download_image(gtin: str) -> str | None:
    log(f"[{gtin}] Начинаем поиск...", "INFO")
    
    for base_url in API_BASE_URLS:
        for method in API_METHODS:
            url = base_url.rstrip('/') + method
            params = {"gtin": gtin}
            headers = {}

            if API_KEY:
                params["apikey"] = API_KEY
            elif BEARER_TOKEN:
                headers["Authorization"] = f"Bearer {BEARER_TOKEN}"
            else:
                log("Нет API_KEY или BEARER_TOKEN", "ERROR")
                return None

            try:
                log(f"[{gtin}] Пробуем {url} с params={params}", "DEBUG")
                response = requests.get(url, params=params, headers=headers, timeout=15)

                if response.status_code == 404:
                    log(f"[{gtin}] {url} вернул 404, пробуем дальше", "WARNING")
                    continue
                elif response.status_code != 200:
                    log(f"[{gtin}] HTTP {response.status_code}: {response.text[:200]}", "ERROR")
                    continue

                data = response.json()
                log(f"[{gtin}] Получен ответ, разбираем...", "DEBUG")

                # Извлекаем карточку
                product = None
                if "result" in data and data["result"]:
                    product = data["result"][0] if isinstance(data["result"], list) else data["result"]
                elif "good" in data:
                    product = data["good"]
                else:
                    product = data

                if not product:
                    log(f"[{gtin}] Товар не найден в ответе", "WARNING")
                    log(f"[{gtin}] Ответ: {json.dumps(data, ensure_ascii=False)[:200]}", "DEBUG")
                    continue

                # Ищем URL изображения
                image_url = None
                if product.get("good_img"):
                    image_url = product["good_img"]
                elif product.get("good_images"):
                    for img in product["good_images"]:
                        if img.get("photo_type") in ("default", "facing"):
                            image_url = img.get("photo_url")
                            break
                    if not image_url:
                        image_url = product["good_images"][0].get("photo_url")

                if not image_url:
                    log(f"[{gtin}] Изображение не найдено", "WARNING")
                    continue

                log(f"[{gtin}] Найдено изображение: {image_url}", "INFO")
                img_resp = requests.get(image_url, timeout=20)
                img_resp.raise_for_status()

                content_type = img_resp.headers.get("content-type", "")
                ext = ".jpg"
                if "png" in content_type:
                    ext = ".png"
                elif "gif" in content_type:
                    ext = ".gif"
                else:
                    parsed = urlparse(image_url)
                    path = parsed.path
                    if path.endswith((".png", ".jpg", ".jpeg", ".gif")):
                        ext = os.path.splitext(path)[1] or ".jpg"

                # Нормализуем GTIN (удаляем ведущие нули)
                normalized_gtin = gtin.lstrip('0') or '0'
                filepath = os.path.join(SAVE_DIR, f"{normalized_gtin}{ext}")
                with open(filepath, "wb") as f:
                    f.write(img_resp.content)

                log(f"[{gtin}] ✅ Сохранено → {filepath}", "SUCCESS")
                return filepath

            except Exception as e:
                log(f"[{gtin}] Ошибка: {e}", "ERROR")
                continue

    log(f"[{gtin}] ❌ Не удалось скачать изображение ни по одному URL", "ERROR")
    return None

# ============= ОСНОВНАЯ ФУНКЦИЯ =============
def sync_today_images():
    gtins = get_gtins_from_aggregates_today()
    log(f"📅 План на {date.today()}: {len(gtins)} GTIN", "INFO")
    if not gtins:
        log("Нет GTIN для обработки", "WARNING")
        return

    downloaded = 0
    skipped = 0
    failed = 0

    for gtin in gtins:
        if image_exists(gtin):
            skipped += 1
            continue

        result = download_image(gtin)
        if result:
            downloaded += 1
        else:
            failed += 1

    log(f"\n--- ИТОГ ---", "INFO")
    log(f"✅ Скачано:   {downloaded}", "INFO")
    log(f"⏩ Пропущено: {skipped}", "INFO")
    log(f"❌ Ошибок:    {failed}", "INFO")

# ============= ТОЧКА ВХОДА =============
if __name__ == "__main__":
    sync_today_images()