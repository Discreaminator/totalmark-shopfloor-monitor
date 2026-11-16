#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ShopfloorMonitor — сторож зависания монитора (Windows 10)

Что делает:
  1. Раз в N секунд опрашивает http://<хост>:5000/api/health
  2. Если poller «молчит» дольше порога (status=stale или age > порога) —
     шлёт письмо на указанные адреса
  3. Когда снова ok — одно письмо «восстановлен»
  4. GUI: ввести SMTP и URL один раз, сохранить, проверить, запустить

Запуск:
  python hang_alert_watcher.py              → окно настроек
  python hang_alert_watcher.py --watch      → только фоновый опрос (для Планировщика)

Зависимости: только стандартная библиотека Python 3.8+
"""

from __future__ import annotations

import os
import sys
import json
import time
import ssl
import smtplib
import traceback
import platform
import subprocess
import urllib.request
import urllib.error
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import tkinter as tk
from tkinter import ttk, messagebox, StringVar, BooleanVar, IntVar

# ============================
# ПУТИ
# ============================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)

APP_NAME = "ShopfloorMonitor — сторож зависания"
CONFIG_FILE = os.path.join(SCRIPT_DIR, "hang_alert_config.json")
STATE_FILE = os.path.join(SCRIPT_DIR, "hang_alert_state.json")
LOG_FILE = os.path.join(SCRIPT_DIR, "hang_alert.log")

DEFAULTS = {
    "health_url": "http://127.0.0.1:5000/api/health",
    "poll_interval_sec": 30,
    "stale_sec": 120,
    "email_cooldown_sec": 900,
    "smtp_server": "",
    "smtp_port": "587",
    "encryption_type": "STARTTLS",  # STARTTLS | SSL
    "login_format": "user@domain",  # domain\\user | user@domain | user_only
    "login": "",
    "password": "",
    "sender_email": "",
    "recipients": "",  # через ;
    "skip_cert_check": False,
    "tls_version": "TLSv1_2",
    "force_starttls_on_443": False,
    "custom_ciphers": "",
}


def log(msg: str) -> None:
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def load_config() -> dict:
    cfg = dict(DEFAULTS)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception as e:
            log(f"Ошибка чтения config: {e}")
    return cfg


def save_config(cfg: dict) -> None:
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=4, ensure_ascii=False)


def load_state() -> dict:
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"was_stale": False, "last_alert_ts": 0.0}


def save_state(state: dict) -> None:
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)
    except Exception as e:
        log(f"state save: {e}")


# ============================
# SMTP (логика как в mail_sender)
# ============================
def send_email(
    smtp_server,
    smtp_port,
    login,
    password,
    sender_email,
    recipient_list,
    subject,
    body,
    encryption_type="STARTTLS",
    skip_cert_check=False,
    login_format="user@domain",
    tls_version="TLSv1_2",
    force_starttls_on_443=False,
    custom_ciphers=None,
):
    """UTF-8 тема/текст (кириллица). Не использовать msg.as_string() без кодировки."""
    from email.message import EmailMessage
    from email.header import Header

    if not recipient_list:
        raise ValueError("Список получателей пуст")
    recipients = [r.strip() for r in str(recipient_list).replace(",", ";").split(";") if r.strip()]
    if not recipients:
        raise ValueError("Нет корректных адресов")

    login_auth = login
    if login_format == "user_only":
        login_auth = login.split("\\")[-1] if "\\" in login else login

    from_addr = (sender_email or login or "").strip()
    # Заголовок From: domain\user ломает ASCII-codec — заменяем слэш для заголовка
    if "\\" in from_addr and "@" not in from_addr:
        from_header = from_addr.replace("\\", "/")
    else:
        from_header = from_addr

    msg = EmailMessage()
    msg.set_content(body or "", charset="utf-8")
    msg["Subject"] = str(Header(subject or "", "utf-8"))
    msg["From"] = from_header
    msg["To"] = ", ".join(recipients)

    context = ssl.create_default_context()
    if skip_cert_check:
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE

    if tls_version == "TLSv1_0":
        context.minimum_version = ssl.TLSVersion.TLSv1
        context.maximum_version = ssl.TLSVersion.TLSv1
    elif tls_version == "TLSv1_1":
        context.minimum_version = ssl.TLSVersion.TLSv1_1
        context.maximum_version = ssl.TLSVersion.TLSv1_1
    else:
        context.minimum_version = ssl.TLSVersion.TLSv1_2

    if custom_ciphers:
        try:
            context.set_ciphers(custom_ciphers)
        except Exception as e:
            log(f"ciphers: {e}")

    port = int(smtp_port)
    use_starttls = (encryption_type == "STARTTLS") or (force_starttls_on_443 and port == 443)

    # envelope from — желательно email с @
    if "@" in from_addr:
        envelope_from = from_addr
    elif "@" in str(login_auth):
        envelope_from = login_auth
    else:
        envelope_from = from_header

    try:
        if use_starttls:
            with smtplib.SMTP(smtp_server, port, timeout=25) as server:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
                server.login(login_auth, password)
                server.send_message(msg, from_addr=envelope_from, to_addrs=recipients)
        else:
            with smtplib.SMTP_SSL(smtp_server, port, context=context, timeout=25) as server:
                server.login(login_auth, password)
                server.send_message(msg, from_addr=envelope_from, to_addrs=recipients)
        return True
    except Exception as e:
        raise Exception(
            f"Ошибка SMTP (login={login_auth}, port={port}, STARTTLS={use_starttls}): {e}\n"
            f"{traceback.format_exc()}"
        )


def send_from_config(cfg: dict, subject: str, body: str) -> bool:
    send_email(
        smtp_server=cfg["smtp_server"],
        smtp_port=cfg["smtp_port"],
        login=cfg["login"],
        password=cfg["password"],
        sender_email=cfg.get("sender_email") or cfg["login"],
        recipient_list=cfg["recipients"],
        subject=subject,
        body=body,
        encryption_type=cfg.get("encryption_type", "STARTTLS"),
        skip_cert_check=bool(cfg.get("skip_cert_check")),
        login_format=cfg.get("login_format", "user@domain"),
        tls_version=cfg.get("tls_version", "TLSv1_2"),
        force_starttls_on_443=bool(cfg.get("force_starttls_on_443")),
        custom_ciphers=cfg.get("custom_ciphers") or None,
    )
    return True


# ============================
# HEALTH CHECK
# ============================
def fetch_health(url: str, timeout: float = 10.0) -> dict:
    """Возвращает dict health или {"_error": "..."}."""
    req = urllib.request.Request(url, headers={"Accept": "application/json", "User-Agent": "ShopfloorMonitorHangWatcher/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
            data = json.loads(raw)
            data["_http_code"] = getattr(resp, "status", 200)
            return data
    except urllib.error.HTTPError as e:
        # 503 при stale — тело всё равно JSON
        try:
            raw = e.read().decode("utf-8", errors="replace")
            data = json.loads(raw)
            data["_http_code"] = e.code
            return data
        except Exception:
            return {"_error": f"HTTP {e.code}: {e.reason}", "_http_code": e.code}
    except Exception as e:
        return {"_error": str(e), "_http_code": 0}


def is_stale(health: dict, stale_sec: int) -> tuple[bool, str]:
    """(stale?, reason)."""
    if health.get("_error") and not health.get("status"):
        return True, f"нет ответа от /api/health: {health.get('_error')}"

    status = (health.get("status") or "").lower()
    age = health.get("last_poll_age_sec")
    try:
        age_f = float(age) if age is not None else None
    except (TypeError, ValueError):
        age_f = None

    if status == "stale":
        return True, f"status=stale, age={age_f}"
    if age_f is not None and age_f > stale_sec:
        return True, f"last_poll_age_sec={age_f} > {stale_sec}"
    if health.get("_http_code") == 503:
        return True, "HTTP 503"
    return False, "ok"


def format_health_body(health: dict, extra: str = "") -> str:
    lines = [
        "ShopfloorMonitor — уведомление сторожа зависания",
        f"Время (локальное ПК): {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}",
        "",
        extra,
        "",
        "--- /api/health ---",
    ]
    for k in (
        "status",
        "last_poll_ok_at",
        "last_poll_at",
        "last_poll_age_sec",
        "last_poll_error",
        "last_poll_duration_sec",
        "report_busy",
        "clients",
        "async_mode",
        "poll_count",
        "last_report_at",
        "last_report_duration_sec",
        "_http_code",
        "_error",
    ):
        if k in health:
            lines.append(f"{k}: {health.get(k)}")
    lines.append("")
    lines.append("Проверьте процесс монитора, логи logs/api или logs/app.log на сервере.")
    return "\n".join(lines)


def run_watch_once(cfg: dict, state: dict) -> dict:
    """Один цикл: опрос + при необходимости письмо. Возвращает обновлённый state."""
    url = (cfg.get("health_url") or "").strip()
    if not url:
        log("Не задан health_url")
        return state

    stale_sec = int(cfg.get("stale_sec") or 120)
    cooldown = int(cfg.get("email_cooldown_sec") or 900)
    health = fetch_health(url)
    stale, reason = is_stale(health, stale_sec)
    now_ts = time.time()
    last_alert = float(state.get("last_alert_ts") or 0)

    if stale:
        log(f"STALE: {reason}")
        if now_ts - last_alert >= cooldown:
            try:
                send_from_config(
                    cfg,
                    subject=f"[ShopfloorMonitor] ЗАВИСАНИЕ монитора · {reason[:80]}",
                    body=format_health_body(health, extra=f"Причина: {reason}"),
                )
                log("Письмо ЗАВИСАНИЕ отправлено")
                state["last_alert_ts"] = now_ts
                state["was_stale"] = True
            except Exception as e:
                log(f"Ошибка отправки письма: {e}")
        else:
            left = int(cooldown - (now_ts - last_alert))
            log(f"Письмо пропущено (cooldown, ещё {left} с)")
        state["was_stale"] = True
    else:
        log(f"OK age={health.get('last_poll_age_sec')} status={health.get('status')}")
        if state.get("was_stale"):
            try:
                send_from_config(
                    cfg,
                    subject="[ShopfloorMonitor] ВОССТАНОВЛЕН — монитор снова отвечает",
                    body=format_health_body(health, extra="Poller снова в норме."),
                )
                log("Письмо ВОССТАНОВЛЕН отправлено")
            except Exception as e:
                log(f"Ошибка письма восстановления: {e}")
            state["was_stale"] = False
            state["last_alert_ts"] = 0.0

    save_state(state)
    return state


def run_watch_loop() -> None:
    log("=== Режим --watch (фоновый опрос) ===")
    while True:
        cfg = load_config()
        state = load_state()
        try:
            run_watch_once(cfg, state)
        except Exception as e:
            log(f"Цикл: {e}\n{traceback.format_exc()}")
        interval = max(10, int(cfg.get("poll_interval_sec") or 30))
        time.sleep(interval)


# ============================
# GUI
# ============================
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("720x780")
        self.root.resizable(False, False)

        self.cfg = load_config()
        self.health_url = StringVar(value=self.cfg.get("health_url", DEFAULTS["health_url"]))
        self.poll_interval = StringVar(value=str(self.cfg.get("poll_interval_sec", 30)))
        self.stale_sec = StringVar(value=str(self.cfg.get("stale_sec", 120)))
        self.cooldown = StringVar(value=str(self.cfg.get("email_cooldown_sec", 900)))

        self.smtp_server = StringVar(value=self.cfg.get("smtp_server", ""))
        self.smtp_port = StringVar(value=str(self.cfg.get("smtp_port", "587")))
        self.encryption_type = StringVar(value=self.cfg.get("encryption_type", "STARTTLS"))
        self.login_format = StringVar(value=self.cfg.get("login_format", "user@domain"))
        self.login = StringVar(value=self.cfg.get("login", ""))
        self.password = StringVar(value=self.cfg.get("password", ""))
        self.sender_email = StringVar(value=self.cfg.get("sender_email", ""))
        self.recipients = StringVar(value=self.cfg.get("recipients", ""))
        self.skip_cert = BooleanVar(value=bool(self.cfg.get("skip_cert_check")))
        self.tls_version = StringVar(value=self.cfg.get("tls_version", "TLSv1_2"))
        self.force_starttls = BooleanVar(value=bool(self.cfg.get("force_starttls_on_443")))
        self.custom_ciphers = StringVar(value=self.cfg.get("custom_ciphers", ""))

        self.status_text = StringVar(value="Готово. Сохраните настройки и нажмите «Проверить».")
        self._watch_job = None
        self._watching = False

        self._build()

    def _build(self):
        f = ttk.Frame(self.root, padding=16)
        f.pack(fill=tk.BOTH, expand=True)

        r = 0
        ttk.Label(f, text="Монитор ShopfloorMonitor", font=("Segoe UI", 12, "bold")).grid(
            row=r, column=0, columnspan=2, sticky=tk.W, pady=(0, 8)
        )
        r += 1
        ttk.Label(f, text="URL /api/health:").grid(row=r, column=0, sticky=tk.W, pady=3)
        ttk.Entry(f, textvariable=self.health_url, width=58).grid(row=r, column=1, sticky=tk.W, pady=3)
        r += 1
        ttk.Label(f, text="Интервал опроса, сек:").grid(row=r, column=0, sticky=tk.W, pady=3)
        ttk.Entry(f, textvariable=self.poll_interval, width=10).grid(row=r, column=1, sticky=tk.W, pady=3)
        r += 1
        ttk.Label(f, text="Порог зависания, сек:").grid(row=r, column=0, sticky=tk.W, pady=3)
        ttk.Entry(f, textvariable=self.stale_sec, width=10).grid(row=r, column=1, sticky=tk.W, pady=3)
        r += 1
        ttk.Label(f, text="Пауза между письмами, сек:").grid(row=r, column=0, sticky=tk.W, pady=3)
        ttk.Entry(f, textvariable=self.cooldown, width=10).grid(row=r, column=1, sticky=tk.W, pady=3)
        r += 1

        ttk.Separator(f, orient=tk.HORIZONTAL).grid(row=r, column=0, columnspan=2, sticky=tk.EW, pady=10)
        r += 1
        ttk.Label(f, text="Почта (SMTP)", font=("Segoe UI", 12, "bold")).grid(
            row=r, column=0, columnspan=2, sticky=tk.W, pady=(0, 8)
        )
        r += 1

        for label, var, width in [
            ("SMTP-сервер:", self.smtp_server, 40),
            ("SMTP-порт:", self.smtp_port, 10),
        ]:
            ttk.Label(f, text=label).grid(row=r, column=0, sticky=tk.W, pady=3)
            ttk.Entry(f, textvariable=var, width=width).grid(row=r, column=1, sticky=tk.W, pady=3)
            r += 1

        ttk.Label(f, text="Шифрование:").grid(row=r, column=0, sticky=tk.W, pady=3)
        ttk.Combobox(
            f, textvariable=self.encryption_type, values=["STARTTLS", "SSL"], state="readonly", width=15
        ).grid(row=r, column=1, sticky=tk.W, pady=3)
        r += 1

        ttk.Label(f, text="Формат логина:").grid(row=r, column=0, sticky=tk.W, pady=3)
        ttk.Combobox(
            f,
            textvariable=self.login_format,
            values=["user@domain", "domain\\user", "user_only"],
            state="readonly",
            width=18,
        ).grid(row=r, column=1, sticky=tk.W, pady=3)
        r += 1

        ttk.Label(f, text="Логин:").grid(row=r, column=0, sticky=tk.W, pady=3)
        ttk.Entry(f, textvariable=self.login, width=40).grid(row=r, column=1, sticky=tk.W, pady=3)
        r += 1
        ttk.Label(f, text="Пароль:").grid(row=r, column=0, sticky=tk.W, pady=3)
        ttk.Entry(f, textvariable=self.password, width=40, show="*").grid(row=r, column=1, sticky=tk.W, pady=3)
        r += 1
        ttk.Label(f, text="From (опционально):").grid(row=r, column=0, sticky=tk.W, pady=3)
        ttk.Entry(f, textvariable=self.sender_email, width=40).grid(row=r, column=1, sticky=tk.W, pady=3)
        r += 1
        ttk.Label(f, text="Кому (через ;):").grid(row=r, column=0, sticky=tk.W, pady=3)
        ttk.Entry(f, textvariable=self.recipients, width=40).grid(row=r, column=1, sticky=tk.W, pady=3)
        r += 1

        ttk.Label(f, text="TLS:").grid(row=r, column=0, sticky=tk.W, pady=3)
        ttk.Combobox(
            f, textvariable=self.tls_version, values=["TLSv1_2", "TLSv1_1", "TLSv1_0"], state="readonly", width=12
        ).grid(row=r, column=1, sticky=tk.W, pady=3)
        r += 1
        ttk.Checkbutton(f, text="Пропустить проверку SSL-сертификата", variable=self.skip_cert).grid(
            row=r, column=0, columnspan=2, sticky=tk.W, pady=2
        )
        r += 1
        ttk.Checkbutton(f, text="Принудительно STARTTLS на порту 443", variable=self.force_starttls).grid(
            row=r, column=0, columnspan=2, sticky=tk.W, pady=2
        )
        r += 1

        btns = ttk.Frame(f)
        btns.grid(row=r, column=0, columnspan=2, pady=12)
        ttk.Button(btns, text="💾 Сохранить", command=self.save).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="🩺 Проверить health", command=self.test_health).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="📧 Тест письма", command=self.test_mail).pack(side=tk.LEFT, padx=4)
        r += 1

        btns2 = ttk.Frame(f)
        btns2.grid(row=r, column=0, columnspan=2, pady=4)
        ttk.Button(btns2, text="▶ Следить в этом окне", command=self.toggle_watch_gui).pack(side=tk.LEFT, padx=4)
        if platform.system() == "Windows":
            ttk.Button(btns2, text="📅 В Планировщик (фон)", command=self.install_task).pack(side=tk.LEFT, padx=4)
            ttk.Button(btns2, text="🗑️ Убрать из Планировщика", command=self.remove_task).pack(side=tk.LEFT, padx=4)
        r += 1

        ttk.Label(f, textvariable=self.status_text, wraplength=660, foreground="#333").grid(
            row=r, column=0, columnspan=2, sticky=tk.W, pady=10
        )
        r += 1
        note = (
            "Как пользоваться:\n"
            "1) Укажите URL монитора (с /api/health) и SMTP — как в авторассылке отчётов.\n"
            "2) «Сохранить» → «Проверить health» → «Тест письма».\n"
            "3) «Следить в этом окне» ИЛИ «В Планировщик» (работает даже без открытого окна).\n"
            f"Лог: {LOG_FILE}"
        )
        ttk.Label(f, text=note, foreground="gray", justify=tk.LEFT).grid(
            row=r, column=0, columnspan=2, sticky=tk.W
        )

    def collect(self) -> dict:
        return {
            "health_url": self.health_url.get().strip(),
            "poll_interval_sec": int(self.poll_interval.get() or 30),
            "stale_sec": int(self.stale_sec.get() or 120),
            "email_cooldown_sec": int(self.cooldown.get() or 900),
            "smtp_server": self.smtp_server.get().strip(),
            "smtp_port": self.smtp_port.get().strip(),
            "encryption_type": self.encryption_type.get(),
            "login_format": self.login_format.get(),
            "login": self.login.get().strip(),
            "password": self.password.get(),
            "sender_email": self.sender_email.get().strip(),
            "recipients": self.recipients.get().strip(),
            "skip_cert_check": self.skip_cert.get(),
            "tls_version": self.tls_version.get(),
            "force_starttls_on_443": self.force_starttls.get(),
            "custom_ciphers": self.custom_ciphers.get().strip(),
        }

    def save(self):
        cfg = self.collect()
        if not cfg["health_url"]:
            messagebox.showerror("Ошибка", "Укажите URL /api/health")
            return
        if not cfg["smtp_server"] or not cfg["login"] or not cfg["password"] or not cfg["recipients"]:
            messagebox.showerror("Ошибка", "Заполните SMTP, логин, пароль и получателей")
            return
        save_config(cfg)
        self.cfg = cfg
        self.status_text.set(f"Сохранено → {CONFIG_FILE}")
        messagebox.showinfo("Ок", "Настройки сохранены")

    def test_health(self):
        try:
            self.save_silent()
        except Exception:
            pass
        cfg = load_config()
        h = fetch_health(cfg["health_url"])
        stale, reason = is_stale(h, int(cfg.get("stale_sec") or 120))
        msg = json.dumps(h, ensure_ascii=False, indent=2)[:1500]
        self.status_text.set(f"health: stale={stale} ({reason})")
        messagebox.showinfo("Health", f"stale={stale}\n{reason}\n\n{msg}")

    def test_mail(self):
        cfg = self.collect()
        save_config(cfg)
        try:
            send_from_config(
                cfg,
                subject="[ShopfloorMonitor] Тест сторожа зависания",
                body="Если вы видите это письмо — SMTP настроен верно.\n"
                f"URL: {cfg.get('health_url')}\n"
                f"Время: {datetime.now()}",
            )
            messagebox.showinfo("Ок", "Тестовое письмо отправлено")
            self.status_text.set("Тестовое письмо отправлено")
        except Exception as e:
            messagebox.showerror("Ошибка SMTP", str(e)[:800])

    def save_silent(self):
        save_config(self.collect())

    def toggle_watch_gui(self):
        if self._watching:
            self._watching = False
            if self._watch_job:
                self.root.after_cancel(self._watch_job)
                self._watch_job = None
            self.status_text.set("Опрос остановлен")
            return
        try:
            self.save_silent()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            return
        self._watching = True
        self.status_text.set("Опрос запущен…")
        self._tick()

    def _tick(self):
        if not self._watching:
            return
        cfg = load_config()
        state = load_state()
        try:
            run_watch_once(cfg, state)
            self.status_text.set(
                f"Последний опрос {datetime.now().strftime('%H:%M:%S')} — см. hang_alert.log"
            )
        except Exception as e:
            self.status_text.set(f"Ошибка: {e}")
            log(str(e))
        interval_ms = max(10, int(cfg.get("poll_interval_sec") or 30)) * 1000
        self._watch_job = self.root.after(interval_ms, self._tick)

    def install_task(self):
        if platform.system() != "Windows":
            messagebox.showerror("Ошибка", "Только Windows")
            return
        try:
            self.save_silent()
        except Exception:
            pass
        script = os.path.abspath(sys.argv[0])
        python_exe = sys.executable
        # pythonw — без чёрного окна; если нет — python
        if python_exe.lower().endswith("python.exe"):
            pythonw = python_exe[:-10] + "pythonw.exe"
            if os.path.isfile(pythonw):
                python_exe = pythonw
        tr = f'"{python_exe}" "{script}" --watch'
        task_name = "ShopfloorMonitorHangAlert"
        cmd = [
            "schtasks",
            "/create",
            "/tn",
            task_name,
            "/tr",
            tr,
            "/sc",
            "onstart",
            "/ru",
            "SYSTEM",
            "/rl",
            "HIGHEST",
            "/f",
        ]
        # Дополнительно: каждые 5 мин на случай если процесс упадёт — отдельная задача-«пин»
        # Основной --watch сам крутится в цикле; onstart достаточно.
        try:
            result = subprocess.run(cmd, capture_output=True)
            err = (result.stderr or b"").decode("cp866", errors="replace")
            if result.returncode != 0:
                # без SYSTEM — от текущего пользователя
                cmd2 = [
                    "schtasks",
                    "/create",
                    "/tn",
                    task_name,
                    "/tr",
                    tr,
                    "/sc",
                    "onlogon",
                    "/f",
                ]
                result = subprocess.run(cmd2, capture_output=True)
                err = (result.stderr or b"").decode("cp866", errors="replace")
                if result.returncode != 0:
                    messagebox.showerror("Ошибка", err or "schtasks failed")
                    return
            messagebox.showinfo(
                "Готово",
                f"Задача «{task_name}» создана.\n"
                f"Команда: {tr}\n\n"
                "Можно сразу запустить:\n"
                f'schtasks /run /tn "{task_name}"',
            )
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def remove_task(self):
        if platform.system() != "Windows":
            return
        try:
            subprocess.run(["schtasks", "/delete", "/tn", "ShopfloorMonitorHangAlert", "/f"], capture_output=True)
            messagebox.showinfo("Готово", "Задача удалена (если была)")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


# ============================
# MAIN
# ============================
if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("--watch", "-w"):
        run_watch_loop()
        sys.exit(0)

    root = tk.Tk()
    try:
        root.iconbitmap(default="")
    except Exception:
        pass
    App(root)
    root.mainloop()
