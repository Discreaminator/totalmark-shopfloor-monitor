#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import smtplib
import ssl
import re
import subprocess
import traceback
import platform
import time
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, StringVar, IntVar, BooleanVar

# ============================
# КОНСТАНТЫ И НАСТРОЙКИ (с абсолютными путями)
# ============================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)  # переходим в папку скрипта

APP_NAME = "Авторассылка отчётов"
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")
SENT_LOG_FILE = os.path.join(SCRIPT_DIR, "sent.log")
REPORTS_DIR = os.path.join(SCRIPT_DIR, "reports")

# Если папки reports нет, создаём
if not os.path.exists(REPORTS_DIR):
    try:
        os.makedirs(REPORTS_DIR)
    except:
        pass

# ============================
# РАБОТА С КОНФИГУРАЦИЕЙ
# ============================
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

# ============================
# ЛОГИРОВАНИЕ ОТПРАВЛЕННЫХ ФАЙЛОВ
# ============================
def load_sent_log():
    if os.path.exists(SENT_LOG_FILE):
        with open(SENT_LOG_FILE, 'r', encoding='utf-8') as f:
            return set(line.strip() for line in f if line.strip())
    return set()

def add_to_sent_log(filename):
    with open(SENT_LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(filename + '\n')

# ============================
# ПОИСК ФАЙЛА ОТЧЁТА (улучшенный)
# ============================
def find_report(shift_type, date_str, reports_dir=REPORTS_DIR, max_age_hours=14):
    """
    Ищет самый свежий отчёт нужного типа за указанную дату.
    Если точного совпадения нет – берёт самый свежий файл этого типа
    не старше max_age_hours (на случай рассинхронизации даты).
    """
    if not os.path.isdir(reports_dir):
        return None

    pattern_exact = re.compile(rf'report_{date_str}_{shift_type}_(\d{{4}})\.xlsx')
    pattern_any   = re.compile(rf'report_\d{{4}}-\d{{2}}-\d{{2}}_{shift_type}_(\d{{4}})\.xlsx')

    candidates = []
    now = time.time()
    for fname in os.listdir(reports_dir):
        filepath = os.path.join(reports_dir, fname)
        if not os.path.isfile(filepath):
            continue
        mtime = os.path.getmtime(filepath)
        if now - mtime > max_age_hours * 3600:
            continue

        if pattern_exact.match(fname):
            candidates.append((mtime, filepath, 0))   # приоритет 0 — точное совпадение
        elif pattern_any.match(fname):
            candidates.append((mtime, filepath, 1))   # запасной вариант

    if not candidates:
        return None

    # Сначала точные, потом самые свежие
    candidates.sort(key=lambda x: (x[2], -x[0]))
    return candidates[0][1]

# ============================
# ОТПРАВКА ПИСЬМА
# ============================
def send_email(smtp_server, smtp_port, login, password, sender_email,
               recipient_list, subject, body, attachment_path=None,
               encryption_type='SSL', skip_cert_check=False,
               login_format='domain\\user', tls_version='TLSv1_2',
               force_starttls_on_443=False, custom_ciphers=None):
    """
    login_format: 'domain\\user', 'user@domain', 'user_only'
    encryption_type: 'SSL' или 'STARTTLS'
    """
    if not recipient_list:
        raise ValueError("Список получателей пуст")
    recipients = [r.strip() for r in recipient_list.split(';') if r.strip()]
    if not recipients:
        raise ValueError("Нет корректных адресов получателей")

    msg = MIMEMultipart()
    msg['From'] = sender_email if sender_email else login
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename="{os.path.basename(attachment_path)}"'
            )
            msg.attach(part)

    # Обработка формата логина
    if login_format == 'user@domain':
        if '@' not in login:
            login = f"{login}@cstream.ru"  # можно сделать настраиваемым
    elif login_format == 'user_only':
        login = login.split('\\')[-1] if '\\' in login else login
    # else: 'domain\\user' - оставляем как есть

    # Настройка SSL-контекста
    context = ssl.create_default_context()
    if skip_cert_check:
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE

    # Устанавливаем версию TLS
    if tls_version == 'TLSv1_0':
        context.minimum_version = ssl.TLSVersion.TLSv1
        context.maximum_version = ssl.TLSVersion.TLSv1
    elif tls_version == 'TLSv1_1':
        context.minimum_version = ssl.TLSVersion.TLSv1_1
        context.maximum_version = ssl.TLSVersion.TLSv1_1
    else:  # TLSv1_2 (по умолчанию)
        context.minimum_version = ssl.TLSVersion.TLSv1_2
        context.maximum_version = ssl.TLSVersion.TLSv1_2

    if custom_ciphers:
        try:
            context.set_ciphers(custom_ciphers)
        except Exception as e:
            print(f"Предупреждение: не удалось установить шифры: {e}")

    try:
        # Определяем, использовать ли STARTTLS
        use_starttls = (encryption_type == 'STARTTLS') or (force_starttls_on_443 and smtp_port == 443)

        if use_starttls:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls(context=context)
                server.login(login, password)
                server.sendmail(sender_email or login, recipients, msg.as_string())
        else:
            with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
                server.login(login, password)
                server.sendmail(sender_email or login, recipients, msg.as_string())
        return True
    except Exception as e:
        raise Exception(f"Ошибка при отправке (логин: {login}, порт: {smtp_port}, TLS: {tls_version}, STARTTLS: {use_starttls}): {str(e)}\n{traceback.format_exc()}")

# ============================
# ОСНОВНАЯ ЛОГИКА ОТПРАВКИ ПО РАСПИСАНИЮ (исправленная)
# ============================
def send_scheduled_report(shift_type):
    config = load_config()
    if not config:
        print("❌ Конфигурация не найдена. Запустите GUI для настройки.")
        return

    # И для day, и для night отчёт называется сегодняшней датой
    date_obj = datetime.now().date()
    date_str = date_obj.strftime('%Y-%m-%d')

    reports_dir = config.get('reports_dir', REPORTS_DIR)
    if not os.path.isdir(reports_dir):
        print(f"⚠️ Папка с отчётами не существует: {reports_dir}")
        return

    # Повторные попытки (до 3 раз с паузой 30 сек)
    report_file = None
    for attempt in range(3):
        report_file = find_report(shift_type, date_str, reports_dir)
        if report_file:
            break
        if attempt < 2:
            print(f"⏳ Отчёт ещё не появился, ждём 30 сек... (попытка {attempt+1}/3)")
            time.sleep(30)

    if not report_file:
        print(f"⚠️ Отчёт за {date_str} ({shift_type}) не найден после 3 попыток.")
        print(f"   Искали в: {os.path.abspath(reports_dir)}")
        try:
            files = os.listdir(reports_dir)
            print(f"   Файлы в папке ({len(files)}): {files[:10]}{'...' if len(files)>10 else ''}")
        except Exception as e:
            print(f"   Не удалось прочитать папку: {e}")
        return

    sent_log = load_sent_log()
    filename = os.path.basename(report_file)
    if filename in sent_log:
        print(f"ℹ️ Отчёт {filename} уже был отправлен ранее.")
        return

    shift_name = "дневную" if shift_type == 'day' else "ночную"
    subject = f"📊 Отчёт по линиям за {date_obj.strftime('%d.%m.%Y')} ({shift_name} смену) по площадке Подольск"
    body = (f"Коллеги, высылаю отчет по линиям за {date_obj.strftime('%d.%m.%Y')} "
            f"\"{shift_name} смену\" по площадке Подольск.\n\n"
            f"Файл отчёта приложен к письму.")

    smtp_server = config['smtp_server']
    smtp_port = int(config['smtp_port'])
    login = config['login']
    password = config['password']
    sender = config.get('sender_email', login)
    recipients = config['recipients']
    skip_cert = config.get('skip_cert_check', False)
    encryption = config.get('encryption_type', 'SSL')
    login_format = config.get('login_format', 'domain\\user')
    tls_version = config.get('tls_version', 'TLSv1_2')
    force_starttls = config.get('force_starttls_on_443', False)
    custom_ciphers = config.get('custom_ciphers', '')

    try:
        send_email(smtp_server, smtp_port, login, password, sender, recipients,
                   subject, body, report_file,
                   encryption_type=encryption, skip_cert_check=skip_cert,
                   login_format=login_format, tls_version=tls_version,
                   force_starttls_on_443=force_starttls, custom_ciphers=custom_ciphers)
        add_to_sent_log(filename)
        print(f"✅ Отчёт {filename} успешно отправлен")
    except Exception as e:
        print(f"❌ Ошибка отправки: {e}")

# ============================
# GUI-ПРИЛОЖЕНИЕ
# ============================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("750x900")
        self.root.resizable(False, False)

        self.config = load_config()

        self.smtp_server = StringVar(value=self.config.get('smtp_server', ''))
        self.smtp_port = StringVar(value=self.config.get('smtp_port', '587'))
        self.login = StringVar(value=self.config.get('login', ''))
        self.password = StringVar(value=self.config.get('password', ''))
        self.sender_email = StringVar(value=self.config.get('sender_email', ''))
        self.recipients = StringVar(value=self.config.get('recipients', ''))
        self.reports_dir = StringVar(value=self.config.get('reports_dir', REPORTS_DIR))
        self.skip_cert_check = BooleanVar(value=self.config.get('skip_cert_check', False))
        self.encryption_type = StringVar(value=self.config.get('encryption_type', 'SSL'))
        self.login_format = StringVar(value=self.config.get('login_format', 'domain\\user'))
        self.tls_version = StringVar(value=self.config.get('tls_version', 'TLSv1_2'))
        self.force_starttls = BooleanVar(value=self.config.get('force_starttls_on_443', False))
        self.custom_ciphers = StringVar(value=self.config.get('custom_ciphers', ''))

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Настройки почтового сервера", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0,10))

        ttk.Label(main_frame, text="SMTP-сервер:").grid(row=1, column=0, sticky=tk.W, pady=4)
        ttk.Entry(main_frame, textvariable=self.smtp_server, width=40).grid(row=1, column=1, sticky=tk.W, pady=4)

        ttk.Label(main_frame, text="SMTP-порт:").grid(row=2, column=0, sticky=tk.W, pady=4)
        ttk.Entry(main_frame, textvariable=self.smtp_port, width=10).grid(row=2, column=1, sticky=tk.W, pady=4)

        ttk.Label(main_frame, text="Тип шифрования:").grid(row=3, column=0, sticky=tk.W, pady=4)
        cb_enc = ttk.Combobox(main_frame, textvariable=self.encryption_type,
                              values=["SSL", "STARTTLS"], state="readonly")
        cb_enc.grid(row=3, column=1, sticky=tk.W, pady=4)
        cb_enc.set(self.encryption_type.get())

        ttk.Label(main_frame, text="Формат логина:").grid(row=4, column=0, sticky=tk.W, pady=4)
        cb_login = ttk.Combobox(main_frame, textvariable=self.login_format,
                                values=["domain\\user", "user@domain", "user_only"], state="readonly")
        cb_login.grid(row=4, column=1, sticky=tk.W, pady=4)
        cb_login.set(self.login_format.get())
        ttk.Label(main_frame, text="(примеры: tehf\\reports, reports@cstream.ru, reports)", foreground="gray", font=('Arial', 8)).grid(row=5, column=1, sticky=tk.W)

        ttk.Label(main_frame, text="Логин:").grid(row=6, column=0, sticky=tk.W, pady=4)
        ttk.Entry(main_frame, textvariable=self.login, width=40).grid(row=6, column=1, sticky=tk.W, pady=4)

        ttk.Label(main_frame, text="Пароль:").grid(row=7, column=0, sticky=tk.W, pady=4)
        ttk.Entry(main_frame, textvariable=self.password, width=40, show="*").grid(row=7, column=1, sticky=tk.W, pady=4)

        ttk.Label(main_frame, text="Email отправителя (опционально):").grid(row=8, column=0, sticky=tk.W, pady=4)
        ttk.Entry(main_frame, textvariable=self.sender_email, width=40).grid(row=8, column=1, sticky=tk.W, pady=4)

        ttk.Label(main_frame, text="Получатели (через ;):").grid(row=9, column=0, sticky=tk.W, pady=4)
        ttk.Entry(main_frame, textvariable=self.recipients, width=40).grid(row=9, column=1, sticky=tk.W, pady=4)

        ttk.Label(main_frame, text="Папка с отчётами:").grid(row=10, column=0, sticky=tk.W, pady=4)
        frm = ttk.Frame(main_frame)
        frm.grid(row=10, column=1, sticky=tk.W)
        ttk.Entry(frm, textvariable=self.reports_dir, width=30).pack(side=tk.LEFT)
        ttk.Button(frm, text="Выбрать", command=self.select_reports_dir).pack(side=tk.LEFT, padx=5)

        ttk.Label(main_frame, text="Дополнительные настройки SSL", font=('Arial', 10, 'bold')).grid(row=11, column=0, columnspan=2, sticky=tk.W, pady=(10,5))
        ttk.Label(main_frame, text="Версия TLS:").grid(row=12, column=0, sticky=tk.W, pady=4)
        cb_tls = ttk.Combobox(main_frame, textvariable=self.tls_version,
                              values=["TLSv1_2", "TLSv1_1", "TLSv1_0"], state="readonly")
        cb_tls.grid(row=12, column=1, sticky=tk.W, pady=4)
        cb_tls.set(self.tls_version.get())

        ttk.Checkbutton(main_frame, text="Принудительно использовать STARTTLS на порту 443",
                        variable=self.force_starttls).grid(row=13, column=0, columnspan=2, sticky=tk.W, pady=4)

        ttk.Label(main_frame, text="Пользовательские шифры (опционально):").grid(row=14, column=0, sticky=tk.W, pady=4)
        ttk.Entry(main_frame, textvariable=self.custom_ciphers, width=40).grid(row=14, column=1, sticky=tk.W, pady=4)
        ttk.Label(main_frame, text="(оставьте пустым для авто-выбора)", foreground="gray", font=('Arial', 8)).grid(row=15, column=1, sticky=tk.W)

        ttk.Checkbutton(main_frame, text="Пропустить проверку SSL-сертификата",
                        variable=self.skip_cert_check).grid(row=16, column=0, columnspan=2, sticky=tk.W, pady=4)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=17, column=0, columnspan=2, pady=15)
        ttk.Button(btn_frame, text="💾 Сохранить настройки", command=self.save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📧 Проверить подключение", command=self.test_connection).pack(side=tk.LEFT, padx=5)

        ttk.Label(main_frame, text="Настройка расписания", font=('Arial', 12, 'bold')).grid(row=18, column=0, columnspan=2, sticky=tk.W, pady=(20,5))

        is_windows = platform.system() == 'Windows'
        if is_windows:
            ttk.Label(main_frame, text="Дневная смена (20:05):").grid(row=19, column=0, sticky=tk.W, pady=2)
            ttk.Button(main_frame, text="✅ Установить", command=lambda: self.setup_schedule('day')).grid(row=19, column=1, sticky=tk.W, pady=2)
            ttk.Label(main_frame, text="Ночная смена (8:05):").grid(row=20, column=0, sticky=tk.W, pady=2)
            ttk.Button(main_frame, text="✅ Установить", command=lambda: self.setup_schedule('night')).grid(row=20, column=1, sticky=tk.W, pady=2)
            ttk.Button(main_frame, text="🗑️ Удалить всё расписание", command=self.remove_schedule).grid(row=21, column=0, columnspan=2, pady=10)
        else:
            ttk.Label(main_frame, text="Настройка расписания поддерживается только в Windows.",
                      foreground="orange").grid(row=19, column=0, columnspan=2, pady=2)

        info = ("Примечание: расписание создаётся в Планировщике задач Windows.\n"
                "Для корректной работы убедитесь, что Python и этот скрипт доступны.\n"
                "Пароль хранится в открытом виде в config.json.")
        ttk.Label(main_frame, text=info, foreground="gray", justify=tk.LEFT).grid(row=22, column=0, columnspan=2, sticky=tk.W, pady=(10,0))

    def select_reports_dir(self):
        path = filedialog.askdirectory(title="Выберите папку с отчётами")
        if path:
            self.reports_dir.set(path)

    def save_settings(self):
        config = {
            'smtp_server': self.smtp_server.get().strip(),
            'smtp_port': self.smtp_port.get().strip(),
            'encryption_type': self.encryption_type.get(),
            'login_format': self.login_format.get(),
            'login': self.login.get().strip(),
            'password': self.password.get().strip(),
            'sender_email': self.sender_email.get().strip(),
            'recipients': self.recipients.get().strip(),
            'reports_dir': self.reports_dir.get().strip(),
            'skip_cert_check': self.skip_cert_check.get(),
            'tls_version': self.tls_version.get(),
            'force_starttls_on_443': self.force_starttls.get(),
            'custom_ciphers': self.custom_ciphers.get().strip()
        }
        if not config['smtp_server'] or not config['login'] or not config['password'] or not config['recipients']:
            messagebox.showerror("Ошибка", "Заполните все обязательные поля (SMTP-сервер, логин, пароль, получатели).")
            return
        save_config(config)
        self.config = config
        messagebox.showinfo("Успех", "Настройки сохранены в config.json")

    def test_connection(self):
        self.save_settings()
        config = load_config()
        if not config:
            return
        try:
            recipients = config['recipients'].split(';')[0].strip()
            if not recipients:
                messagebox.showerror("Ошибка", "Нет получателей для теста.")
                return
            send_email(
                smtp_server=config['smtp_server'],
                smtp_port=int(config['smtp_port']),
                login=config['login'],
                password=config['password'],
                sender_email=config.get('sender_email', config['login']),
                recipient_list=recipients,
                subject="Тестовое письмо от авторассылки",
                body="Это тестовое письмо. Если вы его получили, настройки SMTP верны.",
                encryption_type=config.get('encryption_type', 'SSL'),
                skip_cert_check=config.get('skip_cert_check', False),
                login_format=config.get('login_format', 'domain\\user'),
                tls_version=config.get('tls_version', 'TLSv1_2'),
                force_starttls_on_443=config.get('force_starttls_on_443', False),
                custom_ciphers=config.get('custom_ciphers', '')
            )
            messagebox.showinfo("Успех", f"Тестовое письмо отправлено на {recipients}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось отправить письмо:\n{str(e)}")

    def setup_schedule(self, shift_type):
        if platform.system() != 'Windows':
            messagebox.showerror("Ошибка", "Настройка расписания поддерживается только в Windows.")
            return
        config = load_config()
        if not config:
            messagebox.showerror("Ошибка", "Сначала сохраните настройки.")
            return

        if shift_type == 'day':
            task_name = "SendReportDay"
            time_str = "20:05"
            shift_param = "day"
        else:
            task_name = "SendReportNight"
            time_str = "08:05"
            shift_param = "night"

        script_path = os.path.abspath(sys.argv[0])
        python_exe = sys.executable

        command = f'"{python_exe}" "{script_path}" --send --shift {shift_param}'

        create_cmd = [
            "schtasks", "/create", "/tn", task_name,
            "/tr", command,
            "/sc", "daily",
            "/st", time_str,
            "/ru", "SYSTEM",
            "/rl", "HIGHEST",
            "/f"
        ]

        try:
            # Используем cp866 для декодирования вывода консоли Windows
            result = subprocess.run(create_cmd, capture_output=True)
            stderr = result.stderr.decode('cp866', errors='replace') if result.stderr else ''
            stdout = result.stdout.decode('cp866', errors='replace') if result.stdout else ''

            if result.returncode != 0:
                messagebox.showerror("Ошибка", f"Не удалось создать задачу:\n{stderr}")
            else:
                messagebox.showinfo("Успех", f"Задача '{task_name}' создана на {time_str} ежедневно.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Исключение при создании задачи:\n{str(e)}")

    def remove_schedule(self):
        if platform.system() != 'Windows':
            messagebox.showerror("Ошибка", "Удаление расписания поддерживается только в Windows.")
            return
        tasks = ["SendReportDay", "SendReportNight"]
        for task in tasks:
            try:
                result = subprocess.run(["schtasks", "/delete", "/tn", task, "/f"], capture_output=True)
                # Декодируем вывод, игнорируем ошибки
                stderr = result.stderr.decode('cp866', errors='replace') if result.stderr else ''
                if result.returncode != 0 and "не найдена" not in stderr.lower():
                    # Если задача не найдена - это нормально, игнорируем
                    pass
            except:
                pass
        messagebox.showinfo("Готово", "Задачи удалены (если существовали).")

# ============================
# ТОЧКА ВХОДА
# ============================
if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--send":
        shift = None
        if len(sys.argv) > 2 and sys.argv[2] == "--shift":
            if len(sys.argv) > 3:
                shift = sys.argv[3]
        if shift not in ['day', 'night']:
            print("❌ Укажите смену: --shift day или --shift night")
            sys.exit(1)
        send_scheduled_report(shift)
        sys.exit(0)

    root = tk.Tk()
    app = App(root)
    root.mainloop()