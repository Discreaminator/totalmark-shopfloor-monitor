#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ShopfloorMonitor — графическая панель здоровья сервера.

Отдельный процесс: здесь НЕТ кода сервера (нет Flask, poller, БД).
Только HTTP-клиент к уже запущенному app.py + чтение хвоста лога.

  GET  /api/health
  GET  /api/diag
  GET  /api/downtime
  GET  /api/lines_live   — линии, продукты, выпуск смены
  хвост logs/app.log (или app.log)

Запуск (рядом с работающим монитором):
  python app_gui.py
  python app_gui.py --url http://127.0.0.1:5000
"""

from __future__ import annotations

import argparse
import json
import os
import queue
import threading
import traceback
import urllib.error
import urllib.request
import webbrowser
from collections import deque
from datetime import datetime
from pathlib import Path
from tkinter import (
    BOTH,
    DISABLED,
    END,
    LEFT,
    NORMAL,
    RIGHT,
    WORD,
    BooleanVar,
    StringVar,
    Tk,
    ttk,
    messagebox,
)
from tkinter.scrolledtext import ScrolledText

# ---------------------------------------------------------------------------
# Тема
# ---------------------------------------------------------------------------
COLORS = {
    'bg': '#1a1d23',
    'panel': '#23272f',
    'panel2': '#2a2f39',
    'border': '#3a4050',
    'text': '#e8eaed',
    'muted': '#9aa0a6',
    'ok': '#2A9D8F',
    'warn': '#E9C46A',
    'bad': '#E63946',
    'accent': '#F7572E',
    'log_bg': '#12141a',
    'log_info': '#b0b8c4',
    'log_warn': '#E9C46A',
    'log_error': '#ff6b6b',
    'log_debug': '#6b7280',
}

STATUS_RU = {
    'ok': 'в работе',
    'warn': 'задержка',
    'bad': 'остановка',
    'downtime': 'простой',
    'idle': 'нет данных',
}

DEFAULT_URL = os.environ.get('MONITOR_URL', 'http://127.0.0.1:5000')
REFRESH_MS = int(os.environ.get('GUI_REFRESH_MS', '2000'))
LOG_MAX_LINES = int(os.environ.get('GUI_LOG_MAX_LINES', '500'))
LOG_MAX_LINE_CHARS = int(os.environ.get('GUI_LOG_MAX_LINE_CHARS', '2000'))
LOG_READ_CHUNK = int(os.environ.get('GUI_LOG_READ_CHUNK', '65536'))  # байт за один read
POLL_HISTORY = 60


def _windows_dpi_scale() -> float:
    """Масштаб Windows (1.0 = 96 DPI). На 4K часто 1.5–2.5."""
    try:
        from ctypes import windll
        # Per-monitor DPI aware — до создания окон
        try:
            windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                windll.user32.SetProcessDPIAware()
            except Exception:
                pass
        try:
            dpi = int(windll.user32.GetDpiForSystem())
        except Exception:
            hdc = windll.user32.GetDC(0)
            dpi = int(windll.gdi32.GetDeviceCaps(hdc, 88))  # LOGPIXELSX
            windll.user32.ReleaseDC(0, hdc)
        if dpi > 0:
            return max(1.0, dpi / 96.0)
    except Exception:
        pass
    return 1.0


def _resolve_ui_scale(cli_scale: float | None = None) -> float:
    """
    Масштаб шрифтов (tk scaling).
    = DPI Windows × (--scale | GUI_SCALE).
    Без доп. «comfort» — иначе окно раздувается и не влезает в 4K.
    """
    dpi = _windows_dpi_scale()
    env = os.environ.get('GUI_SCALE', '').strip().replace(',', '.')
    extra = 1.0
    if cli_scale is not None:
        extra = float(cli_scale)
    elif env:
        try:
            extra = float(env)
        except ValueError:
            extra = 1.0
    return max(1.0, min(3.0, dpi * max(0.5, extra)))


def _app_dir() -> Path:
    return Path(__file__).resolve().parent


def _find_log_path() -> Path | None:
    base = _app_dir()
    for p in (base / 'logs' / 'app.log', base / 'app.log'):
        if p.is_file():
            return p
    return base / 'logs' / 'app.log'


def http_json(url: str, method: str = 'GET', timeout: float = 4.0) -> tuple[int, dict | list | None, str]:
    req = urllib.request.Request(url, method=method)
    req.add_header('Accept', 'application/json')
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode('utf-8', errors='replace')
            code = resp.getcode()
            try:
                return code, json.loads(raw), ''
            except json.JSONDecodeError:
                return code, None, raw[:200]
    except urllib.error.HTTPError as e:
        raw = e.read().decode('utf-8', errors='replace') if e.fp else ''
        try:
            body = json.loads(raw) if raw else None
        except json.JSONDecodeError:
            body = None
        return e.code, body, str(e)
    except Exception as e:
        return 0, None, str(e)


class LogTailer:
    def __init__(self, path: Path, out_q: queue.Queue, stop_event: threading.Event):
        self.path = path
        self.out_q = out_q
        self.stop_event = stop_event
        self._pos = 0
        self._inode = None

    def _q_put(self, item):
        """Не блокируем tailer: при переполнении очереди выкидываем старое."""
        try:
            self.out_q.put_nowait(item)
            return
        except queue.Full:
            pass
        try:
            self.out_q.get_nowait()
        except queue.Empty:
            pass
        try:
            self.out_q.put_nowait(item)
        except queue.Full:
            pass

    def run(self):
        while not self.stop_event.is_set():
            try:
                if not self.path.is_file():
                    self._q_put(('meta', f'лог не найден: {self.path}'))
                    self.stop_event.wait(2.0)
                    continue
                st = self.path.stat()
                inode = getattr(st, 'st_ino', None) or st.st_mtime
                size = st.st_size
                if self._inode != inode or size < self._pos:
                    self._inode = inode
                    self._pos = max(0, size - 80_000)
                    self._q_put(('reset', ''))
                with open(self.path, 'r', encoding='utf-8', errors='replace') as f:
                    f.seek(self._pos)
                    # Ограниченный chunk — иначе огромный app.log заливает очередь/Tk
                    chunk = f.read(LOG_READ_CHUNK)
                    self._pos = f.tell()
                if chunk:
                    for line in chunk.splitlines():
                        if line.strip():
                            if len(line) > LOG_MAX_LINE_CHARS:
                                line = line[: LOG_MAX_LINE_CHARS - 1] + '…'
                            self._q_put(('line', line))
            except Exception as e:
                self._q_put(('meta', f'ошибка чтения лога: {e}'))
            self.stop_event.wait(0.6)


class ServerGui:
    def __init__(self, root: Tk, base_url: str, ui_scale: float = 1.0):
        self.root = root
        self.base_url = base_url.rstrip('/')
        self.ui_scale = max(1.0, float(ui_scale))
        self.stop_event = threading.Event()
        self.log_q: queue.Queue = queue.Queue(maxsize=500)
        self.poll_dur_hist: deque[float] = deque(maxlen=POLL_HISTORY)
        self._log_filter_errors = BooleanVar(value=False)
        self._url_var = StringVar(value=self.base_url)
        self._status_var = StringVar(value='подключение…')
        self._metric_vars: dict[str, StringVar] = {}
        self._log_path = _find_log_path()
        self._side_lock = threading.Lock()
        self._side_busy = False
        self._log_line_count = 0
        self._tick_after_id = None
        self._drain_after_id = None

        # Tk: pixels per point. 96 DPI → 96/72 ≈ 1.333
        try:
            root.tk.call('tk', 'scaling', 96.0 / 72.0 * self.ui_scale)
        except Exception:
            pass

        self._build()
        self._start_log_tail()
        self._schedule_tick(200)
        self._schedule_drain(300)
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)

    def _s(self, px: int) -> int:
        """
        Отступы / rowheight. Не умножаем на полный DPI:
        шрифты уже растут через tk scaling, иначе панель раздувается.
        """
        layout = 1.0 + (self.ui_scale - 1.0) * 0.25
        layout = max(1.0, min(1.35, layout))
        return max(1, int(round(px * layout)))

    def _fs(self, pt: int) -> tuple:
        """Пункты шрифта: масштаб уже в tk scaling — не умножать второй раз."""
        return ('Segoe UI', max(9, int(pt)))

    def _fsb(self, pt: int) -> tuple:
        return ('Segoe UI', max(9, int(pt)), 'bold')

    def _mono(self, pt: int) -> tuple:
        return ('Consolas', max(9, int(pt)))

    def _fit_to_screen(self):
        """Окно ≤ ~90% экрана, по центру — независимо от DPI-множителя."""
        r = self.root
        r.update_idletasks()
        sw = int(r.winfo_screenwidth())
        sh = int(r.winfo_screenheight())
        w = min(1400, int(sw * 0.90))
        h = min(900, int(sh * 0.86))
        w = max(800, min(w, sw - 32))
        h = max(600, min(h, sh - 64))
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        r.geometry(f'{w}x{h}+{x}+{y}')
        r.minsize(min(800, sw), min(600, sh))
        self._screen_h = sh

    def _mv(self, key: str, default: str = '—') -> StringVar:
        if key not in self._metric_vars:
            self._metric_vars[key] = StringVar(value=default)
        return self._metric_vars[key]

    def _bump_scale(self, delta: float):
        # --scale это множитель поверх DPI, не абсолютный ui_scale
        dpi = max(_windows_dpi_scale(), 1.0)
        cur_extra = self.ui_scale / dpi
        new_extra = max(0.75, min(2.0, round(cur_extra + delta, 2)))
        messagebox.showinfo(
            'Масштаб',
            f'Сейчас шрифты ×{self.ui_scale:.2f} (DPI {dpi:.2f} × {cur_extra:.2f}).\n\n'
            f'Перезапуск крупнее/мельче:\n'
            f'  python app_gui.py --url {self.base_url} --scale {new_extra:.2f}',
        )

    def _build(self):
        r = self.root
        r.title(f'ShopfloorMonitor · панель сервера · шрифт ×{self.ui_scale:.2f}')
        self._fit_to_screen()
        r.configure(bg=COLORS['bg'])

        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
        style.configure('.', background=COLORS['bg'], foreground=COLORS['text'],
                        fieldbackground=COLORS['panel'], font=self._fs(10))
        style.configure('TFrame', background=COLORS['bg'])
        style.configure('Panel.TFrame', background=COLORS['panel'])
        style.configure('TLabel', background=COLORS['bg'], foreground=COLORS['text'], font=self._fs(10))
        style.configure('Muted.TLabel', background=COLORS['bg'], foreground=COLORS['muted'], font=self._fs(10))
        style.configure('Panel.TLabel', background=COLORS['panel'], foreground=COLORS['text'], font=self._fs(10))
        style.configure('PanelMuted.TLabel', background=COLORS['panel'], foreground=COLORS['muted'], font=self._fs(9))
        style.configure('Header.TLabel', background=COLORS['bg'], foreground=COLORS['text'], font=self._fsb(16))
        style.configure('Health.TLabel', background=COLORS['panel'], foreground=COLORS['ok'], font=self._fsb(22))
        style.configure('TButton', padding=self._s(6), font=self._fs(10))
        style.configure('TCheckbutton', background=COLORS['bg'], foreground=COLORS['text'], font=self._fs(10))
        style.configure('TEntry', font=self._fs(10))
        style.configure(
            'Treeview',
            background=COLORS['panel2'],
            foreground=COLORS['text'],
            fieldbackground=COLORS['panel2'],
            rowheight=self._s(28),
            borderwidth=0,
            font=self._fs(10),
        )
        style.configure('Treeview.Heading', background=COLORS['panel'], foreground=COLORS['muted'],
                        font=self._fsb(10))
        style.map('Treeview', background=[('selected', COLORS['border'])])

        top = ttk.Frame(r)
        top.pack(fill='x', padx=self._s(14), pady=(self._s(12), self._s(6)))
        ttk.Label(top, text='ShopfloorMonitor · сервер', style='Header.TLabel').pack(side=LEFT)
        ttk.Label(top, textvariable=self._status_var, style='Muted.TLabel').pack(side=RIGHT)

        url_row = ttk.Frame(r)
        url_row.pack(fill='x', padx=self._s(14), pady=(0, self._s(8)))
        ttk.Label(url_row, text='URL', style='Muted.TLabel').pack(side=LEFT, padx=(0, self._s(6)))
        ttk.Entry(url_row, textvariable=self._url_var, width=42).pack(side=LEFT, padx=(0, self._s(8)))
        ttk.Button(url_row, text='Применить', command=self._apply_url).pack(side=LEFT, padx=2)
        # Только один refresh — не вызывать _tick_api (иначе плодятся after-цепочки)
        ttk.Button(url_row, text='Обновить', command=self._refresh_once).pack(side=LEFT, padx=2)
        ttk.Button(url_row, text='Монитор', command=self._open_monitor).pack(side=LEFT, padx=2)
        ttk.Button(url_row, text='Dump stacks', command=self._dump_stacks).pack(side=LEFT, padx=2)
        ttk.Button(url_row, text='Папка логов', command=self._open_logs_dir).pack(side=LEFT, padx=2)
        ttk.Button(url_row, text='A−', width=3, command=lambda: self._bump_scale(-0.15)).pack(side=RIGHT, padx=2)
        ttk.Button(url_row, text='A+', width=3, command=lambda: self._bump_scale(0.15)).pack(side=RIGHT, padx=2)

        banner = ttk.Frame(r, style='Panel.TFrame')
        banner.pack(fill='x', padx=self._s(14), pady=self._s(6))
        inner = ttk.Frame(banner, style='Panel.TFrame')
        inner.pack(fill='x', padx=self._s(14), pady=self._s(12))
        self.health_label = ttk.Label(inner, text='…', style='Health.TLabel')
        self.health_label.pack(side=LEFT)
        ttk.Label(inner, textvariable=self._mv('health_detail'), style='PanelMuted.TLabel').pack(
            side=LEFT, padx=self._s(16)
        )

        kpi = ttk.Frame(r)
        kpi.pack(fill='x', padx=self._s(14), pady=self._s(6))
        for key, title in (
            ('kpi_age', 'Возраст poll'),
            ('kpi_clients', 'Клиенты'),
            ('kpi_emit', 'Emit queue'),
            ('kpi_lines_ok', 'Линии в работе'),
            ('kpi_downtime', 'Простои open'),
            ('kpi_report', 'Последний отчёт'),
        ):
            cell = ttk.Frame(kpi, style='Panel.TFrame')
            cell.pack(side=LEFT, expand=True, fill='x', padx=3)
            ttk.Label(cell, text=title, style='PanelMuted.TLabel').pack(
                anchor='w', padx=self._s(10), pady=(self._s(8), 0)
            )
            ttk.Label(
                cell, textvariable=self._mv(key), style='Panel.TLabel', font=self._fsb(14)
            ).pack(anchor='w', padx=self._s(10), pady=(2, self._s(10)))

        # --- линии / продукты ---
        lines_wrap = ttk.Frame(r, style='Panel.TFrame')
        lines_wrap.pack(fill='both', expand=True, padx=self._s(14), pady=self._s(6))
        head = ttk.Frame(lines_wrap, style='Panel.TFrame')
        head.pack(fill='x', padx=self._s(12), pady=(self._s(10), self._s(4)))
        ttk.Label(head, text='Линии и продукты', style='Panel.TLabel',
                  font=self._fsb(11)).pack(side=LEFT)
        ttk.Label(head, textvariable=self._mv('lines_shift'), style='PanelMuted.TLabel').pack(side=RIGHT)

        cols = ('status', 'line', 'product', 'gtin', 'qty', 'dm', 'age', 'types', 'reason')
        self.lines = ttk.Treeview(lines_wrap, columns=cols, show='headings', height=9)
        headings = {
            'status': ('Статус', 100),
            'line': ('Линия', 170),
            'product': ('Продукт', 280),
            'gtin': ('GTIN', 130),
            'qty': ('Смена L1/ТУ', 110),
            'dm': ('Смена DM', 100),
            'age': ('Тишина', 90),
            'types': ('Типы', 100),
            'reason': ('Простой', 130),
        }
        for c, (title, wcol) in headings.items():
            self.lines.heading(c, text=title)
            self.lines.column(
                c, width=self._s(wcol),
                anchor='w' if c in ('line', 'product', 'reason') else 'center',
            )
        self.lines.tag_configure('ok', foreground=COLORS['ok'])
        self.lines.tag_configure('warn', foreground=COLORS['warn'])
        self.lines.tag_configure('bad', foreground=COLORS['bad'])
        self.lines.tag_configure('downtime', foreground=COLORS['accent'])
        self.lines.tag_configure('idle', foreground=COLORS['muted'])
        sy = ttk.Scrollbar(lines_wrap, orient='vertical', command=self.lines.yview)
        self.lines.configure(yscrollcommand=sy.set)
        self.lines.pack(side=LEFT, fill=BOTH, expand=True, padx=(self._s(10), 0), pady=(0, self._s(10)))
        sy.pack(side=RIGHT, fill='y', padx=(0, self._s(10)), pady=(0, self._s(10)))

        # --- mid: poller + threads ---
        mid = ttk.Frame(r)
        mid.pack(fill='x', padx=self._s(14), pady=self._s(6))
        mid.columnconfigure(0, weight=1, uniform='mid')
        mid.columnconfigure(1, weight=1, uniform='mid')

        mid_h = min(self._s(240), max(180, int(getattr(self, '_screen_h', 900) * 0.22)))
        left = ttk.Frame(mid, style='Panel.TFrame', height=mid_h)
        left.grid(row=0, column=0, sticky='nsew', padx=(0, self._s(6)))
        left.grid_propagate(False)
        left.pack_propagate(False)

        ttk.Label(left, text='Poller', style='Panel.TLabel',
                  font=self._fsb(11)).pack(anchor='w', padx=self._s(12), pady=(self._s(10), self._s(4)))
        grid = ttk.Frame(left, style='Panel.TFrame')
        grid.pack(fill='x', padx=self._s(12), pady=(0, self._s(8)))
        metrics = [
            ('poll_count', 'Цикл #'),
            ('poll_gen', 'Generation'),
            ('poll_stage', 'Стадия'),
            ('poll_dur', 'Длит. цикла'),
            ('poll_fetch', 'Fetch'),
            ('poll_process', 'Process'),
            ('poll_rows', 'Строк'),
            ('poll_backlog', 'Backlog'),
            ('poll_alive', 'Поток'),
            ('poll_async', 'async_mode'),
            ('poll_wm', 'Watermarks'),
            ('poll_err', 'Ошибка'),
        ]
        for i, (key, title) in enumerate(metrics):
            row, col = divmod(i, 2)
            cell = ttk.Frame(grid, style='Panel.TFrame')
            cell.grid(row=row, column=col, sticky='ew', padx=4, pady=2)
            ttk.Label(cell, text=title, style='PanelMuted.TLabel').pack(anchor='w')
            ttk.Label(
                cell, textvariable=self._mv(key), style='Panel.TLabel', width=28, anchor='w'
            ).pack(anchor='w')
        grid.columnconfigure(0, weight=1, uniform='m')
        grid.columnconfigure(1, weight=1, uniform='m')
        ttk.Label(left, text='Длительность циклов', style='PanelMuted.TLabel').pack(
            anchor='w', padx=self._s(12)
        )
        self.hist_label = ttk.Label(
            left, text='—', style='Panel.TLabel', font=self._mono(9),
            width=48, anchor='w',
        )
        self.hist_label.pack(anchor='w', padx=self._s(12), pady=(0, self._s(10)))

        right = ttk.Frame(mid, style='Panel.TFrame', height=mid_h)
        right.grid(row=0, column=1, sticky='nsew', padx=(self._s(6), 0))
        right.grid_propagate(False)
        right.pack_propagate(False)

        ttk.Label(right, text='Потоки', style='Panel.TLabel',
                  font=self._fsb(11)).pack(anchor='w', padx=self._s(12), pady=(self._s(10), self._s(4)))
        self.threads = ttk.Treeview(
            right, columns=('name', 'alive', 'daemon'), show='headings', height=8
        )
        self.threads.heading('name', text='Имя')
        self.threads.heading('alive', text='Alive')
        self.threads.heading('daemon', text='Daemon')
        self.threads.column('name', width=self._s(240), stretch=True)
        self.threads.column('alive', width=self._s(70), anchor='center', stretch=False)
        self.threads.column('daemon', width=self._s(80), anchor='center', stretch=False)
        self.threads.pack(fill=BOTH, expand=True, padx=self._s(10), pady=(0, self._s(10)))

        # --- log ---
        log_head = ttk.Frame(r)
        log_head.pack(fill='x', padx=self._s(14), pady=(self._s(4), 2))
        ttk.Label(log_head, text='Лог', font=self._fsb(11)).pack(side=LEFT)
        ttk.Label(log_head, textvariable=self._mv('log_path'), style='Muted.TLabel').pack(
            side=LEFT, padx=self._s(10)
        )
        ttk.Checkbutton(
            log_head, text='только WARN/ERROR', variable=self._log_filter_errors
        ).pack(side=RIGHT)

        # undo=False критично: иначе Tk копит историю insert/delete → RSS растёт до ГБ
        self.log_text = ScrolledText(
            r, height=12, wrap=WORD, bg=COLORS['log_bg'], fg=COLORS['log_info'],
            insertbackground=COLORS['text'], relief='flat', borderwidth=0,
            font=self._mono(10), state=DISABLED, undo=False, maxundo=0,
        )
        self.log_text.pack(fill=BOTH, expand=True, padx=self._s(14), pady=(0, self._s(12)))
        for tag, color in (
            ('INFO', COLORS['log_info']),
            ('WARNING', COLORS['log_warn']),
            ('WARN', COLORS['log_warn']),
            ('ERROR', COLORS['log_error']),
            ('CRITICAL', COLORS['log_error']),
            ('DEBUG', COLORS['log_debug']),
            ('META', COLORS['muted']),
        ):
            self.log_text.tag_configure(tag, foreground=color)

        self._mv('log_path').set(str(self._log_path) if self._log_path else '—')

    def _apply_url(self):
        self.base_url = self._url_var.get().strip().rstrip('/')
        self._status_var.set(f'URL → {self.base_url}')
        self._refresh_once()

    def _open_monitor(self):
        webbrowser.open(self.base_url + '/')

    def _open_logs_dir(self):
        p = self._log_path.parent if self._log_path else _app_dir() / 'logs'
        p.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(p))  # type: ignore[attr-defined]
        except Exception:
            webbrowser.open(p.as_uri())

    def _dump_stacks(self):
        code, data, err = http_json(self.base_url + '/api/diag', method='POST', timeout=8.0)
        if code and data:
            dump = data.get('dump_file') if isinstance(data, dict) else None
            messagebox.showinfo('Dump stacks', f'HTTP {code}\nФайл: {dump or "—"}')
            if isinstance(data, dict):
                self._apply_diag(data)
        else:
            messagebox.showerror('Dump stacks', err or 'нет ответа')

    def _start_log_tail(self):
        if not self._log_path:
            return
        threading.Thread(
            target=LogTailer(self._log_path, self.log_q, self.stop_event).run,
            daemon=True,
            name='log-tail',
        ).start()

    def _on_close(self):
        self.stop_event.set()
        for aid in (self._tick_after_id, self._drain_after_id):
            if aid is not None:
                try:
                    self.root.after_cancel(aid)
                except Exception:
                    pass
        self.root.destroy()

    def _schedule_tick(self, delay_ms: int):
        if self.stop_event.is_set():
            return
        if self._tick_after_id is not None:
            try:
                self.root.after_cancel(self._tick_after_id)
            except Exception:
                pass
        self._tick_after_id = self.root.after(delay_ms, self._tick_api)

    def _schedule_drain(self, delay_ms: int):
        if self.stop_event.is_set():
            return
        if self._drain_after_id is not None:
            try:
                self.root.after_cancel(self._drain_after_id)
            except Exception:
                pass
        self._drain_after_id = self.root.after(delay_ms, self._drain_log_q)

    def _tick_api(self):
        self._tick_after_id = None
        if self.stop_event.is_set():
            return
        try:
            self._refresh_once()
        except Exception:
            self._status_var.set('ошибка обновления')
            traceback.print_exc()
        self._schedule_tick(REFRESH_MS)

    def _refresh_once(self):
        code, health, err = http_json(self.base_url + '/api/health', timeout=3.5)
        now = datetime.now().strftime('%H:%M:%S')
        if not code or not isinstance(health, dict):
            self._set_health_banner('НЕТ СВЯЗИ', COLORS['bad'], err or 'сервер не отвечает')
            self._status_var.set(f'{now} · offline · {err}')
            return

        status = (health.get('status') or '').lower()
        stale = status == 'stale' or code == 503
        if stale:
            self._set_health_banner('STALE', COLORS['bad'], f"poller: {health.get('stale_reason') or 'stale'}")
        else:
            self._set_health_banner('HEALTH OK', COLORS['ok'], 'poller жив')

        self._mv('kpi_age').set(self._fmt_sec(health.get('last_poll_age_sec')))
        self._mv('kpi_clients').set(str(health.get('clients', '—')))
        self._mv('kpi_emit').set(str(health.get('emit_queue', '—')))

        report_at = health.get('last_report_at') or '—'
        report_dur = health.get('last_report_duration_sec')
        if report_dur is not None:
            self._mv('kpi_report').set(f'{self._short_ts(report_at)} · {report_dur} с')
        else:
            self._mv('kpi_report').set(self._short_ts(report_at))

        self._mv('poll_count').set(str(health.get('poll_count', '—')))
        self._mv('poll_gen').set(str(health.get('poller_generation', '—')))
        self._mv('poll_stage').set(str(health.get('poll_stage') or '—'))
        dur = health.get('last_poll_duration_sec')
        self._mv('poll_dur').set(self._fmt_sec(dur))
        if isinstance(dur, (int, float)):
            self.poll_dur_hist.append(float(dur))
            self._render_hist()
        self._mv('poll_fetch').set(self._fmt_sec(health.get('last_fetch_sec')))
        self._mv('poll_process').set(self._fmt_sec(health.get('last_process_sec')))
        self._mv('poll_rows').set(str(health.get('last_poll_rows', '—')))
        self._mv('poll_backlog').set(self._fmt_sec(health.get('backlog_sec')))
        self._mv('poll_alive').set('alive' if health.get('poller_thread_alive') else 'DEAD')
        self._mv('poll_async').set(str(health.get('async_mode') or '—'))
        wm = health.get('watermarks') or {}
        if isinstance(wm, dict) and wm:
            parts = [f'{k} {self._short_ts(v)}' for k, v in list(wm.items())[:4]]
            self._mv('poll_wm').set(self._clip(' · '.join(parts), 40))
        else:
            self._mv('poll_wm').set('—')
        self._mv('poll_err').set(self._clip(str(health.get('last_poll_error') or '—'), 40))
        self._mv('health_detail').set(
            f'HTTP {code} · gen {health.get("poller_generation")} · '
            f'stage {health.get("poll_stage") or "—"} · report_busy={health.get("report_busy")}'
        )
        self._status_var.set(f'{now} · {self.base_url} · refresh {REFRESH_MS} мс')

        with self._side_lock:
            if self._side_busy:
                return
            self._side_busy = True
        threading.Thread(target=self._fetch_side, daemon=True).start()

    def _fetch_side(self):
        try:
            code, diag, _ = http_json(self.base_url + '/api/diag', timeout=3.5)
            if code and isinstance(diag, dict):
                self.root.after(0, lambda d=diag: self._apply_diag(d))

            code2, dt, _ = http_json(self.base_url + '/api/downtime', timeout=3.5)
            if code2 and isinstance(dt, dict):
                self.root.after(0, lambda d=dt: self._apply_downtime(d))

            code3, live, err3 = http_json(self.base_url + '/api/lines_live', timeout=4.0)
            if code3 and isinstance(live, dict):
                self.root.after(0, lambda d=live: self._apply_lines(d))
            elif code3 == 404:
                self.root.after(0, lambda: self._mv('lines_shift').set(
                    'нет /api/lines_live — обновите app.py на сервере'
                ))
            elif err3:
                self.root.after(0, lambda e=err3: self._mv('lines_shift').set(f'lines_live: {e}') )
        finally:
            with self._side_lock:
                self._side_busy = False

    def _apply_diag(self, diag: dict):
        threads = diag.get('threads') or []
        for item in self.threads.get_children():
            self.threads.delete(item)
        for t in threads:
            self.threads.insert(
                '', END,
                values=(
                    t.get('name', ''),
                    'да' if t.get('alive') else 'нет',
                    'да' if t.get('daemon') else 'нет',
                ),
            )

    def _apply_downtime(self, dt: dict):
        open_map = dt.get('open') or {}
        self._mv('kpi_downtime').set(str(len(open_map)))

    def _apply_lines(self, live: dict):
        rows = live.get('lines') or []
        shift = self._short_ts(live.get('shift_start'))
        self._mv('lines_shift').set(f'смена с {shift} · линий {len(rows)}')
        n_ok = sum(1 for r in rows if r.get('status') == 'ok')
        self._mv('kpi_lines_ok').set(f'{n_ok} / {len(rows)}')

        for item in self.lines.get_children():
            self.lines.delete(item)
        for r in rows:
            st = r.get('status') or 'idle'
            age = r.get('age_sec')
            if age is None:
                age_s = '—'
            elif age < 60:
                age_s = f'{int(age)} с'
            else:
                age_s = f'{age / 60:.1f} мин'
            product = r.get('product') or '—'
            gtin = r.get('gtin') or '—'
            types = ','.join(r.get('entity_types') or []) or (r.get('last_entity_type') or '—')
            reason = ''
            if r.get('downtime_open'):
                reason = r.get('downtime_reason') or 'открыт'
            self.lines.insert(
                '', END,
                tags=(st,),
                values=(
                    STATUS_RU.get(st, st),
                    r.get('display_name') or r.get('line_id') or '',
                    product,
                    gtin,
                    r.get('qty_shift', 0),
                    r.get('qty_dm_shift', 0),
                    age_s,
                    types,
                    reason,
                ),
            )

    def _set_health_banner(self, title: str, color: str, detail: str):
        self.health_label.configure(text=title, foreground=color)
        self._mv('health_detail').set(detail)

    def _render_hist(self):
        vals = list(self.poll_dur_hist)
        if not vals:
            self.hist_label.configure(text='—')
            return
        mx = max(vals) or 1.0
        blocks = '▁▂▃▄▅▆▇█'
        # всегда 40 символов — иначе полоска растёт и двигает layout
        n = 40
        recent = list(vals)[-n:]
        if len(recent) < n:
            recent = [0.0] * (n - len(recent)) + recent
        bars = []
        for v in recent:
            h = int(round(7 * (float(v) / mx))) if mx else 0
            bars.append(blocks[max(0, min(7, h))])
        self.hist_label.configure(
            text=''.join(bars) + f'  {vals[-1]:.2f}/{mx:.2f}s'
        )

    @staticmethod
    def _clip(s: str, n: int) -> str:
        s = str(s or '')
        return s if len(s) <= n else s[: n - 1] + '…'

    @staticmethod
    def _fmt_sec(v) -> str:
        if v is None:
            return '—'
        try:
            return f'{float(v):.1f} с'
        except (TypeError, ValueError):
            return str(v)

    @staticmethod
    def _short_ts(v) -> str:
        if not v:
            return '—'
        s = str(v)
        if 'T' in s:
            return s.split('T', 1)[-1][:8]
        return s[-8:]

    def _drain_log_q(self):
        self._drain_after_id = None
        if self.stop_event.is_set():
            return
        # Лимит за тик — не блокируем UI на тысячах строк
        batch = 0
        try:
            while batch < 80:
                kind, payload = self.log_q.get_nowait()
                batch += 1
                if kind == 'reset':
                    self._log_clear()
                elif kind == 'meta':
                    self._log_append(payload, 'META')
                else:
                    self._log_append(payload, self._level_of(payload))
        except queue.Empty:
            pass
        self._schedule_drain(250)

    def _log_clear(self):
        self.log_text.configure(state=NORMAL)
        self.log_text.delete('1.0', END)
        self.log_text.configure(state=DISABLED)
        self._log_line_count = 0

    def _log_append(self, line: str, tag: str):
        """
        Без get('1.0', END) на каждую строку — это была утечка:
        полный снимок Text + фрагментация Tcl/Tk → RSS до ~2 ГБ.
        Счётчик строк + undo=False держат память плоской.
        """
        if self._log_filter_errors.get():
            up = line.upper()
            if 'ERROR' not in up and 'WARN' not in up and 'CRITICAL' not in up:
                if tag not in ('ERROR', 'WARNING', 'WARN', 'CRITICAL', 'META'):
                    return
        if len(line) > LOG_MAX_LINE_CHARS:
            line = line[: LOG_MAX_LINE_CHARS - 1] + '…'
        self.log_text.configure(state=NORMAL)
        self.log_text.insert(END, line + '\n', tag)
        self._log_line_count += 1
        if self._log_line_count > LOG_MAX_LINES:
            cut = self._log_line_count - LOG_MAX_LINES
            self.log_text.delete('1.0', f'{cut + 1}.0')
            self._log_line_count = LOG_MAX_LINES
        # see(END) только периодически — иначе layout на каждой строке
        if self._log_line_count % 20 == 0 or self._log_line_count <= 3:
            self.log_text.see(END)
        self.log_text.configure(state=DISABLED)

    @staticmethod
    def _level_of(line: str) -> str:
        up = line.upper()
        for lvl in ('CRITICAL', 'ERROR', 'WARNING', 'WARN', 'DEBUG', 'INFO'):
            if f' {lvl} ' in f' {up} ' or f'-{lvl}-' in up or f'|{lvl}|' in up:
                return 'WARN' if lvl == 'WARNING' else lvl
        return 'INFO'


def main():
    # DPI awareness выставляется внутри _windows_dpi_scale() до Tk()
    parser = argparse.ArgumentParser(description='ShopfloorMonitor server health GUI (отдельный клиент)')
    parser.add_argument('--url', default=DEFAULT_URL, help='Base URL монитора')
    parser.add_argument(
        '--scale', type=float, default=None,
        help='Множитель поверх DPI Windows (1.0 = как система; 1.25 крупнее шрифт, окно всё равно влезет)',
    )
    args = parser.parse_args()

    ui_scale = _resolve_ui_scale(args.scale)
    root = Tk()
    ServerGui(root, args.url, ui_scale=ui_scale)
    root.mainloop()


if __name__ == '__main__':
    main()
