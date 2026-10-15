# BikeTrack Pro - English-only cleanup build

"""BikeTrack Pro - A bicycle shop customer tracking and receipt printing application."""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import re
import sys
import qrcode
from PIL import Image, ImageTk, ImageDraw, ImageFont
import io
from datetime import datetime, timedelta
import sqlite3
import shutil
import win32print
try:
    import win32ui
    import win32con
    from PIL import ImageWin
except Exception:
    win32ui = None
    win32con = None
    ImageWin = None
import json
from pathlib import Path
import itertools

def _ctx_get_focus_widget(root_widget):
    try:
        return root_widget.focus_get()
    except Exception:
        return None

def _ctx_cut(root_widget):
    w = _ctx_get_focus_widget(root_widget)
    try:
        if w:
            w.event_generate('<<Cut>>')
    except Exception:
        pass

def _ctx_copy(root_widget):
    w = _ctx_get_focus_widget(root_widget)
    try:
        if w:
            w.event_generate('<<Copy>>')
    except Exception:
        pass

def _ctx_paste(root_widget):
    w = _ctx_get_focus_widget(root_widget)
    try:
        if w:
            w.event_generate('<<Paste>>')
    except Exception:
        pass

def _ctx_select_all(root_widget):
    w = _ctx_get_focus_widget(root_widget)
    if not w:
        return
    try:
        if isinstance(w, tk.Entry) or w.winfo_class() in ('Entry', 'TEntry'):
            w.select_range(0, 'end')
            w.icursor('end')
            return
        if isinstance(w, tk.Text) or w.winfo_class() == 'Text':
            w.tag_add('sel', '1.0', 'end-1c')
            return
    except Exception:
        pass

def create_global_context_menu(root_widget):
    """Create a global right-click context menu for text inputs (Entry/TEntry/Text)."""
    menu = tk.Menu(root_widget, tearoff=0)
    menu.add_command(label='Cut', command=lambda: _ctx_cut(root_widget))
    menu.add_command(label='Copy', command=lambda: _ctx_copy(root_widget))
    menu.add_command(label='Paste', command=lambda: _ctx_paste(root_widget))
    menu.add_separator()
    menu.add_command(label='Select All', command=lambda: _ctx_select_all(root_widget))
    root_widget._global_context_menu = menu

    def _popup(event):
        cls = event.widget.winfo_class()
        if cls in ('Entry', 'TEntry', 'Text'):
            try:
                menu.tk.call('tk_popup', menu, event.x_root, event.y_root)
            finally:
                menu.grab_release()
    for cls in ('Entry', 'TEntry', 'Text'):
        root_widget.bind_class(cls, '<Button-3>', _popup, add='+')
import Levenshtein
import tkinter.font as tkfont
from typing import List, Dict, Tuple, Optional, Any
import threading
import queue
import concurrent.futures
import functools
import time
import webbrowser
_thread_pool = concurrent.futures.ThreadPoolExecutor(max_workers=4)
_db_connections = queue.Queue(maxsize=5)
_app_cache = {}
DB_DEFAULT_BUSY_TIMEOUT_MS = 5000
DB_LOCK_RETRY_BUSY_TIMEOUT_MS = 250
DB_LOCK_MAX_WAIT_SECONDS = 5.0
DB_LOCK_RETRY_DELAY_SECONDS = 0.15

def get_db_connection() -> sqlite3.Connection:
    """Get a connection from the connection pool or create a new one."""
    try:
        conn = _db_connections.get(block=False)
    except queue.Empty:
        conn = sqlite3.connect(DATABASE_FILE, timeout=DB_DEFAULT_BUSY_TIMEOUT_MS / 1000, check_same_thread=False)
        conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA journal_mode=DELETE')
    conn.execute(f'PRAGMA busy_timeout={DB_DEFAULT_BUSY_TIMEOUT_MS}')
    conn.execute('PRAGMA synchronous=NORMAL')
    conn.execute('PRAGMA cache_size=10000')
    conn.execute('PRAGMA temp_store=MEMORY')
    conn.execute('PRAGMA foreign_keys=ON')
    return conn

def is_database_locked_error(error: Exception) -> bool:
    """Return True when SQLite reports that the database is locked/busy."""
    message = str(error).lower()
    return 'database is locked' in message or 'database table is locked' in message or 'database schema is locked' in message or ('database is busy' in message) or ('locked' in message and 'database' in message)

def set_database_lock_indicator(locked: bool, auto_clear_ms: Optional[int]=None) -> None:
    """Turn the main input fields red while the shared SQLite database is locked."""

    def _apply() -> None:
        try:
            widgets = []
            if phone_entry is not None:
                widgets.append(phone_entry)
            if model_entry is not None:
                widgets.append(model_entry)
            widgets.extend([w for w in issue_entries if w is not None])
            if locked:
                for widget in widgets:
                    try:
                        if not hasattr(widget, '_normal_database_bg'):
                            widget._normal_database_bg = widget.cget('background')
                        widget.configure(background='#ffb3b3')
                    except Exception:
                        pass
                if status_label is not None:
                    message = 'Database is locked - waiting...'
                    status_label.config(text=message, fg='red')
            else:
                for widget in widgets:
                    try:
                        normal_bg = getattr(widget, '_normal_database_bg', 'white')
                        widget.configure(background=normal_bg)
                    except Exception:
                        pass
                try:
                    check_customer_flag_warning()
                except Exception:
                    pass
            if locked and auto_clear_ms and (root is not None):
                root.after(auto_clear_ms, lambda: set_database_lock_indicator(False))
        except Exception:
            pass
    try:
        if root is not None:
            root.after(0, _apply)
    except Exception:
        pass

def release_db_connection(conn: sqlite3.Connection) -> None:
    """Return a connection to the pool or close it if the pool is full."""
    try:
        _db_connections.put(conn, block=False)
    except queue.Full:
        conn.close()

def run_in_thread(func):
    """Decorator to run a function in a thread from the thread pool."""

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        return _thread_pool.submit(func, *args, **kwargs)
    return wrapper

def check_required_settings() -> None:
    """Check if required settings are configured and show warnings if not."""
    text = ENGLISH_TEXT
    settings = load_settings()
    warnings = []
    backup_folder = resolve_backup_folder_path(settings.get('backup_folder', 'backup_db'))
    if not os.path.exists(backup_folder):
        warnings.append('Backup folder is not set or does not exist.')
    if not (settings.get('customer_printer') or settings.get('shop_printer') or settings.get('printer')):
        warnings.append('No receipt printer is selected.')
    if warnings:
        warning_message = '\n'.join(warnings)
        warning_message += '\n\n' + 'Please configure these settings in the Settings menu.'
        messagebox.showwarning('Configuration Warning', warning_message)

def _toast_host():
    """Return the best toplevel host window for showing toasts."""
    try:
        w = root.focus_get()
        if w is not None:
            return w.winfo_toplevel()
    except Exception:
        pass
    try:
        return tk._default_root or root
    except Exception:
        return root

def show_toast(parent, message: str, duration: int=3000):
    """Display a small temporary message at the bottom of the given window."""
    try:
        host = parent or _toast_host()
    except Exception:
        host = _toast_host()
    if host is None:
        return
    try:
        if hasattr(host, '_vh_toast') and host._vh_toast and host._vh_toast.winfo_exists():
            host._vh_toast.destroy()
    except Exception:
        pass
    toast = tk.Label(host, text=str(message), bg='#333333', fg='#ffffff', padx=10, pady=6, bd=0, relief='flat')
    toast.place(relx=0.5, rely=1.0, anchor='s', y=-10)
    host._vh_toast = toast
    toast.after(max(500, int(duration)), toast.destroy)

def app_info(title: str, message: str, parent=None, duration: int=3000):
    """Replacement for messagebox.showinfo using non-blocking toast."""
    show_toast(parent or _toast_host(), message, duration=duration)


def on_phone_enter(event=None):
    """Handle Enter on the phone field: open/update the separate visit-history window, then move focus on."""
    try:
        _do_phone_autorefresh(force=True, lift=True)
    except Exception:
        pass
    try:
        return focus_next_widget(event)
    except Exception:
        return 'break'


def reload_visit_history_dialog(dialog, phone: str) -> None:
    """Reload an already-open separate Visit History window for a new phone number."""
    try:
        if dialog is None or not dialog.winfo_exists():
            return
        clean = _phone_digits(phone)
        text = ENGLISH_TEXT
        tree = getattr(dialog, '_vh_tree', None)
        status_var = getattr(dialog, '_vh_status_var', None)
        refresh_button = getattr(dialog, '_vh_refresh_button', None)
        if not clean:
            if status_var:
                status_var.set(text.get('visit_history_empty', 'Phone number is empty.'))
            if tree and tree.winfo_exists():
                for item in tree.get_children():
                    tree.delete(item)
            return
        try:
            dialog.title(text.get('visit_history_title', 'Visit History') + ' – ' + clean)
        except Exception:
            pass
        dialog._vh_phone = clean
        if status_var:
            status_var.set('Loading visit history...')
        if refresh_button:
            try:
                refresh_button.config(state=tk.DISABLED)
            except Exception:
                pass

        def _thread():
            rows = []
            error = None
            conn = None
            try:
                conn = get_db_connection()
                cur = conn.cursor()
                cur.execute("""
                    SELECT datetime, phone, model,
                           issue1, issue2, issue3, issue4, issue5,
                           issue6, issue7, issue8, issue9, issue10
                    FROM customer_bikes
                    WHERE REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(phone,' ',''),'-',''),'+',''),'(',''),')','') LIKE ?
                    ORDER BY datetime DESC
                    """, (f'%{clean}%',))
                rows = cur.fetchall()
            except Exception as exc:
                error = exc
            finally:
                if conn:
                    try:
                        release_db_connection(conn)
                    except Exception:
                        pass

            def _populate():
                try:
                    if dialog is None or not dialog.winfo_exists():
                        return
                    current_phone = getattr(dialog, '_vh_phone', '')
                    if current_phone != clean:
                        return
                    tree = getattr(dialog, '_vh_tree', None)
                    if not (tree and tree.winfo_exists()):
                        return
                    for item in tree.get_children():
                        tree.delete(item)
                    if error is not None:
                        if status_var:
                            status_var.set(f"{text.get('visit_history_error', 'Error loading history')}: {error}")
                        return
                    for r in rows:
                        dt, row_phone, model, *issues = r
                        issues_list = [str(i).strip() for i in issues if i and str(i).strip()]
                        issues_str = ' | '.join(issues_list) if issues_list else '-'
                        tree.insert('', 'end', values=(visit_date_label(dt), visit_age_label(dt), row_phone, model, issues_str))
                    count = len(rows)
                    if status_var:
                        if count:
                            status_var.set(f'{count} previous visit(s) found for {clean}. Double-click a row to reuse the bike model.')
                        else:
                            status_var.set(text.get('visit_history_not_found', 'No history found for this number.'))
                finally:
                    if refresh_button:
                        try:
                            refresh_button.config(state=tk.NORMAL)
                        except Exception:
                            pass
            try:
                dialog.after(0, _populate)
            except Exception:
                pass
        threading.Thread(target=_thread, daemon=True).start()
    except Exception:
        pass


def _schedule_phone_autorefresh(delay_ms: int=500):
    """Debounced refresh/open for the separate phone visit-history window."""
    try:
        aft = getattr(root, '_vh_debounce_after_id', None)
        if aft:
            try:
                root.after_cancel(aft)
            except Exception:
                pass
        root._vh_debounce_after_id = root.after(delay_ms, _do_phone_autorefresh)
    except Exception:
        pass

def _phone_digits(value: str) -> str:
    """Return only digits from a phone-like value."""
    return ''.join((ch for ch in str(value or '') if ch.isdigit()))

def _add_months_safe(dt: datetime, months: int) -> datetime:
    """Add whole calendar months while keeping the day valid for the target month."""
    from calendar import monthrange
    month_index = dt.month - 1 + int(months)
    year = dt.year + month_index // 12
    month = month_index % 12 + 1
    day = min(dt.day, monthrange(year, month)[1])
    return dt.replace(year=year, month=month, day=day)

def parse_visit_datetime(dt_value: str) -> Optional[datetime]:
    """Parse stored visit datetime safely."""
    try:
        raw = str(dt_value or '').strip()
        if not raw:
            return None
        try:
            return datetime.strptime(raw, '%Y-%m-%d %H:%M:%S')
        except Exception:
            return datetime.fromisoformat(raw.replace('T', ' ')[:19])
    except Exception:
        return None

def visit_age_label(dt_value: str) -> str:
    """Return a compact calendar age label: years, months and days since dt_value."""
    try:
        start = parse_visit_datetime(dt_value)
        if not start:
            return '-'
        now = datetime.now()
        if start > now:
            return '0 d'
        total_months = (now.year - start.year) * 12 + (now.month - start.month)
        anchor = _add_months_safe(start, total_months)
        if anchor > now:
            total_months -= 1
            anchor = _add_months_safe(start, total_months)
        years = max(0, total_months // 12)
        months = max(0, total_months % 12)
        days = max(0, (now - anchor).days)
        return f'{years}y {months}mo {days}d'
    except Exception:
        return '-'

def visit_date_label(dt_value: str) -> str:
    """Return only the date part for a stored visit datetime."""
    dt = parse_visit_datetime(dt_value)
    return dt.strftime('%Y-%m-%d') if dt else str(dt_value or '').strip().split(' ')[0]

def visit_time_label(dt_value: str) -> str:
    """Return only the time part for a stored visit datetime."""
    dt = parse_visit_datetime(dt_value)
    return dt.strftime('%H:%M') if dt else str(dt_value or '').strip().split(' ')[1][:5] if ' ' in str(dt_value or '').strip() else ''


def _do_phone_autorefresh(force: bool=False, lift: bool=False):
    """Refresh the embedded visit-history panel for the typed phone number."""
    try:
        current = phone_entry.get().strip() if phone_entry else ''
        clean = _phone_digits(current)
        last = getattr(root, '_vh_last_phone', None)
        if not force and clean == last:
            return
        root._vh_last_phone = clean
        reload_main_visit_history(clean)
    except Exception:
        pass

CUSTOMER_FLAG_LEVELS = {'watch': {'title': 'Watch', 'color': '#fff3b0', 'warning': '⚠ WATCH CUSTOMER'}, 'problem': {'title': 'Problem', 'color': '#ffd08a', 'warning': '⚠ PROBLEM CUSTOMER'}, 'high_risk': {'title': 'High Risk', 'color': '#ff8a8a', 'warning': '⚠ HIGH RISK CUSTOMER'}}

def normalize_flag_phone(phone: str) -> str:
    """Normalize phone number for customer flag matching."""
    return ''.join((ch for ch in str(phone or '') if ch.isdigit()))

def get_customer_flag(phone: str) -> Optional[sqlite3.Row]:
    """Return saved customer flag for a phone number, if any."""
    clean_phone = normalize_flag_phone(phone)
    if not clean_phone:
        return None
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('\n            SELECT phone, flag_level, note, created_at, updated_at\n            FROM customer_flags\n            WHERE normalized_phone = ?\n            LIMIT 1\n            ', (clean_phone,))
        return cursor.fetchone()
    except Exception:
        return None
    finally:
        if conn:
            release_db_connection(conn)

def get_all_customer_flags() -> Dict[str, sqlite3.Row]:
    """Return all saved customer flags keyed by normalized phone number."""
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('\n            SELECT phone, normalized_phone, flag_level, note, created_at, updated_at\n            FROM customer_flags\n            ORDER BY updated_at DESC\n            ')
        return {str(row['normalized_phone']): row for row in cursor.fetchall()}
    except Exception:
        return {}
    finally:
        if conn:
            release_db_connection(conn)

def _same_calendar_day_for_year(base_dt: datetime, target_year: int) -> datetime:
    """Return the same month/day in another year; use February 28 for invalid leap dates."""
    try:
        return base_dt.replace(year=target_year)
    except ValueError:
        return base_dt.replace(year=target_year, day=28)

def _count_bike_records_between(start_dt: datetime, end_dt: datetime) -> int:
    """Count bike records in [start_dt, end_dt). Each row counts as one bike record."""
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('\n            SELECT COUNT(*)\n            FROM customer_bikes\n            WHERE datetime >= ? AND datetime < ?\n            ', (start_dt.strftime('%Y-%m-%d %H:%M:%S'), end_dt.strftime('%Y-%m-%d %H:%M:%S')))
        row = cursor.fetchone()
        return int(row[0] if row else 0)
    except Exception:
        return 0
    finally:
        if conn:
            release_db_connection(conn)

def _bike_statistics_for_year_offset(base_today: datetime, years_back: int) -> Dict[str, Tuple[int, str]]:
    """Build today/month/year-to-date bike statistics for the requested year offset."""
    target_day = _same_calendar_day_for_year(base_today, base_today.year - years_back)
    day_start = target_day.replace(hour=0, minute=0, second=0, microsecond=0)
    day_end = day_start + timedelta(days=1)
    month_start = day_start.replace(day=1)
    month_end = day_end
    year_start = day_start.replace(month=1, day=1)
    year_end = day_end
    return {'Today': (_count_bike_records_between(day_start, day_end), f"{day_start.strftime('%Y-%m-%d')}"), 'This month': (_count_bike_records_between(month_start, month_end), f"{month_start.strftime('%Y-%m-%d')} to {target_day.strftime('%Y-%m-%d')}"), 'This year': (_count_bike_records_between(year_start, year_end), f"{year_start.strftime('%Y-%m-%d')} to {target_day.strftime('%Y-%m-%d')}")}

def open_bike_statistics_dialog(parent=None) -> None:
    """Show bike record statistics without changing customer data."""
    global bike_statistics_window
    parent = parent or root
    if bike_statistics_window and tk.Toplevel.winfo_exists(bike_statistics_window):
        bike_statistics_window.lift()
        bike_statistics_window.focus_force()
        return
    dialog = tk.Toplevel(parent)
    bike_statistics_window = dialog
    dialog.title('Bike Statistics')
    dialog.geometry(get_saved_window_geometry('bike_statistics', '620x250'))
    dialog.minsize(560, 220)
    set_window_icon(dialog)

    def _close_dialog():
        global bike_statistics_window
        save_window_geometry(dialog, 'bike_statistics')
        dialog.destroy()
        bike_statistics_window = None
    dialog.protocol('WM_DELETE_WINDOW', _close_dialog)
    container = ttk.Frame(dialog)
    container.pack(fill='both', expand=True, padx=12, pady=12)
    ttk.Label(container, text='Bike Statistics', font=('Arial', 14, 'bold')).pack(anchor='w', pady=(0, 8))
    ttk.Label(container, text='Each saved bike record is counted. Repeated phone numbers are counted as separate records.', wraplength=580).pack(anchor='w', pady=(0, 10))
    today = datetime.now()
    current_stats = _bike_statistics_for_year_offset(today, 0)
    last_year_stats = _bike_statistics_for_year_offset(today, 1)
    two_years_stats = _bike_statistics_for_year_offset(today, 2)
    columns = ('period', 'current', 'last_year', 'two_years')
    tree = ttk.Treeview(container, columns=columns, show='headings', height=3)
    tree.heading('period', text='Period')
    tree.heading('current', text='Current')
    tree.heading('last_year', text='Same period last year')
    tree.heading('two_years', text='Same period 2 years ago')
    tree.column('period', width=110, anchor='w')
    tree.column('current', width=90, anchor='center')
    tree.column('last_year', width=165, anchor='center')
    tree.column('two_years', width=175, anchor='center')
    tree.pack(fill='x', pady=(0, 12))
    for label in ('Today', 'This month', 'This year'):
        tree.insert('', 'end', values=(label, f'{current_stats[label][0]} bikes', f'{last_year_stats[label][0]} bikes', f'{two_years_stats[label][0]} bikes'))
    ttk.Button(container, text='Close', command=_close_dialog).pack(anchor='e', pady=(10, 0))

def open_customer_flags_list_dialog(parent=None, initial_phone: str='') -> None:
    """Show and manage all marked phone numbers in one window."""
    parent = parent or root
    dialog = tk.Toplevel(parent)
    dialog.title('Marked Customers')
    dialog.geometry(get_saved_window_geometry('marked_customers', '840x520'))
    dialog.minsize(720, 420)
    set_window_icon(dialog)

    def _close_dialog():
        save_window_geometry(dialog, 'marked_customers')
        dialog.destroy()
    dialog.protocol('WM_DELETE_WINDOW', _close_dialog)
    frame = ttk.Frame(dialog)
    frame.pack(fill='both', expand=True, padx=10, pady=10)
    columns = ('level', 'phone', 'note', 'updated')
    tree = ttk.Treeview(frame, columns=columns, show='headings', height=12)
    tree.heading('level', text='Mark')
    tree.heading('phone', text='Phone')
    tree.heading('note', text='Note')
    tree.heading('updated', text='Updated')
    tree.column('level', width=110, anchor='center')
    tree.column('phone', width=160, anchor='center')
    tree.column('note', width=400, anchor='w')
    tree.column('updated', width=140, anchor='center')
    yscroll = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=yscroll.set)
    tree.pack(side='left', fill='both', expand=True)
    yscroll.pack(side='right', fill='y')
    for key, info in CUSTOMER_FLAG_LEVELS.items():
        tree.tag_configure(f'flag_{key}', background=info['color'])
    editor = ttk.LabelFrame(dialog, text='Manage Marked Number')
    editor.pack(fill='x', padx=10, pady=(0, 10))
    ttk.Label(editor, text='Phone:').grid(row=0, column=0, sticky='w', padx=8, pady=6)
    phone_var = tk.StringVar()
    try:
        current_phone = str(initial_phone or '').strip()
        if not current_phone:
            current_phone = phone_entry.get().strip() if phone_entry else ''
        if current_phone:
            phone_var.set(current_phone)
    except Exception:
        pass
    phone_box = ttk.Entry(editor, textvariable=phone_var, width=24)
    phone_box.grid(row=0, column=1, sticky='w', padx=8, pady=6)
    ttk.Label(editor, text='Mark:').grid(row=0, column=2, sticky='w', padx=8, pady=6)
    level_keys = ('watch', 'problem', 'high_risk')
    level_titles = [CUSTOMER_FLAG_LEVELS[k]['title'] for k in level_keys]
    title_to_key = {CUSTOMER_FLAG_LEVELS[k]['title']: k for k in level_keys}
    key_to_title = {k: CUSTOMER_FLAG_LEVELS[k]['title'] for k in level_keys}
    selected_title = tk.StringVar(value='Watch')
    level_combo = ttk.Combobox(editor, values=level_titles, textvariable=selected_title, state='readonly', width=18)
    level_combo.grid(row=0, column=3, sticky='w', padx=8, pady=6)
    ttk.Label(editor, text='Note:').grid(row=1, column=0, sticky='nw', padx=8, pady=6)
    note_text = tk.Text(editor, width=70, height=4)
    note_text.grid(row=1, column=1, columnspan=3, sticky='we', padx=8, pady=6)
    editor.columnconfigure(3, weight=1)

    def _refresh_tree():
        for item in tree.get_children():
            tree.delete(item)
        for flag in get_all_customer_flags().values():
            level = str(flag['flag_level'] or 'watch')
            info = CUSTOMER_FLAG_LEVELS.get(level, CUSTOMER_FLAG_LEVELS['watch'])
            note = str(flag['note'] or '')
            tree.insert('', 'end', values=(info['title'], flag['phone'], note, flag['updated_at'] or ''), tags=(f'flag_{level}',))

    def _clear_editor():
        phone_var.set('')
        selected_title.set('Watch')
        note_text.delete('1.0', 'end')

    def _load_selected(event=None):
        selected = tree.selection()
        if not selected:
            return
        values = tree.item(selected[0], 'values')
        if not values:
            return
        selected_title.set(values[0] or 'Watch')
        phone_var.set(values[1] or '')
        note_text.delete('1.0', 'end')
        note_text.insert('1.0', values[2] or '')

    def _save_mark():
        phone = phone_var.get().strip()
        if not normalize_flag_phone(phone):
            return
        level = title_to_key.get(selected_title.get(), 'watch')
        note = note_text.get('1.0', 'end-1c')
        if save_customer_flag(phone, level, note):
            _refresh_tree()
            try:
                check_customer_flag_warning()
            except Exception:
                pass

    def _delete_selected():
        phone = phone_var.get().strip()
        selected = tree.selection()
        if selected:
            values = tree.item(selected[0], 'values')
            if values and values[1]:
                phone = values[1]
        if not normalize_flag_phone(phone):
            return
        if delete_customer_flag(phone):
            _refresh_tree()
            _clear_editor()
            try:
                check_customer_flag_warning()
            except Exception:
                pass
    tree.bind('<<TreeviewSelect>>', _load_selected)
    _refresh_tree()
    if normalize_flag_phone(phone_var.get()):
        existing_flag = get_customer_flag(phone_var.get())
        if existing_flag:
            selected_title.set(key_to_title.get(str(existing_flag['flag_level'] or 'watch'), 'Watch'))
            note_text.delete('1.0', 'end')
            note_text.insert('1.0', str(existing_flag['note'] or ''))
    buttons = ttk.Frame(dialog)
    buttons.pack(fill='x', padx=10, pady=(0, 10))
    ttk.Button(buttons, text='Save / Update Mark', command=_save_mark).pack(side='left', padx=5)
    ttk.Button(buttons, text='Delete Mark', command=_delete_selected).pack(side='left', padx=5)
    ttk.Button(buttons, text='New', command=_clear_editor).pack(side='left', padx=5)
    ttk.Button(buttons, text='Close', command=_close_dialog).pack(side='right', padx=5)

def check_customer_flag_warning(event=None) -> None:
    """Show a visual warning when the current phone number is marked."""
    try:
        if not phone_entry:
            return
        if not hasattr(root, '_normal_phone_bg'):
            try:
                root._normal_phone_bg = phone_entry.cget('background')
            except Exception:
                root._normal_phone_bg = 'white'
        if not hasattr(root, '_normal_phone_fg'):
            try:
                root._normal_phone_fg = phone_entry.cget('foreground')
            except Exception:
                root._normal_phone_fg = 'black'
        flag = get_customer_flag(phone_entry.get())
        if not flag:
            phone_entry.configure(background=root._normal_phone_bg, foreground=root._normal_phone_fg)
            if customer_flag_label:
                customer_flag_label.configure(text='', fg='black')
            return
        level = str(flag['flag_level'] or 'watch')
        info = CUSTOMER_FLAG_LEVELS.get(level, CUSTOMER_FLAG_LEVELS['watch'])
        note = str(flag['note'] or '').strip()
        phone_entry.configure(background=info['color'], foreground='black')
        if customer_flag_label:
            customer_flag_label.configure(text='', fg='#8b0000')
    except Exception:
        pass

def save_customer_flag(phone: str, flag_level: str, note: str='') -> bool:
    """Create or update a customer flag without changing customer records."""
    clean_phone = normalize_flag_phone(phone)
    if not clean_phone:
        return False
    if flag_level not in CUSTOMER_FLAG_LEVELS:
        flag_level = 'watch'
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('\n            INSERT INTO customer_flags (phone, normalized_phone, flag_level, note, created_at, updated_at)\n            VALUES (?, ?, ?, ?, ?, ?)\n            ON CONFLICT(normalized_phone) DO UPDATE SET\n                phone = excluded.phone,\n                flag_level = excluded.flag_level,\n                note = excluded.note,\n                updated_at = excluded.updated_at\n            ', (phone.strip(), clean_phone, flag_level, note.strip(), now, now))
        conn.commit()
        return True
    except Exception as e:
        try:
            messagebox.showerror('Database Error', f'Could not save customer flag: {str(e)}')
        except Exception:
            pass
        return False
    finally:
        if conn:
            release_db_connection(conn)

def delete_customer_flag(phone: str) -> bool:
    """Delete a customer flag for the current phone number."""
    clean_phone = normalize_flag_phone(phone)
    if not clean_phone:
        return False
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM customer_flags WHERE normalized_phone = ?', (clean_phone,))
        conn.commit()
        return True
    except Exception as e:
        try:
            messagebox.showerror('Database Error', f'Could not delete customer flag: {str(e)}')
        except Exception:
            pass
        return False
    finally:
        if conn:
            release_db_connection(conn)

def open_customer_flag_dialog() -> None:
    """Open a small dialog for marking the current phone number."""
    try:
        phone = phone_entry.get().strip() if phone_entry else ''
    except Exception:
        phone = ''
    if not normalize_flag_phone(phone):
        messagebox.showwarning('Customer Flag', 'Enter a phone number first.')
        return
    existing = get_customer_flag(phone)
    dialog = tk.Toplevel(root)
    dialog.title('Customer Flag')
    dialog.resizable(False, False)
    set_window_icon(dialog)
    dialog.transient(root)
    tk.Label(dialog, text='Phone:').grid(row=0, column=0, sticky='w', padx=10, pady=(10, 5))
    tk.Label(dialog, text=phone).grid(row=0, column=1, sticky='w', padx=10, pady=(10, 5))
    tk.Label(dialog, text='Mark:').grid(row=1, column=0, sticky='w', padx=10, pady=5)
    level_keys = ('watch', 'problem', 'high_risk')
    level_titles = [CUSTOMER_FLAG_LEVELS[k]['title'] for k in level_keys]
    title_to_key = {CUSTOMER_FLAG_LEVELS[k]['title']: k for k in level_keys}
    if existing:
        selected_default = CUSTOMER_FLAG_LEVELS.get(existing['flag_level'], CUSTOMER_FLAG_LEVELS['watch'])['title']
    else:
        selected_default = 'Watch'
    selected_title = tk.StringVar(value=selected_default)
    level_combo = ttk.Combobox(dialog, values=level_titles, textvariable=selected_title, state='readonly', width=18)
    level_combo.grid(row=1, column=1, sticky='we', padx=10, pady=5)
    tk.Label(dialog, text='Note:').grid(row=2, column=0, sticky='nw', padx=10, pady=5)
    note_text = tk.Text(dialog, width=42, height=5)
    note_text.grid(row=2, column=1, padx=10, pady=5)
    if existing and existing['note']:
        note_text.insert('1.0', str(existing['note']))
    buttons = ttk.Frame(dialog)
    buttons.grid(row=3, column=0, columnspan=2, pady=10)

    def _save():
        level = title_to_key.get(selected_title.get(), 'watch')
        note = note_text.get('1.0', 'end-1c')
        if save_customer_flag(phone, level, note):
            check_customer_flag_warning()
            dialog.destroy()

    def _delete():
        if delete_customer_flag(phone):
            check_customer_flag_warning()
            dialog.destroy()
    ttk.Button(buttons, text='Save Mark', command=_save).pack(side='left', padx=5)
    ttk.Button(buttons, text='Delete Mark', command=_delete).pack(side='left', padx=5)
    ttk.Button(buttons, text='Cancel', command=dialog.destroy).pack(side='left', padx=5)
    dialog.grab_set()

def start_main_application():
    """Start the main application"""
    global root, phone_entry, model_entry, issue_entries, main_issue_entry, main_issues_tree, main_visit_history_tree, save_button, bike_list_button, edit_cancel_button, main_edit_state, status_label, customer_flag_label
    main_edit_state = None
    root = tk.Tk()
    create_global_context_menu(root)
    root.title('BikeTrack Pro')
    init_database()
    settings = load_settings()
    refresh_issue_autofill_values(settings)
    refresh_bike_brand_autofill_values(settings)
    input_font_sizes = get_input_font_sizes(settings)
    default_main_geometry = get_main_window_default_geometry(settings)
    saved_main_geometry = None
    try:
        saved_geometries = settings.get('window_geometries', {})
        if isinstance(saved_geometries, dict):
            candidate_geometry = saved_geometries.get('main_window')
            if isinstance(candidate_geometry, str) and candidate_geometry.strip():
                saved_main_geometry = candidate_geometry.strip()
    except Exception:
        saved_main_geometry = None
    root._main_window_loaded_saved_geometry = bool(saved_main_geometry)
    root.geometry(saved_main_geometry or default_main_geometry)
    root.minsize(360, 240)
    root.resizable(True, True)
    set_window_icon(root)
    create_menu()
    check_required_settings()
    text = ENGLISH_TEXT
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.rowconfigure(1, weight=0)

    form_container = ttk.Frame(root, padding=(6, 6, 6, 0))
    form_container.grid(row=0, column=0, sticky='nsew')
    form_container.columnconfigure(0, weight=1)
    form_container.rowconfigure(0, weight=1)

    form_canvas = tk.Canvas(form_container, borderwidth=0, highlightthickness=0, background=root.cget('background'))
    form_scrollbar = ttk.Scrollbar(form_container, orient='vertical', command=form_canvas.yview)
    form_canvas.configure(yscrollcommand=form_scrollbar.set)
    form_canvas.grid(row=0, column=0, sticky='nsew')
    form_scrollbar.grid(row=0, column=1, sticky='ns')

    main_frame = ttk.Frame(form_canvas, padding=(0, 0, 0, 0))
    main_window_id = form_canvas.create_window((0, 0), window=main_frame, anchor='nw')
    main_frame.columnconfigure(1, weight=1)
    main_frame.columnconfigure(3, weight=1)
    root._main_form_canvas = form_canvas
    root._main_form_required_height = 0

    def _update_main_scrollbar_visibility():
        try:
            form_canvas.update_idletasks()
            bbox = form_canvas.bbox('all')
            if not bbox:
                return
            content_height = bbox[3] - bbox[1]
            canvas_height = form_canvas.winfo_height()
            root._main_form_required_height = content_height
            if content_height > canvas_height + 2:
                form_scrollbar.grid(row=0, column=1, sticky='ns')
            else:
                form_scrollbar.grid_remove()
                form_canvas.yview_moveto(0)
        except Exception:
            pass

    def _on_main_frame_configure(event=None):
        try:
            form_canvas.configure(scrollregion=form_canvas.bbox('all'))
            _update_main_scrollbar_visibility()
        except Exception:
            pass

    def _layout_issue_fields(two_columns=True):
        try:
            if getattr(root, '_modern_issue_table_enabled', False):
                return
            if not issue_entries or len(issue_entries) < 10:
                return
            mode = 'two' if two_columns else 'one'
            if getattr(root, '_main_issue_layout_mode', None) == mode:
                return
            root._main_issue_layout_mode = mode

            # Keep the normal label/input columns for Phone Number and Bike Model,
            # but use a separate issue frame with equal-width columns for the issue fields.
            issue_fields_frame.columnconfigure(0, weight=1, uniform='issue_columns')
            issue_fields_frame.columnconfigure(1, weight=1, uniform='issue_columns')

            for entry in issue_entries:
                entry.grid_forget()

            if two_columns:
                for i, entry in enumerate(issue_entries):
                    issue_row = i % 5
                    entry_col = 0 if i < 5 else 1
                    entry_padx = (0, 5) if i < 5 else (5, 0)
                    entry.grid(row=issue_row, column=entry_col, padx=entry_padx, pady=2, sticky='we')
            else:
                for i, entry in enumerate(issue_entries):
                    entry.grid(row=i, column=0, columnspan=2, padx=(0, 0), pady=1, sticky='we')
            form_canvas.configure(scrollregion=form_canvas.bbox('all'))
            _update_main_scrollbar_visibility()
        except Exception:
            pass

    def _apply_responsive_issue_layout(width=None):
        try:
            if width is None:
                width = form_canvas.winfo_width()
            _layout_issue_fields(two_columns=(int(width) >= 410))
        except Exception:
            pass

    def _on_form_canvas_configure(event):
        try:
            form_canvas.itemconfigure(main_window_id, width=event.width)
            _apply_responsive_issue_layout(event.width)
            _update_main_scrollbar_visibility()
        except Exception:
            pass

    def _on_main_mousewheel(event):
        try:
            bbox = form_canvas.bbox('all')
            if not bbox:
                return None
            content_height = bbox[3] - bbox[1]
            if content_height <= form_canvas.winfo_height() + 2:
                return None
            delta = getattr(event, 'delta', 0)
            if delta:
                form_canvas.yview_scroll(int(-1 * (delta / 120)), 'units')
                return 'break'
        except Exception:
            return None
        return None

    main_frame.bind('<Configure>', _on_main_frame_configure)
    form_canvas.bind('<Configure>', _on_form_canvas_configure)
    form_canvas.bind('<Enter>', lambda event: form_canvas.bind_all('<MouseWheel>', _on_main_mousewheel))
    form_canvas.bind('<Leave>', lambda event: form_canvas.unbind_all('<MouseWheel>'))

    try:
        main_style = ttk.Style(root)
        main_style.configure('MainInput.Treeview', font=('Arial', 9), rowheight=22)
        main_style.configure('MainInput.Treeview.Heading', font=('Arial', 9, 'bold'), padding=(3, 2))
    except Exception:
        pass

    root._main_form_summary_var = tk.StringVar(value='New visit  |  0 issues  |  Phone + bike needed')
    summary_label = tk.Label(
        main_frame,
        textvariable=root._main_form_summary_var,
        font=('Arial', 9, 'bold'),
        anchor='w',
        padx=8,
        pady=4,
        bg='#eef5ff',
        fg='#17324d'
    )
    summary_label.grid(row=0, column=0, columnspan=4, padx=(0, 0), pady=(0, 4), sticky='we')
    root._main_summary_label = summary_label

    edit_banner_label = tk.Label(
        main_frame,
        text='',
        font=('Arial', 9, 'bold'),
        anchor='w',
        padx=8,
        pady=4,
        bg='#fff3cd',
        fg='#6b4e00'
    )
    edit_banner_label.grid(row=1, column=0, columnspan=4, padx=(0, 0), pady=(0, 4), sticky='we')
    edit_banner_label.grid_remove()
    root._main_edit_banner_label = edit_banner_label

    phone_label = ttk.Label(main_frame, text=text['tel_label'])
    phone_label.grid(row=2, column=0, padx=(0, 8), pady=(0, 6), sticky='e')
    phone_entry = tk.Entry(main_frame, font=('Arial', input_font_sizes['phone']), width=10)
    phone_entry.grid(row=2, column=1, columnspan=3, padx=(0, 0), pady=(0, 4), sticky='we')
    phone_entry.bind('<Return>', on_phone_enter)
    phone_entry.bind('<KeyRelease>', lambda e: (_schedule_phone_autorefresh(), check_customer_flag_warning(), _update_main_form_summary()))
    root._vh_last_phone = phone_entry.get().strip()
    customer_flag_label = tk.Label(main_frame, text='', font=('Arial', 9, 'bold'), anchor='w')
    check_customer_flag_warning()
    model_label = ttk.Label(main_frame, text=text['bike_label'])
    model_label.grid(row=3, column=0, padx=(0, 8), pady=(0, 6), sticky='e')
    model_entry = tk.Entry(main_frame, font=('Arial', input_font_sizes['bike']), width=14)
    model_entry.grid(row=3, column=1, columnspan=3, padx=(0, 0), pady=(0, 6), sticky='we')
    model_entry.bind('<Return>', focus_next_widget)
    model_entry.bind('<KeyRelease>', lambda e: _update_main_form_summary(), add='+')
    setup_bike_brand_autofill(model_entry)
    issue_title_label = ttk.Label(main_frame, text='Add issue:', font=('Arial', 9, 'bold'))
    issue_title_label.grid(row=4, column=0, columnspan=4, padx=(0, 0), pady=(0, 2), sticky='w')

    issue_input_frame = ttk.Frame(main_frame)
    issue_input_frame.grid(row=5, column=0, columnspan=4, padx=(0, 0), pady=(0, 4), sticky='we')
    issue_input_frame.columnconfigure(0, weight=1)
    main_issue_entry = tk.Entry(issue_input_frame, font=('Arial', input_font_sizes['issue']), width=14)
    main_issue_entry.grid(row=0, column=0, sticky='we', padx=(0, 5))
    setup_issue_autofill(main_issue_entry)
    main_issue_entry.bind('<Return>', accept_issue_autofill_or_add_to_table)
    main_issue_entry.bind('<KP_Enter>', accept_issue_autofill_or_add_to_table)
    add_issue_button = ttk.Button(issue_input_frame, text='+ Add', command=add_issue_from_main_entry)
    add_issue_button.grid(row=0, column=1, sticky='e')
    root._main_issue_entry = main_issue_entry
    root._main_issue_add_button = add_issue_button
    root._main_issue_edit_item = ''

    issues_table_frame = ttk.Frame(main_frame)
    issues_table_frame.grid(row=6, column=0, columnspan=4, padx=(0, 0), pady=(0, 4), sticky='nsew')
    issues_table_frame.columnconfigure(0, weight=1)
    issues_table_frame.rowconfigure(0, weight=1)
    main_issues_tree = ttk.Treeview(issues_table_frame, columns=('issue',), show='headings', height=4, selectmode='extended', style='MainInput.Treeview')
    main_issues_tree.heading('issue', text='Selected repairs / issues')
    main_issues_tree.column('issue', width=330, minwidth=160, anchor='w', stretch=True)
    main_issues_scrollbar = ttk.Scrollbar(issues_table_frame, orient='vertical', command=main_issues_tree.yview)
    main_issues_tree.configure(yscrollcommand=main_issues_scrollbar.set)
    main_issues_tree.grid(row=0, column=0, sticky='nsew')
    main_issues_scrollbar.grid(row=0, column=1, sticky='ns')
    main_issues_tree.bind('<Double-1>', edit_selected_main_issue)
    main_issues_tree.bind('<Delete>', remove_selected_main_issue)
    main_issues_tree.bind('<Alt-Up>', move_selected_main_issue_up)
    main_issues_tree.bind('<Alt-Down>', move_selected_main_issue_down)
    root._main_issues_tree = main_issues_tree

    issue_tools_frame = ttk.Frame(main_frame)
    issue_tools_frame.grid(row=7, column=0, columnspan=4, padx=(0, 0), pady=(0, 6), sticky='we')
    main_issue_count_label = ttk.Label(issue_tools_frame, text='0/10 selected repairs', anchor='w')
    main_issue_count_label.pack(side='left')
    root._main_issue_count_label = main_issue_count_label
    ttk.Button(issue_tools_frame, text='Clear', command=clear_main_issue_table).pack(side='right', padx=(4, 0))
    ttk.Button(issue_tools_frame, text='Remove', command=remove_selected_main_issue).pack(side='right', padx=(4, 0))
    ttk.Button(issue_tools_frame, text='Down', command=move_selected_main_issue_down).pack(side='right', padx=(4, 0))
    ttk.Button(issue_tools_frame, text='Up', command=move_selected_main_issue_up).pack(side='right', padx=(4, 0))
    ttk.Button(issue_tools_frame, text='Edit', command=edit_selected_main_issue).pack(side='right', padx=(4, 0))

    hidden_issue_frame = ttk.Frame(main_frame)
    for i in range(10):
        entry = tk.Entry(hidden_issue_frame, font=('Arial', input_font_sizes['issue']), width=8)
        issue_entries.append(entry)

    shop_note_label = ttk.Label(main_frame, text='Shop Note (printed only on shop receipt):', font=('Arial', 9, 'bold'))
    shop_note_label.grid(row=8, column=0, columnspan=4, padx=(0, 0), pady=(0, 2), sticky='w')
    shop_note_entry = tk.Text(main_frame, height=2, wrap='word', font=('Arial', input_font_sizes['issue']))
    shop_note_entry.grid(row=9, column=0, columnspan=4, padx=(0, 0), pady=(0, 6), sticky='we')
    root._shop_note_entry = shop_note_entry

    history_header_frame = ttk.Frame(main_frame)
    history_header_frame.grid(row=10, column=0, columnspan=4, padx=(0, 0), pady=(0, 2), sticky='we')
    history_header_frame.columnconfigure(0, weight=1)
    history_title_label = ttk.Label(history_header_frame, text='Visit history:', font=('Arial', 9, 'bold'))
    history_title_label.grid(row=0, column=0, sticky='w')
    history_more_button = ttk.Button(history_header_frame, text='More', width=6, command=toggle_main_visit_history_size)
    history_more_button.grid(row=0, column=1, padx=(4, 0), sticky='e')
    history_toggle_button = ttk.Button(history_header_frame, text='Hide', width=6, command=toggle_main_visit_history)
    history_toggle_button.grid(row=0, column=2, padx=(4, 0), sticky='e')
    root._main_visit_history_more_button = history_more_button
    root._main_visit_history_toggle_button = history_toggle_button
    root._main_visit_history_collapsed = False
    root._main_visit_history_tall = False

    history_frame = ttk.Frame(main_frame)
    history_frame.grid(row=11, column=0, columnspan=4, padx=(0, 0), pady=(0, 4), sticky='nsew')
    history_frame.columnconfigure(0, weight=1)
    history_frame.rowconfigure(0, weight=1)
    root._main_visit_history_frame = history_frame
    main_visit_history_tree = ttk.Treeview(history_frame, columns=('date', 'ago', 'phone', 'model', 'issues'), show='headings', height=4, selectmode='browse', style='MainInput.Treeview')
    main_visit_history_tree.configure(displaycolumns=('date', 'ago', 'model', 'issues'))
    main_visit_history_tree.heading('date', text='Date')
    main_visit_history_tree.heading('ago', text='Ago')
    main_visit_history_tree.heading('phone', text='Phone')
    main_visit_history_tree.heading('model', text='Bike')
    main_visit_history_tree.heading('issues', text='Issues')
    main_visit_history_tree.column('date', width=82, minwidth=70, anchor='center', stretch=False)
    main_visit_history_tree.column('ago', width=78, minwidth=66, anchor='center', stretch=False)
    main_visit_history_tree.column('phone', width=98, minwidth=80, anchor='center', stretch=False)
    main_visit_history_tree.column('model', width=150, minwidth=90, anchor='w', stretch=True)
    main_visit_history_tree.column('issues', width=250, minwidth=120, anchor='w', stretch=True)
    history_scrollbar_y = ttk.Scrollbar(history_frame, orient='vertical', command=main_visit_history_tree.yview)
    history_scrollbar_x = ttk.Scrollbar(history_frame, orient='horizontal', command=main_visit_history_tree.xview)
    main_visit_history_tree.configure(yscrollcommand=history_scrollbar_y.set, xscrollcommand=history_scrollbar_x.set)
    main_visit_history_tree.grid(row=0, column=0, sticky='nsew')
    history_scrollbar_y.grid(row=0, column=1, sticky='ns')
    history_scrollbar_x.grid(row=1, column=0, sticky='ew')
    main_visit_history_tree.bind('<Double-1>', use_selected_main_history_bike)
    main_visit_history_tree.bind('<Return>', use_selected_main_history_bike)
    root._main_visit_history_tree = main_visit_history_tree
    root._main_visit_history_status_var = tk.StringVar(value='Enter phone number to see previous visits.')
    history_status_label = ttk.Label(main_frame, textvariable=root._main_visit_history_status_var, anchor='w')
    history_status_label.grid(row=12, column=0, columnspan=4, padx=(0, 0), pady=(0, 2), sticky='we')
    root._main_visit_history_status_widget = history_status_label

    root._main_labels = {'phone': phone_label, 'model': model_label, 'issue_title': issue_title_label}
    root._main_issue_layout_mode = None
    root._modern_issue_table_enabled = True
    _update_main_issue_count_label()
    _refresh_main_input_visual_state()

    bottom_bar = ttk.Frame(root, padding=(6, 5, 6, 4))
    bottom_bar.grid(row=1, column=0, sticky='ew')
    bottom_bar.columnconfigure(0, weight=1)

    status_label = tk.Label(bottom_bar, text='', anchor='w')
    status_label.grid(row=0, column=0, padx=(0, 8), sticky='we')

    button_frame = ttk.Frame(bottom_bar)
    button_frame.grid(row=0, column=1, sticky='e')
    save_text = text['save_button']
    save_button = ttk.Button(button_frame, text=save_text, command=add_customer)
    save_button.pack(side='left', padx=4)
    edit_cancel_button = ttk.Button(button_frame, text='Cancel Edit', command=cancel_main_edit_mode)

    def on_ctrl_s(event=None):
        try:
            if save_button is not None:
                save_button.invoke()
            else:
                add_customer()
        except Exception:
            add_customer()
        return 'break'

    def on_save_shortcut(event=None):
        ctrl_pressed = bool(getattr(event, 'state', 0) & 4)
        keysym = str(getattr(event, 'keysym', '')).lower()
        keycode = getattr(event, 'keycode', None)
        if ctrl_pressed and (keysym == 's' or keycode == 83):
            return on_ctrl_s(event)
        return None
    root.bind_all('<KeyPress>', on_save_shortcut)
    bike_list_button = ttk.Button(button_frame, text=text['bike_list_button'], command=open_bike_list)
    bike_list_button.pack(side='left', padx=4)

    def close_main_window():
        if not messagebox.askyesno('Exit BikeTrack Pro', 'Are you sure you want to close BikeTrack Pro?', parent=root):
            return
        save_window_geometry(root, 'main_window')
        root.destroy()
    root.protocol('WM_DELETE_WINDOW', close_main_window)
    phone_entry.focus_set()
    root.update_idletasks()
    if not getattr(root, '_main_window_loaded_saved_geometry', False):
        try:
            initial_width = max(520, int(root.winfo_width() or 520))
            initial_height = max(420, min(int(root.winfo_reqheight() + 4), int((root.winfo_screenheight() or 900) - 80)))
            root.geometry(f'{initial_width}x{initial_height}')
        except Exception:
            pass
    root.minsize(360, 240)
    root.resizable(True, True)
    root.after(100, _update_main_scrollbar_visibility)
    root.mainloop()
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_FILE = os.path.join(BASE_DIR, 'customer_bikes.db')
ICON_PATH = os.path.join(BASE_DIR, 'icon.ico')
LOCAL_SETTINGS_DIR = os.path.join(os.environ.get('LOCALAPPDATA') or os.path.expanduser('~'), 'BikeTrackPro')
LOCAL_SETTINGS_FILE = os.path.join(LOCAL_SETTINGS_DIR, 'local_settings.json')
LOCAL_ONLY_SETTING_KEYS = {'printer', 'customer_printer', 'shop_printer'}

def get_database_folder() -> str:
    """Return the folder that contains the shared SQLite database."""
    return os.path.dirname(os.path.abspath(DATABASE_FILE))

def resolve_backup_folder_path(backup_folder: Optional[str]=None) -> str:
    """Resolve a backup folder setting relative to the shared database folder."""
    folder = (backup_folder or 'backup_db').strip() or 'backup_db'
    folder = os.path.expandvars(os.path.expanduser(folder))
    if os.path.isabs(folder):
        return folder
    return os.path.abspath(os.path.join(get_database_folder(), folder))

def make_backup_folder_setting(folder_path: Optional[str]) -> str:
    """Store backup folder settings as relative paths when possible."""
    folder = (folder_path or 'backup_db').strip() or 'backup_db'
    folder = os.path.expandvars(os.path.expanduser(folder))
    if not os.path.isabs(folder):
        return folder
    try:
        rel_path = os.path.relpath(folder, get_database_folder())
        if rel_path and rel_path != '.':
            return rel_path
        return 'backup_db'
    except ValueError:
        return folder

def load_local_settings() -> Dict:
    """Load settings that must stay local to this Windows user/computer."""
    try:
        if os.path.exists(LOCAL_SETTINGS_FILE):
            with open(LOCAL_SETTINGS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
    except Exception:
        pass
    return {}

def save_local_settings(local_settings: Dict) -> None:
    """Save local-only settings outside the shared database."""
    try:
        os.makedirs(LOCAL_SETTINGS_DIR, exist_ok=True)
        with open(LOCAL_SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(local_settings, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showerror('Local Settings Error', f'Could not save local settings: {str(e)}')

def save_local_printer(printer_name: Optional[str]) -> None:
    """Save a legacy/single selected printer only for this computer.

    Kept for older parts of the app. If separate receipt printers are not already
    configured, this also becomes the customer/shop printer fallback.
    """
    local_settings = load_local_settings()
    printer_name = printer_name or None
    local_settings['printer'] = printer_name
    local_settings['shop_printer'] = printer_name
    if not local_settings.get('customer_printer'):
        local_settings['customer_printer'] = printer_name
    save_local_settings(local_settings)

def save_local_receipt_printers(customer_printer: Optional[str], shop_printer: Optional[str]) -> None:
    """Save separate customer/shop receipt printers only for this computer."""
    local_settings = load_local_settings()
    customer_printer = customer_printer or None
    shop_printer = shop_printer or None
    local_settings['customer_printer'] = customer_printer
    local_settings['shop_printer'] = shop_printer
    local_settings['printer'] = shop_printer or customer_printer
    save_local_settings(local_settings)

def merge_local_settings(settings: Dict) -> Dict:
    """Overlay local-only settings on top of shared database settings."""
    merged = dict(settings)
    local_settings = load_local_settings()
    legacy_printer = local_settings.get('printer') or None
    customer_printer = local_settings.get('customer_printer') or legacy_printer
    shop_printer = local_settings.get('shop_printer') or legacy_printer
    merged['customer_printer'] = customer_printer or None
    merged['shop_printer'] = shop_printer or None
    merged['printer'] = shop_printer or customer_printer or legacy_printer or None
    return merged

def strip_local_only_settings(settings: Dict) -> Dict:
    """Remove local-only settings before saving shared settings to the database."""
    shared_settings = dict(settings)
    for key in LOCAL_ONLY_SETTING_KEYS:
        shared_settings[key] = DEFAULT_SETTINGS.get(key)
    shared_settings.pop('language', None)
    return shared_settings
RECEIPT_80MM_DEFAULT_WIDTH_PX = 576
RECEIPT_COMMON_PAPER_WIDTHS = (
    ('58', 384),
    ('80S', 512),
    ('80', 576),
    ('80W', 640),
)
DEFAULT_CUSTOMER_TITLE_TEXT = 'VÄSTRA HAMNENS CYKLAR'
DEFAULT_SHOP_TITLE_TEXT = 'BUTIKSKVITTO'
DEFAULT_THANK_YOU_TEXT = 'Tack!'
DEFAULT_CUSTOMER_RECEIPT_TEMPLATE = '{logo}\n[center][title]{customer_title}\n{line}\n[bold]Telefon: {phone}\n[bold]Cykel: {bike}\nDatum: {date}\nTid: {time}\n\n{shop_info}\n\n{qr}\n[center][small]{thanks}'
DEFAULT_SHOP_RECEIPT_TEMPLATE = '[center][title]{shop_title}\n{line}\n[bold]Telefon: {phone}\n[bold]Cykel: {bike}\nDatum: {date}\nTid: {time}\n{line}\n[bold]Fel/Brist:\n{issues}\n\n{qr}'
DEFAULT_RECEIPT_DESIGNER = {'paper_width': RECEIPT_80MM_DEFAULT_WIDTH_PX, 'margin': 8, 'line_spacing': 3, 'bottom_feed': 28, 'customer_top_space': 0, 'customer_bottom_feed': 28, 'shop_top_space': 0, 'shop_bottom_feed': 28, 'font_family': 'Segoe UI', 'text_align': 'left', 'logo_align': 'center', 'title_align': 'center', 'phone_align': 'left', 'bike_align': 'left', 'datetime_align': 'left', 'shop_info_align': 'left', 'issue_align': 'left', 'qr_align': 'center', 'thanks_align': 'center', 'title_size': 27, 'body_size': 21, 'small_size': 16, 'phone_size': 23, 'bike_size': 23, 'datetime_size': 19, 'shop_info_size': 18, 'issue_size': 22, 'title_bold': True, 'body_bold': False, 'phone_bold': True, 'bike_bold': True, 'datetime_bold': False, 'shop_info_bold': False, 'issue_bold': False, 'compact_mode': True, 'center_qr': True, 'logo_width': 260, 'logo_aspect': 100, 'logo_after_margin': 6, 'customer_qr_width': 220, 'shop_qr_width': 220, 'customer_qr_aspect': 100, 'shop_qr_aspect': 100, 'simple_designer': True, 'customer_show_logo': True, 'customer_show_shop_info': True, 'customer_show_qr': True, 'customer_show_title': True, 'customer_show_phone': True, 'customer_show_phone_label': True, 'customer_show_bike': True, 'customer_show_bike_label': True, 'customer_show_datetime': True, 'customer_show_date': True, 'customer_show_date_label': True, 'customer_show_time': True, 'customer_show_time_label': True, 'customer_show_issues': False, 'customer_show_issues_label': True, 'customer_show_thanks': True, 'shop_show_logo': False, 'shop_show_shop_info': False, 'shop_show_qr': True, 'shop_show_phone': True, 'shop_show_phone_label': True, 'shop_show_bike': True, 'shop_show_bike_label': True, 'shop_show_datetime': True, 'shop_show_date': True, 'shop_show_date_label': True, 'shop_show_time': True, 'shop_show_time_label': True, 'shop_show_issues': True, 'shop_show_issues_label': True, 'shop_show_title': True, 'shop_show_note': True, 'shop_show_note_label': True, 'shop_note_label_text': 'Shop Note', 'shop_note_size': 21, 'shop_note_bold': False, 'shop_note_align': 'left', 'customer_title_text': DEFAULT_CUSTOMER_TITLE_TEXT, 'shop_title_text': DEFAULT_SHOP_TITLE_TEXT, 'thank_you_text': DEFAULT_THANK_YOU_TEXT, 'customer_template': DEFAULT_CUSTOMER_RECEIPT_TEMPLATE, 'shop_template': DEFAULT_SHOP_RECEIPT_TEMPLATE}
PRINT_FONT_CHOICES = ('Segoe UI', 'Segoe UI Semibold', 'Tahoma', 'Verdana', 'Bahnschrift', 'Trebuchet MS', 'Arial', 'Calibri', 'Consolas', 'Courier New')
PRINT_FONT_FILE_MAP = {'arial': ('arialbd.ttf', 'arial.ttf'), 'segoe ui': ('segoeuib.ttf', 'segoeui.ttf'), 'segoe ui semibold': ('seguisb.ttf', 'seguisb.ttf'), 'tahoma': ('tahomabd.ttf', 'tahoma.ttf'), 'verdana': ('verdanab.ttf', 'verdana.ttf'), 'calibri': ('calibrib.ttf', 'calibri.ttf'), 'bahnschrift': ('bahnschrift.ttf', 'bahnschrift.ttf'), 'trebuchet ms': ('trebucbd.ttf', 'trebuc.ttf'), 'consolas': ('consolab.ttf', 'consola.ttf'), 'courier new': ('courbd.ttf', 'cour.ttf')}
PRINT_ALIGNMENT_CHOICES = ('Left', 'Center', 'Right')
PRINT_ALIGNMENT_DISPLAY = {'left': 'Left', 'center': 'Center', 'right': 'Right'}

# Receipt Designer values that can be configured separately for the customer
# receipt and the shop receipt. Legacy un-prefixed keys are kept as fallbacks
# so older saved settings continue to work.
RECEIPT_KIND_DESIGN_KEYS = (
    'paper_width', 'margin', 'line_spacing', 'font_family',
    'text_align', 'logo_align', 'title_align', 'phone_align', 'bike_align',
    'datetime_align', 'shop_info_align', 'issue_align', 'qr_align', 'thanks_align',
    'title_size', 'body_size', 'small_size', 'phone_size', 'bike_size',
    'datetime_size', 'shop_info_size', 'issue_size',
    'title_bold', 'body_bold', 'phone_bold', 'bike_bold', 'datetime_bold',
    'shop_info_bold', 'issue_bold', 'compact_mode',
    'logo_width', 'logo_aspect', 'logo_after_margin',
    'qr_width', 'qr_aspect', 'title_text', 'thank_you_text',
    'shop_info_text', 'logo_path',
)


def _receipt_kind_prefix(receipt_kind: str) -> str:
    return 'shop' if str(receipt_kind).strip().lower() == 'shop' else 'customer'


def _receipt_kind_key(receipt_kind: str, key: str) -> str:
    return f'{_receipt_kind_prefix(receipt_kind)}_{key}'


def _receipt_design_value_for_kind(design: Dict, receipt_kind: str, key: str, default: Any=None) -> Any:
    """Return a receipt designer value with per-receipt override support."""
    prefix = _receipt_kind_prefix(receipt_kind)
    direct_key = f'{prefix}_{key}'
    if direct_key in design:
        return design.get(direct_key)
    # Backward-compatible mapping for keys that already existed separately.
    if key == 'qr_width':
        legacy_key = f'{prefix}_qr_width'
        if legacy_key in design:
            return design.get(legacy_key)
    if key == 'qr_aspect':
        legacy_key = f'{prefix}_qr_aspect'
        if legacy_key in design:
            return design.get(legacy_key)
    if key == 'title_text':
        legacy_key = f'{prefix}_title_text'
        if legacy_key in design:
            return design.get(legacy_key)
    if key == 'thank_you_text':
        legacy_key = f'{prefix}_thank_you_text'
        if legacy_key in design:
            return design.get(legacy_key)
    if key == 'shop_info_text':
        legacy_key = f'{prefix}_shop_info_text'
        if legacy_key in design:
            return design.get(legacy_key)
    return design.get(key, default)


def _receipt_design_for_kind(design: Dict, receipt_kind: str, settings: Optional[Dict]=None) -> Dict:
    """Return a single-receipt view of Receipt Designer settings."""
    scoped = dict(design or {})
    prefix = _receipt_kind_prefix(receipt_kind)
    for key in RECEIPT_KIND_DESIGN_KEYS:
        value = _receipt_design_value_for_kind(scoped, prefix, key, scoped.get(key))
        if value is not None:
            scoped[key] = value
    scoped['paper_width'] = _clamp_int(scoped.get('paper_width'), RECEIPT_80MM_DEFAULT_WIDTH_PX, 240, 832)
    scoped['margin'] = _clamp_int(scoped.get('margin'), 8, 0, 80)
    scoped['line_spacing'] = _clamp_int(scoped.get('line_spacing'), 3, 0, 18)
    scoped['top_space'] = _clamp_int(scoped.get(f'{prefix}_top_space', scoped.get('top_space', 0)), 0, 0, 240)
    scoped['bottom_feed'] = _clamp_int(scoped.get(f'{prefix}_bottom_feed', scoped.get('bottom_feed', 28)), 28, 0, 240)
    scoped['font_family'] = str(scoped.get('font_family') or 'Segoe UI').strip() or 'Segoe UI'
    scoped['text_align'] = _normalize_receipt_align(scoped.get('text_align', 'left'), 'left')
    for align_key in ('logo_align', 'title_align', 'phone_align', 'bike_align', 'datetime_align', 'shop_info_align', 'issue_align', 'qr_align', 'thanks_align'):
        scoped[align_key] = _normalize_receipt_align(scoped.get(align_key, scoped['text_align']), scoped['text_align'])
    scoped['title_size'] = _clamp_int(scoped.get('title_size'), 27, 10, 72)
    scoped['body_size'] = _clamp_int(scoped.get('body_size'), 21, 8, 60)
    scoped['small_size'] = _clamp_int(scoped.get('small_size'), 16, 7, 48)
    scoped['phone_size'] = _clamp_int(scoped.get('phone_size'), 23, 8, 72)
    scoped['bike_size'] = _clamp_int(scoped.get('bike_size'), 23, 8, 72)
    scoped['datetime_size'] = _clamp_int(scoped.get('datetime_size'), 19, 7, 60)
    scoped['shop_info_size'] = _clamp_int(scoped.get('shop_info_size'), 18, 7, 60)
    scoped['issue_size'] = _clamp_int(scoped.get('issue_size'), 22, 8, 72)
    max_content = max(90, scoped['paper_width'] - 2 * scoped['margin'])
    scoped['logo_width'] = _clamp_int(scoped.get('logo_width'), 260, 40, max_content)
    scoped['logo_aspect'] = _clamp_int(scoped.get('logo_aspect'), 100, 50, 200)
    scoped['logo_after_margin'] = _clamp_int(scoped.get('logo_after_margin'), 6, 0, 120)
    scoped['qr_width'] = _clamp_int(scoped.get('qr_width'), 220, 80, max_content)
    scoped['qr_aspect'] = _clamp_int(scoped.get('qr_aspect'), 100, 50, 200)
    if prefix == 'customer':
        scoped['customer_title_text'] = _clean_receipt_single_line_text(scoped.get('title_text', scoped.get('customer_title_text')), DEFAULT_CUSTOMER_TITLE_TEXT)
        scoped['thank_you_text'] = _clean_receipt_single_line_text(scoped.get('thank_you_text'), DEFAULT_THANK_YOU_TEXT)
    else:
        scoped['shop_title_text'] = _clean_receipt_single_line_text(scoped.get('title_text', scoped.get('shop_title_text')), DEFAULT_SHOP_TITLE_TEXT)
        scoped['thank_you_text'] = _clean_receipt_single_line_text(scoped.get('thank_you_text'), DEFAULT_THANK_YOU_TEXT)
    legacy_logo_path = ''
    try:
        legacy_logo_path = str(((settings or {}).get('customer_logo') or {}).get('path', '') or '').strip()
    except Exception:
        legacy_logo_path = ''
    scoped['logo_path'] = str(scoped.get('logo_path') or legacy_logo_path or '').strip()
    return scoped


def _receipt_logo_settings_for_kind(settings: Dict, design: Dict, receipt_kind: str) -> Dict:
    """Build logo settings for exactly one receipt type."""
    scoped = _receipt_design_for_kind(design, receipt_kind, settings)
    logo_path = str(scoped.get('logo_path', '') or '').strip()
    return {
        'enabled': bool(logo_path),
        'path': logo_path,
        'width': scoped.get('logo_width', 260),
        'height_percent': scoped.get('logo_aspect', 100),
    }


def _normalize_receipt_align(value: Any, default: str='left') -> str:
    """Normalize receipt text alignment to left/center/right."""
    normalized_default = str(default or 'left').strip().lower()
    if normalized_default not in ('left', 'center', 'right'):
        normalized_default = 'left'
    raw_value = str(value or normalized_default).strip().lower()
    aliases = {'left': 'left', 'l': 'left', 'چپ': 'left', 'چپچین': 'left', 'center': 'center', 'centre': 'center', 'middle': 'center', 'c': 'center', 'وسط': 'center', 'وسطچین': 'center', 'right': 'right', 'r': 'right', 'راست': 'right', 'راستچین': 'right'}
    return aliases.get(raw_value, normalized_default)

def _receipt_alignment_display(value: Any) -> str:
    return PRINT_ALIGNMENT_DISPLAY.get(_normalize_receipt_align(value), 'Left')

def _receipt_alignment_marker(value: Any, default: str='left') -> str:
    """Return template markup for alignment, keeping left alignment markup-free."""
    align = _normalize_receipt_align(value, default)
    return '' if align == 'left' else f'[{align}]'

def _receipt_design_align(design: Dict, key: str, default: str='left') -> str:
    """Read one per-element alignment setting with the global text alignment as fallback."""
    fallback = _normalize_receipt_align(design.get('text_align', default), default)
    return _normalize_receipt_align(design.get(key, fallback), fallback)

def build_simple_receipt_template(receipt_kind: str, design: Dict) -> str:
    """Create a clean fixed receipt template from simple show/hide switches."""

    def marker(key: str, default: str='left') -> str:
        return _receipt_alignment_marker(_receipt_design_align(design, key, default), default)
    logo_prefix = marker('logo_align', 'center')
    title_prefix = marker('title_align', 'center')
    phone_prefix = marker('phone_align', 'left')
    bike_prefix = marker('bike_align', 'left')
    datetime_prefix = marker('datetime_align', 'left')
    shop_info_prefix = marker('shop_info_align', 'left')
    issue_prefix = marker('issue_align', 'left')
    qr_prefix = marker('qr_align', 'center')
    note_prefix = _receipt_alignment_marker(design.get('shop_note_align', 'left'))
    thanks_prefix = marker('thanks_align', 'center')
    thank_you_text = _clean_receipt_single_line_text(design.get('thank_you_text'), DEFAULT_THANK_YOU_TEXT)
    customer_title_text = _clean_receipt_single_line_text(design.get('customer_title_text'), DEFAULT_CUSTOMER_TITLE_TEXT)
    shop_title_text = _clean_receipt_single_line_text(design.get('shop_title_text'), DEFAULT_SHOP_TITLE_TEXT)

    def value_line(prefix: str, markup: str, label: str, placeholder: str, show_label: bool) -> str:
        if show_label:
            return f'{prefix}{markup}{label}: {{{placeholder}}}'
        return f'{prefix}{markup}{{{placeholder}}}'
    lines: List[str] = []
    if receipt_kind == 'customer':
        if design.get('customer_show_logo', True):
            lines.append(f'{logo_prefix}{{logo}}')
        if design.get('customer_show_title', True):
            lines.append(f'{title_prefix}[title]{customer_title_text}')
            lines.append('{line}')
        if design.get('customer_show_phone', True):
            lines.append(value_line(phone_prefix, '[bold]', 'Telefon', 'phone', design.get('customer_show_phone_label', True)))
        if design.get('customer_show_bike', True):
            lines.append(value_line(bike_prefix, '[bold]', 'Cykel', 'bike', design.get('customer_show_bike_label', True)))
        if design.get('customer_show_date', design.get('customer_show_datetime', True)):
            lines.append(value_line(datetime_prefix, '', 'Datum', 'date', design.get('customer_show_date_label', True)))
        if design.get('customer_show_time', design.get('customer_show_datetime', True)):
            lines.append(value_line(datetime_prefix, '', 'Tid', 'time', design.get('customer_show_time_label', True)))
        if design.get('customer_show_issues', False):
            lines.append('{line}')
            if design.get('customer_show_issues_label', True):
                lines.append(f'{issue_prefix}[bold]Fel/Brist:')
            lines.append(f'{issue_prefix}{{issues}}')
        if design.get('customer_show_shop_info', True):
            lines.append('')
            lines.append(f'{shop_info_prefix}{{shop_info}}')
        if design.get('customer_show_qr', True):
            lines.append('')
            lines.append(f'{qr_prefix}{{qr}}')
        if design.get('customer_show_thanks', True) and thank_you_text:
            lines.append(f'{thanks_prefix}[small]{thank_you_text}')
    else:
        if design.get('shop_show_logo', False):
            lines.append(f'{logo_prefix}{{logo}}')
        if design.get('shop_show_title', True):
            lines.append(f'{title_prefix}[title]{shop_title_text}')
            lines.append('{line}')
        if design.get('shop_show_phone', True):
            lines.append(value_line(phone_prefix, '[bold]', 'Telefon', 'phone', design.get('shop_show_phone_label', True)))
        if design.get('shop_show_bike', True):
            lines.append(value_line(bike_prefix, '[bold]', 'Cykel', 'bike', design.get('shop_show_bike_label', True)))
        if design.get('shop_show_date', design.get('shop_show_datetime', True)):
            lines.append(value_line(datetime_prefix, '', 'Datum', 'date', design.get('shop_show_date_label', True)))
        if design.get('shop_show_time', design.get('shop_show_datetime', True)):
            lines.append(value_line(datetime_prefix, '', 'Tid', 'time', design.get('shop_show_time_label', True)))
        if design.get('shop_show_issues', True):
            lines.append('{line}')
            if design.get('shop_show_issues_label', True):
                lines.append(f'{issue_prefix}[bold]Fel/Brist:')
            lines.append(f'{issue_prefix}{{issues}}')
        if design.get('shop_show_shop_info', False):
            lines.append('')
            lines.append(f'{shop_info_prefix}{{shop_info}}')
        if design.get('shop_show_note', True):
            if design.get('shop_show_qr', True):
                lines.append('')
            lines.append(f'{note_prefix}{{shop_note}}')
        if design.get('shop_show_qr', True):
            if not design.get('shop_show_note', True):
                lines.append('')
            lines.append(f'{qr_prefix}{{qr}}')
    return '\n'.join(lines).strip()
DEFAULT_SETTINGS = {'shop_info': 'Shop Name: Example Bike Shop\nTelefonnummer: 123-456-789\nAddress: 123 Example Street, City', 'customer_qr': 'Customer QR Code Content', 'shop_qr': 'Shop QR Code Content', 'printer': None, 'customer_printer': None, 'shop_printer': None, 'font_settings': {'customer_phone': 'Normal', 'customer_model': 'Normal', 'customer_datetime': 'Normal', 'shop_info': 'Double Height', 'shop_issue': 'Double Height'}, 'receipt_designer': DEFAULT_RECEIPT_DESIGNER.copy(), 'qr_width_settings': {'customer_qr_width': 576, 'shop_qr_width': 576}, 'qr_enabled': {'customer_qr': True, 'shop_qr': True}, 'customer_logo': {'enabled': False, 'path': '', 'width': 384, 'height_percent': 100}, 'sms_message': 'Hello, this is a message from Bike Shop.', 'empty_lines': 3, 'treeview_font_size': 10, 'bike_list_visible_columns': [], 'bike_list_column_widths': {}, 'input_font_size': 14, 'phone_input_font_size': 26, 'bike_input_font_size': 18, 'issue_input_font_size': 12, 'issue_autofill_enabled': True, 'issue_autofill_custom_items': '', 'issue_autofill_include_history': True, 'issue_autofill_auto_add_new': True, 'bike_brand_autofill_enabled': True, 'bike_brand_autofill_custom_items': '', 'bike_color_autofill_custom_items': '', 'bike_model_autofill_custom_items': '', 'bike_brand_autofill_include_history': True, 'bike_brand_autofill_auto_add_new': True, 'bike_model_autofill_auto_add_new': True, 'backup_folder': 'backup_db', 'window_geometries': {}}
FONT_SIZE_COMMANDS = {'Normal': b'\x1b!\x00', 'Double Height': b'\x1b!\x10', 'Double Width': b'\x1b! ', 'Double Width and Height': b'\x1b!0'}
ALIGN_CENTER = b'\x1ba\x01'
ALIGN_LEFT = b'\x1ba\x00'
ESC_INIT = b'\x1b@'
ESC_CODEPAGE_CP850 = b'\x1bt\x02'
ESC_NORMAL_TEXT = FONT_SIZE_COMMANDS['Normal'] + ALIGN_LEFT

def escpos_text(value: Any) -> bytes:
    """Encode plain receipt text for thermal ESC/POS printers without crashing on odd characters."""
    return str(value if value is not None else '').encode('cp850', errors='replace')

def escpos_line(value: Any, size_command: bytes=FONT_SIZE_COMMANDS['Normal']) -> bytes:
    """Return one encoded receipt line with the requested font size."""
    return size_command + escpos_text(value) + b'\n'

def escpos_qr_native(qr_content: Any, module_size: int=6, error_correction: int=49) -> bytes:
    """
    Print a QR code using the printer's native ESC/POS QR command instead of sending it as a raster image.
    This is much faster and cleaner on Star TSP100 printers running ESC/POS emulation.
    """
    payload = str(qr_content if qr_content is not None else '').encode('utf-8', errors='replace')
    module_size = max(3, min(int(module_size or 6), 8))
    error_correction = max(48, min(int(error_correction or 49), 51))

    def gs_k(command_payload: bytes) -> bytes:
        length = len(command_payload)
        return b'\x1d(k' + length.to_bytes(2, 'little') + command_payload
    return b''.join([ALIGN_CENTER, gs_k(b'1A2\x00'), gs_k(b'1C' + bytes([module_size])), gs_k(b'1E' + bytes([error_correction])), gs_k(b'1P0' + payload), gs_k(b'1Q0'), b'\n', ALIGN_LEFT, FONT_SIZE_COMMANDS['Normal']])
settings_window = None
about_window = None
help_window = None
bike_list_window = None
bike_statistics_window = None
bike_shop_app = None
root = None
phone_entry = None
model_entry = None
issue_entries = []
main_issue_entry = None
main_issues_tree = None
main_visit_history_tree = None
save_button = None
bike_list_button = None
edit_cancel_button = None
main_edit_state = None
status_label = None
customer_flag_label = None
menu_bar = None
FIXED_INPUT_FONT_SIZES = {'phone': 26, 'bike': 18, 'issue': 12}

def _clamp_input_font_size(value: Any, default: int=12) -> int:
    """Return a safe font size for legacy callers."""
    try:
        size = int(value)
    except Exception:
        size = int(default)
    return max(6, min(size, 30))

def get_input_font_size(settings: Optional[Dict]=None) -> int:
    """Return the fixed fallback font size for issue input fields."""
    return FIXED_INPUT_FONT_SIZES['issue']


def _update_main_form_summary() -> None:
    """Refresh the small visual summary strip in the main input window only."""
    try:
        if root is None:
            return
        summary_var = getattr(root, '_main_form_summary_var', None)
        if summary_var is None:
            return
        try:
            issue_count = len(_main_issue_tree_items())
        except Exception:
            issue_count = 0
        phone_ready = bool(phone_entry is not None and phone_entry.get().strip())
        bike_ready = bool(model_entry is not None and model_entry.get().strip())
        mode_text = 'Editing record' if main_edit_state else 'New visit'
        ready_text = 'Ready to save' if (phone_ready and bike_ready) else 'Phone + bike needed'
        issue_text = f'{issue_count} issue' + ('' if issue_count == 1 else 's')
        summary_var.set(f'{mode_text}  |  {issue_text}  |  {ready_text}')
    except Exception:
        pass


def _update_main_edit_banner() -> None:
    """Show or hide the visual edit-mode banner without changing save logic."""
    try:
        if root is None:
            return
        banner = getattr(root, '_main_edit_banner_label', None)
        if banner is None or not banner.winfo_exists():
            return
        if main_edit_state:
            dt = str(main_edit_state.get('datetime', '') or '').strip()
            phone = str(main_edit_state.get('phone', '') or '').strip()
            model = str(main_edit_state.get('model', '') or '').strip()
            details = ' | '.join([part for part in (dt, phone, model) if part])
            banner.config(text=f'EDITING EXISTING RECORD  —  {details}' if details else 'EDITING EXISTING RECORD')
            banner.grid()
        else:
            banner.grid_remove()
    except Exception:
        pass


def _refresh_main_input_visual_state() -> None:
    """Update only main-window visual indicators."""
    _update_main_edit_banner()
    _update_main_form_summary()


def _set_main_visit_history_expanded(expanded: bool=True) -> None:
    """Collapse/expand the embedded visit history area visually only."""
    try:
        if root is None:
            return
        root._main_visit_history_collapsed = not expanded
        frame = getattr(root, '_main_visit_history_frame', None)
        status = getattr(root, '_main_visit_history_status_widget', None)
        toggle = getattr(root, '_main_visit_history_toggle_button', None)
        more_button = getattr(root, '_main_visit_history_more_button', None)
        if expanded:
            if frame is not None and frame.winfo_exists():
                frame.grid()
            if status is not None and status.winfo_exists():
                status.grid()
            if more_button is not None and more_button.winfo_exists():
                more_button.config(state=tk.NORMAL)
            if toggle is not None and toggle.winfo_exists():
                toggle.config(text='Hide')
        else:
            if frame is not None and frame.winfo_exists():
                frame.grid_remove()
            if status is not None and status.winfo_exists():
                status.grid_remove()
            if more_button is not None and more_button.winfo_exists():
                more_button.config(state=tk.DISABLED)
            if toggle is not None and toggle.winfo_exists():
                toggle.config(text='Show')
    except Exception:
        pass


def toggle_main_visit_history() -> None:
    """Toggle the embedded visit history area."""
    try:
        collapsed = bool(getattr(root, '_main_visit_history_collapsed', False))
        _set_main_visit_history_expanded(expanded=collapsed)
    except Exception:
        pass


def toggle_main_visit_history_size() -> None:
    """Switch the embedded visit history tree between compact and taller view."""
    try:
        tree = getattr(root, '_main_visit_history_tree', None)
        button = getattr(root, '_main_visit_history_more_button', None)
        if tree is None or not tree.winfo_exists():
            return
        expanded = not bool(getattr(root, '_main_visit_history_tall', False))
        root._main_visit_history_tall = expanded
        tree.configure(height=8 if expanded else 4)
        if button is not None and button.winfo_exists():
            button.config(text='Less' if expanded else 'More')
    except Exception:
        pass


def move_selected_main_issue(direction: int) -> str:
    """Move selected visible issue rows up/down; hidden fields are synced afterwards."""
    try:
        tree = getattr(root, '_main_issues_tree', None)
        if tree is None or not tree.winfo_exists():
            return 'break'
        selected = list(tree.selection())
        if not selected:
            return 'break'
        children = list(tree.get_children(''))
        selected_set = set(selected)
        if int(direction) < 0:
            ordered = [item for item in children if item in selected_set]
            for item in ordered:
                index = list(tree.get_children('')).index(item)
                if index > 0:
                    tree.move(item, '', index - 1)
        else:
            ordered = [item for item in children if item in selected_set]
            for item in reversed(ordered):
                current_children = list(tree.get_children(''))
                index = current_children.index(item)
                if index < len(current_children) - 1:
                    tree.move(item, '', index + 1)
        tree.selection_set(selected)
        tree.focus(selected[0])
        _sync_issue_entries_from_main_tree()
        _refresh_main_input_visual_state()
    except Exception:
        pass
    return 'break'


def move_selected_main_issue_up(event=None) -> str:
    return move_selected_main_issue(-1)


def move_selected_main_issue_down(event=None) -> str:
    return move_selected_main_issue(1)

def get_input_font_sizes(settings: Optional[Dict]=None) -> Dict[str, int]:
    """Return fixed font sizes for phone, bike model, and issue fields."""
    return dict(FIXED_INPUT_FONT_SIZES)

def get_main_window_default_geometry(settings: Optional[Dict]=None) -> str:
    """Return the default main input window size."""
    return '520x560'

def auto_resize_main_window_height() -> None:
    """Adjust only the startup/content height while keeping the current responsive width."""
    try:
        if root is None:
            return
        root.update_idletasks()
        current_width = max(360, int(root.winfo_width() or 430))
        form_required_height = int(getattr(root, '_main_form_required_height', 0) or 0)
        base_required_height = root.winfo_reqheight() + 6
        if form_required_height:
            canvas_reqheight = getattr(root, '_main_form_canvas', root).winfo_reqheight()
            non_form_height = max(0, base_required_height - max(1, canvas_reqheight))
            required_height = form_required_height + non_form_height + 12
        else:
            required_height = base_required_height
        screen_height = root.winfo_screenheight() or 900
        max_height = max(260, screen_height - 80)
        new_height = max(240, min(required_height, max_height))
        geometry = root.geometry()
        x = y = None
        try:
            import re
            match = re.match('\\d+x\\d+([+-]\\d+)([+-]\\d+)', geometry)
            if match:
                x = int(match.group(1))
                y = int(match.group(2))
        except Exception:
            x = y = None
        if x is not None and y is not None:
            root.geometry(f'{int(current_width)}x{int(new_height)}{x:+d}{y:+d}')
        else:
            root.geometry(f'{int(current_width)}x{int(new_height)}')
        root.update_idletasks()
    except Exception:
        pass

def apply_input_field_font_size(sizes: Optional[Any]=None) -> None:
    """Apply the fixed font sizes to the main input fields."""
    try:
        field_sizes = get_input_font_sizes()
        if phone_entry is not None:
            phone_entry.configure(font=('Arial', field_sizes['phone']))
        if model_entry is not None:
            model_entry.configure(font=('Arial', field_sizes['bike']))
        quick_issue_entry = getattr(root, '_main_issue_entry', None) if root is not None else None
        if quick_issue_entry is not None:
            quick_issue_entry.configure(font=('Arial', field_sizes['issue']))
        for widget in list(issue_entries):
            if widget is not None:
                widget.configure(font=('Arial', field_sizes['issue']))
        if root is not None:
            root.update_idletasks()
            auto_resize_main_window_height()
    except Exception:
        pass
ISSUE_AUTOFILL_MAX_DB_VALUES = 400
ISSUE_AUTOFILL_VISIBLE_ROWS = 9
DEFAULT_ISSUE_AUTOFILL_ITEMS = ['Punktering', 'Byta slang', 'Byta däck', 'Justera broms', 'Byta bromsbelägg', 'Justera växlar', 'Byta växelvajer', 'Byta bromsvajer', 'Byta kedja', 'Byta kassett', 'Byta handtag', 'Service']
issue_autofill_values: List[str] = []

def _split_issue_autofill_custom_items(raw_value: Any) -> List[str]:
    """Return clean custom issue suggestions from Settings."""
    try:
        if raw_value is None:
            return []
        if isinstance(raw_value, (list, tuple)):
            candidates = raw_value
        else:
            candidates = str(raw_value).replace(';', '\n').splitlines()
        cleaned = []
        seen = set()
        for item in candidates:
            value = str(item or '').strip()
            key = value.casefold()
            if value and key not in seen:
                seen.add(key)
                cleaned.append(value)
        return cleaned
    except Exception:
        return []

def refresh_issue_autofill_values(settings: Optional[Dict]=None) -> None:
    """Load issue suggestions from Settings and, optionally, saved issue history."""
    global issue_autofill_values
    try:
        settings = settings or load_settings()
        if not settings.get('issue_autofill_enabled', True):
            issue_autofill_values = []
            return
        values = []
        seen = set()

        def _add(value: Any) -> None:
            value_text = str(value or '').strip()
            key = value_text.casefold()
            if value_text and key not in seen:
                seen.add(key)
                values.append(value_text)
        for item in _split_issue_autofill_custom_items(settings.get('issue_autofill_custom_items', '')):
            _add(item)
        if settings.get('issue_autofill_include_history', True):
            conn = None
            try:
                conn = get_db_connection()
                cursor = conn.cursor()
                union_parts = [f"SELECT TRIM(issue{i}) AS issue FROM customer_bikes WHERE TRIM(COALESCE(issue{i}, '')) <> ''" for i in range(1, 11)]
                query = '\n                    SELECT issue, COUNT(*) AS usage_count\n                    FROM (\n                        {union_sql}\n                    )\n                    GROUP BY LOWER(issue)\n                    ORDER BY usage_count DESC, issue COLLATE NOCASE ASC\n                    LIMIT ?\n                '.format(union_sql=' UNION ALL '.join(union_parts))
                cursor.execute(query, (ISSUE_AUTOFILL_MAX_DB_VALUES,))
                for row in cursor.fetchall():
                    try:
                        _add(row['issue'])
                    except Exception:
                        _add(row[0])
            except Exception:
                pass
            finally:
                if conn:
                    release_db_connection(conn)
        if not values:
            for item in DEFAULT_ISSUE_AUTOFILL_ITEMS:
                _add(item)
        issue_autofill_values = values[:ISSUE_AUTOFILL_MAX_DB_VALUES]
    except Exception:
        issue_autofill_values = issue_autofill_values if issue_autofill_values else list(DEFAULT_ISSUE_AUTOFILL_ITEMS)

def _get_issue_autofill_matches(query: str) -> List[str]:
    """Return issue suggestions. Prefix matches are shown before contains matches."""
    try:
        query = str(query or '').strip().casefold()
        values = list(issue_autofill_values)
        if not values:
            return []
        if not query:
            return values[:ISSUE_AUTOFILL_VISIBLE_ROWS]
        prefix_matches = [v for v in values if v.casefold().startswith(query)]
        contains_matches = [v for v in values if query in v.casefold() and (not v.casefold().startswith(query))]
        exact_key = query
        matches = [v for v in prefix_matches + contains_matches if v.casefold() != exact_key]
        return matches[:ISSUE_AUTOFILL_VISIBLE_ROWS]
    except Exception:
        return []

def _get_issue_autofill_settings_cached() -> Dict:
    """Return cached settings so autocomplete never waits on the database while typing."""
    try:
        cached = _app_cache.get('settings')
        if isinstance(cached, dict):
            return cached
    except Exception:
        pass
    try:
        return load_settings()
    except Exception:
        return DEFAULT_SETTINGS.copy()

def _ensure_issue_autofill_values_ready() -> None:
    """Keep a small default suggestion list available without database access."""
    global issue_autofill_values
    try:
        if not issue_autofill_values:
            issue_autofill_values = list(DEFAULT_ISSUE_AUTOFILL_ITEMS)
    except Exception:
        issue_autofill_values = list(DEFAULT_ISSUE_AUTOFILL_ITEMS)

def hide_issue_autofill() -> None:
    """Hide the active issue suggestion popup."""
    try:
        popup = getattr(root, '_issue_autofill_popup', None)
        if popup is not None and popup.winfo_exists():
            popup.destroy()
        if root is not None:
            root._issue_autofill_popup = None
            root._issue_autofill_entry = None
            root._issue_autofill_listbox = None
    except Exception:
        pass

def _accept_issue_autofill_value(entry: tk.Entry, value: str) -> str:
    """Insert the selected suggestion into an issue field and move to the next field."""
    try:
        entry.delete(0, tk.END)
        entry.insert(0, value)
        entry.icursor(tk.END)
    except Exception:
        pass
    hide_issue_autofill()
    try:
        if root is not None and entry is getattr(root, '_main_issue_entry', None):
            entry.focus_set()
        else:
            entry.tk_focusNext().focus()
    except Exception:
        try:
            entry.focus_set()
        except Exception:
            pass
    return 'break'

def _show_issue_autofill(entry: tk.Entry, force: bool=False) -> None:
    """Show matching issue suggestions directly under the active issue entry."""
    try:
        if root is None or entry is None:
            return
        settings = _get_issue_autofill_settings_cached()
        if not settings.get('issue_autofill_enabled', True):
            hide_issue_autofill()
            return
        _ensure_issue_autofill_values_ready()
        matches = _get_issue_autofill_matches(entry.get())
        if not matches and (not force):
            hide_issue_autofill()
            return
        if not matches:
            matches = issue_autofill_values[:ISSUE_AUTOFILL_VISIBLE_ROWS]
        if not matches:
            hide_issue_autofill()
            return
        root.update_idletasks()
        entry.update_idletasks()
        popup = getattr(root, '_issue_autofill_popup', None)
        listbox = getattr(root, '_issue_autofill_listbox', None)
        if popup is None or not popup.winfo_exists():
            popup = tk.Toplevel(root)
            popup.withdraw()
            popup.overrideredirect(True)
            popup.resizable(False, False)
            try:
                popup.attributes('-toolwindow', True)
            except Exception:
                pass
            listbox = tk.Listbox(popup, height=min(len(matches), ISSUE_AUTOFILL_VISIBLE_ROWS), activestyle='dotbox', exportselection=False, font=('Arial', FIXED_INPUT_FONT_SIZES['issue']), borderwidth=1, relief='solid')
            listbox.pack(fill='both', expand=True)
            root._issue_autofill_popup = popup
            root._issue_autofill_listbox = listbox
        else:
            popup.withdraw()
            listbox.delete(0, tk.END)
            listbox.config(height=min(len(matches), ISSUE_AUTOFILL_VISIBLE_ROWS))
        for value in matches:
            listbox.insert(tk.END, value)
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(0)
        listbox.activate(0)

        def _pick_from_list(event=None):
            try:
                selected = listbox.curselection()
                index = int(selected[0]) if selected else int(listbox.index('active'))
                value = listbox.get(index)
                return _accept_issue_autofill_value(entry, value)
            except Exception:
                return 'break'

        def _return_focus(event=None):
            try:
                entry.focus_set()
            except Exception:
                pass
            return 'break'
        listbox.bind('<ButtonRelease-1>', _pick_from_list)
        listbox.bind('<Return>', _pick_from_list)
        listbox.bind('<KP_Enter>', _pick_from_list)
        listbox.bind('<Escape>', lambda event=None: (hide_issue_autofill(), entry.focus_set(), 'break')[-1])
        listbox.bind('<Left>', _return_focus)
        listbox.bind('<Right>', _return_focus)
        listbox.bind('<Up>', lambda event=None: (_select_issue_autofill_index(listbox, (int(listbox.curselection()[0]) if listbox.curselection() else 0) - 1), 'break')[-1])
        listbox.bind('<Down>', lambda event=None: (_select_issue_autofill_index(listbox, (int(listbox.curselection()[0]) if listbox.curselection() else 0) + 1), 'break')[-1])
        width = max(int(entry.winfo_width()), 180)
        x = int(entry.winfo_rootx())
        y = int(entry.winfo_rooty() + entry.winfo_height() + 1)
        try:
            font_obj = tkfont.Font(font=listbox.cget('font'))
            row_height = max(22, int(font_obj.metrics('linespace')) + 6)
        except Exception:
            row_height = 24
        popup_height = max(24, min(len(matches), ISSUE_AUTOFILL_VISIBLE_ROWS) * row_height + 2)
        popup.geometry(f'{width}x{popup_height}+{x}+{y}')
        popup.deiconify()
        try:
            popup.lift(root)
        except Exception:
            popup.lift()
        root._issue_autofill_entry = entry
    except Exception:
        hide_issue_autofill()

def _select_issue_autofill_index(listbox: tk.Listbox, index: int) -> None:
    """Select one suggestion row without moving focus out of the active issue field."""
    try:
        if listbox is None or not listbox.winfo_exists() or listbox.size() <= 0:
            return
        index = max(0, min(int(index), listbox.size() - 1))
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(index)
        listbox.activate(index)
        listbox.see(index)
    except Exception:
        pass

def _move_issue_autofill_selection(entry: tk.Entry, direction: int, event=None) -> str:
    """Move the highlighted autocomplete item with Up/Down while focus stays in the entry."""
    try:
        popup = getattr(root, '_issue_autofill_popup', None)
        listbox = getattr(root, '_issue_autofill_listbox', None)
        popup_visible = popup is not None and popup.winfo_exists() and popup.winfo_viewable()
        if not popup_visible or listbox is None or (not listbox.winfo_exists()) or (listbox.size() <= 0):
            _show_issue_autofill(entry, force=True)
            listbox = getattr(root, '_issue_autofill_listbox', None)
            if listbox is not None and listbox.winfo_exists() and (listbox.size() > 0):
                _select_issue_autofill_index(listbox, 0)
            return 'break'
        selected = listbox.curselection()
        if selected:
            current = int(selected[0])
        else:
            try:
                current = int(listbox.index('active'))
            except Exception:
                current = 0
        _select_issue_autofill_index(listbox, current + int(direction))
    except Exception:
        pass
    return 'break'

def _focus_issue_autofill_list(entry: tk.Entry, event=None) -> Optional[str]:
    """Backward-compatible alias: open/navigate suggestions without stealing focus."""
    return _move_issue_autofill_selection(entry, 0, event)

def _schedule_hide_issue_autofill(event=None) -> None:
    """Hide popup after focus changes, but allow mouse clicks inside it first."""
    try:
        if root is None:
            return

        def _hide_if_focus_left():
            try:
                focus_widget = root.focus_get()
                popup = getattr(root, '_issue_autofill_popup', None)
                listbox = getattr(root, '_issue_autofill_listbox', None)
                active_entry = getattr(root, '_issue_autofill_entry', None)
                if focus_widget in (listbox, active_entry):
                    return
                if popup is not None and popup.winfo_exists() and (focus_widget is not None):
                    try:
                        if focus_widget.winfo_toplevel() == popup:
                            return
                    except Exception:
                        pass
                hide_issue_autofill()
            except Exception:
                hide_issue_autofill()
        root.after(220, _hide_if_focus_left)
    except Exception:
        pass

def _on_issue_autofill_keyrelease(entry: tk.Entry, event=None) -> None:
    """Refresh suggestions while typing in issue fields."""
    try:
        ignored = {'Up', 'Down', 'Left', 'Right', 'Return', 'KP_Enter', 'Tab', 'Escape', 'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R'}
        if event is not None and getattr(event, 'keysym', '') in ignored:
            return
        _show_issue_autofill(entry)
    except Exception:
        pass

def _focus_next_from_issue_entry(entry: tk.Entry) -> str:
    """Move focus to the next widget from an issue entry."""
    try:
        entry.tk_focusNext().focus()
    except Exception:
        pass
    hide_issue_autofill()
    return 'break'

def _accept_issue_autofill_or_focus_next(entry: tk.Entry, event=None) -> str:
    """Accept the visible/first issue suggestion, otherwise keep the old Enter behavior."""
    try:
        popup = getattr(root, '_issue_autofill_popup', None)
        listbox = getattr(root, '_issue_autofill_listbox', None)
        popup_visible = popup is not None and popup.winfo_exists()
        if popup_visible and listbox is not None and listbox.winfo_exists() and (listbox.size() > 0):
            selected = listbox.curselection()
            index = int(selected[0]) if selected else 0
            return _accept_issue_autofill_value(entry, listbox.get(index))
        matches = _get_issue_autofill_matches(entry.get())
        if matches and str(entry.get() or '').strip():
            return _accept_issue_autofill_value(entry, matches[0])
    except Exception:
        pass
    return _focus_next_from_issue_entry(entry)

def setup_issue_autofill(entry: tk.Entry) -> None:
    """Attach autofill behavior to a main issue input field."""
    try:
        entry.bind('<KeyRelease>', lambda event, e=entry: _on_issue_autofill_keyrelease(e, event), add='+')
        entry.bind('<Down>', lambda event, e=entry: _move_issue_autofill_selection(e, 1, event))
        entry.bind('<Up>', lambda event, e=entry: _move_issue_autofill_selection(e, -1, event))
        entry.bind('<Control-space>', lambda event, e=entry: (_show_issue_autofill(e, force=True), 'break')[-1])
        entry.bind('<FocusIn>', lambda event, e=entry: _ensure_issue_autofill_values_ready(), add='+')
        entry.bind('<Return>', lambda event, e=entry: _accept_issue_autofill_or_focus_next(e, event))
        entry.bind('<KP_Enter>', lambda event, e=entry: _accept_issue_autofill_or_focus_next(e, event))
        entry.bind('<FocusOut>', _schedule_hide_issue_autofill, add='+')
        entry.bind('<Escape>', lambda event=None: (hide_issue_autofill(), 'break')[-1])
    except Exception:
        pass


def _main_issue_tree_items() -> List[str]:
    """Return issue values from the visible main-window issue table."""
    try:
        tree = getattr(root, '_main_issues_tree', None)
        if tree is not None and tree.winfo_exists():
            values = []
            for item in tree.get_children(''):
                row_values = tree.item(item, 'values') or []
                issue = str(row_values[0] if row_values else '').strip()
                if issue:
                    values.append(issue)
            return values
    except Exception:
        pass
    try:
        return [entry.get().strip() for entry in issue_entries if entry.get().strip()]
    except Exception:
        return []


def _sync_issue_entries_from_main_tree() -> None:
    """Keep the original ten database issue fields in sync with the visible table."""
    try:
        values = _main_issue_tree_items()[:10]
        for i, entry in enumerate(issue_entries[:10]):
            entry.delete(0, tk.END)
            if i < len(values):
                entry.insert(0, values[i])
    except Exception:
        pass


def refresh_main_issue_table_from_entries() -> None:
    """Rebuild the visible issue table from the hidden database fields."""
    try:
        tree = getattr(root, '_main_issues_tree', None)
        if tree is None or not tree.winfo_exists():
            return
        for item in tree.get_children(''):
            tree.delete(item)
        for entry in issue_entries[:10]:
            issue = entry.get().strip()
            if issue:
                tree.insert('', 'end', values=(issue,))
        _update_main_issue_count_label()
    except Exception:
        pass


def _update_main_issue_count_label() -> None:
    """Update the small issue counter under the issue table."""
    try:
        label = getattr(root, '_main_issue_count_label', None)
        if label is None or not label.winfo_exists():
            return
        count = len(_main_issue_tree_items())
        label.config(text=f'{count}/10 selected repairs')
        _update_main_form_summary()
    except Exception:
        pass


def _set_main_issue_edit_item(item_id: str='') -> None:
    """Switch the quick issue input between add and edit mode."""
    try:
        root._main_issue_edit_item = item_id or ''
        button = getattr(root, '_main_issue_add_button', None)
        if button is not None and button.winfo_exists():
            button.config(text='Update' if item_id else '+ Add')
    except Exception:
        pass


def add_issue_from_main_entry(event=None) -> str:
    """Add or update one issue from the single quick issue input."""
    try:
        entry = getattr(root, '_main_issue_entry', None)
        tree = getattr(root, '_main_issues_tree', None)
        if entry is None or tree is None or not tree.winfo_exists():
            return 'break'
        value = ' '.join(str(entry.get() or '').strip().split())
        if not value:
            entry.focus_set()
            return 'break'
        edit_item = getattr(root, '_main_issue_edit_item', '')
        if edit_item and edit_item in tree.get_children(''):
            tree.item(edit_item, values=(value,))
            _set_main_issue_edit_item('')
        else:
            if len(tree.get_children('')) >= 10:
                try:
                    status_label.config(text='Maximum 10 issues per receipt.', fg='red')
                except Exception:
                    pass
                return 'break'
            tree.insert('', 'end', values=(value,))
        entry.delete(0, tk.END)
        entry.focus_set()
        _sync_issue_entries_from_main_tree()
        _update_main_issue_count_label()
        try:
            status_label.config(text='', fg='black')
        except Exception:
            pass
        _refresh_main_input_visual_state()
    except Exception:
        pass
    return 'break'


def accept_issue_autofill_or_add_to_table(event=None) -> str:
    """Enter in the quick issue input accepts the visible suggestion and adds it to the table."""
    try:
        entry = getattr(root, '_main_issue_entry', None)
        popup = getattr(root, '_issue_autofill_popup', None)
        listbox = getattr(root, '_issue_autofill_listbox', None)
        popup_visible = popup is not None and popup.winfo_exists() and popup.winfo_viewable()
        if entry is not None and popup_visible and listbox is not None and listbox.winfo_exists() and listbox.size() > 0:
            selected = listbox.curselection()
            index = int(selected[0]) if selected else 0
            value = listbox.get(index)
            entry.delete(0, tk.END)
            entry.insert(0, value)
            entry.icursor(tk.END)
            hide_issue_autofill()
        return add_issue_from_main_entry(event)
    except Exception:
        return 'break'


def edit_selected_main_issue(event=None) -> str:
    """Load the selected issue table row into the quick input for editing."""
    try:
        tree = getattr(root, '_main_issues_tree', None)
        entry = getattr(root, '_main_issue_entry', None)
        if tree is None or entry is None:
            return 'break'
        selected = tree.selection()
        if not selected:
            return 'break'
        item = selected[0]
        row_values = tree.item(item, 'values') or []
        issue = str(row_values[0] if row_values else '').strip()
        entry.delete(0, tk.END)
        entry.insert(0, issue)
        entry.focus_set()
        entry.icursor(tk.END)
        _set_main_issue_edit_item(item)
    except Exception:
        pass
    return 'break'


def remove_selected_main_issue(event=None) -> str:
    """Remove selected rows from the issue table."""
    try:
        tree = getattr(root, '_main_issues_tree', None)
        if tree is None:
            return 'break'
        for item in tree.selection():
            tree.delete(item)
        _set_main_issue_edit_item('')
        _sync_issue_entries_from_main_tree()
        _update_main_issue_count_label()
    except Exception:
        pass
    return 'break'


def clear_main_issue_table() -> None:
    """Clear the visible issue table and the hidden issue fields."""
    try:
        tree = getattr(root, '_main_issues_tree', None)
        if tree is not None and tree.winfo_exists():
            for item in tree.get_children(''):
                tree.delete(item)
        entry = getattr(root, '_main_issue_entry', None)
        if entry is not None and entry.winfo_exists():
            entry.delete(0, tk.END)
        _set_main_issue_edit_item('')
        for issue_entry in issue_entries:
            issue_entry.delete(0, tk.END)
        _update_main_issue_count_label()
    except Exception:
        pass


def reload_main_visit_history(phone: Optional[str]=None) -> None:
    """Reload the embedded visit-history table below the main input fields."""
    try:
        if root is None:
            return
        tree = getattr(root, '_main_visit_history_tree', None)
        status_var = getattr(root, '_main_visit_history_status_var', None)
        if tree is None or not tree.winfo_exists():
            return
        clean = _phone_digits(phone if phone is not None else (phone_entry.get().strip() if phone_entry else ''))
        root._main_visit_history_phone = clean
        for item in tree.get_children(''):
            tree.delete(item)
        if len(clean) < 6:
            if status_var:
                status_var.set('Enter phone number to see previous visits.')
            return
        if status_var:
            status_var.set('Loading visit history...')

        def _thread() -> None:
            rows = []
            error = None
            conn = None
            try:
                conn = get_db_connection()
                cur = conn.cursor()
                cur.execute("""
                    SELECT datetime, phone, model,
                           issue1, issue2, issue3, issue4, issue5,
                           issue6, issue7, issue8, issue9, issue10
                    FROM customer_bikes
                    WHERE REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(phone,' ',''),'-',''),'+',''),'(',''),')','') LIKE ?
                    ORDER BY datetime DESC
                    LIMIT 12
                """, (f'%{clean}%',))
                rows = cur.fetchall()
            except Exception as exc:
                error = exc
            finally:
                if conn:
                    try:
                        release_db_connection(conn)
                    except Exception:
                        pass

            def _populate() -> None:
                try:
                    current_clean = getattr(root, '_main_visit_history_phone', '')
                    if current_clean != clean:
                        return
                    tree = getattr(root, '_main_visit_history_tree', None)
                    status_var = getattr(root, '_main_visit_history_status_var', None)
                    if tree is None or not tree.winfo_exists():
                        return
                    for item in tree.get_children(''):
                        tree.delete(item)
                    if error is not None:
                        if status_var:
                            status_var.set(f'Error loading history: {error}')
                        return
                    for r in rows:
                        dt, row_phone, model, *issues = r
                        issues_list = [str(i).strip() for i in issues if i and str(i).strip()]
                        issues_str = ' | '.join(issues_list) if issues_list else '-'
                        tree.insert('', 'end', values=(visit_date_label(dt), visit_age_label(dt), row_phone, model, issues_str))
                    if status_var:
                        if rows:
                            status_var.set(f'{len(rows)} previous visit(s). Double-click a row to reuse the bike model.')
                        else:
                            status_var.set('No history found for this phone number.')
                except Exception:
                    pass
            try:
                root.after(0, _populate)
            except Exception:
                pass
        threading.Thread(target=_thread, daemon=True).start()
    except Exception:
        pass


def use_selected_main_history_bike(event=None) -> str:
    """Copy the selected embedded history row's bike model into the main bike field."""
    try:
        tree = getattr(root, '_main_visit_history_tree', None)
        if tree is None:
            return 'break'
        selected = tree.selection()
        if not selected:
            return 'break'
        values = tree.item(selected[0], 'values') or []
        if len(values) < 4:
            return 'break'
        row_phone = str(values[2]).strip()
        model = str(values[3]).strip()
        if phone_entry is not None and row_phone:
            phone_entry.delete(0, tk.END)
            phone_entry.insert(0, row_phone)
        if model_entry is not None:
            model_entry.delete(0, tk.END)
            model_entry.insert(0, model)
        entry = getattr(root, '_main_issue_entry', None)
        if entry is not None and entry.winfo_exists():
            entry.focus_set()
        show_toast(root, 'Bike model copied from visit history.', 2000)
    except Exception:
        pass
    return 'break'
BIKE_BRAND_AUTOFILL_MAX_DB_VALUES = 300
BIKE_COLOR_AUTOFILL_MAX_DB_VALUES = 200
BIKE_MODEL_AUTOFILL_MAX_DB_VALUES = 500
BIKE_BRAND_AUTOFILL_VISIBLE_ROWS = 9
DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS = ['Batavus', 'Bergamont', 'Bianchi', 'BMC', 'Brompton', 'Cannondale', 'Cargobike', 'Canyon', 'Centurion', 'Cervélo', 'Crescent', 'Cube', 'Ecoride', 'Electra', 'Felt', 'Focus', 'Fuji', 'Gazelle', 'Giant', 'Ghost', 'Haibike', 'Kalkhoff', 'Kildemoes', 'Kona', 'Lapierre', 'Liv', 'Merida', 'Monark', 'Moustache', 'Nishiki', 'Norco', 'Orbea', 'Peugeot', 'Puky', 'Raleigh', 'Riese & Müller', 'Rose', 'Santa Cruz', 'Scott', 'Sensa', 'Sjösala', 'Skeppshult', 'Specialized', 'Stevens', 'Trek', 'Tunturi', 'Urban Arrow', 'VanMoof', 'Winther', 'Woom', 'Yamaha']
DEFAULT_BIKE_COLOR_AUTOFILL_ITEMS = ['Svart', 'Matt Svart', 'Vit', 'Grå', 'Silver', 'Röd', 'Blå', 'Mörkblå', 'Ljusblå', 'Grön', 'Mörkgrön', 'Gul', 'Orange', 'Rosa', 'Lila', 'Brun', 'Beige', 'Turkos', 'Guld', 'Bordeaux']
DEFAULT_BIKE_MODEL_AUTOFILL_ITEMS = ['Damcykel', 'Herrcykel', 'Unisexcykel', 'Citybike', 'Elcykel', 'Hybridcykel', 'Mountainbike', 'MTB', 'Racercykel', 'Gravelbike', 'Barncykel', 'Juniorcykel', 'Lastcykel', 'Lådcykel', 'Trekkingcykel', 'Comfortcykel', 'Sportcykel', 'Cargobike']
bike_brand_autofill_values: List[str] = []
bike_color_autofill_values: List[str] = []
bike_model_autofill_values: List[str] = []

def _split_bike_autofill_custom_items(raw_value: Any, normalizer=None) -> List[str]:
    """Return clean custom suggestions from Settings."""
    try:
        if raw_value is None:
            return []
        if isinstance(raw_value, (list, tuple)):
            candidates = raw_value
        else:
            candidates = str(raw_value).replace(';', '\n').splitlines()
        cleaned = []
        seen = set()
        normalizer = normalizer or (lambda v: str(v or '').strip().casefold())
        for item in candidates:
            value = str(item or '').strip()
            key = normalizer(value)
            if value and key and (key not in seen):
                seen.add(key)
                cleaned.append(value)
        return cleaned
    except Exception:
        return []

def _split_bike_brand_autofill_custom_items(raw_value: Any) -> List[str]:
    """Return clean custom bike brand suggestions from Settings."""
    return _split_bike_autofill_custom_items(raw_value, _normalize_bike_brand_key)

def _split_bike_color_autofill_custom_items(raw_value: Any) -> List[str]:
    """Return clean custom bike color suggestions from Settings."""
    return _split_bike_autofill_custom_items(raw_value, _normalize_bike_brand_key)

def _split_bike_model_autofill_custom_items(raw_value: Any) -> List[str]:
    """Return clean custom bike model suggestions from Settings."""
    return _split_bike_autofill_custom_items(raw_value, _normalize_bike_brand_key)

def _normalize_bike_brand_key(value: Any) -> str:
    """Return a comparison key for bike descriptor suggestions."""
    return str(value or '').strip().casefold().replace('ü', 'u').replace('ä', 'a').replace('å', 'a').replace('ö', 'o').replace('é', 'e').replace('è', 'e').replace('&', 'and')

def _compact_bike_descriptor(value: Any) -> str:
    """Normalize spaces in the one-line bike descriptor field."""
    try:
        return ' '.join(str(value or '').replace('/', ' ').split())
    except Exception:
        return str(value or '').strip()

def _known_values_for_matching(values: Optional[List[str]], defaults: List[str]) -> List[str]:
    """Return a non-empty list for prefix matching."""
    try:
        result = [str(v).strip() for v in values or [] if str(v).strip()]
        return result or list(defaults)
    except Exception:
        return list(defaults)

def _split_after_known_prefix(text_value: Any, known_values: Optional[List[str]]) -> Tuple[str, str]:
    """Return (known prefix, remaining text) when text starts with a known suggestion."""
    compact = _compact_bike_descriptor(text_value)
    if not compact:
        return ('', '')
    words = compact.split()
    try:
        sorted_known = sorted([str(v).strip() for v in known_values or [] if str(v).strip()], key=lambda item: (len(_compact_bike_descriptor(item).split()), len(str(item))), reverse=True)
        for value in sorted_known:
            value_words = _compact_bike_descriptor(value).split()
            if not value_words or len(value_words) > len(words):
                continue
            candidate = ' '.join(words[:len(value_words)])
            if _normalize_bike_brand_key(candidate) == _normalize_bike_brand_key(value):
                return (value, ' '.join(words[len(value_words):]).strip())
    except Exception:
        pass
    return ('', compact)

def _extract_bike_brand_from_model(model_value: Any, known_brands: Optional[List[str]]=None) -> str:
    """Extract a useful brand-like suggestion from a saved bike model field."""
    value = _compact_bike_descriptor(model_value)
    if not value:
        return ''
    try:
        known = known_brands or DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS
        brand, _rest = _split_after_known_prefix(value, known)
        if brand:
            return str(brand).strip()
    except Exception:
        pass
    first = value.split(' ', 1)[0].strip(' ,.;:-_()[]{}')
    if 2 <= len(first) <= 24:
        return first
    return ''

def _extract_bike_color_from_model(model_value: Any, known_brands: Optional[List[str]]=None, known_colors: Optional[List[str]]=None) -> str:
    """Extract a bike color suggestion from a saved one-field bike description."""
    try:
        value = _compact_bike_descriptor(model_value)
        if not value:
            return ''
        brands = known_brands or DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS
        colors = known_colors or DEFAULT_BIKE_COLOR_AUTOFILL_ITEMS
        _brand, after_brand = _split_after_known_prefix(value, brands)
        if not after_brand:
            return ''
        color, _after_color = _split_after_known_prefix(after_brand, colors)
        return str(color).strip()
    except Exception:
        return ''

def _extract_bike_model_from_model(model_value: Any, known_brands: Optional[List[str]]=None, known_colors: Optional[List[str]]=None) -> str:
    """Extract the model part after brand and color from a saved one-field bike description."""
    try:
        value = _compact_bike_descriptor(model_value)
        if not value:
            return ''
        brands = known_brands or DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS
        colors = known_colors or DEFAULT_BIKE_COLOR_AUTOFILL_ITEMS
        _brand, after_brand = _split_after_known_prefix(value, brands)
        if not after_brand:
            return ''
        _color, after_color = _split_after_known_prefix(after_brand, colors)
        model_part = after_color.strip(' ,.;:-_()[]{}')
        if 2 <= len(model_part) <= 60:
            return model_part
    except Exception:
        pass
    return ''

def refresh_bike_brand_autofill_values(settings: Optional[Dict]=None) -> None:
    """Load bike brand, color, and model suggestions from Settings and saved history."""
    global bike_brand_autofill_values, bike_color_autofill_values, bike_model_autofill_values
    try:
        settings = settings or load_settings()
        if not settings.get('bike_brand_autofill_enabled', True):
            bike_brand_autofill_values = []
            bike_color_autofill_values = []
            bike_model_autofill_values = []
            return
        brand_values: List[str] = []
        color_values: List[str] = []
        model_values: List[str] = []
        seen_brands = set()
        seen_colors = set()
        seen_models = set()

        def _add_to(target: List[str], seen: set, value: Any, limit_key=None) -> None:
            value_text = str(value or '').strip()
            key = _normalize_bike_brand_key(value_text)
            if value_text and key and (key not in seen):
                seen.add(key)
                target.append(value_text)
        brand_custom_items = _split_bike_brand_autofill_custom_items(settings.get('bike_brand_autofill_custom_items', ''))
        color_custom_items = _split_bike_color_autofill_custom_items(settings.get('bike_color_autofill_custom_items', ''))
        model_custom_items = _split_bike_model_autofill_custom_items(settings.get('bike_model_autofill_custom_items', ''))
        for brand in brand_custom_items or DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS:
            _add_to(brand_values, seen_brands, brand)
        for color in color_custom_items or DEFAULT_BIKE_COLOR_AUTOFILL_ITEMS:
            _add_to(color_values, seen_colors, color)
        for model in model_custom_items or DEFAULT_BIKE_MODEL_AUTOFILL_ITEMS:
            _add_to(model_values, seen_models, model)
        known_brands = _known_values_for_matching(brand_values, DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS)
        known_colors = _known_values_for_matching(color_values, DEFAULT_BIKE_COLOR_AUTOFILL_ITEMS)
        if settings.get('bike_brand_autofill_include_history', True):
            conn = None
            try:
                conn = get_db_connection()
                cursor = conn.cursor()
                cursor.execute("\n                    SELECT TRIM(model) AS model, COUNT(*) AS usage_count\n                    FROM customer_bikes\n                    WHERE TRIM(COALESCE(model, '')) <> ''\n                    GROUP BY LOWER(model)\n                    ORDER BY usage_count DESC, model COLLATE NOCASE ASC\n                    LIMIT ?\n                ", (max(BIKE_BRAND_AUTOFILL_MAX_DB_VALUES, BIKE_MODEL_AUTOFILL_MAX_DB_VALUES),))
                for row in cursor.fetchall():
                    try:
                        model_value = row['model']
                    except Exception:
                        model_value = row[0]
                    _add_to(brand_values, seen_brands, _extract_bike_brand_from_model(model_value, known_brands))
                    _add_to(color_values, seen_colors, _extract_bike_color_from_model(model_value, known_brands, known_colors))
                    _add_to(model_values, seen_models, _extract_bike_model_from_model(model_value, known_brands, known_colors))
            except Exception:
                pass
            finally:
                if conn:
                    release_db_connection(conn)
        bike_brand_autofill_values = brand_values[:BIKE_BRAND_AUTOFILL_MAX_DB_VALUES]
        bike_color_autofill_values = color_values[:BIKE_COLOR_AUTOFILL_MAX_DB_VALUES]
        bike_model_autofill_values = model_values[:BIKE_MODEL_AUTOFILL_MAX_DB_VALUES]
    except Exception:
        bike_brand_autofill_values = bike_brand_autofill_values if bike_brand_autofill_values else list(DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS)
        bike_color_autofill_values = bike_color_autofill_values if bike_color_autofill_values else list(DEFAULT_BIKE_COLOR_AUTOFILL_ITEMS)
        bike_model_autofill_values = bike_model_autofill_values if bike_model_autofill_values else list(DEFAULT_BIKE_MODEL_AUTOFILL_ITEMS)

def _ensure_bike_brand_autofill_values_ready() -> None:
    """Keep default brand/color/model lists available without database access while typing."""
    global bike_brand_autofill_values, bike_color_autofill_values, bike_model_autofill_values
    try:
        if not bike_brand_autofill_values:
            bike_brand_autofill_values = list(DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS)
        if not bike_color_autofill_values:
            bike_color_autofill_values = list(DEFAULT_BIKE_COLOR_AUTOFILL_ITEMS)
        if not bike_model_autofill_values:
            bike_model_autofill_values = list(DEFAULT_BIKE_MODEL_AUTOFILL_ITEMS)
    except Exception:
        bike_brand_autofill_values = list(DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS)
        bike_color_autofill_values = list(DEFAULT_BIKE_COLOR_AUTOFILL_ITEMS)
        bike_model_autofill_values = list(DEFAULT_BIKE_MODEL_AUTOFILL_ITEMS)

def _get_bike_brand_autofill_settings_cached() -> Dict:
    """Return cached settings so bike descriptor autocomplete never waits on the database while typing."""
    try:
        cached = _app_cache.get('settings')
        if isinstance(cached, dict):
            return cached
    except Exception:
        pass
    try:
        return load_settings()
    except Exception:
        return DEFAULT_SETTINGS.copy()

def _get_bike_descriptor_context(query: str) -> Dict[str, str]:
    """Detect which part of the one-field bike descriptor is currently being typed.

    Known brands are detected from the saved brand list. Unknown/new brands are
    also supported: after the user types a space after a new brand, the field
    moves to the color stage, and after a color it moves to the model/type stage.
    """
    raw_text = str(query or '')
    text_value = _compact_bike_descriptor(raw_text)
    brands = _known_values_for_matching(bike_brand_autofill_values, DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS)
    colors = _known_values_for_matching(bike_color_autofill_values, DEFAULT_BIKE_COLOR_AUTOFILL_ITEMS)
    brand, after_brand = _split_after_known_prefix(text_value, brands)
    if brand:
        if not after_brand:
            return {'stage': 'color', 'brand': brand, 'color': '', 'query': ''}
        color, after_color = _split_after_known_prefix(after_brand, colors)
        if not color:
            return {'stage': 'color', 'brand': brand, 'color': '', 'query': after_brand}
        return {'stage': 'model', 'brand': brand, 'color': color, 'query': after_color}
    if text_value:
        has_trailing_separator = bool(raw_text and raw_text[-1].isspace())
        parts = text_value.split(' ', 1)
        possible_brand = parts[0].strip(' ,.;:-_()[]{}')
        rest = parts[1].strip() if len(parts) > 1 else ''
        if possible_brand and 2 <= len(possible_brand) <= 40 and (not _is_existing_bike_non_brand_value(possible_brand)) and (rest or has_trailing_separator):
            if not rest:
                return {'stage': 'color', 'brand': possible_brand, 'color': '', 'query': ''}
            color, after_color = _split_after_known_prefix(rest, colors)
            if not color:
                return {'stage': 'color', 'brand': possible_brand, 'color': '', 'query': rest}
            return {'stage': 'model', 'brand': possible_brand, 'color': color, 'query': after_color}
    return {'stage': 'brand', 'brand': '', 'color': '', 'query': text_value}

def _get_bike_values_for_stage(stage: str) -> List[str]:
    """Return suggestion values for the active bike descriptor stage."""
    if stage == 'color':
        return list(bike_color_autofill_values or DEFAULT_BIKE_COLOR_AUTOFILL_ITEMS)
    if stage == 'model':
        return list(bike_model_autofill_values or DEFAULT_BIKE_MODEL_AUTOFILL_ITEMS)
    return list(bike_brand_autofill_values or DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS)

def _get_bike_stage_label(stage: str) -> str:
    """Return a short label for the active suggestion list."""
    if stage == 'color':
        return 'Color'
    if stage == 'model':
        return 'Model'
    return 'Brand'

def _get_bike_brand_autofill_matches(query: str) -> List[str]:
    """Return matching bike brand/color/model suggestions based on the current stage."""
    try:
        context = _get_bike_descriptor_context(query)
        query_key = _normalize_bike_brand_key(context.get('query', ''))
        values = _get_bike_values_for_stage(context.get('stage', 'brand'))
        if not values:
            return []
        if not query_key:
            return values[:BIKE_BRAND_AUTOFILL_VISIBLE_ROWS]
        prefix_matches = [v for v in values if _normalize_bike_brand_key(v).startswith(query_key)]
        contains_matches = [v for v in values if query_key in _normalize_bike_brand_key(v) and (not _normalize_bike_brand_key(v).startswith(query_key))]
        matches = [v for v in prefix_matches + contains_matches if _normalize_bike_brand_key(v) != query_key]
        return matches[:BIKE_BRAND_AUTOFILL_VISIBLE_ROWS]
    except Exception:
        return []

def hide_bike_brand_autofill() -> None:
    """Hide the active bike descriptor suggestion popup."""
    try:
        popup = getattr(root, '_bike_brand_autofill_popup', None)
        if popup is not None and popup.winfo_exists():
            popup.destroy()
        if root is not None:
            root._bike_brand_autofill_popup = None
            root._bike_brand_autofill_entry = None
            root._bike_brand_autofill_listbox = None
            root._bike_brand_autofill_label = None
    except Exception:
        pass

def _accept_bike_brand_autofill_value(entry: tk.Entry, value: str) -> str:
    """Insert the selected brand/color/model in the same Bike Model field."""
    try:
        context = _get_bike_descriptor_context(entry.get())
        stage = context.get('stage', 'brand')
        selected_value = str(value or '').strip()
        if not selected_value:
            return 'break'
        if stage == 'brand':
            new_text = f'{selected_value} '
            show_next = True
        elif stage == 'color':
            brand = context.get('brand', '').strip()
            new_text = f'{brand} {selected_value} '.strip() + ' '
            show_next = True
        else:
            brand = context.get('brand', '').strip()
            color = context.get('color', '').strip()
            prefix = ' '.join((part for part in (brand, color) if part))
            new_text = f'{prefix} {selected_value}'.strip()
            show_next = False
        entry.delete(0, tk.END)
        entry.insert(0, new_text)
        entry.icursor(tk.END)
        entry.focus_set()
    except Exception:
        show_next = False
    hide_bike_brand_autofill()
    try:
        if show_next and root is not None:
            root.after(60, lambda e=entry: _show_bike_brand_autofill(e, force=True))
    except Exception:
        pass
    return 'break'

def _show_bike_brand_autofill(entry: tk.Entry, force: bool=False) -> None:
    """Show matching bike descriptor suggestions directly under the Bike Model field."""
    try:
        if root is None or entry is None:
            return
        hide_issue_autofill()
        settings = _get_bike_brand_autofill_settings_cached()
        if not settings.get('bike_brand_autofill_enabled', True):
            hide_bike_brand_autofill()
            return
        _ensure_bike_brand_autofill_values_ready()
        context = _get_bike_descriptor_context(entry.get())
        stage = context.get('stage', 'brand')
        matches = _get_bike_brand_autofill_matches(entry.get())
        if not matches and (not force):
            hide_bike_brand_autofill()
            return
        if not matches:
            matches = _get_bike_values_for_stage(stage)[:BIKE_BRAND_AUTOFILL_VISIBLE_ROWS]
        if not matches:
            hide_bike_brand_autofill()
            return
        root.update_idletasks()
        entry.update_idletasks()
        popup = getattr(root, '_bike_brand_autofill_popup', None)
        listbox = getattr(root, '_bike_brand_autofill_listbox', None)
        title_label = getattr(root, '_bike_brand_autofill_label', None)
        if popup is None or not popup.winfo_exists():
            popup = tk.Toplevel(root)
            popup.withdraw()
            popup.overrideredirect(True)
            popup.resizable(False, False)
            try:
                popup.attributes('-toolwindow', True)
            except Exception:
                pass
            title_label = tk.Label(popup, text=_get_bike_stage_label(stage), anchor='w', font=('Arial', 9, 'bold'), padx=6, pady=2, background='#f3f4f6')
            title_label.pack(fill='x')
            listbox = tk.Listbox(popup, height=min(len(matches), BIKE_BRAND_AUTOFILL_VISIBLE_ROWS), activestyle='dotbox', exportselection=False, font=('Arial', FIXED_INPUT_FONT_SIZES['bike']), borderwidth=1, relief='solid')
            listbox.pack(fill='both', expand=True)
            root._bike_brand_autofill_popup = popup
            root._bike_brand_autofill_listbox = listbox
            root._bike_brand_autofill_label = title_label
        else:
            popup.withdraw()
            if title_label is not None and title_label.winfo_exists():
                title_label.config(text=_get_bike_stage_label(stage))
            listbox.delete(0, tk.END)
            listbox.config(height=min(len(matches), BIKE_BRAND_AUTOFILL_VISIBLE_ROWS))
        for value in matches:
            listbox.insert(tk.END, value)
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(0)
        listbox.activate(0)

        def _pick_from_list(event=None):
            try:
                selected = listbox.curselection()
                index = int(selected[0]) if selected else int(listbox.index('active'))
                value = listbox.get(index)
                return _accept_bike_brand_autofill_value(entry, value)
            except Exception:
                return 'break'

        def _return_focus(event=None):
            try:
                entry.focus_set()
            except Exception:
                pass
            return 'break'
        listbox.bind('<ButtonRelease-1>', _pick_from_list)
        listbox.bind('<Return>', _pick_from_list)
        listbox.bind('<KP_Enter>', _pick_from_list)
        listbox.bind('<Escape>', lambda event=None: (hide_bike_brand_autofill(), entry.focus_set(), 'break')[-1])
        listbox.bind('<Left>', _return_focus)
        listbox.bind('<Right>', _return_focus)
        listbox.bind('<Up>', lambda event=None: (_select_bike_brand_autofill_index(listbox, (int(listbox.curselection()[0]) if listbox.curselection() else 0) - 1), 'break')[-1])
        listbox.bind('<Down>', lambda event=None: (_select_bike_brand_autofill_index(listbox, (int(listbox.curselection()[0]) if listbox.curselection() else 0) + 1), 'break')[-1])
        width = max(int(entry.winfo_width()), 220)
        x = int(entry.winfo_rootx())
        y = int(entry.winfo_rooty() + entry.winfo_height() + 1)
        try:
            font_obj = tkfont.Font(font=listbox.cget('font'))
            row_height = max(22, int(font_obj.metrics('linespace')) + 6)
        except Exception:
            row_height = 24
        popup_height = max(48, min(len(matches), BIKE_BRAND_AUTOFILL_VISIBLE_ROWS) * row_height + 24)
        popup.geometry(f'{width}x{popup_height}+{x}+{y}')
        popup.deiconify()
        try:
            popup.lift(root)
        except Exception:
            popup.lift()
        root._bike_brand_autofill_entry = entry
    except Exception:
        hide_bike_brand_autofill()

def _select_bike_brand_autofill_index(listbox: tk.Listbox, index: int) -> None:
    """Select one bike descriptor suggestion row without moving focus out of the model field."""
    try:
        if listbox is None or not listbox.winfo_exists() or listbox.size() <= 0:
            return
        index = max(0, min(int(index), listbox.size() - 1))
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(index)
        listbox.activate(index)
        listbox.see(index)
    except Exception:
        pass

def _move_bike_brand_autofill_selection(entry: tk.Entry, direction: int, event=None) -> str:
    """Move the highlighted bike descriptor suggestion with Up/Down while focus stays in the model field."""
    try:
        popup = getattr(root, '_bike_brand_autofill_popup', None)
        listbox = getattr(root, '_bike_brand_autofill_listbox', None)
        popup_visible = popup is not None and popup.winfo_exists() and popup.winfo_viewable()
        if not popup_visible or listbox is None or (not listbox.winfo_exists()) or (listbox.size() <= 0):
            _show_bike_brand_autofill(entry, force=True)
            listbox = getattr(root, '_bike_brand_autofill_listbox', None)
            if listbox is not None and listbox.winfo_exists() and (listbox.size() > 0):
                _select_bike_brand_autofill_index(listbox, 0)
            return 'break'
        selected = listbox.curselection()
        if selected:
            current = int(selected[0])
        else:
            try:
                current = int(listbox.index('active'))
            except Exception:
                current = 0
        _select_bike_brand_autofill_index(listbox, current + int(direction))
    except Exception:
        pass
    return 'break'

def _schedule_hide_bike_brand_autofill(event=None) -> None:
    """Hide bike descriptor popup after focus changes, while allowing mouse clicks inside it first."""
    try:
        if root is None:
            return

        def _hide_if_focus_left():
            try:
                focus_widget = root.focus_get()
                popup = getattr(root, '_bike_brand_autofill_popup', None)
                listbox = getattr(root, '_bike_brand_autofill_listbox', None)
                title_label = getattr(root, '_bike_brand_autofill_label', None)
                active_entry = getattr(root, '_bike_brand_autofill_entry', None)
                if focus_widget in (listbox, title_label, active_entry):
                    return
                if popup is not None and popup.winfo_exists() and (focus_widget is not None):
                    try:
                        if focus_widget.winfo_toplevel() == popup:
                            return
                    except Exception:
                        pass
                hide_bike_brand_autofill()
            except Exception:
                hide_bike_brand_autofill()
        root.after(220, _hide_if_focus_left)
    except Exception:
        pass

def _on_bike_brand_autofill_keyrelease(entry: tk.Entry, event=None) -> None:
    """Refresh suggestions while typing brand, color, and model in the same Bike Model field."""
    try:
        ignored = {'Up', 'Down', 'Left', 'Right', 'Return', 'KP_Enter', 'Tab', 'Escape', 'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R'}
        if event is not None and getattr(event, 'keysym', '') in ignored:
            return
        _show_bike_brand_autofill(entry)
    except Exception:
        pass

def _accept_bike_brand_autofill_or_focus_next(entry: tk.Entry, event=None) -> str:
    """Accept visible/first bike descriptor suggestion; otherwise confirm a new brand."""
    try:
        settings = _get_bike_brand_autofill_settings_cached()
        if not settings.get('bike_brand_autofill_enabled', True):
            hide_bike_brand_autofill()
            try:
                entry.tk_focusNext().focus()
            except Exception:
                pass
            return 'break'
        popup = getattr(root, '_bike_brand_autofill_popup', None)
        listbox = getattr(root, '_bike_brand_autofill_listbox', None)
        popup_visible = popup is not None and popup.winfo_exists()
        if popup_visible and listbox is not None and listbox.winfo_exists() and (listbox.size() > 0):
            selected = listbox.curselection()
            index = int(selected[0]) if selected else 0
            return _accept_bike_brand_autofill_value(entry, listbox.get(index))
        current_text = str(entry.get() or '')
        matches = _get_bike_brand_autofill_matches(current_text)
        if matches and current_text.strip():
            return _accept_bike_brand_autofill_value(entry, matches[0])
        context = _get_bike_descriptor_context(current_text)
        candidate_brand = _compact_bike_descriptor(current_text)
        if context.get('stage') == 'brand' and candidate_brand and (2 <= len(candidate_brand) <= 40) and (' ' not in candidate_brand) and (not _is_existing_bike_non_brand_value(candidate_brand)):
            entry.delete(0, tk.END)
            entry.insert(0, candidate_brand + ' ')
            entry.icursor(tk.END)
            entry.focus_set()
            hide_bike_brand_autofill()
            try:
                if root is not None:
                    root.after(60, lambda e=entry: _show_bike_brand_autofill(e, force=True))
            except Exception:
                pass
            return 'break'
    except Exception:
        pass
    hide_bike_brand_autofill()
    try:
        entry.tk_focusNext().focus()
    except Exception:
        pass
    return 'break'

def setup_bike_brand_autofill(entry: tk.Entry) -> None:
    """Attach staged brand/color/model autofill behavior to the Bike Model field."""
    try:
        entry.bind('<KeyRelease>', lambda event, e=entry: _on_bike_brand_autofill_keyrelease(e, event), add='+')
        entry.bind('<Down>', lambda event, e=entry: _move_bike_brand_autofill_selection(e, 1, event))
        entry.bind('<Up>', lambda event, e=entry: _move_bike_brand_autofill_selection(e, -1, event))
        entry.bind('<Control-space>', lambda event, e=entry: (_show_bike_brand_autofill(e, force=True), 'break')[-1])
        entry.bind('<FocusIn>', lambda event, e=entry: _ensure_bike_brand_autofill_values_ready(), add='+')
        entry.bind('<Return>', lambda event, e=entry: _accept_bike_brand_autofill_or_focus_next(e, event))
        entry.bind('<KP_Enter>', lambda event, e=entry: _accept_bike_brand_autofill_or_focus_next(e, event))
        entry.bind('<FocusOut>', _schedule_hide_bike_brand_autofill, add='+')
        entry.bind('<Escape>', lambda event=None: (hide_bike_brand_autofill(), 'break')[-1])
    except Exception:
        pass
ENGLISH_TEXT = {'window_title': 'BikeTrack Pro',
 'tel_label': 'Phone Number:',
 'bike_label': 'Bike Model:',
 'issue_label': 'Issue {0}:',
 'visit_history_title': 'Visit History',
 'visit_history_empty': 'Phone number is empty.',
 'visit_history_not_found': 'No history found for this number.',
 'visit_history_error': 'Error loading history',
 'history_col_datetime': 'Date/Time',
 'history_col_model': 'Model',
 'history_col_issues': 'Issues',
 'history_col_ago': 'Ago',
 'close_button': 'Close',
 'save_button': 'Save',
 'bike_list_button': 'Open Bike List',
 'menu_settings': 'Settings',
 'menu_about': 'About',
 'menu_help': 'Help',
 'status_empty': 'Phone number and bike model cannot be empty!',
 'status_saved': 'Customer data saved and receipt printed!',
 'status_printer': 'No printer selected! Please set a printer in Settings.',
 'status_opened': 'Bike list window opened.',
 'status_open_error': 'Failed to open bike list: {0}',
 'db_error_title': 'Database Error',
 'db_error_message': 'Failed to save to database: {0}',
 'backup_error_title': 'Backup Error',
 'backup_error_message': 'Failed to create backup: {0}',
 'settings_title': 'Settings',
 'shop_info_tab': 'Shop Info',
 'shop_info_label': 'Shop Information:',
 'customer_qr_tab': 'Customer QR',
 'customer_qr_label': 'Customer QR Code Content:',
 'shop_qr_tab': 'Shop QR',
 'shop_qr_label': 'Shop QR Code SMS Message:',
 'printer_tab': 'Printer',
 'printer_label': 'Select Printer:',
 'printer_info': 'Select the Windows printer driver for receipts. Recommended for Star TSP100/futurePRNT.',
 'font_tab': 'Receipt Designer',
 'font_label': 'Receipt Designer:',
 'font_phone': 'Phone Number',
 'font_bike': 'Bike Model',
 'font_datetime': 'Date & Time',
 'font_shop_info': 'Shop Info',
 'font_issue': 'Issues',
 'qr_width_tab': 'QR Width',
 'qr_width_label': 'QR Code Width Settings:',
 'customer_qr_width_label': 'Customer QR Width (pixels):',
 'shop_qr_width_label': 'Shop QR Width (pixels):',
 'no_printer_title': 'Printer Warning',
 'no_printer_message': 'No printer selected. Printing will not work until a printer is chosen.',
 'width_positive': 'QR widths must be positive numbers!',
 'error_title': 'Error',
 'invalid_value': 'Invalid input: {0}',
 'saved_title': 'Settings Saved',
 'about_title': 'About BikeTrack Pro',
 'about_text': 'BikeTrack Pro\n'
               'Version 5.96\n'
               'Developed by Aidin594\n'
               'Girra Sweden AB\n'
               'Email: info@vhcyklar.se\n'
               '\n'
               'A simple tool to track bike shop customers, manage records, and print receipts efficiently.',
 'help_title': 'Help',
 'help_text': 'BikeTrack Pro Help\n'
              '\n'
              '1. Enter customer details and issues.\n'
              "2. Click 'Save' or Ctrl+S to save and print.\n"
              "3. Use 'Open Bike List' to view and manage records.\n"
              '4. Configure settings via the Settings menu.',
 'bike_list_title': 'Bike List',
 'db_load_error': 'Failed to load database: {0}',
 'db_save_error': 'Failed to save to database: {0}',
 'treeview_error': 'Failed to populate list: {0}',
 'delete_info': 'Please select one record to delete.',
 'confirm_delete_title': 'Confirm Delete',
 'confirm_delete_message': 'Are you sure you want to delete this record?',
 'delete_success': 'Record deleted successfully.',
 'undo_title': 'Undo Complete',
 'undo_message': 'Last action has been undone.',
 'search_info': 'Please enter a phone number to search.',
 'select_record_info': 'Please select a record.',
 'select_phone_info': 'Please select a record with a phone number.',
 'phone_label': 'Phone Number:',
 'search_button': 'Search',
 'refresh_button': 'Refresh',
 'settings_button': 'Settings',
 'delete_button': 'Delete',
 'undo_button': 'Undo',
 'context_delete': 'Delete Record',
 'context_generate': 'Generate QR Code',
 'context_search': 'Search Similar Phones',
 'context_copy': 'Copy Phone Number',
 'qr_title': 'QR Code',
 'print_qr_button': 'Print QR Code',
 'qr_error': 'Failed to generate QR code: {0}',
 'printer_warning_title': 'Printer Not Set',
 'printer_warning_message': 'Please select a printer in settings.',
 'print_success_title': 'Print Success',
 'print_success_message': 'QR code printed successfully.',
 'print_error': 'Failed to print QR code: {0}',
 'copy_success': 'Phone number copied to clipboard.',
 'bike_settings_title': 'Bike List Settings',
 'sms_label': 'SMS Message:',
 'bike_printer_label': 'Printer:',
 'empty_lines_label': 'Empty Lines After QR:',
 'lines_error': 'Number of lines must be non-negative!',
 'number_error': 'Please enter a valid number.',
 'bike_saved_title': 'Settings Saved',
 'bike_saved_message': 'Bike list settings saved.',
 'loading_message': 'Loading...',
 'searching_message': 'Searching...',
 'qr_enabled_label': 'Enable QR Code:',
 'customer_qr_enabled': 'Enable Customer QR Code',
 'shop_qr_enabled': 'Enable Shop QR Code',
 'customer_logo_tab': 'Customer Logo',
 'customer_logo_label': 'Logo file:',
 'customer_logo_browse': 'Browse',
 'customer_logo_clear': 'Clear',
 'customer_logo_filetypes': 'Image files',
 'reprint_button': 'Reprint Receipt',
 'reprint_title': 'Reprint Receipt',
 'reprint_select': 'Select Entry to Reprint',
 'reprint_success': 'Receipt reprinted successfully!',
 'reprint_error': 'Error reprinting receipt: {0}',
 'no_entries': 'No entries found in database.'}

def set_window_icon(window: tk.Tk) -> None:
    """Set the window icon if the icon file exists."""
    if os.path.exists(ICON_PATH):
        try:
            window.iconbitmap(ICON_PATH)
        except tk.TclError as e:
            print(f'Failed to set icon: {e}')
    else:
        print(f'Icon file not found at {ICON_PATH}')

def init_database() -> None:
    """Initialize the SQLite database with required tables and indices."""
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('CREATE TABLE IF NOT EXISTS customer_bikes (\n            id INTEGER PRIMARY KEY AUTOINCREMENT,\n            datetime TEXT NOT NULL,\n            phone TEXT NOT NULL,\n            model TEXT NOT NULL,\n            issue1 TEXT, issue2 TEXT, issue3 TEXT, issue4 TEXT, issue5 TEXT,\n            issue6 TEXT, issue7 TEXT, issue8 TEXT, issue9 TEXT, issue10 TEXT\n        )')
        cursor.execute("PRAGMA table_info(customer_bikes)")
        if 'shop_note' not in {col[1] for col in cursor.fetchall()}:
            cursor.execute("ALTER TABLE customer_bikes ADD COLUMN shop_note TEXT NOT NULL DEFAULT ''")
        cursor.execute('CREATE TABLE IF NOT EXISTS settings (\n            id INTEGER PRIMARY KEY AUTOINCREMENT,\n            settings_json TEXT NOT NULL\n        )')
        cursor.execute('CREATE TABLE IF NOT EXISTS customer_flags (\n            id INTEGER PRIMARY KEY AUTOINCREMENT,\n            phone TEXT NOT NULL,\n            normalized_phone TEXT NOT NULL UNIQUE,\n            flag_level TEXT NOT NULL,\n            note TEXT,\n            created_at TEXT NOT NULL,\n            updated_at TEXT NOT NULL\n        )')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_phone ON customer_bikes(phone)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_datetime ON customer_bikes(datetime)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_model ON customer_bikes(model)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_customer_flags_phone ON customer_flags(normalized_phone)')
        cursor.execute('SELECT COUNT(*) FROM settings')
        if cursor.fetchone()[0] == 0:
            cursor.execute('INSERT INTO settings (settings_json) VALUES (?)', (json.dumps(DEFAULT_SETTINGS),))
        conn.commit()
    except sqlite3.Error as e:
        messagebox.showerror('Database Error', f'Could not initialize database: {str(e)}')
    finally:
        if conn:
            release_db_connection(conn)

def load_settings() -> Dict:
    """Load settings from the database, falling back to defaults if unavailable."""
    conn = None
    try:
        if 'settings' in _app_cache and time.time() - _app_cache.get('settings_time', 0) < 300:
            settings = merge_local_settings(_app_cache['settings'])
            settings.pop('language', None)
            return settings
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT settings_json FROM settings WHERE id = 1')
        result = cursor.fetchone()
        if result:
            settings = json.loads(result[0])
            settings = merge_local_settings(settings)
            settings.pop('language', None)
            _app_cache['settings'] = settings
            _app_cache['settings_time'] = time.time()
            return settings
        settings = merge_local_settings(DEFAULT_SETTINGS.copy())
        settings.pop('language', None)
        return settings
    except sqlite3.Error as e:
        messagebox.showerror('Database Error', f'Could not load settings: {str(e)}')
        settings = merge_local_settings(DEFAULT_SETTINGS.copy())
        settings.pop('language', None)
        return settings
    finally:
        if conn:
            release_db_connection(conn)

def save_settings(settings: Dict) -> None:
    """Save settings to the database."""
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        shared_settings = strip_local_only_settings(settings)
        settings_json = json.dumps(shared_settings, ensure_ascii=False)
        cursor.execute('INSERT OR REPLACE INTO settings (id, settings_json) VALUES (1, ?)', (settings_json,))
        conn.commit()
        _app_cache['settings'] = merge_local_settings(shared_settings)
        _app_cache['settings_time'] = time.time()
    except sqlite3.Error as e:
        messagebox.showerror('Database Error', f'Could not save settings: {str(e)}')
    finally:
        if conn:
            release_db_connection(conn)

def get_saved_window_geometry(key: str, default_geometry: str) -> str:
    """Return saved geometry for a window, or the default geometry."""
    try:
        settings = load_settings()
        geometries = settings.get('window_geometries', {})
        geometry = geometries.get(key)
        if isinstance(geometry, str) and geometry:
            return geometry
    except Exception:
        pass
    return default_geometry

def get_safe_window_geometry(window, key: str, default_geometry: str, max_width_ratio: float=0.92, max_height_ratio: float=0.82) -> str:
    """Return a saved geometry capped to the visible screen so action buttons stay reachable."""
    raw_geometry = get_saved_window_geometry(key, default_geometry)
    match = re.match('^(\\d+)x(\\d+)([+-]\\d+)?([+-]\\d+)?$', str(raw_geometry).strip())
    if match:
        width = int(match.group(1))
        height = int(match.group(2))
    else:
        fallback = re.match('^(\\d+)x(\\d+)', default_geometry)
        width = int(fallback.group(1)) if fallback else 1080
        height = int(fallback.group(2)) if fallback else 680
    try:
        screen_width = max(800, int(window.winfo_screenwidth()))
        screen_height = max(600, int(window.winfo_screenheight()))
    except Exception:
        screen_width, screen_height = (1366, 768)
    max_width = max(760, int(screen_width * max_width_ratio))
    max_height = max(500, int(screen_height * max_height_ratio))
    width = max(760, min(width, max_width))
    height = max(500, min(height, max_height))
    x = max(0, (screen_width - width) // 2)
    y = max(0, min(40, (screen_height - height) // 3))
    return f'{width}x{height}+{x}+{y}'

def save_window_geometry(window, key: str) -> None:
    """Save current window geometry."""
    try:
        geometry = window.geometry()
        if not geometry:
            return
        settings = load_settings()
        geometries = settings.get('window_geometries', {})
        if not isinstance(geometries, dict):
            geometries = {}
        geometries[key] = geometry
        settings['window_geometries'] = geometries
        save_settings(settings)
    except Exception:
        pass

def _autofill_unique_lines_from_raw(raw_value: Any, defaults: Optional[List[str]], normalizer=None) -> List[str]:
    """Return clean unique lines, using defaults when the stored list is empty."""
    try:
        normalizer = normalizer or (lambda v: str(v or '').strip().casefold())
        if raw_value is None or not str(raw_value).strip():
            candidates = list(defaults or [])
        elif isinstance(raw_value, (list, tuple)):
            candidates = list(raw_value)
        else:
            candidates = str(raw_value).replace(';', '\n').splitlines()
        result: List[str] = []
        seen = set()
        for item in candidates:
            value = ' '.join(str(item or '').strip().split())
            key = normalizer(value)
            if value and key and (key not in seen):
                seen.add(key)
                result.append(value)
        return result
    except Exception:
        return list(defaults or [])

def _append_unique_autofill_items(existing_items: List[str], new_items: List[Any], normalizer=None, max_items: int=1000) -> Tuple[List[str], bool]:
    """Append clean new suggestions without duplicates."""
    normalizer = normalizer or (lambda v: str(v or '').strip().casefold())
    result = list(existing_items or [])
    seen = {normalizer(item) for item in result if normalizer(item)}
    changed = False
    for item in new_items or []:
        value = ' '.join(str(item or '').strip().split())
        key = normalizer(value)
        if value and key and (key not in seen):
            seen.add(key)
            result.append(value)
            changed = True
    return (result[:max_items], changed)

def _is_existing_bike_non_brand_value(value: Any) -> bool:
    """Avoid accidentally storing a known color or bike type as a bike brand."""
    try:
        key = _normalize_bike_brand_key(value)
        color_keys = {_normalize_bike_brand_key(v) for v in bike_color_autofill_values or DEFAULT_BIKE_COLOR_AUTOFILL_ITEMS}
        model_keys = {_normalize_bike_brand_key(v) for v in bike_model_autofill_values or DEFAULT_BIKE_MODEL_AUTOFILL_ITEMS}
        return bool(key and (key in color_keys or key in model_keys))
    except Exception:
        return False

def _extract_new_bike_brand_suggestion(model_value: Any) -> str:
    """Extract the brand part from the one-field bike descriptor after Save."""
    try:
        compact = _compact_bike_descriptor(model_value)
        if not compact:
            return ''
        known_brands = _known_values_for_matching(bike_brand_autofill_values, DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS)
        known_colors = _known_values_for_matching(bike_color_autofill_values, DEFAULT_BIKE_COLOR_AUTOFILL_ITEMS)
        brand = _extract_bike_brand_from_model(compact, known_brands)
        candidate = ' '.join(str(brand or '').strip().split())
        known_brand_keys = {_normalize_bike_brand_key(v) for v in known_brands}
        if _normalize_bike_brand_key(candidate) in known_brand_keys:
            return ''
        if _is_existing_bike_non_brand_value(candidate):
            return ''
        first, rest = _split_after_known_prefix(compact, known_colors)
        if first and (not rest):
            return ''
        if first and _normalize_bike_brand_key(first) == _normalize_bike_brand_key(candidate):
            return ''
        if 2 <= len(candidate) <= 40:
            return candidate
    except Exception:
        pass
    return ''

def update_autofill_settings_from_saved_record(model_value: Any, issue_values: Optional[List[Any]]=None) -> None:
    """Persist newly typed bike brand and issue suggestions into Settings after a successful Save."""
    try:
        settings = load_settings()
        changed = False
        if settings.get('issue_autofill_auto_add_new', True):
            issue_existing = _autofill_unique_lines_from_raw(settings.get('issue_autofill_custom_items', ''), DEFAULT_ISSUE_AUTOFILL_ITEMS, lambda v: str(v or '').strip().casefold())
            clean_issues = [' '.join(str(issue or '').strip().split()) for issue in issue_values or []]
            issue_existing, issue_changed = _append_unique_autofill_items(issue_existing, clean_issues, lambda v: str(v or '').strip().casefold(), ISSUE_AUTOFILL_MAX_DB_VALUES)
            if issue_changed:
                settings['issue_autofill_custom_items'] = '\n'.join(issue_existing)
                changed = True
        auto_add_new_brand = settings.get('bike_brand_autofill_auto_add_new', settings.get('bike_model_autofill_auto_add_new', True))
        if auto_add_new_brand:
            brand_existing = _autofill_unique_lines_from_raw(settings.get('bike_brand_autofill_custom_items', ''), DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS, _normalize_bike_brand_key)
            new_brand = _extract_new_bike_brand_suggestion(model_value)
            brand_existing, brand_changed = _append_unique_autofill_items(brand_existing, [new_brand], _normalize_bike_brand_key, BIKE_BRAND_AUTOFILL_MAX_DB_VALUES)
            if brand_changed:
                settings['bike_brand_autofill_custom_items'] = '\n'.join(brand_existing)
                changed = True
        if changed:
            save_settings(settings)
    except Exception:
        pass


def _main_form_values_from_entries() -> Tuple[str, str, List[str], str, List[str]]:
    """Read current main-form values in the same format used by normal Save."""
    raw_phone = phone_entry.get().strip() if phone_entry else ''
    phone = raw_phone.upper()
    raw_model = model_entry.get().strip() if model_entry else ''
    model = raw_model.upper()
    _sync_issue_entries_from_main_tree()
    raw_issues = [entry.get().strip() for entry in issue_entries if entry.get().strip()]
    issues = [issue.upper() for issue in raw_issues]
    issues.extend([''] * (10 - len(issues)))
    return phone, model, issues, raw_model, raw_issues

def _set_main_edit_button_state(enabled: bool=True) -> None:
    """Enable or disable main edit/save buttons safely."""
    state = tk.NORMAL if enabled else tk.DISABLED
    try:
        if save_button is not None:
            save_button.config(state=state)
    except Exception:
        pass
    try:
        if bike_list_button is not None:
            bike_list_button.config(state=state)
    except Exception:
        pass
    try:
        if edit_cancel_button is not None:
            edit_cancel_button.config(state=state)
    except Exception:
        pass

def enter_main_edit_mode(row: Dict[str, Any]) -> None:
    """Load one existing bike-list record into the main input window for editing without printing."""
    global main_edit_state
    text = ENGLISH_TEXT
    if not row:
        return
    try:
        main_edit_state = {
            'db_id': row.get('_db_id'),
            'datetime': row.get('datetime', ''),
            'phone': row.get('phone', ''),
            'model': row.get('model', ''),
            'issues': [row.get(f'issue{i}', '') or '' for i in range(1, 11)],
            'shop_note': row.get('shop_note', '') or ''
        }
        phone_entry.delete(0, tk.END)
        phone_entry.insert(0, str(row.get('phone', '') or ''))
        model_entry.delete(0, tk.END)
        model_entry.insert(0, str(row.get('model', '') or ''))
        for i, entry in enumerate(issue_entries[:10], start=1):
            entry.delete(0, tk.END)
            entry.insert(0, str(row.get(f'issue{i}', '') or ''))
        refresh_main_issue_table_from_entries()
        note_entry = getattr(root, '_shop_note_entry', None)
        if note_entry is not None:
            note_entry.delete('1.0', 'end')
            note_entry.insert('1.0', str(row.get('shop_note', '') or ''))
        if save_button is not None:
            save_button.config(text='Save Edit', command=update_customer_record_from_main)
        if edit_cancel_button is not None and not edit_cancel_button.winfo_ismapped():
            edit_cancel_button.pack(side='left', padx=4)
        try:
            root.title('BikeTrack Pro - Editing Record')
        except Exception:
            pass
        try:
            status_label.config(text='Editing old record. Save Edit saves without printing.', fg='blue')
        except Exception:
            pass
        _refresh_main_input_visual_state()
        try:
            root.lift()
            root.focus_force()
        except Exception:
            pass
        try:
            if phone_entry is not None:
                phone_entry.focus_set()
        except Exception:
            pass
    except Exception as e:
        messagebox.showerror(text['error_title'], str(e))

def exit_main_edit_mode(clear_status: bool=False) -> None:
    """Return the main input window to normal new-record save mode."""
    global main_edit_state
    main_edit_state = None
    try:
        if save_button is not None:
            save_button.config(text=ENGLISH_TEXT['save_button'], command=add_customer)
    except Exception:
        pass
    try:
        if edit_cancel_button is not None and edit_cancel_button.winfo_ismapped():
            edit_cancel_button.pack_forget()
    except Exception:
        pass
    try:
        root.title('BikeTrack Pro')
    except Exception:
        pass
    if clear_status:
        try:
            status_label.config(text='', fg='black')
        except Exception:
            pass
    _refresh_main_input_visual_state()

def cancel_main_edit_mode() -> None:
    """Cancel editing an old record and clear the edit form to avoid accidental duplicate saving."""
    exit_main_edit_mode(clear_status=True)
    try:
        clear_customer_form()
    except Exception:
        pass
    try:
        status_label.config(text='Edit cancelled.', fg='gray')
    except Exception:
        pass

def update_customer_record_from_main() -> None:
    """Update the selected old record from the main form without printing a receipt."""
    global main_edit_state
    text = ENGLISH_TEXT
    if not main_edit_state:
        add_customer()
        return
    phone, model, issues, raw_model, raw_issues = _main_form_values_from_entries()
    if not phone or not model:
        status_label.config(text=text['status_empty'], fg='red')
        return
    edit_state = dict(main_edit_state)
    shop_note = get_main_shop_note()
    _set_main_edit_button_state(False)
    status_label.config(text='Updating old record...', fg='blue')
    root.update()
    threading.Thread(
        target=lambda: threaded_update_customer_record(edit_state, phone, model, issues, raw_model, raw_issues, shop_note),
        daemon=True
    ).start()

def threaded_update_customer_record(edit_state: Dict[str, Any], phone: str, model: str, issues: List[str], raw_model: str, raw_issues: List[str], shop_note: str='') -> None:
    """Worker for saving an edited old record without printing."""
    if update_customer_record_in_database(edit_state, phone, model, issues, shop_note):
        update_autofill_settings_from_saved_record(raw_model or model, raw_issues or issues)
        refresh_issue_autofill_values()
        refresh_bike_brand_autofill_values()
        backup_database()
        root.after(0, lambda: exit_main_edit_mode(clear_status=False))
        root.after(0, lambda: clear_customer_form())
        if bike_list_window and tk.Toplevel.winfo_exists(bike_list_window) and bike_shop_app:
            root.after(0, lambda: bike_shop_app.register_edit_undo_from_main(edit_state, phone, model, issues, shop_note))
            root.after(0, lambda: bike_shop_app.refresh_treeview())
        root.after(0, lambda: status_label.config(text='Record updated. No receipt printed.', fg='green'))
    root.after(0, lambda: _set_main_edit_button_state(True))

def update_customer_record_in_database(edit_state: Dict[str, Any], phone: str, model: str, issues: List[str], shop_note: str='') -> bool:
    """Update one existing customer_bikes row. Never prints."""
    text = ENGLISH_TEXT
    deadline = time.monotonic() + DB_LOCK_MAX_WAIT_SECONDS
    lock_indicator_on = False
    last_lock_error = None
    set_values = [phone, model] + list(issues[:10]) + [shop_note]
    while True:
        conn = None
        try:
            conn = get_db_connection()
            conn.execute(f'PRAGMA busy_timeout={DB_LOCK_RETRY_BUSY_TIMEOUT_MS}')
            cursor = conn.cursor()
            db_id = edit_state.get('db_id')
            if db_id not in (None, ''):
                cursor.execute(
                    "UPDATE customer_bikes "
                    "SET phone=?, model=?, issue1=?, issue2=?, issue3=?, issue4=?, issue5=?, "
                    "issue6=?, issue7=?, issue8=?, issue9=?, issue10=?, shop_note=? "
                    "WHERE id=?",
                    tuple(set_values) + (db_id,)
                )
            else:
                original_issues = list(edit_state.get('issues') or [])
                original_issues.extend([''] * (10 - len(original_issues)))
                cursor.execute(
                    "UPDATE customer_bikes "
                    "SET phone=?, model=?, issue1=?, issue2=?, issue3=?, issue4=?, issue5=?, "
                    "issue6=?, issue7=?, issue8=?, issue9=?, issue10=?, shop_note=? "
                    "WHERE datetime=? AND phone=? AND model=? "
                    "AND COALESCE(issue1,'')=? AND COALESCE(issue2,'')=? AND COALESCE(issue3,'')=? "
                    "AND COALESCE(issue4,'')=? AND COALESCE(issue5,'')=? AND COALESCE(issue6,'')=? "
                    "AND COALESCE(issue7,'')=? AND COALESCE(issue8,'')=? AND COALESCE(issue9,'')=? "
                    "AND COALESCE(issue10,'')=?",
                    tuple(set_values) + (
                        edit_state.get('datetime', ''), edit_state.get('phone', ''), edit_state.get('model', ''),
                        *original_issues[:10]
                    )
                )
            if cursor.rowcount <= 0:
                raise sqlite3.Error('The original record was not found. Refresh Bike List and try again.')
            conn.commit()
            try:
                _app_cache.pop('bike_data', None)
                _app_cache.pop('bike_headers', None)
                _app_cache.pop('bike_data_time', None)
            except Exception:
                pass
            if lock_indicator_on:
                set_database_lock_indicator(False)
            return True
        except sqlite3.OperationalError as e:
            if is_database_locked_error(e):
                last_lock_error = e
                try:
                    if conn:
                        conn.rollback()
                except Exception:
                    pass
                if not lock_indicator_on:
                    set_database_lock_indicator(True)
                    lock_indicator_on = True
                if time.monotonic() < deadline:
                    time.sleep(DB_LOCK_RETRY_DELAY_SECONDS)
                    continue
                set_database_lock_indicator(True, auto_clear_ms=5000)
                try:
                    root.after(0, lambda: status_label.config(text='Database is still locked. Try again.', fg='red'))
                    root.after(0, lambda: messagebox.showerror(text['db_error_title'], text['db_error_message'].format(str(last_lock_error))))
                except Exception:
                    pass
                return False
            try:
                root.after(0, lambda: messagebox.showerror(text['db_error_title'], text['db_error_message'].format(str(e))))
            except Exception:
                pass
            return False
        except sqlite3.Error as e:
            try:
                root.after(0, lambda: messagebox.showerror(text['db_error_title'], text['db_error_message'].format(str(e))))
                root.after(0, lambda: status_label.config(text='Update failed.', fg='red'))
            except Exception:
                pass
            return False
        finally:
            if conn:
                try:
                    conn.execute(f'PRAGMA busy_timeout={DB_DEFAULT_BUSY_TIMEOUT_MS}')
                except Exception:
                    pass
                release_db_connection(conn)

def get_main_shop_note() -> str:
    """Read the optional private shop note from the main form."""
    note_entry = getattr(root, '_shop_note_entry', None)
    return note_entry.get('1.0', 'end-1c').strip() if note_entry is not None else ''


def add_customer() -> None:
    """Add a new customer to the database and print a receipt."""
    global bike_list_window, bike_shop_app
    text = ENGLISH_TEXT
    phone = phone_entry.get().strip().upper()
    raw_model = model_entry.get().strip()
    model = raw_model.upper()
    if not phone or not model:
        status_label.config(text=text['status_empty'], fg='red')
        return
    save_button.config(state=tk.DISABLED)
    bike_list_button.config(state=tk.DISABLED)
    status_label.config(text='Processing...', fg='blue')
    root.update()
    current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    _sync_issue_entries_from_main_tree()
    raw_issues = [entry.get().strip() for entry in issue_entries if entry.get().strip()]
    issues = [issue.upper() for issue in raw_issues]
    issues.extend([''] * (10 - len(issues)))
    shop_note = get_main_shop_note()
    threading.Thread(target=lambda: threaded_add_customer(current_datetime, phone, model, issues, raw_model, raw_issues, shop_note), daemon=True).start()

def threaded_add_customer(current_datetime, phone, model, issues, raw_model: Optional[str]=None, raw_issues: Optional[List[str]]=None, shop_note: str=''):
    """Process customer data in a separate thread."""
    text = ENGLISH_TEXT
    global bike_list_window, bike_shop_app
    if save_to_database(current_datetime, phone, model, issues, shop_note):
        update_autofill_settings_from_saved_record(raw_model or model, raw_issues or issues)
        refresh_issue_autofill_values()
        refresh_bike_brand_autofill_values()
        backup_database()
        root.after(0, lambda: clear_customer_form())
        print_receipt([current_datetime, phone, model] + issues[:10] + [shop_note])
        if bike_list_window and tk.Toplevel.winfo_exists(bike_list_window) and bike_shop_app:
            root.after(0, lambda: bike_shop_app.refresh_treeview())
        root.after(0, lambda: status_label.config(text=text['status_saved'], fg='green'))
    root.after(0, lambda: save_button.config(state=tk.NORMAL))
    root.after(0, lambda: bike_list_button.config(state=tk.NORMAL))

def clear_customer_form():
    """Clear all form fields after saving."""
    hide_bike_brand_autofill()
    phone_entry.delete(0, tk.END)
    model_entry.delete(0, tk.END)
    clear_main_issue_table()
    note_entry = getattr(root, '_shop_note_entry', None)
    if note_entry is not None:
        note_entry.delete('1.0', 'end')
    try:
        reload_main_visit_history('')
    except Exception:
        pass
    phone_entry.focus_set()

def save_to_database(datetime: str, phone: str, model: str, issues: List[str], shop_note: str='') -> bool:
    """Save customer data to the database."""
    text = ENGLISH_TEXT
    deadline = time.monotonic() + DB_LOCK_MAX_WAIT_SECONDS
    lock_indicator_on = False
    last_lock_error = None
    while True:
        conn = None
        try:
            conn = get_db_connection()
            conn.execute(f'PRAGMA busy_timeout={DB_LOCK_RETRY_BUSY_TIMEOUT_MS}')
            cursor = conn.cursor()
            cursor.execute('INSERT INTO customer_bikes (\n                datetime, phone, model, issue1, issue2, issue3, issue4, issue5,\n                issue6, issue7, issue8, issue9, issue10, shop_note\n            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', (datetime, phone, model) + tuple(issues[:10]) + (shop_note,))
            conn.commit()
            try:
                _app_cache.pop('bike_data', None)
                _app_cache.pop('bike_headers', None)
                _app_cache.pop('bike_data_time', None)
            except Exception:
                pass
            if lock_indicator_on:
                set_database_lock_indicator(False)
            return True
        except sqlite3.OperationalError as e:
            if is_database_locked_error(e):
                last_lock_error = e
                try:
                    if conn:
                        conn.rollback()
                except Exception:
                    pass
                if not lock_indicator_on:
                    set_database_lock_indicator(True)
                    lock_indicator_on = True
                if time.monotonic() < deadline:
                    time.sleep(DB_LOCK_RETRY_DELAY_SECONDS)
                    continue
                set_database_lock_indicator(True, auto_clear_ms=5000)
                try:
                    if status_label is not None and root is not None:
                        message = 'Database is still locked. Try again.'
                        root.after(0, lambda: status_label.config(text=message, fg='red'))
                except Exception:
                    pass
                try:
                    if root is not None:
                        root.after(0, lambda: messagebox.showerror(text['db_error_title'], text['db_error_message'].format(str(last_lock_error))))
                    else:
                        messagebox.showerror(text['db_error_title'], text['db_error_message'].format(str(last_lock_error)))
                except Exception:
                    pass
                return False
            try:
                if root is not None:
                    root.after(0, lambda: messagebox.showerror(text['db_error_title'], text['db_error_message'].format(str(e))))
                else:
                    messagebox.showerror(text['db_error_title'], text['db_error_message'].format(str(e)))
            except Exception:
                pass
            return False
        except sqlite3.Error as e:
            try:
                if root is not None:
                    root.after(0, lambda: messagebox.showerror(text['db_error_title'], text['db_error_message'].format(str(e))))
                else:
                    messagebox.showerror(text['db_error_title'], text['db_error_message'].format(str(e)))
            except Exception:
                pass
            return False
        finally:
            if conn:
                try:
                    conn.execute(f'PRAGMA busy_timeout={DB_DEFAULT_BUSY_TIMEOUT_MS}')
                except Exception:
                    pass
                release_db_connection(conn)

def backup_database() -> None:
    """Backup the database to the configured folder."""
    _thread_pool.submit(threaded_backup_database)

def threaded_backup_database() -> None:
    """Perform database backup in a background thread."""
    text = ENGLISH_TEXT
    settings = load_settings()
    backup_folder = resolve_backup_folder_path(settings.get('backup_folder', 'backup_db'))
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(backup_folder, f'customer_bikes_backup_{timestamp}.db')
    try:
        conn = get_db_connection()
        conn.execute('PRAGMA wal_checkpoint(FULL)')
        release_db_connection(conn)
        shutil.copy(DATABASE_FILE, backup_path)
    except Exception as e:
        root.after(0, lambda: messagebox.showerror(text['backup_error_title'], text['backup_error_message'].format(str(e))))

def generate_qr_code(qr_content: str, target_width: int) -> Image.Image:
    """Generate a QR code image with the specified content and width."""
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_H, border=4)
    qr.add_data(qr_content)
    qr.make(fit=True)
    modules = qr.modules_count
    box_size = max(1, target_width // modules)
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=box_size, border=4)
    qr.add_data(qr_content)
    qr.make(fit=True)
    return qr.make_image(fill='black', back_color='white')

def image_to_escpos(image: Image.Image, target_width: int, height_percent: int=100) -> bytes:
    """Convert an image to ESC/POS commands for printing."""
    im = image.convert('1')
    if im.width != target_width:
        ratio = target_width / im.width
        new_height = int(im.height * ratio)
        im = im.resize((target_width, new_height), Image.LANCZOS)
    height_percent = _clamp_int(height_percent, 100, 50, 200)
    if height_percent != 100:
        corrected_height = max(1, int(round(im.height * height_percent / 100.0)))
        im = im.resize((target_width, corrected_height), Image.LANCZOS)
    width, height = im.size
    width_bytes = (width + 7) // 8
    buffer = io.BytesIO()
    buffer.write(b'\x1b@')
    buffer.write(ALIGN_CENTER)
    buffer.write(b'\x1dv0\x00')
    buffer.write(width_bytes.to_bytes(2, 'little'))
    buffer.write(height.to_bytes(2, 'little'))
    for y in range(height):
        for x in range(0, width, 8):
            byte = 0
            for bit in range(8):
                if x + bit < width and im.getpixel((x + bit, y)) == 0:
                    byte |= 1 << 7 - bit
            buffer.write(bytes([byte]))
    buffer.write(b'\n')
    buffer.write(ALIGN_LEFT)
    return buffer.getvalue()

def receipt_logo_to_escpos(logo_settings: Dict) -> bytes:
    """Convert the selected customer receipt logo to ESC/POS commands."""
    try:
        if not logo_settings.get('enabled', False):
            return b''
        logo_path = str(logo_settings.get('path', '')).strip()
        if not logo_path or not os.path.exists(logo_path):
            return b''
        target_width = int(logo_settings.get('width', 384))
        target_width = max(64, min(target_width, 576))
        target_height_percent = _clamp_int(logo_settings.get('height_percent', logo_settings.get('aspect', 100)), 100, 50, 200)
        with Image.open(logo_path) as img:
            img = img.convert('RGBA')
            background = Image.new('RGBA', img.size, 'WHITE')
            background.alpha_composite(img)
            return image_to_escpos(background.convert('RGB'), target_width, target_height_percent) + b'\n'
    except Exception:
        return b''
GDI_RECEIPT_WIDTH = RECEIPT_80MM_DEFAULT_WIDTH_PX
GDI_RECEIPT_MARGIN = 8
GDI_LINE_SPACING = 3
GDI_BOTTOM_FEED = 28

def _clamp_int(value: Any, default: int, minimum: int, maximum: int) -> int:
    try:
        number = int(float(str(value).strip()))
    except Exception:
        number = default
    return max(minimum, min(maximum, number))

def _clean_receipt_single_line_text(value: Any, default: str=DEFAULT_THANK_YOU_TEXT) -> str:
    """Return a safe single-line text value for receipt designer fields."""
    try:
        cleaned = str(value if value is not None else '').replace('\r', ' ').replace('\n', ' ').strip()
    except Exception:
        cleaned = ''
    return cleaned or default

def get_receipt_designer_settings(settings: Optional[Dict]=None) -> Dict:
    """Return receipt designer settings with safe defaults for older databases."""
    base = DEFAULT_RECEIPT_DESIGNER.copy()
    try:
        incoming = (settings or {}).get('receipt_designer', {})
        if isinstance(incoming, dict):
            base.update(incoming)
    except Exception:
        pass
    try:
        qr_width_settings = (settings or {}).get('qr_width_settings', {})
        if isinstance(qr_width_settings, dict):
            if not base.get('customer_qr_width'):
                base['customer_qr_width'] = qr_width_settings.get('customer_qr_width', 220)
            if not base.get('shop_qr_width'):
                base['shop_qr_width'] = qr_width_settings.get('shop_qr_width', 220)
    except Exception:
        pass
    base['paper_width'] = _clamp_int(base.get('paper_width'), RECEIPT_80MM_DEFAULT_WIDTH_PX, 240, 832)
    base['margin'] = _clamp_int(base.get('margin'), 8, 0, 80)
    base['line_spacing'] = _clamp_int(base.get('line_spacing'), 3, 0, 18)
    base['bottom_feed'] = _clamp_int(base.get('bottom_feed'), 28, 0, 160)
    base['customer_top_space'] = _clamp_int(base.get('customer_top_space'), 0, 0, 240)
    base['shop_top_space'] = _clamp_int(base.get('shop_top_space'), 0, 0, 240)
    base['customer_bottom_feed'] = _clamp_int(base.get('customer_bottom_feed'), base['bottom_feed'], 0, 240)
    base['shop_bottom_feed'] = _clamp_int(base.get('shop_bottom_feed'), base['bottom_feed'], 0, 240)
    base['title_size'] = _clamp_int(base.get('title_size'), 27, 10, 72)
    base['body_size'] = _clamp_int(base.get('body_size'), 21, 8, 60)
    base['small_size'] = _clamp_int(base.get('small_size'), 16, 7, 48)
    base['phone_size'] = _clamp_int(base.get('phone_size'), 23, 8, 72)
    base['bike_size'] = _clamp_int(base.get('bike_size'), 23, 8, 72)
    base['datetime_size'] = _clamp_int(base.get('datetime_size'), 19, 7, 60)
    base['shop_info_size'] = _clamp_int(base.get('shop_info_size'), 18, 7, 60)
    base['issue_size'] = _clamp_int(base.get('issue_size'), 22, 8, 72)
    base['logo_width'] = _clamp_int(base.get('logo_width'), 260, 40, max(80, base['paper_width'] - 2 * base['margin']))
    # Logo aspect correction matches the QR controls: 100 keeps the original logo ratio.
    # Lower/higher values change logo height only while paper width stays unchanged.
    base['logo_aspect'] = _clamp_int(base.get('logo_aspect'), 100, 50, 200)
    try:
        incoming = (settings or {}).get('receipt_designer', {})
        if isinstance(incoming, dict):
            for prefix in ('customer', 'shop'):
                old_dt = bool(base.get(f'{prefix}_show_datetime', True))
                if f'{prefix}_show_date' not in incoming:
                    base[f'{prefix}_show_date'] = old_dt
                if f'{prefix}_show_time' not in incoming:
                    base[f'{prefix}_show_time'] = old_dt
    except Exception:
        pass
    base['customer_qr_width'] = _clamp_int(base.get('customer_qr_width'), 220, 80, 832)
    base['shop_qr_width'] = _clamp_int(base.get('shop_qr_width'), 220, 80, 832)
    # QR aspect correction is a pre-distortion control for printers/paper that squeeze the bitmap.
    # 100 means normal square QR. Lower/higher values change QR height only while paper width stays unchanged.
    base['customer_qr_aspect'] = _clamp_int(base.get('customer_qr_aspect'), 100, 50, 200)
    base['shop_qr_aspect'] = _clamp_int(base.get('shop_qr_aspect'), 100, 50, 200)
    font_family = str(base.get('font_family') or 'Segoe UI').strip()
    base['font_family'] = font_family or 'Segoe UI'
    base['text_align'] = _normalize_receipt_align(base.get('text_align', DEFAULT_RECEIPT_DESIGNER.get('text_align', 'left')), 'left')
    for _align_key in ('logo_align', 'title_align', 'phone_align', 'bike_align', 'datetime_align', 'shop_info_align', 'issue_align', 'qr_align', 'thanks_align'):
        _align_default = DEFAULT_RECEIPT_DESIGNER.get(_align_key, base['text_align'])
        base[_align_key] = _normalize_receipt_align(base.get(_align_key, _align_default), _align_default)
    base['customer_title_text'] = _clean_receipt_single_line_text(base.get('customer_title_text'), DEFAULT_CUSTOMER_TITLE_TEXT)
    base['shop_title_text'] = _clean_receipt_single_line_text(base.get('shop_title_text'), DEFAULT_SHOP_TITLE_TEXT)
    base['thank_you_text'] = _clean_receipt_single_line_text(base.get('thank_you_text'), DEFAULT_THANK_YOU_TEXT)
    if not str(base.get('customer_template') or '').strip():
        base['customer_template'] = DEFAULT_CUSTOMER_RECEIPT_TEMPLATE
    if not str(base.get('shop_template') or '').strip():
        base['shop_template'] = DEFAULT_SHOP_RECEIPT_TEMPLATE
    bool_keys = ('title_bold', 'body_bold', 'phone_bold', 'bike_bold', 'datetime_bold', 'shop_info_bold', 'issue_bold', 'compact_mode', 'center_qr', 'simple_designer', 'customer_show_logo', 'customer_show_shop_info', 'customer_show_qr', 'customer_show_title', 'customer_show_phone', 'customer_show_phone_label', 'customer_show_bike', 'customer_show_bike_label', 'customer_show_datetime', 'customer_show_date', 'customer_show_date_label', 'customer_show_time', 'customer_show_time_label', 'customer_show_issues', 'customer_show_issues_label', 'customer_show_thanks', 'shop_show_logo', 'shop_show_shop_info', 'shop_show_qr', 'shop_show_phone', 'shop_show_phone_label', 'shop_show_bike', 'shop_show_bike_label', 'shop_show_datetime', 'shop_show_date', 'shop_show_date_label', 'shop_show_time', 'shop_show_time_label', 'shop_show_issues', 'shop_show_issues_label', 'shop_show_title')
    for key in bool_keys:
        base[key] = bool(base.get(key, DEFAULT_RECEIPT_DESIGNER.get(key, False)))
    try:
        legacy_logo_path = str(((settings or {}).get('customer_logo') or {}).get('path', '') or '').strip()
    except Exception:
        legacy_logo_path = ''
    for prefix in ('customer', 'shop'):
        base[f'{prefix}_paper_width'] = _clamp_int(base.get(f'{prefix}_paper_width', base.get('paper_width')), base['paper_width'], 240, 832)
        base[f'{prefix}_margin'] = _clamp_int(base.get(f'{prefix}_margin', base.get('margin')), base['margin'], 0, 80)
        base[f'{prefix}_line_spacing'] = _clamp_int(base.get(f'{prefix}_line_spacing', base.get('line_spacing')), base['line_spacing'], 0, 18)
        base[f'{prefix}_font_family'] = str(base.get(f'{prefix}_font_family', base.get('font_family', 'Segoe UI')) or 'Segoe UI').strip() or 'Segoe UI'
        for _align_key in ('text_align', 'logo_align', 'title_align', 'phone_align', 'bike_align', 'datetime_align', 'shop_info_align', 'issue_align', 'qr_align', 'thanks_align'):
            _fallback = base.get(_align_key, 'left')
            base[f'{prefix}_{_align_key}'] = _normalize_receipt_align(base.get(f'{prefix}_{_align_key}', _fallback), _fallback)
        base[f'{prefix}_title_size'] = _clamp_int(base.get(f'{prefix}_title_size', base.get('title_size')), base['title_size'], 10, 72)
        base[f'{prefix}_body_size'] = _clamp_int(base.get(f'{prefix}_body_size', base.get('body_size')), base['body_size'], 8, 60)
        base[f'{prefix}_small_size'] = _clamp_int(base.get(f'{prefix}_small_size', base.get('small_size')), base['small_size'], 7, 48)
        base[f'{prefix}_phone_size'] = _clamp_int(base.get(f'{prefix}_phone_size', base.get('phone_size')), base['phone_size'], 8, 72)
        base[f'{prefix}_bike_size'] = _clamp_int(base.get(f'{prefix}_bike_size', base.get('bike_size')), base['bike_size'], 8, 72)
        base[f'{prefix}_datetime_size'] = _clamp_int(base.get(f'{prefix}_datetime_size', base.get('datetime_size')), base['datetime_size'], 7, 60)
        base[f'{prefix}_shop_info_size'] = _clamp_int(base.get(f'{prefix}_shop_info_size', base.get('shop_info_size')), base['shop_info_size'], 7, 60)
        base[f'{prefix}_issue_size'] = _clamp_int(base.get(f'{prefix}_issue_size', base.get('issue_size')), base['issue_size'], 8, 72)
        max_kind_content = max(90, base[f'{prefix}_paper_width'] - 2 * base[f'{prefix}_margin'])
        base[f'{prefix}_logo_width'] = _clamp_int(base.get(f'{prefix}_logo_width', base.get('logo_width')), base['logo_width'], 40, max_kind_content)
        base[f'{prefix}_logo_aspect'] = _clamp_int(base.get(f'{prefix}_logo_aspect', base.get('logo_aspect')), base['logo_aspect'], 50, 200)
        base[f'{prefix}_logo_after_margin'] = _clamp_int(base.get(f'{prefix}_logo_after_margin', base.get('logo_after_margin', 6)), base.get('logo_after_margin', 6), 0, 120)
        base[f'{prefix}_qr_width'] = _clamp_int(base.get(f'{prefix}_qr_width', base.get(f'{prefix}_qr_width', base.get('customer_qr_width' if prefix == 'customer' else 'shop_qr_width', 220))), 220, 80, max_kind_content)
        base[f'{prefix}_qr_aspect'] = _clamp_int(base.get(f'{prefix}_qr_aspect', base.get('customer_qr_aspect' if prefix == 'customer' else 'shop_qr_aspect', 100)), 100, 50, 200)
        for _bold_key in ('title_bold', 'body_bold', 'phone_bold', 'bike_bold', 'datetime_bold', 'shop_info_bold', 'issue_bold', 'compact_mode'):
            base[f'{prefix}_{_bold_key}'] = bool(base.get(f'{prefix}_{_bold_key}', base.get(_bold_key, DEFAULT_RECEIPT_DESIGNER.get(_bold_key, False))))
        _title_default = DEFAULT_CUSTOMER_TITLE_TEXT if prefix == 'customer' else DEFAULT_SHOP_TITLE_TEXT
        base[f'{prefix}_title_text'] = _clean_receipt_single_line_text(base.get(f'{prefix}_title_text', base.get(f'{prefix}_title_text', _title_default)), _title_default)
        base[f'{prefix}_thank_you_text'] = _clean_receipt_single_line_text(base.get(f'{prefix}_thank_you_text', base.get('thank_you_text', DEFAULT_THANK_YOU_TEXT)), DEFAULT_THANK_YOU_TEXT)
        base[f'{prefix}_shop_info_text'] = str(base.get(f'{prefix}_shop_info_text', (settings or {}).get('shop_info', '')) or '').strip()
        base[f'{prefix}_logo_path'] = str(base.get(f'{prefix}_logo_path', base.get('logo_path', legacy_logo_path)) or '').strip()
    base['shop_show_note'] = bool(base.get('shop_show_note', True))
    base['shop_show_note_label'] = bool(base.get('shop_show_note_label', True))
    base['shop_note_label_text'] = str(base.get('shop_note_label_text', 'Shop Note') or '').strip()
    base['shop_note_size'] = _clamp_int(base.get('shop_note_size'), 21, 7, 72)
    base['shop_note_bold'] = bool(base.get('shop_note_bold', False))
    base['shop_note_align'] = _normalize_receipt_align(base.get('shop_note_align', 'left'), 'left')
    return base

def _gdi_font_path(bold: bool=False, family: str='Segoe UI') -> Optional[str]:
    """Return a Windows font path when available, otherwise None for PIL's fallback font."""
    family_key = str(family or 'Segoe UI').strip().lower()
    candidates = []
    if os.name == 'nt':
        windir = os.environ.get('WINDIR', 'C:\\Windows')
        fonts_dir = os.path.join(windir, 'Fonts')
        preferred = PRINT_FONT_FILE_MAP.get(family_key) or PRINT_FONT_FILE_MAP['segoe ui']
        candidates.append(os.path.join(fonts_dir, preferred[0] if bold else preferred[1]))
        for font_name in PRINT_FONT_CHOICES:
            font_pair = PRINT_FONT_FILE_MAP.get(font_name.lower())
            if font_pair:
                candidates.append(os.path.join(fonts_dir, font_pair[0] if bold else font_pair[1]))
        arial_pair = PRINT_FONT_FILE_MAP['arial']
        candidates.append(os.path.join(fonts_dir, arial_pair[0] if bold else arial_pair[1]))
    candidates.extend(['/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf' if bold else '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', '/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf' if bold else '/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf'])
    for path in candidates:
        if path and os.path.exists(path):
            return path
    return None

def _gdi_font(size: int, bold: bool=False, family: str='Segoe UI') -> ImageFont.ImageFont:
    """Create a Unicode-capable font for receipt image rendering."""
    path = _gdi_font_path(bold, family)
    try:
        if path:
            return ImageFont.truetype(path, size=size)
    except Exception:
        pass
    return ImageFont.load_default()

def _gdi_font_from_setting(setting_name: str) -> ImageFont.ImageFont:
    """Map old ESC/POS font-size setting names to practical bitmap font sizes for backward compatibility."""
    mapping = {'Normal': (26, False), 'Double Height': (36, False), 'Double Width': (30, True), 'Double Width and Height': (42, True)}
    size, bold = mapping.get(setting_name or 'Normal', mapping['Normal'])
    return _gdi_font(size, bold)

def _text_bbox(draw: ImageDraw.ImageDraw, text_value: str, font: ImageFont.ImageFont) -> Tuple[int, int, int, int]:
    try:
        return draw.textbbox((0, 0), text_value, font=font)
    except Exception:
        width, height = draw.textsize(text_value, font=font)
        return (0, 0, width, height)

def _text_width(draw: ImageDraw.ImageDraw, text_value: str, font: ImageFont.ImageFont) -> int:
    bbox = _text_bbox(draw, text_value, font)
    return max(0, bbox[2] - bbox[0])

def _line_height(draw: ImageDraw.ImageDraw, font: ImageFont.ImageFont, line_spacing: Optional[int]=None) -> int:
    bbox = _text_bbox(draw, 'AgÅÄÖgjpqy', font)
    spacing = GDI_LINE_SPACING if line_spacing is None else int(line_spacing)
    return max(14, bbox[3] - bbox[1] + max(0, spacing))

def _wrap_text(draw: ImageDraw.ImageDraw, text_value: Any, font: ImageFont.ImageFont, max_width: int) -> List[str]:
    """Word-wrap text for a receipt width without breaking extended Latin characters."""
    value = str(text_value if text_value is not None else '').replace('\r\n', '\n').replace('\r', '\n')
    output: List[str] = []
    for raw_line in value.split('\n'):
        line = raw_line.strip()
        if not line:
            output.append('')
            continue
        words = line.split()
        current = ''
        for word in words:
            candidate = word if not current else f'{current} {word}'
            if _text_width(draw, candidate, font) <= max_width:
                current = candidate
            elif current:
                output.append(current)
                current = word
            else:
                chunk = ''
                for ch in word:
                    candidate_chunk = chunk + ch
                    if _text_width(draw, candidate_chunk, font) <= max_width:
                        chunk = candidate_chunk
                    else:
                        if chunk:
                            output.append(chunk)
                        chunk = ch
                current = chunk
        if current:
            output.append(current)
    return output

def _draw_wrapped_text(draw: ImageDraw.ImageDraw, y: int, text_value: Any, font: ImageFont.ImageFont, max_width: int, x: int=GDI_RECEIPT_MARGIN, align: str='left', extra_after: int=0, receipt_width: Optional[int]=None, line_spacing: Optional[int]=None) -> int:
    width = receipt_width or GDI_RECEIPT_WIDTH
    for line in _wrap_text(draw, text_value, font, max_width):
        line_height = _line_height(draw, font, line_spacing)
        if line:
            if align == 'center':
                line_width = _text_width(draw, line, font)
                draw_x = max(0, (width - line_width) // 2)
            elif align == 'right':
                line_width = _text_width(draw, line, font)
                draw_x = max(0, x + max_width - line_width)
            else:
                draw_x = x
            draw.text((draw_x, y), line, font=font, fill=0)
        y += line_height
    return y + extra_after

def _paste_aligned(base: Image.Image, image: Image.Image, y: int, max_width: int, align: str='center', x: int=GDI_RECEIPT_MARGIN, extra_after: int=10) -> int:
    """Paste logo/QR according to the selected receipt alignment."""
    if image.mode != 'RGB':
        background = Image.new('RGB', image.size, 'WHITE')
        if image.mode == 'RGBA':
            background.paste(image, mask=image.split()[-1])
        else:
            background.paste(image.convert('RGB'))
        image = background
    if image.width > max_width:
        ratio = max_width / float(image.width)
        image = image.resize((max_width, max(1, int(image.height * ratio))), Image.Resampling.LANCZOS)
    align = _normalize_receipt_align(align, 'center')
    if align == 'left':
        paste_x = max(0, x)
    elif align == 'right':
        paste_x = max(0, x + max_width - image.width)
    else:
        paste_x = max(0, (base.width - image.width) // 2)
    base.paste(image, (paste_x, y))
    return y + image.height + extra_after

def _paste_centered(base: Image.Image, image: Image.Image, y: int, max_width: int, extra_after: int=10) -> int:
    return _paste_aligned(base, image, y, max_width, align='center', extra_after=extra_after)

def _load_receipt_logo_image(logo_settings: Dict, max_width: Optional[int]=None) -> Optional[Image.Image]:
    try:
        logo_path = str(logo_settings.get('path', '')).strip()
        if not logo_path or not os.path.exists(logo_path):
            return None
        target_width = int(logo_settings.get('width', 260))
        logo_height_percent = _clamp_int(logo_settings.get('height_percent', logo_settings.get('aspect', 100)), 100, 50, 200)
        if max_width is None:
            max_width = GDI_RECEIPT_WIDTH - 2 * GDI_RECEIPT_MARGIN
        target_width = max(40, min(target_width, max_width))
        with Image.open(logo_path) as img:
            img = img.convert('RGBA')
            background = Image.new('RGBA', img.size, 'WHITE')
            background.alpha_composite(img)
            img = background.convert('RGB')
            if img.width != target_width:
                ratio = target_width / float(img.width)
                img = img.resize((target_width, max(1, int(img.height * ratio))), Image.Resampling.LANCZOS)
            if logo_height_percent != 100:
                corrected_height = max(1, int(round(img.height * logo_height_percent / 100.0)))
                img = img.resize((target_width, corrected_height), Image.Resampling.LANCZOS)
            return img
    except Exception:
        return None

def _qr_image_for_gdi(content: str, target_width: int, max_allowed_width: Optional[int]=None, qr_height_percent: int=100) -> Optional[Image.Image]:
    try:
        if not str(content or '').strip():
            return None
        if max_allowed_width is None:
            max_allowed_width = GDI_RECEIPT_WIDTH - 2 * GDI_RECEIPT_MARGIN
        target_width = max(80, min(int(target_width or 220), max_allowed_width))
        qr_height_percent = _clamp_int(qr_height_percent, 100, 50, 200)
        target_height = max(40, int(round(target_width * qr_height_percent / 100.0)))
        qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=8, border=4)
        qr.add_data(content)
        qr.make(fit=True)
        img = qr.make_image(fill_color='black', back_color='white').convert('RGB')
        # Deliberately allow non-square bitmap output here. This compensates for
        # thermal printer/paper combinations that physically squeeze the QR on print.
        img = img.resize((target_width, target_height), Image.Resampling.NEAREST)
        return img
    except Exception:
        return None

def _split_datetime_for_receipt(value: Any) -> Tuple[str, str, str]:
    text_value = str(value or '').strip()
    if not text_value:
        now = datetime.now()
        return (now.strftime('%Y-%m-%d'), now.strftime('%H:%M'), now.strftime('%Y-%m-%d %H:%M'))
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%dT%H:%M'):
        try:
            dt = datetime.strptime(text_value[:19], fmt)
            return (dt.strftime('%Y-%m-%d'), dt.strftime('%H:%M'), dt.strftime('%Y-%m-%d %H:%M'))
        except Exception:
            pass
    parts = text_value.replace('T', ' ').split()
    date_part = parts[0] if parts else text_value
    time_part = parts[1][:5] if len(parts) > 1 else ''
    return (date_part, time_part, text_value)

def _receipt_template_context(data: List[str], settings: Dict, receipt_kind: str) -> Dict[str, Any]:
    date_part, time_part, datetime_part = _split_datetime_for_receipt(data[0] if data else '')
    phone = data[1] if len(data) > 1 else ''
    bike = data[2] if len(data) > 2 else ''
    issues = [str(issue).strip() for issue in data[3:13] if str(issue or '').strip()]
    shop_note = str(data[13] or '').strip() if receipt_kind == 'shop' and len(data) > 13 else ''
    design = _receipt_design_for_kind(get_receipt_designer_settings(settings), receipt_kind, settings)
    thanks_text = _clean_receipt_single_line_text(design.get('thank_you_text'), DEFAULT_THANK_YOU_TEXT)
    customer_title_text = _clean_receipt_single_line_text(design.get('customer_title_text'), DEFAULT_CUSTOMER_TITLE_TEXT)
    shop_title_text = _clean_receipt_single_line_text(design.get('shop_title_text'), DEFAULT_SHOP_TITLE_TEXT)
    shop_info_text = str(design.get('shop_info_text', settings.get('shop_info', '')) or '').strip()
    return {'phone': phone, 'bike': bike, 'date': date_part, 'time': time_part, 'datetime': datetime_part, 'shop_info': shop_info_text, 'issues': issues, 'shop_note': shop_note, 'receipt_type': 'Kundkvitto' if receipt_kind == 'customer' else 'Butikskvitto', 'thanks': thanks_text, 'customer_title': customer_title_text, 'shop_title': shop_title_text}

def _replace_receipt_placeholders(line: str, context: Dict[str, Any]) -> str:
    replacements = {'{phone}': str(context.get('phone', '')), '{bike}': str(context.get('bike', '')), '{date}': str(context.get('date', '')), '{time}': str(context.get('time', '')), '{datetime}': str(context.get('datetime', '')), '{receipt_type}': str(context.get('receipt_type', '')), '{customer_title}': str(context.get('customer_title', DEFAULT_CUSTOMER_TITLE_TEXT)), '{shop_title}': str(context.get('shop_title', DEFAULT_SHOP_TITLE_TEXT)), '{thanks}': str(context.get('thanks', DEFAULT_THANK_YOU_TEXT))}
    output = line
    for key, value in replacements.items():
        output = output.replace(key, value)
    return output

def _parse_receipt_line_markup(raw_line: str, default_align: str='left') -> Tuple[str, str, Dict[str, bool]]:
    """Parse light template markup like [center][bold][small]Text."""
    line = str(raw_line or '')
    align = _normalize_receipt_align(default_align, 'left')
    flags = {'bold': False, 'small': False, 'title': False}
    changed = True
    while changed:
        changed = False
        stripped = line.lstrip()
        prefix_spaces = line[:len(line) - len(stripped)]
        lowered = stripped.lower()
        for tag in ('[center]', '[right]', '[left]', '[bold]', '[small]', '[title]'):
            if lowered.startswith(tag):
                if tag == '[center]':
                    align = 'center'
                elif tag == '[right]':
                    align = 'right'
                elif tag == '[left]':
                    align = 'left'
                elif tag == '[bold]':
                    flags['bold'] = True
                elif tag == '[small]':
                    flags['small'] = True
                elif tag == '[title]':
                    flags['title'] = True
                    flags['bold'] = True
                line = prefix_spaces + stripped[len(tag):]
                changed = True
                break
    stripped_line = line.strip()
    if stripped_line.startswith('**') and stripped_line.endswith('**') and (len(stripped_line) >= 4):
        flags['bold'] = True
        line = stripped_line[2:-2]
    return (line, align, flags)

def _receipt_fonts_from_design(design: Dict) -> Dict[str, ImageFont.ImageFont]:
    family = design.get('font_family', 'Segoe UI')
    return {'title': _gdi_font(design['title_size'], bool(design.get('title_bold', True)), family), 'body': _gdi_font(design['body_size'], bool(design.get('body_bold', False)), family), 'body_bold': _gdi_font(design['body_size'], True, family), 'small': _gdi_font(design['small_size'], False, family), 'small_bold': _gdi_font(design['small_size'], True, family), 'phone': _gdi_font(design['phone_size'], bool(design.get('phone_bold', True)), family), 'bike': _gdi_font(design['bike_size'], bool(design.get('bike_bold', True)), family), 'datetime': _gdi_font(design['datetime_size'], bool(design.get('datetime_bold', False)), family), 'shop_info': _gdi_font(design['shop_info_size'], bool(design.get('shop_info_bold', False)), family), 'issue': _gdi_font(design['issue_size'], bool(design.get('issue_bold', False)), family), 'issue_bold': _gdi_font(design['issue_size'], True, family), 'note': _gdi_font(_clamp_int(design.get('shop_note_size'), 21, 7, 72), bool(design.get('shop_note_bold', False)), family)}

def _choose_receipt_font(raw_line: str, flags: Dict[str, bool], fonts: Dict[str, ImageFont.ImageFont]) -> ImageFont.ImageFont:
    low = raw_line.lower()
    if flags.get('title'):
        return fonts['title']
    if flags.get('small'):
        return fonts['small_bold'] if flags.get('bold') else fonts['small']
    if '{phone}' in low:
        return fonts['phone']
    if '{bike}' in low:
        return fonts['bike']
    if '{date}' in low or '{time}' in low or '{datetime}' in low:
        return fonts['datetime']
    if flags.get('bold'):
        return fonts['body_bold']
    return fonts['body']

def build_receipt_image_gdi(data: List[str], settings: Dict, receipt_kind: str) -> Image.Image:
    """Build one customer/shop receipt as a bitmap using the Receipt Designer template."""
    design = _receipt_design_for_kind(get_receipt_designer_settings(settings), receipt_kind, settings)
    receipt_width = design['paper_width']
    margin = design['margin']
    line_spacing = 1 if design.get('compact_mode') else design['line_spacing']
    default_text_align = _normalize_receipt_align(design.get('text_align', 'left'), 'left')
    max_text_width = max(80, receipt_width - 2 * margin)
    fonts = _receipt_fonts_from_design(design)
    # QR visibility is controlled by the receipt template/Visible content settings only.
    # Legacy settings['qr_enabled'] is intentionally ignored to avoid a second hidden QR disable switch.
    context = _receipt_template_context(data, settings, receipt_kind)
    if design.get('simple_designer', True):
        template = build_simple_receipt_template(receipt_kind, design)
    else:
        template = design.get('customer_template') if receipt_kind == 'customer' else design.get('shop_template')
        if not str(template or '').strip():
            template = DEFAULT_CUSTOMER_RECEIPT_TEMPLATE if receipt_kind == 'customer' else DEFAULT_SHOP_RECEIPT_TEMPLATE
    # Add the private note to the shop receipt, even with a custom designer template.
    # Customer receipt layout and content remain completely unchanged.
    if receipt_kind == 'shop' and design.get('shop_show_note', True) and context.get('shop_note') and '{shop_note}' not in template.lower():
        if '{qr}' in template:
            template = template.replace('{qr}', '{shop_note}\n{qr}', 1)
        else:
            template += '\n{shop_note}'
    logo_settings = _receipt_logo_settings_for_kind(settings, design, receipt_kind)
    canvas = Image.new('RGB', (receipt_width, 4200), 'white')
    draw = ImageDraw.Draw(canvas)
    top_space = design.get('top_space', design.get('customer_top_space', 0) if receipt_kind == 'customer' else design.get('shop_top_space', 0))
    y = margin + _clamp_int(top_space, 0, 0, 240)
    first_visible_text = True

    def draw_line_text(line_text: str, raw_source: str, align: str, flags: Dict[str, bool], extra_after: int=0) -> None:
        nonlocal y, first_visible_text
        font = _choose_receipt_font(raw_source, flags, fonts)
        if first_visible_text and line_text.strip() and (not flags.get('small')) and (not any((token in raw_source.lower() for token in ('{phone}', '{bike}', '{date}', '{time}', '{datetime}', '{shop_info}', '{issues}')))):
            font = fonts['title']
            if align == 'left':
                align = 'center'
        y = _draw_wrapped_text(draw, y, line_text, font, max_text_width, x=margin, align=align, extra_after=extra_after, receipt_width=receipt_width, line_spacing=line_spacing)
        if line_text.strip():
            first_visible_text = False
    for raw_line in str(template).replace('\r\n', '\n').replace('\r', '\n').split('\n'):
        raw_stripped = raw_line.strip()
        raw_lower = raw_stripped.lower()
        if not raw_stripped:
            y += max(8, _line_height(draw, fonts['body'], line_spacing) // 2)
            continue
        if '{shop_note}' in raw_lower:
            if receipt_kind == 'shop' and design.get('shop_show_note', True) and context.get('shop_note'):
                note_align = _receipt_design_align(design, 'shop_note_align', 'left')
                note_markup = raw_line.replace('{shop_note}', '').strip()
                if note_markup:
                    _ignored_note, note_align, _note_flags = _parse_receipt_line_markup(note_markup, note_align)
                note_label = str(design.get('shop_note_label_text', 'Shop Note') or '').strip()
                if design.get('shop_show_note_label', True) and note_label:
                    y = _draw_wrapped_text(draw, y, f'{note_label}:', fonts['body_bold'], max_text_width, x=margin, align=note_align, extra_after=2, receipt_width=receipt_width, line_spacing=line_spacing)
                for note_line in context['shop_note'].splitlines():
                    if note_line.strip():
                        y = _draw_wrapped_text(draw, y, note_line, fonts['note'], max_text_width, x=margin, align=note_align, extra_after=2, receipt_width=receipt_width, line_spacing=line_spacing)
                first_visible_text = False
            continue
        if '{logo}' in raw_lower or '{{logo}}' in raw_lower:
            logo_markup = raw_line.replace('{{logo}}', '').replace('{logo}', '').strip()
            _parsed_logo, logo_align, _logo_flags = _parse_receipt_line_markup(logo_markup, _receipt_design_align(design, 'logo_align', 'center'))
            logo_img = _load_receipt_logo_image(logo_settings, max_text_width)
            if logo_img:
                y = _paste_aligned(canvas, logo_img, y, max_text_width, align=logo_align, x=margin, extra_after=_clamp_int(design.get('logo_after_margin', 6), 6, 0, 120))
            continue
        if '{qr}' in raw_lower or '{{qr}}' in raw_lower:
            qr_markup = raw_line.replace('{{qr}}', '').replace('{qr}', '').strip()
            _parsed_qr, qr_align, _qr_flags = _parse_receipt_line_markup(qr_markup, _receipt_design_align(design, 'qr_align', 'center'))
            qr_img = None
            if receipt_kind == 'customer':
                qr_img = _qr_image_for_gdi(str(settings.get('customer_qr', '') or ''), design.get('qr_width', design.get('customer_qr_width', 220)), max_text_width, design.get('qr_aspect', design.get('customer_qr_aspect', 100)))
            elif receipt_kind == 'shop':
                phone = context.get('phone', '')
                sms_text = str(settings.get('shop_qr', '') or '')
                qr_img = _qr_image_for_gdi(f'sms:{phone}?body={sms_text}', design.get('qr_width', design.get('shop_qr_width', 220)), max_text_width, design.get('qr_aspect', design.get('shop_qr_aspect', 100)))
            if qr_img:
                y = _paste_aligned(canvas, qr_img, y, max_text_width, align=qr_align, x=margin, extra_after=6 if design.get('compact_mode') else 10)
            continue
        if raw_lower in ('{line}', '---', '—'):
            y += 4
            draw.line((margin, y, receipt_width - margin, y), fill=0, width=1)
            y += 7 if design.get('compact_mode') else 10
            continue
        if '{shop_info}' in raw_lower:
            shop_info = str(context.get('shop_info', '') or '').strip()
            if shop_info:
                prefix = raw_line.replace('{shop_info}', '').strip()
                shop_info_default_align = _receipt_design_align(design, 'shop_info_align', default_text_align)
                shop_info_align = shop_info_default_align
                if prefix:
                    parsed_prefix, shop_info_align, flags = _parse_receipt_line_markup(prefix, shop_info_default_align)
                    rendered_prefix = _replace_receipt_placeholders(parsed_prefix, context)
                    if rendered_prefix.strip():
                        draw_line_text(rendered_prefix, prefix, shop_info_align, flags)
                for info_line in shop_info.split('\n'):
                    parsed, align, flags = _parse_receipt_line_markup(info_line, shop_info_align)
                    y = _draw_wrapped_text(draw, y, parsed, fonts['shop_info'], max_text_width, x=margin, align=align, extra_after=1, receipt_width=receipt_width, line_spacing=line_spacing)
                    if parsed.strip():
                        first_visible_text = False
            continue
        if '{issues}' in raw_lower:
            issue_markup = raw_line.replace('{issues}', '').strip()
            _parsed_issue_markup, issue_align, _issue_flags = _parse_receipt_line_markup(issue_markup, _receipt_design_align(design, 'issue_align', default_text_align))
            issues = context.get('issues') or []
            if issues:
                for issue in issues:
                    y = _draw_wrapped_text(draw, y, f'• {issue}', fonts['issue'], max_text_width, x=margin, align=issue_align, extra_after=2, receipt_width=receipt_width, line_spacing=line_spacing)
                    first_visible_text = False
            else:
                y = _draw_wrapped_text(draw, y, '• Inga fel angivna', fonts['issue'], max_text_width, x=margin, align=issue_align, extra_after=2, receipt_width=receipt_width, line_spacing=line_spacing)
                first_visible_text = False
            continue
        parsed_line, align, flags = _parse_receipt_line_markup(raw_line, default_text_align)
        rendered_line = _replace_receipt_placeholders(parsed_line, context)
        draw_line_text(rendered_line, raw_line, align, flags)
    bottom_feed = design.get('bottom_feed', design.get('customer_bottom_feed', 28) if receipt_kind == 'customer' else design.get('shop_bottom_feed', 28))
    y += _clamp_int(bottom_feed, 28, 0, 240)
    y = max(160, min(y, canvas.height))
    return canvas.crop((0, 0, receipt_width, y))

def build_qr_record_image_gdi(qr_image: Image.Image, headers: List[str], values: Tuple, settings: Dict) -> Image.Image:
    """Build the QR detail printout from the bike-list window using Windows-driver rendering."""
    design = _receipt_design_for_kind(get_receipt_designer_settings(settings), 'shop', settings)
    receipt_width = design['paper_width']
    margin = design['margin']
    line_spacing = 1 if design.get('compact_mode') else design['line_spacing']
    max_text_width = max(80, receipt_width - 2 * margin)
    qr_width = max(80, min(int(design.get('qr_width', design.get('shop_qr_width', 220))), max_text_width))
    qr_height_percent = _clamp_int(design.get('qr_aspect', design.get('shop_qr_aspect', 100)), 100, 50, 200)
    qr_height = max(40, int(round(qr_width * qr_height_percent / 100.0)))
    family = design.get('font_family', 'Segoe UI')
    normal_font = _gdi_font(design.get('body_size', 21), False, family)
    bold_font = _gdi_font(design.get('body_size', 21), True, family)
    canvas = Image.new('RGB', (receipt_width, 3600), 'white')
    draw = ImageDraw.Draw(canvas)
    y = margin
    qr = qr_image.convert('RGB').resize((qr_width, qr_height), Image.Resampling.NEAREST)
    y = _paste_centered(canvas, qr, y, max_text_width, extra_after=6)
    y += 4
    for i, header in enumerate(headers):
        if i >= len(values):
            continue
        value = values[i] if values[i] else 'Not specified'
        y = _draw_wrapped_text(draw, y, f'{header}:', bold_font, max_text_width, x=margin, extra_after=0, receipt_width=receipt_width, line_spacing=line_spacing)
        y = _draw_wrapped_text(draw, y, str(value), normal_font, max_text_width, x=margin, extra_after=3, receipt_width=receipt_width, line_spacing=line_spacing)
    empty_lines = int(settings.get('empty_lines', DEFAULT_SETTINGS.get('empty_lines', 5)))
    y += max(0, empty_lines) * max(12, _line_height(draw, normal_font, line_spacing) // 2)
    y += design.get('bottom_feed', GDI_BOTTOM_FEED)
    y = max(180, min(y, canvas.height))
    return canvas.crop((0, 0, receipt_width, y))

def print_images_with_windows_driver(printer_name: str, images: List[Image.Image], doc_name: str='BikeTrack Receipt') -> None:
    """Print one or more receipt bitmap pages through the Windows printer driver/GDI."""
    if win32ui is None or win32con is None or ImageWin is None:
        raise RuntimeError('Windows GDI printing is not available. Please install/repair pywin32.')
    if not images:
        return
    dc = win32ui.CreateDC()
    try:
        dc.CreatePrinterDC(printer_name)
        printable_width = int(dc.GetDeviceCaps(win32con.HORZRES))
        printable_height = int(dc.GetDeviceCaps(win32con.VERTRES))
        if printable_width <= 0:
            printable_width = GDI_RECEIPT_WIDTH
        if printable_height <= 0:
            printable_height = 3000
        dc.StartDoc(doc_name)
        try:
            for image in images:
                page = image.convert('RGB')
                target_width = printable_width
                target_height = int(page.height * (target_width / float(page.width)))
                if target_height > printable_height and printable_height > 0:
                    scale = min(printable_width / float(page.width), printable_height / float(page.height))
                    target_width = max(1, int(page.width * scale))
                    target_height = max(1, int(page.height * scale))
                left = max(0, (printable_width - target_width) // 2)
                dc.StartPage()
                dib = ImageWin.Dib(page)
                dib.draw(dc.GetHandleOutput(), (left, 0, left + target_width, target_height))
                dc.EndPage()
        finally:
            dc.EndDoc()
    finally:
        dc.DeleteDC()

def _get_customer_receipt_printer(settings: Dict) -> Optional[str]:
    """Return the configured printer for the customer copy, with legacy fallback."""
    return settings.get('customer_printer') or settings.get('printer') or settings.get('shop_printer')

def _get_shop_receipt_printer(settings: Dict) -> Optional[str]:
    """Return the configured printer for the shop copy, with legacy fallback."""
    return settings.get('shop_printer') or settings.get('printer') or settings.get('customer_printer')

def print_receipt(data: List[str]) -> None:
    """Print customer and shop receipts with QR codes, optionally on two printers."""
    text = ENGLISH_TEXT
    settings = load_settings()
    customer_printer = _get_customer_receipt_printer(settings)
    shop_printer = _get_shop_receipt_printer(settings)
    if not customer_printer and (not shop_printer):
        status_label.config(text=text['status_printer'], fg='red')
        return
    customer_printer = customer_printer or shop_printer
    shop_printer = shop_printer or customer_printer
    threading.Thread(target=lambda: threaded_print_receipt(data, customer_printer, shop_printer, settings), daemon=True).start()

def threaded_print_receipt(data: List[str], customer_printer: str, shop_printer: str, settings: Dict) -> None:
    """Print receipts in a background thread using the Windows printer driver/GDI."""
    try:
        customer_image = build_receipt_image_gdi(data, settings, 'customer')
        shop_image = build_receipt_image_gdi(data, settings, 'shop')
        if customer_printer == shop_printer:
            print_images_with_windows_driver(customer_printer, [customer_image, shop_image], doc_name='BikeTrack Receipt')
        else:
            print_images_with_windows_driver(customer_printer, [customer_image], doc_name='BikeTrack Customer Receipt')
            print_images_with_windows_driver(shop_printer, [shop_image], doc_name='BikeTrack Shop Receipt')
    except Exception as e:
        root.after(0, lambda: messagebox.showerror('Printer Error', f'Failed to print receipt: {str(e)}'))

def list_printers() -> List[str]:
    """List all available printers on the system."""
    local_printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    network_printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_NETWORK)
    return [printer[2] for printer in local_printers + network_printers]

def update_ui_text(refresh_settings_window: bool=True) -> None:
    """Refresh static English UI text without corrupting the main menu."""
    text = ENGLISH_TEXT
    root.title(text['window_title'])
    main_labels = getattr(root, '_main_labels', None)
    if main_labels:
        main_labels['phone'].config(text=text['tel_label'])
        main_labels['model'].config(text=text['bike_label'])
        if 'issue_title' in main_labels:
            main_labels['issue_title'].config(text='Issue:')
        elif 'issues' in main_labels:
            for index, label in enumerate(main_labels['issues'], start=1):
                label.config(text=text['issue_label'].format(index))
    else:
        for widget in root.grid_slaves():
            if isinstance(widget, tk.Label) and widget.grid_info()['row'] == 0:
                widget.config(text=text['tel_label'])
            elif isinstance(widget, tk.Label) and widget.grid_info()['row'] == 1:
                widget.config(text=text['bike_label'])
            elif isinstance(widget, tk.Label) and widget.grid_info()['row'] in range(2, 12):
                row = widget.grid_info()['row'] - 1
                widget.config(text=text['issue_label'].format(row))
    save_button.config(text=text['save_button'])
    bike_list_button.config(text=text['bike_list_button'])
    apply_input_field_font_size()
    rebuild_main_menu()
    if refresh_settings_window and settings_window and tk.Toplevel.winfo_exists(settings_window):
        open_settings(update=True)
    if about_window and tk.Toplevel.winfo_exists(about_window):
        open_about(update=True)
    if help_window and tk.Toplevel.winfo_exists(help_window):
        open_help(update=True)
    if bike_list_window and tk.Toplevel.winfo_exists(bike_list_window):
        bike_list_window.destroy()
        open_bike_list()


def open_settings(update: bool=False) -> None:
    """Open the settings window with all configuration options."""
    global settings_window
    text = ENGLISH_TEXT
    if settings_window and tk.Toplevel.winfo_exists(settings_window) and (not update):
        settings_window.lift()
        return
    if update and settings_window:
        settings_window.destroy()
    settings_window = tk.Toplevel(root)
    settings_window.title(text['settings_title'])
    settings_window.geometry(get_safe_window_geometry(settings_window, 'main_settings', '1320x760'))
    settings_window.minsize(1180, 620)
    settings_window.resizable(True, True)
    set_window_icon(settings_window)
    settings_window.grid_rowconfigure(0, weight=1)
    settings_window.grid_rowconfigure(1, weight=0)
    settings_window.grid_columnconfigure(0, weight=1)
    settings = load_settings()
    style = ttk.Style()
    style.configure('TNotebook.Tab', padding=(10, 6), font=('Arial', 9))
    notebook = ttk.Notebook(settings_window, style='TNotebook')
    notebook.grid(row=0, column=0, sticky='nsew', padx=6, pady=(6, 3))
    shop_info_text = tk.Text()
    customer_qr_text = tk.Text()
    shop_qr_text = tk.Text()
    printer_combo = ttk.Combobox()
    receipt_design = get_receipt_designer_settings(settings)
    receipt_paper_width_var = tk.IntVar(value=receipt_design['paper_width'])
    receipt_margin_var = tk.IntVar(value=receipt_design['margin'])
    receipt_line_spacing_var = tk.IntVar(value=receipt_design['line_spacing'])
    receipt_bottom_feed_var = tk.IntVar(value=receipt_design['bottom_feed'])
    receipt_customer_top_space_var = tk.IntVar(value=receipt_design.get('customer_top_space', 0))
    receipt_customer_bottom_feed_var = tk.IntVar(value=receipt_design.get('customer_bottom_feed', receipt_design.get('bottom_feed', 28)))
    receipt_shop_top_space_var = tk.IntVar(value=receipt_design.get('shop_top_space', 0))
    receipt_shop_bottom_feed_var = tk.IntVar(value=receipt_design.get('shop_bottom_feed', receipt_design.get('bottom_feed', 28)))
    receipt_font_family_var = tk.StringVar(value=receipt_design.get('font_family', 'Segoe UI'))
    receipt_text_align_var = tk.StringVar(value=_receipt_alignment_display(receipt_design.get('text_align', 'left')))
    receipt_logo_align_var = tk.StringVar(value=_receipt_alignment_display(receipt_design.get('logo_align', 'center')))
    receipt_title_align_var = tk.StringVar(value=_receipt_alignment_display(receipt_design.get('title_align', 'center')))
    receipt_phone_align_var = tk.StringVar(value=_receipt_alignment_display(receipt_design.get('phone_align', 'left')))
    receipt_bike_align_var = tk.StringVar(value=_receipt_alignment_display(receipt_design.get('bike_align', 'left')))
    receipt_datetime_align_var = tk.StringVar(value=_receipt_alignment_display(receipt_design.get('datetime_align', 'left')))
    receipt_shop_info_align_var = tk.StringVar(value=_receipt_alignment_display(receipt_design.get('shop_info_align', 'left')))
    receipt_issue_align_var = tk.StringVar(value=_receipt_alignment_display(receipt_design.get('issue_align', 'left')))
    receipt_qr_align_var = tk.StringVar(value=_receipt_alignment_display(receipt_design.get('qr_align', 'center')))
    receipt_thanks_align_var = tk.StringVar(value=_receipt_alignment_display(receipt_design.get('thanks_align', 'center')))
    receipt_customer_title_text_var = tk.StringVar(value=receipt_design.get('customer_title_text', DEFAULT_CUSTOMER_TITLE_TEXT))
    receipt_shop_title_text_var = tk.StringVar(value=receipt_design.get('shop_title_text', DEFAULT_SHOP_TITLE_TEXT))
    receipt_thank_you_text_var = tk.StringVar(value=receipt_design.get('thank_you_text', DEFAULT_THANK_YOU_TEXT))
    receipt_title_size_var = tk.IntVar(value=receipt_design['title_size'])
    receipt_body_size_var = tk.IntVar(value=receipt_design['body_size'])
    receipt_small_size_var = tk.IntVar(value=receipt_design['small_size'])
    receipt_phone_size_var = tk.IntVar(value=receipt_design['phone_size'])
    receipt_bike_size_var = tk.IntVar(value=receipt_design['bike_size'])
    receipt_datetime_size_var = tk.IntVar(value=receipt_design['datetime_size'])
    receipt_shop_info_size_var = tk.IntVar(value=receipt_design['shop_info_size'])
    receipt_issue_size_var = tk.IntVar(value=receipt_design['issue_size'])
    receipt_logo_width_var = tk.IntVar(value=receipt_design['logo_width'])
    receipt_logo_aspect_var = tk.IntVar(value=receipt_design.get('logo_aspect', 100))
    receipt_logo_after_margin_var = tk.IntVar(value=receipt_design.get('logo_after_margin', 6))
    receipt_customer_qr_width_var = tk.IntVar(value=receipt_design.get('customer_qr_width', settings.get('qr_width_settings', {}).get('customer_qr_width', 220)))
    receipt_shop_qr_width_var = tk.IntVar(value=receipt_design.get('shop_qr_width', settings.get('qr_width_settings', {}).get('shop_qr_width', 220)))
    receipt_customer_qr_aspect_var = tk.IntVar(value=receipt_design.get('customer_qr_aspect', 100))
    receipt_shop_qr_aspect_var = tk.IntVar(value=receipt_design.get('shop_qr_aspect', 100))
    receipt_title_bold_var = tk.BooleanVar(value=receipt_design.get('title_bold', True))
    receipt_phone_bold_var = tk.BooleanVar(value=receipt_design.get('phone_bold', True))
    receipt_bike_bold_var = tk.BooleanVar(value=receipt_design.get('bike_bold', True))
    receipt_datetime_bold_var = tk.BooleanVar(value=receipt_design.get('datetime_bold', False))
    receipt_shop_info_bold_var = tk.BooleanVar(value=receipt_design.get('shop_info_bold', False))
    receipt_issue_bold_var = tk.BooleanVar(value=receipt_design.get('issue_bold', False))
    receipt_compact_mode_var = tk.BooleanVar(value=receipt_design.get('compact_mode', True))
    receipt_preview_kind_var = tk.StringVar(value='customer')
    receipt_customer_show_logo_var = tk.BooleanVar(value=receipt_design.get('customer_show_logo', True))
    receipt_customer_show_shop_info_var = tk.BooleanVar(value=receipt_design.get('customer_show_shop_info', True))
    receipt_customer_show_qr_var = tk.BooleanVar(value=receipt_design.get('customer_show_qr', True))
    receipt_customer_show_title_var = tk.BooleanVar(value=receipt_design.get('customer_show_title', True))
    receipt_customer_show_phone_var = tk.BooleanVar(value=receipt_design.get('customer_show_phone', True))
    receipt_customer_show_phone_label_var = tk.BooleanVar(value=receipt_design.get('customer_show_phone_label', True))
    receipt_customer_show_bike_var = tk.BooleanVar(value=receipt_design.get('customer_show_bike', True))
    receipt_customer_show_bike_label_var = tk.BooleanVar(value=receipt_design.get('customer_show_bike_label', True))
    receipt_customer_show_datetime_var = tk.BooleanVar(value=receipt_design.get('customer_show_datetime', True))
    receipt_customer_show_date_var = tk.BooleanVar(value=receipt_design.get('customer_show_date', receipt_design.get('customer_show_datetime', True)))
    receipt_customer_show_date_label_var = tk.BooleanVar(value=receipt_design.get('customer_show_date_label', True))
    receipt_customer_show_time_var = tk.BooleanVar(value=receipt_design.get('customer_show_time', receipt_design.get('customer_show_datetime', True)))
    receipt_customer_show_time_label_var = tk.BooleanVar(value=receipt_design.get('customer_show_time_label', True))
    receipt_customer_show_issues_var = tk.BooleanVar(value=receipt_design.get('customer_show_issues', False))
    receipt_customer_show_issues_label_var = tk.BooleanVar(value=receipt_design.get('customer_show_issues_label', True))
    receipt_customer_show_thanks_var = tk.BooleanVar(value=receipt_design.get('customer_show_thanks', True))
    receipt_shop_show_logo_var = tk.BooleanVar(value=receipt_design.get('shop_show_logo', False))
    receipt_shop_show_shop_info_var = tk.BooleanVar(value=receipt_design.get('shop_show_shop_info', False))
    receipt_shop_show_qr_var = tk.BooleanVar(value=receipt_design.get('shop_show_qr', True))
    receipt_shop_show_phone_var = tk.BooleanVar(value=receipt_design.get('shop_show_phone', True))
    receipt_shop_show_phone_label_var = tk.BooleanVar(value=receipt_design.get('shop_show_phone_label', True))
    receipt_shop_show_bike_var = tk.BooleanVar(value=receipt_design.get('shop_show_bike', True))
    receipt_shop_show_bike_label_var = tk.BooleanVar(value=receipt_design.get('shop_show_bike_label', True))
    receipt_shop_show_datetime_var = tk.BooleanVar(value=receipt_design.get('shop_show_datetime', True))
    receipt_shop_show_date_var = tk.BooleanVar(value=receipt_design.get('shop_show_date', receipt_design.get('shop_show_datetime', True)))
    receipt_shop_show_date_label_var = tk.BooleanVar(value=receipt_design.get('shop_show_date_label', True))
    receipt_shop_show_time_var = tk.BooleanVar(value=receipt_design.get('shop_show_time', receipt_design.get('shop_show_datetime', True)))
    receipt_shop_show_time_label_var = tk.BooleanVar(value=receipt_design.get('shop_show_time_label', True))
    receipt_shop_show_issues_var = tk.BooleanVar(value=receipt_design.get('shop_show_issues', True))
    receipt_shop_show_issues_label_var = tk.BooleanVar(value=receipt_design.get('shop_show_issues_label', True))
    receipt_shop_show_title_var = tk.BooleanVar(value=receipt_design.get('shop_show_title', True))
    receipt_shop_show_note_var = tk.BooleanVar(value=receipt_design.get('shop_show_note', True))
    receipt_shop_show_note_label_var = tk.BooleanVar(value=receipt_design.get('shop_show_note_label', True))
    backup_folder_var = tk.StringVar(value=make_backup_folder_setting(settings.get('backup_folder', 'backup_db')))
    # QR visibility is now controlled only from Visible content > QR code.
    # These aliases keep older save/preview code compatible without a second independent disable switch.
    customer_qr_enabled_var = receipt_customer_show_qr_var
    shop_qr_enabled_var = receipt_shop_show_qr_var
    customer_logo_settings = settings.get('customer_logo', DEFAULT_SETTINGS['customer_logo'])
    customer_logo_path_var = tk.StringVar(value=customer_logo_settings.get('path', ''))
    issue_autofill_enabled_var = tk.BooleanVar(value=settings.get('issue_autofill_enabled', True))
    issue_autofill_history_var = tk.BooleanVar(value=settings.get('issue_autofill_include_history', True))
    issue_autofill_auto_add_var = tk.BooleanVar(value=settings.get('issue_autofill_auto_add_new', True))
    bike_brand_autofill_enabled_var = tk.BooleanVar(value=settings.get('bike_brand_autofill_enabled', True))
    bike_brand_autofill_history_var = tk.BooleanVar(value=settings.get('bike_brand_autofill_include_history', True))
    bike_brand_autofill_auto_add_var = tk.BooleanVar(value=settings.get('bike_brand_autofill_auto_add_new', settings.get('bike_model_autofill_auto_add_new', True)))
    issue_autofill_frame = ttk.Frame(notebook)
    notebook.add(issue_autofill_frame, text='Issue Autofill')
    tk.Checkbutton(issue_autofill_frame, text='Enable autofill for issue fields', variable=issue_autofill_enabled_var).pack(anchor='w', padx=10, pady=(10, 4))
    tk.Checkbutton(issue_autofill_frame, text='Use saved issue history automatically', variable=issue_autofill_history_var).pack(anchor='w', padx=10, pady=(0, 4))
    tk.Checkbutton(issue_autofill_frame, text='Add new issues to this list after Save', variable=issue_autofill_auto_add_var).pack(anchor='w', padx=10, pady=(0, 8))
    tk.Label(issue_autofill_frame, text='Custom suggestions, one per line:').pack(anchor='w', padx=10, pady=(4, 2))
    issue_autofill_text = tk.Text(issue_autofill_frame, width=42, height=10)
    issue_autofill_text.pack(anchor='w', padx=10, pady=5, fill='both', expand=True)
    issue_autofill_text.insert(tk.END, settings.get('issue_autofill_custom_items', ''))
    tk.Label(issue_autofill_frame, text='Tip: type in an Issue field, press Down to choose, Enter to insert.', wraplength=330, justify='left').pack(anchor='w', padx=10, pady=(4, 8))
    bike_brand_autofill_frame = ttk.Frame(notebook)
    notebook.add(bike_brand_autofill_frame, text='Bike Autofill')
    tk.Checkbutton(bike_brand_autofill_frame, text='Enable staged autofill for brand, color, and model in the same field', variable=bike_brand_autofill_enabled_var).pack(anchor='w', padx=10, pady=(10, 4))
    tk.Checkbutton(bike_brand_autofill_frame, text='Use saved bike history automatically', variable=bike_brand_autofill_history_var).pack(anchor='w', padx=10, pady=(0, 4))
    tk.Checkbutton(bike_brand_autofill_frame, text='Add new bike brands to the Brands list after Save', variable=bike_brand_autofill_auto_add_var).pack(anchor='w', padx=10, pady=(0, 8))
    bike_autofill_lists_frame = ttk.Frame(bike_brand_autofill_frame)
    bike_autofill_lists_frame.pack(anchor='w', padx=10, pady=5, fill='both', expand=True)
    tk.Label(bike_autofill_lists_frame, text='Brands').grid(row=0, column=0, sticky='w', padx=(0, 6), pady=(0, 2))
    tk.Label(bike_autofill_lists_frame, text='Colors').grid(row=0, column=1, sticky='w', padx=6, pady=(0, 2))
    tk.Label(bike_autofill_lists_frame, text='Models').grid(row=0, column=2, sticky='w', padx=(6, 0), pady=(0, 2))
    bike_brand_autofill_text = tk.Text(bike_autofill_lists_frame, width=24, height=12)
    bike_brand_autofill_text.grid(row=1, column=0, sticky='nsew', padx=(0, 6), pady=5)
    bike_color_autofill_text = tk.Text(bike_autofill_lists_frame, width=24, height=12)
    bike_color_autofill_text.grid(row=1, column=1, sticky='nsew', padx=6, pady=5)
    bike_model_autofill_text = tk.Text(bike_autofill_lists_frame, width=24, height=12)
    bike_model_autofill_text.grid(row=1, column=2, sticky='nsew', padx=(6, 0), pady=5)
    saved_bike_brand_items = settings.get('bike_brand_autofill_custom_items', '').strip()
    saved_bike_color_items = settings.get('bike_color_autofill_custom_items', '').strip()
    saved_bike_model_items = settings.get('bike_model_autofill_custom_items', '').strip()
    bike_brand_autofill_text.insert(tk.END, saved_bike_brand_items or '\n'.join(DEFAULT_BIKE_BRAND_AUTOFILL_ITEMS))
    bike_color_autofill_text.insert(tk.END, saved_bike_color_items or '\n'.join(DEFAULT_BIKE_COLOR_AUTOFILL_ITEMS))
    bike_model_autofill_text.insert(tk.END, saved_bike_model_items or '\n'.join(DEFAULT_BIKE_MODEL_AUTOFILL_ITEMS))
    tk.Label(bike_brand_autofill_frame, text='Tip: in the Bike Model field choose brand first, then color, then model. Use Down/Enter or mouse click.', wraplength=520, justify='left').pack(anchor='w', padx=10, pady=(4, 8))
    printer_frame = ttk.Frame(notebook)
    notebook.add(printer_frame, text=text['printer_tab'])
    available_printers = list_printers()
    tk.Label(printer_frame, text='Customer receipt printer:').pack(anchor='w', padx=10, pady=(10, 2))
    customer_printer_combo = ttk.Combobox(printer_frame, values=available_printers, width=45)
    customer_printer_combo.pack(anchor='w', padx=10, pady=(0, 8))
    selected_customer_printer = settings.get('customer_printer') or settings.get('printer')
    if selected_customer_printer and selected_customer_printer in customer_printer_combo['values']:
        customer_printer_combo.set(selected_customer_printer)
    tk.Label(printer_frame, text='Shop receipt printer:').pack(anchor='w', padx=10, pady=(6, 2))
    shop_printer_combo = ttk.Combobox(printer_frame, values=available_printers, width=45)
    shop_printer_combo.pack(anchor='w', padx=10, pady=(0, 8))
    selected_shop_printer = settings.get('shop_printer') or settings.get('printer')
    if selected_shop_printer and selected_shop_printer in shop_printer_combo['values']:
        shop_printer_combo.set(selected_shop_printer)
    tk.Label(printer_frame, text='Choose two different Windows printers, or choose the same printer for both copies. If one is empty, the other one is used as fallback.', wraplength=420, justify='left').pack(anchor='w', padx=10, pady=5)
    designer_frame = ttk.Frame(notebook)
    notebook.add(designer_frame, text=text['font_tab'])
    designer_frame.columnconfigure(0, weight=1)
    designer_frame.rowconfigure(0, weight=1)

    # Receipt Designer is now split into two fully independent profiles.
    # Only receipt-designer UI/settings are changed here; the save/print flow remains unchanged.
    designer_canvas = tk.Canvas(designer_frame, highlightthickness=0, borderwidth=0)
    designer_scroll_y = ttk.Scrollbar(designer_frame, orient='vertical', command=designer_canvas.yview)
    designer_canvas.configure(yscrollcommand=designer_scroll_y.set)
    designer_canvas.grid(row=0, column=0, sticky='nsew')
    designer_scroll_y.grid(row=0, column=1, sticky='ns')

    designer_inner = ttk.Frame(designer_canvas)
    designer_canvas_window = designer_canvas.create_window((0, 0), window=designer_inner, anchor='nw')
    designer_inner.columnconfigure(0, weight=1)
    designer_inner.rowconfigure(0, weight=1)

    def _update_designer_scroll_region(_event=None) -> None:
        try:
            designer_canvas.configure(scrollregion=designer_canvas.bbox('all'))
        except Exception:
            pass

    def _resize_designer_inner(event) -> None:
        try:
            designer_canvas.itemconfigure(
                designer_canvas_window,
                width=max(event.width, designer_inner.winfo_reqwidth()),
                height=max(event.height, designer_inner.winfo_reqheight()),
            )
            _update_designer_scroll_region()
        except Exception:
            pass

    def _on_designer_mousewheel(event):
        try:
            if getattr(event, 'num', None) == 4:
                designer_canvas.yview_scroll(-3, 'units')
            elif getattr(event, 'num', None) == 5:
                designer_canvas.yview_scroll(3, 'units')
            else:
                delta = int(-1 * (event.delta / 120)) if getattr(event, 'delta', 0) else 0
                if delta:
                    designer_canvas.yview_scroll(delta, 'units')
        except Exception:
            pass
        return 'break'

    def _bind_designer_mousewheel(_event=None) -> None:
        designer_canvas.bind_all('<MouseWheel>', _on_designer_mousewheel)
        designer_canvas.bind_all('<Button-4>', _on_designer_mousewheel)
        designer_canvas.bind_all('<Button-5>', _on_designer_mousewheel)

    def _unbind_designer_mousewheel(_event=None) -> None:
        designer_canvas.unbind_all('<MouseWheel>')
        designer_canvas.unbind_all('<Button-4>')
        designer_canvas.unbind_all('<Button-5>')

    designer_inner.bind('<Configure>', _update_designer_scroll_region)
    designer_canvas.bind('<Configure>', _resize_designer_inner)
    designer_canvas.bind('<Enter>', _bind_designer_mousewheel)
    designer_canvas.bind('<Leave>', _unbind_designer_mousewheel)

    receipt_tabs = ttk.Notebook(designer_inner)
    receipt_tabs.grid(row=0, column=0, sticky='nsew', padx=8, pady=8)

    def _safe_var_get(variable, default=None):
        try:
            return variable.get()
        except Exception:
            return default

    def _read_text_widget(widget, default: str='') -> str:
        try:
            return widget.get('1.0', tk.END).strip()
        except Exception:
            return default

    def _receipt_var_profile(prefix: str) -> Dict[str, Any]:
        title_default = DEFAULT_CUSTOMER_TITLE_TEXT if prefix == 'customer' else DEFAULT_SHOP_TITLE_TEXT
        return {
            'paper_width': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_paper_width', receipt_design.get('paper_width', RECEIPT_80MM_DEFAULT_WIDTH_PX)), RECEIPT_80MM_DEFAULT_WIDTH_PX, 240, 832)),
            'margin': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_margin', receipt_design.get('margin', 8)), 8, 0, 80)),
            'line_spacing': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_line_spacing', receipt_design.get('line_spacing', 3)), 3, 0, 18)),
            'top_space': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_top_space', 0), 0, 0, 240)),
            'bottom_feed': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_bottom_feed', receipt_design.get('bottom_feed', 28)), 28, 0, 240)),
            'font_family': tk.StringVar(value=str(receipt_design.get(f'{prefix}_font_family', receipt_design.get('font_family', 'Segoe UI')) or 'Segoe UI')),
            'text_align': tk.StringVar(value=_receipt_alignment_display(receipt_design.get(f'{prefix}_text_align', receipt_design.get('text_align', 'left')))),
            'logo_align': tk.StringVar(value=_receipt_alignment_display(receipt_design.get(f'{prefix}_logo_align', receipt_design.get('logo_align', 'center')))),
            'title_align': tk.StringVar(value=_receipt_alignment_display(receipt_design.get(f'{prefix}_title_align', receipt_design.get('title_align', 'center')))),
            'phone_align': tk.StringVar(value=_receipt_alignment_display(receipt_design.get(f'{prefix}_phone_align', receipt_design.get('phone_align', 'left')))),
            'bike_align': tk.StringVar(value=_receipt_alignment_display(receipt_design.get(f'{prefix}_bike_align', receipt_design.get('bike_align', 'left')))),
            'datetime_align': tk.StringVar(value=_receipt_alignment_display(receipt_design.get(f'{prefix}_datetime_align', receipt_design.get('datetime_align', 'left')))),
            'shop_info_align': tk.StringVar(value=_receipt_alignment_display(receipt_design.get(f'{prefix}_shop_info_align', receipt_design.get('shop_info_align', 'left')))),
            'issue_align': tk.StringVar(value=_receipt_alignment_display(receipt_design.get(f'{prefix}_issue_align', receipt_design.get('issue_align', 'left')))),
            'qr_align': tk.StringVar(value=_receipt_alignment_display(receipt_design.get(f'{prefix}_qr_align', receipt_design.get('qr_align', 'center')))),
            'thanks_align': tk.StringVar(value=_receipt_alignment_display(receipt_design.get(f'{prefix}_thanks_align', receipt_design.get('thanks_align', 'center')))),
            'title_size': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_title_size', receipt_design.get('title_size', 27)), 27, 10, 72)),
            'body_size': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_body_size', receipt_design.get('body_size', 21)), 21, 8, 60)),
            'small_size': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_small_size', receipt_design.get('small_size', 16)), 16, 7, 48)),
            'phone_size': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_phone_size', receipt_design.get('phone_size', 23)), 23, 8, 72)),
            'bike_size': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_bike_size', receipt_design.get('bike_size', 23)), 23, 8, 72)),
            'datetime_size': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_datetime_size', receipt_design.get('datetime_size', 19)), 19, 7, 60)),
            'shop_info_size': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_shop_info_size', receipt_design.get('shop_info_size', 18)), 18, 7, 60)),
            'issue_size': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_issue_size', receipt_design.get('issue_size', 22)), 22, 8, 72)),
            'note_size': tk.IntVar(value=_clamp_int(receipt_design.get('shop_note_size', 21), 21, 7, 72)) if prefix == 'shop' else None,
            'note_bold': tk.BooleanVar(value=bool(receipt_design.get('shop_note_bold', False))) if prefix == 'shop' else None,
            'note_align': tk.StringVar(value=_receipt_alignment_display(receipt_design.get('shop_note_align', 'left'))) if prefix == 'shop' else None,
            'note_label_text': tk.StringVar(value=str(receipt_design.get('shop_note_label_text', 'Shop Note') or '')) if prefix == 'shop' else None,
            'title_bold': tk.BooleanVar(value=bool(receipt_design.get(f'{prefix}_title_bold', receipt_design.get('title_bold', True)))),
            'phone_bold': tk.BooleanVar(value=bool(receipt_design.get(f'{prefix}_phone_bold', receipt_design.get('phone_bold', True)))),
            'bike_bold': tk.BooleanVar(value=bool(receipt_design.get(f'{prefix}_bike_bold', receipt_design.get('bike_bold', True)))),
            'datetime_bold': tk.BooleanVar(value=bool(receipt_design.get(f'{prefix}_datetime_bold', receipt_design.get('datetime_bold', False)))),
            'shop_info_bold': tk.BooleanVar(value=bool(receipt_design.get(f'{prefix}_shop_info_bold', receipt_design.get('shop_info_bold', False)))),
            'issue_bold': tk.BooleanVar(value=bool(receipt_design.get(f'{prefix}_issue_bold', receipt_design.get('issue_bold', False)))),
            'compact_mode': tk.BooleanVar(value=bool(receipt_design.get(f'{prefix}_compact_mode', receipt_design.get('compact_mode', True)))),
            'logo_width': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_logo_width', receipt_design.get('logo_width', 260)), 260, 40, 832)),
            'logo_aspect': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_logo_aspect', receipt_design.get('logo_aspect', 100)), 100, 50, 200)),
            'logo_after_margin': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_logo_after_margin', receipt_design.get('logo_after_margin', 6)), 6, 0, 120)),
            'qr_width': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_qr_width', receipt_design.get('customer_qr_width' if prefix == 'customer' else 'shop_qr_width', 220)), 220, 80, 832)),
            'qr_aspect': tk.IntVar(value=_clamp_int(receipt_design.get(f'{prefix}_qr_aspect', receipt_design.get('customer_qr_aspect' if prefix == 'customer' else 'shop_qr_aspect', 100)), 100, 50, 200)),
            'title_text': tk.StringVar(value=_clean_receipt_single_line_text(receipt_design.get(f'{prefix}_title_text', receipt_design.get(f'{prefix}_title_text', title_default)), title_default)),
            'thank_you_text': tk.StringVar(value=_clean_receipt_single_line_text(receipt_design.get(f'{prefix}_thank_you_text', receipt_design.get('thank_you_text', DEFAULT_THANK_YOU_TEXT)), DEFAULT_THANK_YOU_TEXT)),
            'logo_path': tk.StringVar(value=str(receipt_design.get(f'{prefix}_logo_path', customer_logo_settings.get('path', '')) or '').strip()),
        }

    receipt_profiles = {'customer': _receipt_var_profile('customer'), 'shop': _receipt_var_profile('shop')}
    customer_logo_path_var = receipt_profiles['customer']['logo_path']
    preview_widgets: Dict[str, Dict[str, Any]] = {}
    preview_after_ids = {'customer': None, 'shop': None}

    def _compact_spin_cell(parent, label_text: str, variable, index: int, from_: int=0, to: int=999, width: int=5, columns: int=2):
        cell = ttk.Frame(parent)
        cell.grid(row=index // columns, column=index % columns, sticky='w', padx=4, pady=3)
        ttk.Label(cell, text=label_text).pack(anchor='w')
        spin = tk.Spinbox(cell, from_=from_, to=to, textvariable=variable, width=width)
        spin.pack(anchor='w', pady=(1, 0))
        return spin

    def _set_profile_paper_width(prefix: str, width_px: int):
        receipt_profiles[prefix]['paper_width'].set(_clamp_int(width_px, RECEIPT_80MM_DEFAULT_WIDTH_PX, 240, 832))

    def _alignment_cell(parent, label_text: str, variable, index: int, columns: int=2):
        cell = ttk.Frame(parent)
        cell.grid(row=index // columns, column=index % columns, sticky='w', padx=4, pady=2)
        ttk.Label(cell, text=label_text).pack(side='left', padx=(0, 3))
        combo = ttk.Combobox(cell, textvariable=variable, values=PRINT_ALIGNMENT_CHOICES, width=7, state='readonly')
        combo.pack(side='left')
        return combo

    def _font_size_cell(parent, label_text: str, size_var, bold_var, index: int):
        cell = ttk.Frame(parent)
        cell.grid(row=index // 2, column=index % 2, sticky='w', padx=4, pady=3)
        ttk.Label(cell, text=label_text).pack(anchor='w')
        controls = ttk.Frame(cell)
        controls.pack(anchor='w', pady=(1, 0))
        spin = tk.Spinbox(controls, from_=7, to=72, textvariable=size_var, width=5)
        spin.pack(side='left')
        if bold_var is not None:
            tk.Checkbutton(controls, text='Bold', variable=bold_var).pack(side='left', padx=(3, 0))
        return spin

    def _content_text(parent, label_text: str, value: str, height: int=7):
        ttk.Label(parent, text=label_text).pack(anchor='w', padx=8, pady=(6, 2))
        txt = tk.Text(parent, width=34, height=height, wrap='word')
        txt.pack(anchor='nw', fill='x', padx=8, pady=(0, 6))
        txt.insert(tk.END, value)
        return txt

    def _visible_headers(parent) -> None:
        ttk.Label(parent, text='Show').grid(row=0, column=0, sticky='w', padx=6, pady=(2, 4))
        ttk.Label(parent, text='Label').grid(row=0, column=1, sticky='w', padx=6, pady=(2, 4))
        for grid_col in range(2):
            parent.columnconfigure(grid_col, weight=1)

    def _add_visible_item(parent, prefix: str, show_text: str, show_var, index: int, label_text: Optional[str]=None, label_var=None):
        row = 1 + index
        tk.Checkbutton(parent, text=show_text, variable=show_var, command=lambda p=prefix: schedule_receipt_preview(p)).grid(row=row, column=0, sticky='w', padx=6, pady=2)
        if label_text and label_var is not None:
            tk.Checkbutton(parent, text=label_text, variable=label_var, command=lambda p=prefix: schedule_receipt_preview(p)).grid(row=row, column=1, sticky='w', padx=6, pady=2)

    def _build_one_receipt_tab(prefix: str, title: str, qr_label: str, qr_initial: str, show_vars: Dict[str, Any]):
        tab = ttk.Frame(receipt_tabs)
        receipt_tabs.add(tab, text=title)
        # The three Receipt Designer areas are resizable with draggable sashes.
        # This changes only the settings UI layout; saved print values and print logic are untouched.
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        profile = receipt_profiles[prefix]

        panes = tk.PanedWindow(
            tab,
            orient=tk.HORIZONTAL,
            sashwidth=7,
            sashrelief='raised',
            showhandle=True,
            handlesize=10,
            borderwidth=0,
            opaqueresize=True,
        )
        panes.grid(row=0, column=0, sticky='nsew', padx=6, pady=8)

        left = ttk.LabelFrame(panes, text=f'{title} settings')
        left.configure(width=255)
        panes.add(left, minsize=220)
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)
        controls = ttk.Notebook(left)
        controls.grid(row=0, column=0, sticky='nsew', padx=6, pady=6)
        layout_tab = ttk.Frame(controls)
        fonts_tab = ttk.Frame(controls)
        images_tab = ttk.Frame(controls)
        visible_tab = ttk.Frame(controls)
        controls.add(layout_tab, text='Layout')
        controls.add(fonts_tab, text='Fonts')
        controls.add(images_tab, text='Images / QR')
        controls.add(visible_tab, text='Visible')

        layout_box = ttk.LabelFrame(layout_tab, text='Paper / spacing')
        layout_box.pack(anchor='nw', fill='x', pady=(0, 8))
        for grid_col in range(2):
            layout_box.columnconfigure(grid_col, weight=1)
        paper_spin = _compact_spin_cell(layout_box, 'Paper width px', profile['paper_width'], 0, from_=240, to=832)
        ttk.Label(paper_spin.master, text='80mm default: 576', foreground='#555555').pack(anchor='w', pady=(2, 0))
        paper_buttons = ttk.Frame(paper_spin.master)
        paper_buttons.pack(anchor='w', pady=(2, 0))
        for button_index, (button_text, button_width) in enumerate(RECEIPT_COMMON_PAPER_WIDTHS):
            ttk.Button(
                paper_buttons,
                text=button_text,
                width=4,
                command=lambda p=prefix, width=button_width: _set_profile_paper_width(p, width),
            ).grid(row=button_index // 2, column=button_index % 2, sticky='w', padx=(0, 2), pady=(0, 2))
        _compact_spin_cell(layout_box, 'Margin px', profile['margin'], 1, from_=0, to=80)
        _compact_spin_cell(layout_box, 'Line spacing', profile['line_spacing'], 2, from_=0, to=18)
        _compact_spin_cell(layout_box, 'Top space', profile['top_space'], 3, from_=0, to=240)
        _compact_spin_cell(layout_box, 'Bottom space', profile['bottom_feed'], 4, from_=0, to=240)
        default_align_cell = ttk.Frame(layout_box)
        default_align_cell.grid(row=2, column=1, sticky='w', padx=4, pady=3)
        ttk.Label(default_align_cell, text='Default alignment').pack(anchor='w')
        ttk.Combobox(default_align_cell, textvariable=profile['text_align'], values=PRINT_ALIGNMENT_CHOICES, width=10, state='readonly').pack(anchor='w', pady=(1, 0))
        mode_cell = ttk.Frame(layout_box)
        mode_cell.grid(row=3, column=0, sticky='w', padx=4, pady=3)
        ttk.Label(mode_cell, text='Mode').pack(anchor='w')
        tk.Checkbutton(mode_cell, text='Compact mode', variable=profile['compact_mode'], command=lambda p=prefix: schedule_receipt_preview(p)).pack(anchor='w', pady=(1, 0))

        font_box = ttk.LabelFrame(fonts_tab, text='Font sizes')
        font_box.pack(anchor='nw', fill='x', pady=(0, 8))
        for grid_col in range(2):
            font_box.columnconfigure(grid_col, weight=1)
        family_cell = ttk.Frame(font_box)
        family_cell.grid(row=0, column=0, sticky='w', padx=4, pady=3)
        ttk.Label(family_cell, text='Family').pack(anchor='w')
        ttk.Combobox(family_cell, textvariable=profile['font_family'], values=PRINT_FONT_CHOICES, width=13, state='readonly').pack(anchor='w', pady=(1, 0))
        _font_size_cell(font_box, 'Title', profile['title_size'], profile['title_bold'], 1)
        _font_size_cell(font_box, 'Body', profile['body_size'], None, 2)
        _font_size_cell(font_box, 'Small', profile['small_size'], None, 3)
        _font_size_cell(font_box, 'Phone', profile['phone_size'], profile['phone_bold'], 4)
        _font_size_cell(font_box, 'Bike', profile['bike_size'], profile['bike_bold'], 5)
        _font_size_cell(font_box, 'Date/time', profile['datetime_size'], profile['datetime_bold'], 6)
        _font_size_cell(font_box, 'Shop info', profile['shop_info_size'], profile['shop_info_bold'], 7)
        _font_size_cell(font_box, 'Issues', profile['issue_size'], profile['issue_bold'], 8)
        if prefix == 'shop':
            _font_size_cell(font_box, 'Shop Note', profile['note_size'], profile['note_bold'], 9)

        alignment_box = ttk.LabelFrame(fonts_tab, text='Element alignment')
        alignment_box.pack(anchor='nw', fill='x', pady=(0, 8))
        for grid_col in range(2):
            alignment_box.columnconfigure(grid_col, weight=1)
        _alignment_cell(alignment_box, 'Logo', profile['logo_align'], 0)
        _alignment_cell(alignment_box, 'Title', profile['title_align'], 1)
        _alignment_cell(alignment_box, 'Phone', profile['phone_align'], 2)
        _alignment_cell(alignment_box, 'Bike', profile['bike_align'], 3)
        _alignment_cell(alignment_box, 'Date/Time', profile['datetime_align'], 4)
        _alignment_cell(alignment_box, 'Issues', profile['issue_align'], 5)
        _alignment_cell(alignment_box, 'Shop info', profile['shop_info_align'], 6)
        _alignment_cell(alignment_box, 'QR', profile['qr_align'], 7)
        _alignment_cell(alignment_box, 'Thanks', profile['thanks_align'], 8)
        if prefix == 'shop':
            _alignment_cell(alignment_box, 'Shop Note', profile['note_align'], 9)

        image_box = ttk.LabelFrame(images_tab, text='Logo and QR sizing')
        image_box.pack(anchor='nw', fill='x', pady=(0, 8))
        for grid_col in range(2):
            image_box.columnconfigure(grid_col, weight=1)
        _compact_spin_cell(image_box, 'Logo width', profile['logo_width'], 0, from_=40, to=832)
        _compact_spin_cell(image_box, 'Logo height %', profile['logo_aspect'], 1, from_=50, to=200)
        _compact_spin_cell(image_box, 'After logo px', profile['logo_after_margin'], 2, from_=0, to=120)
        _compact_spin_cell(image_box, 'QR width', profile['qr_width'], 3, from_=80, to=832)
        _compact_spin_cell(image_box, 'QR height %', profile['qr_aspect'], 4, from_=50, to=200)

        visible_box = ttk.LabelFrame(visible_tab, text='Visible content')
        visible_box.pack(anchor='nw', fill='x', pady=(0, 8))
        _visible_headers(visible_box)
        _add_visible_item(visible_box, prefix, 'Title', show_vars['title'], 0)
        _add_visible_item(visible_box, prefix, 'Logo', show_vars['logo'], 1)
        _add_visible_item(visible_box, prefix, 'Shop info', show_vars['shop_info'], 2)
        _add_visible_item(visible_box, prefix, 'QR code', show_vars['qr'], 3)
        _add_visible_item(visible_box, prefix, 'Phone', show_vars['phone'], 4, 'Phone:', show_vars['phone_label'])
        _add_visible_item(visible_box, prefix, 'Bike', show_vars['bike'], 5, 'Cykel:', show_vars['bike_label'])
        _add_visible_item(visible_box, prefix, 'Date', show_vars['date'], 6, 'Datum:', show_vars['date_label'])
        _add_visible_item(visible_box, prefix, 'Time', show_vars['time'], 7, 'Tid:', show_vars['time_label'])
        _add_visible_item(visible_box, prefix, 'Issues', show_vars['issues'], 8, 'Fel/Brist:', show_vars['issues_label'])
        if prefix == 'customer':
            _add_visible_item(visible_box, prefix, 'Thanks text', show_vars['thanks'], 9)
        else:
            _add_visible_item(visible_box, prefix, 'Shop Note', show_vars['note'], 9, 'Note heading', show_vars['note_label'])

        preview_box = ttk.LabelFrame(panes, text='Live preview')
        preview_box.configure(width=340)
        panes.add(preview_box, minsize=260)
        preview_box.columnconfigure(0, weight=1)
        preview_box.rowconfigure(1, weight=1)
        preview_header = ttk.Frame(preview_box)
        preview_header.grid(row=0, column=0, sticky='ew', padx=8, pady=(6, 0))
        ttk.Label(preview_header, text=title).pack(side='left')
        ttk.Button(preview_header, text='Print test', command=lambda p=prefix: print_test_receipt(p)).pack(side='right', padx=2)
        preview_area = ttk.Frame(preview_box)
        preview_area.grid(row=1, column=0, sticky='nsew', padx=8, pady=(6, 8))
        preview_canvas = tk.Canvas(preview_area, background='#eeeeee', highlightthickness=1, highlightbackground='#c9c9c9')
        preview_scroll_y = ttk.Scrollbar(preview_area, orient='vertical', command=preview_canvas.yview)
        preview_canvas.configure(yscrollcommand=preview_scroll_y.set)
        preview_scroll_y.pack(side='right', fill='y', pady=(6, 0))
        preview_canvas.pack(side='left', fill='both', expand=True, pady=(6, 0))
        preview_label = tk.Label(preview_canvas, background='#eeeeee')
        preview_window = preview_canvas.create_window((12, 12), window=preview_label, anchor='nw')
        preview_widgets[prefix] = {'canvas': preview_canvas, 'label': preview_label, 'window': preview_window}

        content = ttk.LabelFrame(panes, text=f'{title} content')
        content.configure(width=380)
        panes.add(content, minsize=280)
        content.columnconfigure(0, weight=1)
        content_inner = ttk.Frame(content)
        content_inner.grid(row=0, column=0, sticky='nsew', padx=6, pady=6)
        content_inner.columnconfigure(0, weight=1)
        ttk.Label(content_inner, text='Title text').pack(anchor='w', padx=8, pady=(4, 2))
        tk.Entry(content_inner, textvariable=profile['title_text'], width=34).pack(anchor='nw', fill='x', padx=8, pady=(0, 6))
        ttk.Label(content_inner, text='Thank you text').pack(anchor='w', padx=8, pady=(4, 2))
        tk.Entry(content_inner, textvariable=profile['thank_you_text'], width=34).pack(anchor='nw', fill='x', padx=8, pady=(0, 6))
        if prefix == 'shop':
            ttk.Label(content_inner, text='Shop Note heading (shop receipt only)').pack(anchor='w', padx=8, pady=(4, 2))
            tk.Entry(content_inner, textvariable=profile['note_label_text'], width=34).pack(anchor='nw', fill='x', padx=8, pady=(0, 6))
        profile['shop_info_text_widget'] = _content_text(content_inner, 'Shop info used on this receipt', receipt_design.get(f'{prefix}_shop_info_text', settings.get('shop_info', '')), height=6)
        profile['qr_text_widget'] = _content_text(content_inner, qr_label, qr_initial, height=6)
        ttk.Label(content_inner, text='Logo file').pack(anchor='w', padx=8, pady=(4, 2))
        logo_entry = tk.Entry(content_inner, textvariable=profile['logo_path'], width=34)
        logo_entry.pack(anchor='nw', fill='x', padx=8, pady=(0, 4))
        logo_buttons = ttk.Frame(content_inner)
        logo_buttons.pack(anchor='w', padx=8, pady=(0, 8))

        def browse_logo(p=prefix):
            filetypes = [(text['customer_logo_filetypes'], '*.png *.jpg *.jpeg *.bmp *.gif'), ('PNG', '*.png'), ('JPEG', '*.jpg *.jpeg'), ('All files', '*.*')]
            selected = filedialog.askopenfilename(filetypes=filetypes)
            if selected:
                receipt_profiles[p]['logo_path'].set(selected)
        ttk.Button(logo_buttons, text=text['customer_logo_browse'], command=browse_logo).pack(side='left', padx=(0, 5))
        ttk.Button(logo_buttons, text=text['customer_logo_clear'], command=lambda p=prefix: receipt_profiles[p]['logo_path'].set('')).pack(side='left')

        def _on_preview_canvas_configure(_event=None, p=prefix):
            try:
                _center_preview_window(p)
                schedule_receipt_preview(p)
            except Exception:
                pass
        preview_canvas.bind('<Configure>', _on_preview_canvas_configure)

        def _set_initial_receipt_sashes(paned=panes):
            try:
                paned.update_idletasks()
                total_width = max(1, paned.winfo_width())
                if total_width > 760:
                    first_sash = min(270, max(220, int(total_width * 0.28)))
                    second_sash = min(total_width - 300, first_sash + min(360, max(280, int(total_width * 0.36))))
                    if second_sash > first_sash + 120:
                        paned.sash_place(0, first_sash, 0)
                        paned.sash_place(1, second_sash, 0)
            except Exception:
                pass
        settings_window.after(350, _set_initial_receipt_sashes)

    customer_show_vars = {
        'logo': receipt_customer_show_logo_var,
        'shop_info': receipt_customer_show_shop_info_var,
        'qr': receipt_customer_show_qr_var,
        'title': receipt_customer_show_title_var,
        'phone': receipt_customer_show_phone_var,
        'phone_label': receipt_customer_show_phone_label_var,
        'bike': receipt_customer_show_bike_var,
        'bike_label': receipt_customer_show_bike_label_var,
        'date': receipt_customer_show_date_var,
        'date_label': receipt_customer_show_date_label_var,
        'time': receipt_customer_show_time_var,
        'time_label': receipt_customer_show_time_label_var,
        'issues': receipt_customer_show_issues_var,
        'issues_label': receipt_customer_show_issues_label_var,
        'thanks': receipt_customer_show_thanks_var,
    }
    shop_show_vars = {
        'logo': receipt_shop_show_logo_var,
        'shop_info': receipt_shop_show_shop_info_var,
        'qr': receipt_shop_show_qr_var,
        'title': receipt_shop_show_title_var,
        'phone': receipt_shop_show_phone_var,
        'phone_label': receipt_shop_show_phone_label_var,
        'bike': receipt_shop_show_bike_var,
        'bike_label': receipt_shop_show_bike_label_var,
        'date': receipt_shop_show_date_var,
        'date_label': receipt_shop_show_date_label_var,
        'time': receipt_shop_show_time_var,
        'time_label': receipt_shop_show_time_label_var,
        'issues': receipt_shop_show_issues_var,
        'issues_label': receipt_shop_show_issues_label_var,
        'note': receipt_shop_show_note_var,
        'note_label': receipt_shop_show_note_label_var,
    }

    _build_one_receipt_tab('customer', 'Customer Receipt', text['customer_qr_label'], settings.get('customer_qr', ''), customer_show_vars)
    _build_one_receipt_tab('shop', 'Shop Receipt', text['shop_qr_label'], settings.get('shop_qr', ''), shop_show_vars)
    shop_info_text = receipt_profiles['customer']['shop_info_text_widget']
    customer_qr_text = receipt_profiles['customer']['qr_text_widget']
    shop_qr_text = receipt_profiles['shop']['qr_text_widget']

    def _current_receipt_tab_prefix() -> str:
        try:
            return 'customer' if receipt_tabs.index(receipt_tabs.select()) == 0 else 'shop'
        except Exception:
            return 'customer'

    def _center_preview_window(prefix: str) -> None:
        try:
            widgets = preview_widgets[prefix]
            preview_canvas = widgets['canvas']
            preview_label = widgets['label']
            preview_window = widgets['window']
            preview_canvas.update_idletasks()
            canvas_width = max(1, preview_canvas.winfo_width())
            label_width = max(1, preview_label.winfo_reqwidth())
            x = max(12, int((canvas_width - label_width) / 2))
            preview_canvas.coords(preview_window, x, 12)
            preview_canvas.configure(scrollregion=preview_canvas.bbox('all'))
        except Exception:
            pass

    def _collect_one_receipt_profile(prefix: str) -> Dict[str, Any]:
        profile = receipt_profiles[prefix]
        paper_width = _clamp_int(_safe_var_get(profile['paper_width'], RECEIPT_80MM_DEFAULT_WIDTH_PX), RECEIPT_80MM_DEFAULT_WIDTH_PX, 240, 832)
        margin = _clamp_int(_safe_var_get(profile['margin'], 8), 8, 0, 80)
        max_content = max(90, paper_width - 2 * margin)
        title_default = DEFAULT_CUSTOMER_TITLE_TEXT if prefix == 'customer' else DEFAULT_SHOP_TITLE_TEXT
        result = {
            f'{prefix}_paper_width': paper_width,
            f'{prefix}_margin': margin,
            f'{prefix}_line_spacing': _clamp_int(_safe_var_get(profile['line_spacing'], 3), 3, 0, 18),
            f'{prefix}_top_space': _clamp_int(_safe_var_get(profile['top_space'], 0), 0, 0, 240),
            f'{prefix}_bottom_feed': _clamp_int(_safe_var_get(profile['bottom_feed'], 28), 28, 0, 240),
            f'{prefix}_font_family': str(_safe_var_get(profile['font_family'], 'Segoe UI') or 'Segoe UI').strip() or 'Segoe UI',
            f'{prefix}_text_align': _normalize_receipt_align(_safe_var_get(profile['text_align'], 'Left'), 'left'),
            f'{prefix}_logo_align': _normalize_receipt_align(_safe_var_get(profile['logo_align'], 'Center'), 'center'),
            f'{prefix}_title_align': _normalize_receipt_align(_safe_var_get(profile['title_align'], 'Center'), 'center'),
            f'{prefix}_phone_align': _normalize_receipt_align(_safe_var_get(profile['phone_align'], 'Left'), 'left'),
            f'{prefix}_bike_align': _normalize_receipt_align(_safe_var_get(profile['bike_align'], 'Left'), 'left'),
            f'{prefix}_datetime_align': _normalize_receipt_align(_safe_var_get(profile['datetime_align'], 'Left'), 'left'),
            f'{prefix}_shop_info_align': _normalize_receipt_align(_safe_var_get(profile['shop_info_align'], 'Left'), 'left'),
            f'{prefix}_issue_align': _normalize_receipt_align(_safe_var_get(profile['issue_align'], 'Left'), 'left'),
            f'{prefix}_qr_align': _normalize_receipt_align(_safe_var_get(profile['qr_align'], 'Center'), 'center'),
            f'{prefix}_thanks_align': _normalize_receipt_align(_safe_var_get(profile['thanks_align'], 'Center'), 'center'),
            f'{prefix}_title_size': _clamp_int(_safe_var_get(profile['title_size'], 27), 27, 10, 72),
            f'{prefix}_body_size': _clamp_int(_safe_var_get(profile['body_size'], 21), 21, 8, 60),
            f'{prefix}_small_size': _clamp_int(_safe_var_get(profile['small_size'], 16), 16, 7, 48),
            f'{prefix}_phone_size': _clamp_int(_safe_var_get(profile['phone_size'], 23), 23, 8, 72),
            f'{prefix}_bike_size': _clamp_int(_safe_var_get(profile['bike_size'], 23), 23, 8, 72),
            f'{prefix}_datetime_size': _clamp_int(_safe_var_get(profile['datetime_size'], 19), 19, 7, 60),
            f'{prefix}_shop_info_size': _clamp_int(_safe_var_get(profile['shop_info_size'], 18), 18, 7, 60),
            f'{prefix}_issue_size': _clamp_int(_safe_var_get(profile['issue_size'], 22), 22, 8, 72),
            f'{prefix}_title_bold': bool(profile['title_bold'].get()),
            f'{prefix}_body_bold': False,
            f'{prefix}_phone_bold': bool(profile['phone_bold'].get()),
            f'{prefix}_bike_bold': bool(profile['bike_bold'].get()),
            f'{prefix}_datetime_bold': bool(profile['datetime_bold'].get()),
            f'{prefix}_shop_info_bold': bool(profile['shop_info_bold'].get()),
            f'{prefix}_issue_bold': bool(profile['issue_bold'].get()),
            f'{prefix}_compact_mode': bool(profile['compact_mode'].get()),
            f'{prefix}_logo_width': _clamp_int(_safe_var_get(profile['logo_width'], 260), 260, 40, max_content),
            f'{prefix}_logo_aspect': _clamp_int(_safe_var_get(profile['logo_aspect'], 100), 100, 50, 200),
            f'{prefix}_logo_after_margin': _clamp_int(_safe_var_get(profile['logo_after_margin'], 6), 6, 0, 120),
            f'{prefix}_qr_width': _clamp_int(_safe_var_get(profile['qr_width'], 220), 220, 80, max_content),
            f'{prefix}_qr_aspect': _clamp_int(_safe_var_get(profile['qr_aspect'], 100), 100, 50, 200),
            f'{prefix}_title_text': _clean_receipt_single_line_text(_safe_var_get(profile['title_text'], title_default), title_default),
            f'{prefix}_thank_you_text': _clean_receipt_single_line_text(_safe_var_get(profile['thank_you_text'], DEFAULT_THANK_YOU_TEXT), DEFAULT_THANK_YOU_TEXT),
            f'{prefix}_shop_info_text': _read_text_widget(profile.get('shop_info_text_widget'), settings.get('shop_info', '')),
            f'{prefix}_logo_path': str(_safe_var_get(profile['logo_path'], '') or '').strip(),
        }
        if prefix == 'shop':
            result.update({
                'shop_note_size': _clamp_int(_safe_var_get(profile['note_size'], 21), 21, 7, 72),
                'shop_note_bold': bool(_safe_var_get(profile['note_bold'], False)),
                'shop_note_align': _normalize_receipt_align(_safe_var_get(profile['note_align'], 'Left'), 'left'),
                'shop_note_label_text': str(_safe_var_get(profile['note_label_text'], 'Shop Note') or '').strip().replace('\n', ' '),
            })
        return result

    def _collect_receipt_designer_from_controls() -> Dict:
        customer_design = _collect_one_receipt_profile('customer')
        shop_design = _collect_one_receipt_profile('shop')
        data = DEFAULT_RECEIPT_DESIGNER.copy()
        data.update(receipt_design if isinstance(receipt_design, dict) else {})
        data.update(customer_design)
        data.update(shop_design)
        # Legacy/global keys mirror the customer receipt for backward compatibility.
        for key in ('paper_width', 'margin', 'line_spacing', 'font_family', 'text_align', 'logo_align', 'title_align', 'phone_align', 'bike_align', 'datetime_align', 'shop_info_align', 'issue_align', 'qr_align', 'thanks_align', 'title_size', 'body_size', 'small_size', 'phone_size', 'bike_size', 'datetime_size', 'shop_info_size', 'issue_size', 'title_bold', 'body_bold', 'phone_bold', 'bike_bold', 'datetime_bold', 'shop_info_bold', 'issue_bold', 'compact_mode', 'logo_width', 'logo_aspect', 'logo_after_margin'):
            data[key] = customer_design.get(f'customer_{key}', data.get(key))
        data['bottom_feed'] = customer_design['customer_bottom_feed']
        data['customer_top_space'] = customer_design['customer_top_space']
        data['customer_bottom_feed'] = customer_design['customer_bottom_feed']
        data['shop_top_space'] = shop_design['shop_top_space']
        data['shop_bottom_feed'] = shop_design['shop_bottom_feed']
        data['customer_qr_width'] = customer_design['customer_qr_width']
        data['shop_qr_width'] = shop_design['shop_qr_width']
        data['customer_qr_aspect'] = customer_design['customer_qr_aspect']
        data['shop_qr_aspect'] = shop_design['shop_qr_aspect']
        data['customer_title_text'] = customer_design['customer_title_text']
        data['shop_title_text'] = shop_design['shop_title_text']
        data['thank_you_text'] = customer_design['customer_thank_you_text']
        data['simple_designer'] = True
        data['center_qr'] = True
        data['customer_show_logo'] = bool(receipt_customer_show_logo_var.get())
        data['customer_show_shop_info'] = bool(receipt_customer_show_shop_info_var.get())
        data['customer_show_qr'] = bool(receipt_customer_show_qr_var.get())
        data['customer_show_title'] = bool(receipt_customer_show_title_var.get())
        data['customer_show_phone'] = bool(receipt_customer_show_phone_var.get())
        data['customer_show_phone_label'] = bool(receipt_customer_show_phone_label_var.get())
        data['customer_show_bike'] = bool(receipt_customer_show_bike_var.get())
        data['customer_show_bike_label'] = bool(receipt_customer_show_bike_label_var.get())
        data['customer_show_datetime'] = bool(receipt_customer_show_date_var.get() or receipt_customer_show_time_var.get())
        data['customer_show_date'] = bool(receipt_customer_show_date_var.get())
        data['customer_show_date_label'] = bool(receipt_customer_show_date_label_var.get())
        data['customer_show_time'] = bool(receipt_customer_show_time_var.get())
        data['customer_show_time_label'] = bool(receipt_customer_show_time_label_var.get())
        data['customer_show_issues'] = bool(receipt_customer_show_issues_var.get())
        data['customer_show_issues_label'] = bool(receipt_customer_show_issues_label_var.get())
        data['customer_show_thanks'] = bool(receipt_customer_show_thanks_var.get())
        data['shop_show_logo'] = bool(receipt_shop_show_logo_var.get())
        data['shop_show_shop_info'] = bool(receipt_shop_show_shop_info_var.get())
        data['shop_show_qr'] = bool(receipt_shop_show_qr_var.get())
        data['shop_show_title'] = bool(receipt_shop_show_title_var.get())
        data['shop_show_phone'] = bool(receipt_shop_show_phone_var.get())
        data['shop_show_phone_label'] = bool(receipt_shop_show_phone_label_var.get())
        data['shop_show_bike'] = bool(receipt_shop_show_bike_var.get())
        data['shop_show_bike_label'] = bool(receipt_shop_show_bike_label_var.get())
        data['shop_show_datetime'] = bool(receipt_shop_show_date_var.get() or receipt_shop_show_time_var.get())
        data['shop_show_date'] = bool(receipt_shop_show_date_var.get())
        data['shop_show_date_label'] = bool(receipt_shop_show_date_label_var.get())
        data['shop_show_time'] = bool(receipt_shop_show_time_var.get())
        data['shop_show_time_label'] = bool(receipt_shop_show_time_label_var.get())
        data['shop_show_issues'] = bool(receipt_shop_show_issues_var.get())
        data['shop_show_issues_label'] = bool(receipt_shop_show_issues_label_var.get())
        data['shop_show_note'] = bool(receipt_shop_show_note_var.get())
        data['shop_show_note_label'] = bool(receipt_shop_show_note_label_var.get())
        data['customer_template'] = DEFAULT_CUSTOMER_RECEIPT_TEMPLATE
        data['shop_template'] = DEFAULT_SHOP_RECEIPT_TEMPLATE
        return data

    def _collect_preview_settings() -> Dict:
        preview_settings = dict(settings)
        preview_design = _collect_receipt_designer_from_controls()
        preview_settings['receipt_designer'] = preview_design
        preview_settings['qr_width_settings'] = {'customer_qr_width': preview_design['customer_qr_width'], 'shop_qr_width': preview_design['shop_qr_width']}
        preview_settings['shop_info'] = _read_text_widget(receipt_profiles['customer'].get('shop_info_text_widget'), settings.get('shop_info', ''))
        preview_settings['customer_qr'] = _read_text_widget(customer_qr_text, settings.get('customer_qr', ''))
        preview_settings['shop_qr'] = _read_text_widget(shop_qr_text, settings.get('shop_qr', ''))
        preview_settings['qr_enabled'] = {'customer_qr': bool(preview_design.get('customer_show_qr', True)), 'shop_qr': bool(preview_design.get('shop_show_qr', True))}
        preview_settings['customer_logo'] = {'enabled': bool(preview_design.get('customer_logo_path', '')), 'path': preview_design.get('customer_logo_path', ''), 'width': preview_design['customer_logo_width'], 'height_percent': preview_design.get('customer_logo_aspect', 100)}
        return preview_settings

    def update_receipt_preview(prefix: Optional[str]=None) -> None:
        prefix = _receipt_kind_prefix(prefix or _current_receipt_tab_prefix())
        preview_after_ids[prefix] = None
        try:
            sample_data = [datetime.now().strftime('%Y-%m-%d %H:%M'), '0700000000', 'Skeppshult Svart Damcykel', 'Punktering bakhjul', 'Bromsen låter', 'Justera växlar'] + [''] * 7 + ['Ring kunden innan reparation.']
            preview_settings = _collect_preview_settings()
            img = build_receipt_image_gdi(sample_data, preview_settings, prefix)
            widgets = preview_widgets[prefix]
            preview_canvas = widgets['canvas']
            preview_label = widgets['label']
            preview_canvas.update_idletasks()
            canvas_width = preview_canvas.winfo_width()
            canvas_height = preview_canvas.winfo_height()
            if canvas_width <= 40:
                canvas_width = 520
            if canvas_height <= 40:
                canvas_height = 720
            max_preview_width = max(260, min(460, canvas_width - 32))
            max_preview_height = max(360, canvas_height - 32)
            ratio = min(1.0, max_preview_width / float(max(1, img.width)), max_preview_height / float(max(1, img.height)))
            preview_img = img.resize((max(1, int(img.width * ratio)), max(1, int(img.height * ratio))), Image.Resampling.LANCZOS) if ratio != 1.0 else img
            photo = ImageTk.PhotoImage(preview_img)
            preview_label.configure(image=photo, text='')
            preview_label.image = photo
            _center_preview_window(prefix)
        except Exception as exc:
            try:
                preview_widgets[prefix]['label'].configure(text=f'Preview error: {exc}', image='')
                preview_widgets[prefix]['label'].image = None
            except Exception:
                pass

    def schedule_receipt_preview(prefix: Optional[str]=None, *_args) -> None:
        prefix = _receipt_kind_prefix(prefix or _current_receipt_tab_prefix())
        try:
            if preview_after_ids.get(prefix) is not None:
                settings_window.after_cancel(preview_after_ids[prefix])
            preview_after_ids[prefix] = settings_window.after(250, lambda p=prefix: update_receipt_preview(p))
        except Exception:
            pass

    def print_test_receipt(prefix: Optional[str]=None) -> None:
        prefix = _receipt_kind_prefix(prefix or _current_receipt_tab_prefix())
        try:
            preview_settings = _collect_preview_settings()
            sample_data = [datetime.now().strftime('%Y-%m-%d %H:%M'), '0700000000', 'Skeppshult Svart Damcykel', 'Punktering bakhjul', 'Bromsen låter', 'Justera växlar'] + [''] * 7 + ['Ring kunden innan reparation.']
            image = build_receipt_image_gdi(sample_data, preview_settings, prefix)
            printer_name = customer_printer_combo.get().strip() if prefix == 'customer' else shop_printer_combo.get().strip()
            if not printer_name:
                printer_name = customer_printer_combo.get().strip() or shop_printer_combo.get().strip()
            if not printer_name:
                messagebox.showwarning(text['no_printer_title'], text['no_printer_message'])
                return
            threading.Thread(target=lambda: print_images_with_windows_driver(printer_name, [image], doc_name=f'BikeTrack {prefix.title()} Test Receipt'), daemon=True).start()
        except Exception as exc:
            messagebox.showerror(text['error_title'], str(exc))

    def _trace_profile_vars(prefix: str) -> None:
        for var in receipt_profiles[prefix].values():
            if isinstance(var, (tk.IntVar, tk.StringVar, tk.BooleanVar)):
                try:
                    var.trace_add('write', lambda *_args, p=prefix: schedule_receipt_preview(p))
                except Exception:
                    pass

    _trace_profile_vars('customer')
    _trace_profile_vars('shop')
    for var in [receipt_customer_show_logo_var, receipt_customer_show_shop_info_var, receipt_customer_show_qr_var, receipt_customer_show_title_var, receipt_customer_show_phone_var, receipt_customer_show_phone_label_var, receipt_customer_show_bike_var, receipt_customer_show_bike_label_var, receipt_customer_show_date_var, receipt_customer_show_date_label_var, receipt_customer_show_time_var, receipt_customer_show_time_label_var, receipt_customer_show_issues_var, receipt_customer_show_issues_label_var, receipt_customer_show_thanks_var]:
        try:
            var.trace_add('write', lambda *_args: schedule_receipt_preview('customer'))
        except Exception:
            pass
    for var in [receipt_shop_show_logo_var, receipt_shop_show_shop_info_var, receipt_shop_show_qr_var, receipt_shop_show_title_var, receipt_shop_show_phone_var, receipt_shop_show_phone_label_var, receipt_shop_show_bike_var, receipt_shop_show_bike_label_var, receipt_shop_show_date_var, receipt_shop_show_date_label_var, receipt_shop_show_time_var, receipt_shop_show_time_label_var, receipt_shop_show_issues_var, receipt_shop_show_issues_label_var, receipt_shop_show_note_var, receipt_shop_show_note_label_var]:
        try:
            var.trace_add('write', lambda *_args: schedule_receipt_preview('shop'))
        except Exception:
            pass
    for p in ('customer', 'shop'):
        for txt in (receipt_profiles[p].get('shop_info_text_widget'), receipt_profiles[p].get('qr_text_widget')):
            try:
                txt.bind('<KeyRelease>', lambda _event, prefix=p: schedule_receipt_preview(prefix))
            except Exception:
                pass

    def _on_receipt_tab_changed(_event=None):
        current = _current_receipt_tab_prefix()
        receipt_preview_kind_var.set(current)
        schedule_receipt_preview(current)

    receipt_tabs.bind('<<NotebookTabChanged>>', _on_receipt_tab_changed)
    settings_window.after(100, lambda: update_receipt_preview('customer'))
    settings_window.after(150, lambda: update_receipt_preview('shop'))
    backup_frame = ttk.Frame(notebook)
    notebook.add(backup_frame, text='Backup Folder')
    tk.Label(backup_frame, text='Select Backup Folder:').pack(anchor='w', padx=10, pady=10)
    backup_folder_entry = tk.Entry(backup_frame, textvariable=backup_folder_var, width=40)
    backup_folder_entry.pack(padx=10, pady=5)
    backup_buttons_frame = ttk.Frame(backup_frame)
    backup_buttons_frame.pack(padx=10, pady=5)

    def browse_backup_folder():
        selected = filedialog.askdirectory(initialdir=resolve_backup_folder_path(backup_folder_var.get()))
        if selected:
            backup_folder_var.set(make_backup_folder_setting(selected))
    tk.Button(backup_buttons_frame, text='Browse', command=browse_backup_folder).pack(side='left', padx=5)

    def open_backup_folder():
        folder_path = resolve_backup_folder_path(backup_folder_var.get() or 'backup_db')
        if not os.path.exists(folder_path):
            os.makedirs(folder_path, exist_ok=True)
        os.startfile(folder_path)
    tk.Button(backup_buttons_frame, text='Open Folder', command=open_backup_folder).pack(side='left', padx=5)
    settings_status_var = tk.StringVar(value='')

    def apply_all_settings(close_after: bool=False) -> bool:
        """Apply settings. Apply keeps the dialog open; Save closes it after a successful apply."""
        text = ENGLISH_TEXT
        try:
            settings['shop_info'] = shop_info_text.get('1.0', tk.END).strip()
            settings['customer_qr'] = customer_qr_text.get('1.0', tk.END).strip()
            settings['shop_qr'] = shop_qr_text.get('1.0', tk.END).strip()
            customer_printer = customer_printer_combo.get().strip()
            shop_printer = shop_printer_combo.get().strip()
            if not customer_printer and (not shop_printer):
                messagebox.showwarning(text['no_printer_title'], text['no_printer_message'])
            save_local_receipt_printers(customer_printer, shop_printer)
            settings['customer_printer'] = customer_printer or None
            settings['shop_printer'] = shop_printer or None
            settings['printer'] = shop_printer or customer_printer or None
            receipt_design_to_save = _collect_receipt_designer_from_controls()
            settings['receipt_designer'] = receipt_design_to_save
            settings['font_settings'] = DEFAULT_SETTINGS['font_settings'].copy()
            settings['qr_width_settings'] = {'customer_qr_width': receipt_design_to_save['customer_qr_width'], 'shop_qr_width': receipt_design_to_save['shop_qr_width']}
            settings['qr_enabled'] = {'customer_qr': bool(receipt_design_to_save.get('customer_show_qr', True)), 'shop_qr': bool(receipt_design_to_save.get('shop_show_qr', True))}
            settings['customer_logo'] = {'enabled': bool(customer_logo_path_var.get().strip()), 'path': customer_logo_path_var.get().strip(), 'width': receipt_design_to_save['logo_width'], 'height_percent': receipt_design_to_save.get('logo_aspect', 100)}
            settings['issue_autofill_enabled'] = issue_autofill_enabled_var.get()
            settings['issue_autofill_include_history'] = issue_autofill_history_var.get()
            settings['issue_autofill_auto_add_new'] = issue_autofill_auto_add_var.get()
            settings['issue_autofill_custom_items'] = issue_autofill_text.get('1.0', tk.END).strip()
            settings['bike_brand_autofill_enabled'] = bike_brand_autofill_enabled_var.get()
            settings['bike_brand_autofill_include_history'] = bike_brand_autofill_history_var.get()
            settings['bike_brand_autofill_auto_add_new'] = bike_brand_autofill_auto_add_var.get()
            settings['bike_model_autofill_auto_add_new'] = bike_brand_autofill_auto_add_var.get()
            settings['bike_brand_autofill_custom_items'] = bike_brand_autofill_text.get('1.0', tk.END).strip()
            settings['bike_color_autofill_custom_items'] = bike_color_autofill_text.get('1.0', tk.END).strip()
            settings['bike_model_autofill_custom_items'] = bike_model_autofill_text.get('1.0', tk.END).strip()
            settings['backup_folder'] = make_backup_folder_setting(backup_folder_var.get() or 'backup_db')
            save_settings(settings)
            save_window_geometry(settings_window, 'main_settings')
            update_ui_text(refresh_settings_window=False)
            refresh_issue_autofill_values(settings)
            refresh_bike_brand_autofill_values(settings)
            hide_issue_autofill()
            hide_bike_brand_autofill()
            apply_input_field_font_size(settings)
            try:
                if root is not None:
                    root.minsize(360, 240)
                    root.resizable(True, True)
                    auto_resize_main_window_height()
            except Exception:
                pass
            try:
                update_receipt_preview()
            except Exception:
                pass
            if close_after:
                settings_window.destroy()
                globals().update(settings_window=None)
            else:
                settings_status_var.set('Applied')
                settings_window.after(2500, lambda: settings_status_var.set(''))
            return True
        except ValueError as e:
            messagebox.showerror(text['error_title'], text['invalid_value'].format(str(e)))
            return False
    bottom_bar = ttk.Frame(settings_window)
    bottom_bar.grid(row=1, column=0, sticky='ew', padx=10, pady=(4, 8))
    ttk.Label(bottom_bar, textvariable=settings_status_var, foreground='#1b7f3a').pack(side='left', padx=(0, 8))
    ttk.Button(bottom_bar, text='Apply', command=lambda: apply_all_settings(close_after=False)).pack(side='right', padx=4)
    ttk.Button(bottom_bar, text='Save & Close', command=lambda: apply_all_settings(close_after=True)).pack(side='right', padx=4)
    ttk.Button(bottom_bar, text='Close', command=lambda: _close_settings_window()).pack(side='right', padx=4)

    def _close_settings_window():
        save_window_geometry(settings_window, 'main_settings')
        settings_window.destroy()
        globals().update(settings_window=None)
    settings_window.protocol('WM_DELETE_WINDOW', _close_settings_window)

def open_about(update: bool=False) -> None:
    """Show application details in a self-contained About dialog."""
    global about_window
    if about_window is not None:
        try:
            if about_window.winfo_exists():
                if update:
                    about_window.destroy()
                else:
                    about_window.lift()
                    about_window.focus_set()
                    return
        except tk.TclError:
            pass

    dialog = tk.Toplevel(root)
    about_window = dialog
    dialog.title('About BikeTrack Pro')
    dialog.geometry('520x365')
    dialog.minsize(440, 330)
    dialog.resizable(True, True)
    set_window_icon(dialog)

    panel = ttk.Frame(dialog, padding=24)
    panel.pack(fill='both', expand=True)
    ttk.Label(panel, text='BikeTrack Pro', font=('Segoe UI', 23, 'bold')).pack(anchor='w')
    ttk.Label(panel, text='Bicycle workshop management for Windows',
              font=('Segoe UI', 10)).pack(anchor='w', pady=(2, 16))
    ttk.Separator(panel).pack(fill='x', pady=(0, 14))
    details = (
        'Version 5.96 (source version)\n'
        'Developed by Aidin594\n'
        'Girra Sweden AB\n'
        'Customer records • Repair tracking • Visit history\n'
        'Customer and workshop receipt printing'
    )
    ttk.Label(panel, text=details, justify='left', wraplength=460).pack(anchor='w')
    ttk.Label(panel, text='Contact: info@vhcyklar.se').pack(anchor='w', pady=(12, 0))

    link = ttk.Label(panel, text='GitHub: aidin594/BikeTrack-Pro',
                     foreground='#0969da', cursor='hand2')
    link.pack(anchor='w', pady=(4, 0))
    link.bind('<Button-1>', lambda event: webbrowser.open(
        'https://github.com/aidin594/BikeTrack-Pro'))

    def close_dialog():
        global about_window
        about_window = None
        dialog.destroy()

    ttk.Button(panel, text='Close', command=close_dialog).pack(side='bottom', anchor='e', pady=(15, 0))
    dialog.protocol('WM_DELETE_WINDOW', close_dialog)


def open_help(update: bool=False) -> None:
    """Show built-in help without requiring an external help.html file."""
    global help_window
    if help_window is not None:
        try:
            if help_window.winfo_exists():
                if update:
                    help_window.destroy()
                else:
                    help_window.lift()
                    help_window.focus_set()
                    return
        except tk.TclError:
            pass

    dialog = tk.Toplevel(root)
    help_window = dialog
    dialog.title('BikeTrack Pro - Help')
    dialog.geometry('760x610')
    dialog.minsize(560, 420)
    set_window_icon(dialog)

    container = ttk.Frame(dialog, padding=12)
    container.pack(fill='both', expand=True)
    ttk.Label(container, text='BikeTrack Pro Help',
              font=('Segoe UI', 17, 'bold')).pack(anchor='w', pady=(0, 5))
    ttk.Label(container, text='Quick guide to the main functions. All help is available offline.',
              wraplength=700).pack(anchor='w', pady=(0, 12))

    notebook = ttk.Notebook(container)
    notebook.pack(fill='both', expand=True)

    topics = [
        ('Getting started',
         '1. Enter the customer phone number and the bicycle model.\n'
         '2. Type a repair issue and press Enter or click + Add.\n'
         '3. Repeat for additional repairs (up to 10).\n'
         '4. Click Save or press Ctrl+S to save the visit.\n'
         '5. Use Bike List to find, edit, or manage saved records.\n\n'
         'Tip: Test printing and printer settings before using the application for live jobs.'),
        ('Repairs & history',
         'Add issues using the repair field on the main window.\n'
         'Select an issue and use Edit, Remove, Up, or Down to manage it.\n'
         'Double-click an issue to edit it.\n\n'
         'Enter a phone number to see previous visits. The visit history can help you '
         'reuse an existing bicycle model. Use Bike List for the full list of records.\n\n'
         'You can mark customer numbers using the customer flag functions.'),
        ('Receipts & notes',
         'The Shop Note field is intended for internal workshop information. '
         'It is printed only on the shop receipt, not the customer receipt.\n\n'
         'Configure Customer Printer and Shop Printer separately in Settings.\n'
         'Printing requires compatible Windows printer drivers.\n'
         'Check a test receipt before relying on it for customer orders.\n\n'
         'Printer selection is stored locally for each Windows computer.'),
        ('Settings & backups',
         'Open Settings from the main menu to configure receipt layouts, '
         'printers, backup location and other available options.\n\n'
         'The application stores customer visits in a local SQLite database '
         '(customer_bikes.db, beside the application).\n'
         'Make regular backups and keep a copy outside the program folder.\n'
         'Do not upload real customer databases or local settings to GitHub.'),
        ('Troubleshooting',
         'Application does not start:\n'
         '  Check whether Windows security blocked the downloaded EXE.\n\n'
         'Receipts do not print:\n'
         '  Confirm the correct printers are selected in Settings, then test Windows printing.\n\n'
         'Visit history seems empty:\n'
         '  Check the entered phone number and the saved records in Bike List.\n\n'
         'Database is locked:\n'
         '  Wait for another database operation to finish before retrying.\n\n'
         'Project: https://github.com/aidin594/BikeTrack-Pro'),
    ]
    for heading, body in topics:
        page = ttk.Frame(notebook, padding=12)
        notebook.add(page, text=heading)
        scrollbar = ttk.Scrollbar(page, orient='vertical')
        scrollbar.pack(side='right', fill='y')
        viewer = tk.Text(page, wrap='word', font=('Segoe UI', 10),
                         padx=12, pady=12, borderwidth=0,
                         yscrollcommand=scrollbar.set)
        viewer.pack(side='left', fill='both', expand=True)
        scrollbar.configure(command=viewer.yview)
        viewer.insert('1.0', body)
        viewer.configure(state='disabled')

    def close_dialog():
        global help_window
        help_window = None
        dialog.destroy()

    ttk.Button(container, text='Close', command=close_dialog).pack(anchor='e', pady=(10, 0))
    dialog.protocol('WM_DELETE_WINDOW', close_dialog)


def rebuild_main_menu() -> None:
    """Rebuild the main menu from scratch so labels never shift or duplicate after settings changes."""
    global menu_bar
    if root is None:
        return
    text = ENGLISH_TEXT
    if menu_bar is None:
        menu_bar = tk.Menu(root)
    else:
        try:
            menu_bar.delete(0, tk.END)
        except Exception:
            menu_bar = tk.Menu(root)
    menu_bar.add_command(label=text['menu_settings'], command=open_settings)
    menu_bar.add_command(label='Bike Statistics', command=lambda: open_bike_statistics_dialog(root))
    menu_bar.add_command(label=text['menu_about'], command=open_about)
    menu_bar.add_command(label=text['menu_help'], command=open_help)
    root.config(menu=menu_bar)

def create_menu() -> None:
    """Create the main menu bar."""
    rebuild_main_menu()

def focus_next_widget(event: tk.Event) -> str:
    """Move focus to the next widget on Enter key press."""
    event.widget.tk_focusNext().focus()
    return 'break'

class BikeShopApp:
    """Application class for managing the bike list window."""
    WINDOW_SIZE = '800x400'
    QR_SIZE = (300, 300)

    def __init__(self, master: tk.Toplevel, initial_phone_history: str=''):
        self.root = master
        text = ENGLISH_TEXT
        self.root.title(text['bike_list_title'])
        self.root.geometry(get_saved_window_geometry('bike_list', self.WINDOW_SIZE))
        set_window_icon(self.root)
        self.qr_window = None
        self.undo_buffer = None
        self.treeview_font_size = self.load_treeview_font_size()
        self.current_phone_history = _phone_digits(initial_phone_history)
        self.history_mode = False
        self.current_view_data = []
        self.sort_column = None
        self.sort_reverse = False
        self._hover_item = None
        self._hover_original_tags = {}
        self.loading_label = tk.Label(self.root, text=text['loading_message'], fg='blue')
        self.loading_label.pack(pady=10)
        self.root.update()
        self.headers = []
        self.display_headers = []
        self.data = []
        threading.Thread(target=self.threaded_load_db_data, daemon=True).start()

    def threaded_load_db_data(self) -> None:
        """Load database data in a separate thread."""
        headers, data = self.load_db_data()
        self.root.after(0, lambda: self.finish_loading(headers, data))

    def finish_loading(self, headers, data) -> None:
        """Finish loading by updating the UI in the main thread."""
        self.headers = headers
        self.display_headers = self._make_display_headers(headers)
        self.data = data
        if not self.headers or not self.data:
            self.root.destroy()
            return
        self.loading_label.destroy()
        self.setup_ui()
        if self.current_phone_history:
            self.root.after(100, lambda: self.show_phone_history(self.current_phone_history, update_entry=True, show_if_empty=True))

    def load_db_data(self) -> Tuple[List[str], List[Dict]]:
        """Load data from the database."""
        text = ENGLISH_TEXT
        conn = None
        try:
            if 'bike_data' in _app_cache and time.time() - _app_cache.get('bike_data_time', 0) < 5:
                return (_app_cache['bike_headers'], _app_cache['bike_data'])
            if not Path(DATABASE_FILE).exists():
                init_database()
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute('PRAGMA table_info(customer_bikes)')
            columns = [col[1] for col in cursor.fetchall()]
            headers = columns[1:]
            cursor.execute(f"SELECT {', '.join(columns)} FROM customer_bikes ORDER BY datetime DESC")
            rows = cursor.fetchall()
            data = []
            for row in rows:
                row_dict = {}
                for i, col in enumerate(columns):
                    row_dict[col] = row[i]
                public_row = {k: v for k, v in row_dict.items() if k != 'id'}
                public_row['_db_id'] = row_dict.get('id')
                data.append(public_row)
            _app_cache['bike_headers'] = headers
            _app_cache['bike_data'] = data
            _app_cache['bike_data_time'] = time.time()
            return (headers, data)
        except sqlite3.Error as e:
            self.root.after(0, lambda: messagebox.showerror(text['error_title'], text['db_load_error'].format(str(e))))
            return ([], [])
        finally:
            if conn:
                release_db_connection(conn)

    def save_data_to_db(self, data: List[Dict], headers: List[str]) -> None:
        """Save data back to the database."""
        text = ENGLISH_TEXT
        self.save_loading_label = tk.Label(self.root, text='Saving...', fg='blue')
        self.save_loading_label.pack(pady=10)
        self.root.update()
        for widget in self.search_frame.winfo_children():
            if isinstance(widget, (tk.Button, ttk.Button)):
                widget.config(state='disabled')
        threading.Thread(target=lambda: self.threaded_save_data(data, headers), daemon=True).start()

    def threaded_save_data(self, data: List[Dict], headers: List[str]) -> None:
        """Save data to the database in a separate thread (batched, single transaction)."""
        text = ENGLISH_TEXT
        try:
            self.backup_db()
            conn = get_db_connection()
            try:
                cursor = conn.cursor()
                cursor.execute('DELETE FROM customer_bikes')
                columns = headers
                placeholders = ','.join(['?' for _ in columns])
                query = f"INSERT INTO customer_bikes ({','.join(columns)}) VALUES ({placeholders})"
                batch_size = 1000

                def gen_rows():
                    for row in data:
                        yield [row.get(col, '') for col in columns]
                it = iter(gen_rows())
                while True:
                    chunk = list(itertools.islice(it, batch_size))
                    if not chunk:
                        break
                    cursor.executemany(query, chunk)
                conn.commit()
            finally:
                release_db_connection(conn)
            self.root.after(0, self.finish_save_data)
        except Exception as e:
            try:
                self.root.after(0, lambda: messagebox.showerror(text['error_title'], text['db_save_error'].format(str(e))))
            except Exception:
                pass
            try:
                self.root.after(0, self.finish_save_data)
            except Exception:
                pass

    def finish_save_data(self) -> None:
        """Clean up after database save operation."""
        if hasattr(self, 'save_loading_label') and self.save_loading_label.winfo_exists():
            self.save_loading_label.destroy()
        if hasattr(self, 'search_frame') and self.search_frame.winfo_exists():
            for widget in self.search_frame.winfo_children():
                if isinstance(widget, (tk.Button, ttk.Button)):
                    widget.config(state='normal')
        try:
            show_toast(self.root, 'Saved.', 2000)
        except Exception:
            pass

    def backup_db(self) -> None:
        """Backup the database before modifications."""
        _thread_pool.submit(threaded_backup_database)

    def load_sms_message(self) -> str:
        """Load the SMS message from settings."""
        settings = load_settings()
        return settings.get('sms_message', DEFAULT_SETTINGS['sms_message'])

    def load_treeview_font_size(self) -> int:
        """Load the Treeview font size from settings."""
        settings = load_settings()
        return settings.get('treeview_font_size', DEFAULT_SETTINGS['treeview_font_size'])

    def _make_display_headers(self, headers: List[str]) -> List[str]:
        """Display-only columns first, followed by database columns."""
        display = ['age_since_visit', 'visit_date', 'visit_time']
        for h in headers:
            if h not in display:
                display.append(h)
        return display

    def _default_visible_columns(self) -> List[str]:
        """Default visible order: Ago first, Date/Time split, then useful database fields."""
        visible = ['age_since_visit', 'visit_date', 'visit_time']
        for h in self.headers:
            if h == 'datetime':
                continue
            if h not in visible:
                visible.append(h)
        return [c for c in visible if c in self.display_headers]

    def _load_visible_columns(self) -> List[str]:
        """Load and validate visible columns from settings."""
        settings = load_settings()
        stored = settings.get('bike_list_visible_columns', [])
        if not isinstance(stored, list):
            stored = []
        visible = [c for c in stored if c in self.display_headers]
        if not visible:
            visible = self._default_visible_columns()
        return visible

    def _column_label(self, col: str) -> str:
        """Human label for bike-list columns."""
        english_headers = {'age_since_visit': 'Ago', 'visit_date': 'Date', 'visit_time': 'Time', 'datetime': 'Date and Time', 'phone': 'Phone', 'model': 'Model', 'issue1': 'Issue 1', 'issue2': 'Issue 2', 'issue3': 'Issue 3', 'issue4': 'Issue 4', 'issue5': 'Issue 5', 'issue6': 'Issue 6', 'issue7': 'Issue 7', 'issue8': 'Issue 8', 'issue9': 'Issue 9', 'issue10': 'Issue 10'}
        labels = english_headers
        return labels.get(col, col)

    def _heading_text(self, col: str) -> str:
        """Column heading text with a small sort indicator."""
        label = self._column_label(col)
        if getattr(self, 'sort_column', None) == col:
            return label + (' ↓' if getattr(self, 'sort_reverse', False) else ' ↑')
        return label

    def _visible_tree_columns(self) -> List[str]:
        """Return actual Treeview display columns in visible order."""
        try:
            display = list(self.tree['displaycolumns'])
            if not display or display == ['#all']:
                return list(getattr(self, 'display_headers', self.headers))
            return [str(c) for c in display]
        except Exception:
            return list(getattr(self, 'visible_columns', []) or getattr(self, 'display_headers', self.headers))

    def _column_from_event_x(self, x: int) -> Optional[str]:
        """Map a mouse x-coordinate to the visible Treeview column id."""
        try:
            col_id = self.tree.identify_column(x)
            if not col_id or not col_id.startswith('#'):
                return None
            idx = int(col_id[1:]) - 1
            visible = self._visible_tree_columns()
            if 0 <= idx < len(visible):
                return visible[idx]
        except Exception:
            pass
        return None

    def _display_value_for_column(self, row: Dict, col: str) -> str:
        """Return the display value for one Treeview column."""
        dt_value = row.get('datetime', '')
        if col == 'age_since_visit':
            return visit_age_label(dt_value)
        if col == 'visit_date':
            return visit_date_label(dt_value)
        if col == 'visit_time':
            return visit_time_label(dt_value)
        return row.get(col, '')

    def _tree_values_for_row(self, row: Dict) -> List[str]:
        """Build Treeview display values without adding calculated fields to the database."""
        return [self._display_value_for_column(row, header) for header in self.display_headers]

    def _db_values_for_row(self, row: Dict) -> List[str]:
        """Return database values in original database-column order."""
        return [row.get(header, '') for header in self.headers]

    def _apply_visible_columns(self) -> None:
        """Apply current visible-column settings to the Treeview."""
        if not hasattr(self, 'tree'):
            return
        self.visible_columns = self._load_visible_columns()
        self.tree['displaycolumns'] = self.visible_columns
        self._refresh_heading_labels()
        self._apply_saved_or_measured_column_widths(self.tree, save_missing=True)

    def _remember_item_row(self, item_id: str, row: Dict) -> None:
        """Keep a stable link between a Treeview row and the original record dict."""
        if not hasattr(self, 'tree_item_rows'):
            self.tree_item_rows = {}
        self.tree_item_rows[item_id] = row

    def _insert_restored_row(self, row: Dict) -> None:
        """Insert a restored row and keep the Treeview-to-data mapping."""
        item_id = self.tree.insert('', 'end', values=self._tree_values_for_row(row))
        self._remember_item_row(item_id, row)
        self._apply_visible_columns()

    def _get_row_for_item(self, item_id: str) -> Optional[Dict]:
        """Return the original record dict for a Treeview item."""
        if hasattr(self, 'tree_item_rows') and item_id in self.tree_item_rows:
            return self.tree_item_rows[item_id]
        try:
            values = self.tree.item(item_id, 'values')
            value_map = dict(zip(self.display_headers, values))
            dt = value_map.get('datetime')
            phone = value_map.get('phone')
            model = value_map.get('model')
            for row in self.data:
                if (not dt or row.get('datetime') == dt) and (not phone or row.get('phone') == phone) and (not model or row.get('model') == model):
                    return row
        except Exception:
            pass
        return None

    def _get_selected_row(self) -> Optional[Dict]:
        """Return the selected record dict."""
        selected_item = self.tree.selection()
        if not selected_item:
            return None
        return self._get_row_for_item(selected_item[0])

    def _column_width_limits(self, col: str) -> Tuple[int, int]:
        """Minimum and maximum width for a bike-list column."""
        limits = {'age_since_visit': (105, 150), 'visit_date': (92, 115), 'visit_time': (62, 85), 'datetime': (135, 170), 'phone': (105, 150), 'model': (160, 330)}
        if col.startswith('issue'):
            return (120, 360)
        return limits.get(col, (85, 260))

    def _load_column_widths(self) -> Dict[str, int]:
        """Load saved bike-list column widths from Settings."""
        try:
            settings = load_settings()
            stored = settings.get('bike_list_column_widths', {})
            if not isinstance(stored, dict):
                return {}
            widths: Dict[str, int] = {}
            valid_columns = set(getattr(self, 'display_headers', self.headers))
            for col, value in stored.items():
                if col not in valid_columns:
                    continue
                try:
                    width = int(value)
                except Exception:
                    continue
                min_width, _ = self._column_width_limits(col)
                widths[col] = max(min_width, min(width, 900))
            return widths
        except Exception:
            return {}

    def save_current_column_widths(self) -> None:
        """Persist current visible Treeview column widths so they are restored next time."""
        try:
            if not hasattr(self, 'tree'):
                return
            settings = load_settings()
            stored = settings.get('bike_list_column_widths', {})
            if not isinstance(stored, dict):
                stored = {}
            valid_columns = set(getattr(self, 'display_headers', self.headers))
            widths = {str(k): int(v) for k, v in stored.items() if str(k) in valid_columns and str(v).isdigit()}
            for col in self._visible_tree_columns():
                if col not in valid_columns:
                    continue
                try:
                    width = int(self.tree.column(col, 'width'))
                except Exception:
                    continue
                min_width, _ = self._column_width_limits(col)
                widths[col] = max(min_width, min(width, 900))
            settings['bike_list_column_widths'] = widths
            save_settings(settings)
        except Exception:
            pass

    def schedule_save_column_widths(self, event: Optional[tk.Event]=None) -> None:
        """Debounce column-width saving after manual resizing or toolbar actions."""
        try:
            if hasattr(self, '_column_width_save_job') and self._column_width_save_job:
                self.root.after_cancel(self._column_width_save_job)
        except Exception:
            pass
        try:
            self._column_width_save_job = self.root.after(350, self.save_current_column_widths)
        except Exception:
            pass

    def _apply_saved_or_measured_column_widths(self, tree: Optional[ttk.Treeview]=None, save_missing: bool=False) -> None:
        """Apply saved widths when available; measure only new/unsaved visible columns."""
        try:
            tree = tree or self.tree
            visible = list(getattr(self, 'visible_columns', []) or self._load_visible_columns())
            visible = [c for c in visible if c in getattr(self, 'display_headers', self.headers)]
            saved = self._load_column_widths()
            rows = [self._get_row_for_item(item) for item in tree.get_children()]
            rows = [r for r in rows if r]
            if not rows:
                rows = list(getattr(self, 'current_view_data', []) or self.data)
            changed = False
            for col in visible:
                anchor = 'center' if col in ('age_since_visit', 'visit_date', 'visit_time', 'datetime', 'phone') else 'w'
                if col in saved:
                    width = saved[col]
                else:
                    width = self._measure_column_width(col, rows)
                    changed = True
                tree.column(col, width=width, anchor=anchor)
            if save_missing and changed:
                self.schedule_save_column_widths()
        except Exception:
            pass

    def _measure_column_width(self, col: str, rows: Optional[List[Dict]]=None) -> int:
        """Measure the best readable width for one column with safe limits."""
        font = tkfont.Font(size=self.treeview_font_size)
        min_width, max_width = self._column_width_limits(col)
        width = max(min_width, font.measure(self._column_label(col)) + 22)
        rows = rows if rows is not None else list(getattr(self, 'current_view_data', []) or self.data)
        for row in rows[:500]:
            txt = str(self._display_value_for_column(row, col))
            if len(txt) > 80:
                txt = txt[:80] + '…'
            width = max(width, font.measure(txt) + 22)
        return max(min_width, min(width, max_width))

    def auto_fit_column(self, col: str) -> None:
        """Autofit a single column, usually from double-clicking a header/separator."""
        try:
            if col not in getattr(self, 'display_headers', self.headers):
                return
            rows = [self._get_row_for_item(item) for item in self.tree.get_children()]
            rows = [r for r in rows if r]
            width = self._measure_column_width(col, rows)
            anchor = 'center' if col in ('age_since_visit', 'visit_date', 'visit_time', 'datetime', 'phone') else 'w'
            self.tree.column(col, width=width, anchor=anchor)
            self.schedule_save_column_widths()
        except Exception:
            pass

    def auto_adjust_columns(self, tree: Optional[ttk.Treeview]=None) -> None:
        """Automatically adjust visible Treeview column widths with min/max limits and remember them."""
        try:
            tree = tree or self.tree
            visible = list(getattr(self, 'visible_columns', []) or self._load_visible_columns())
            visible = [c for c in visible if c in getattr(self, 'display_headers', self.headers)]
            if not visible:
                visible = [c for c in getattr(self, 'display_headers', self.headers) if c != 'id']
            rows = [self._get_row_for_item(item) for item in tree.get_children()]
            rows = [r for r in rows if r]
            if not rows:
                rows = list(getattr(self, 'current_view_data', []) or self.data)
            for col in visible:
                anchor = 'center' if col in ('age_since_visit', 'visit_date', 'visit_time', 'datetime', 'phone') else 'w'
                tree.column(col, width=self._measure_column_width(col, rows), anchor=anchor)
            self.schedule_save_column_widths()
        except Exception:
            pass

    def _refresh_heading_labels(self) -> None:
        """Refresh heading labels, including sort arrows."""
        try:
            for h in getattr(self, 'display_headers', self.headers):
                self.tree.heading(h, text=self._heading_text(h), command=lambda col=h: self.sort_tree_by_column(col))
        except Exception:
            pass

    def _sort_key_for_row(self, row: Dict, col: str):
        """Stable sort key for Treeview columns."""
        if col in ('age_since_visit', 'visit_date', 'visit_time', 'datetime'):
            return parse_visit_datetime(row.get('datetime', '')) or datetime.min
        value = str(self._display_value_for_column(row, col) or '').strip()
        if col == 'phone':
            digits = _phone_digits(value)
            return (0, int(digits)) if digits else (1, value.lower())
        parts = re.split('(\\d+)', value.lower())
        return [int(part) if part.isdigit() else part for part in parts]

    def sort_tree_by_column(self, col: str) -> None:
        """Sort the currently visible Treeview rows by a column."""
        try:
            if col not in getattr(self, 'display_headers', self.headers):
                return
            if getattr(self, 'sort_column', None) == col:
                self.sort_reverse = not getattr(self, 'sort_reverse', False)
            else:
                self.sort_column = col
                self.sort_reverse = col in ('age_since_visit', 'visit_date', 'visit_time', 'datetime')
            items = list(self.tree.get_children())
            items.sort(key=lambda item: self._sort_key_for_row(self._get_row_for_item(item) or {}, col), reverse=self.sort_reverse)
            for index, item in enumerate(items):
                self.tree.move(item, '', index)
            self._refresh_heading_labels()
            self.refresh_current_tree_flag_colors()
        except Exception:
            pass

    def _hide_tree_column(self, col: str) -> None:
        """Hide one visible column and save the preference."""
        try:
            visible = [c for c in self._visible_tree_columns() if c in getattr(self, 'display_headers', self.headers)]
            if col not in visible or len(visible) <= 1:
                return
            self.save_current_column_widths()
            visible.remove(col)
            settings = load_settings()
            settings['bike_list_visible_columns'] = visible
            save_settings(settings)
            self.visible_columns = visible
            self.tree['displaycolumns'] = visible
            self._refresh_heading_labels()
            self._apply_saved_or_measured_column_widths(self.tree, save_missing=True)
        except Exception:
            pass

    def show_header_menu(self, event: tk.Event) -> None:
        """Right-click menu for Treeview column headers."""
        col = self._column_from_event_x(event.x)
        if not col:
            return
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label='Autofit this column', command=lambda c=col: self.auto_fit_column(c))
        menu.add_command(label='Hide this column', command=lambda c=col: self._hide_tree_column(c))
        menu.add_separator()
        menu.add_command(label='Autofit all columns', command=lambda: self.auto_adjust_columns(self.tree))
        menu.add_command(label='Columns...', command=self.open_column_visibility_dialog)
        menu.post(event.x_root, event.y_root)

    def populate_treeview(self, tree: ttk.Treeview, data: List[Dict]) -> None:
        """Populate the Treeview with data (chunked inserts to keep UI responsive)."""
        text = ENGLISH_TEXT
        try:
            if tree is getattr(self, 'tree', None):
                self.current_view_data = list(data)
            tree.tag_configure('evenrow', background='#f7f9fc')
            tree.tag_configure('oddrow', background='white')
            tree.tag_configure('hoverrow', background='#eaf3ff')
            for key, info in CUSTOMER_FLAG_LEVELS.items():
                tree.tag_configure(f'flag_{key}', background=info['color'])
            flag_map = get_all_customer_flags()
            for i in tree.get_children():
                tree.delete(i)
            self.tree_item_rows = {}
            batch = 500
            for start in range(0, len(data), batch):
                chunk = data[start:start + batch]
                for i, row in enumerate(chunk, start=start):
                    values = self._tree_values_for_row(row)
                    phone_key = normalize_flag_phone(row.get('phone', ''))
                    flag = flag_map.get(phone_key)
                    if flag:
                        level = str(flag['flag_level'] or 'watch')
                        tag = f'flag_{level}' if level in CUSTOMER_FLAG_LEVELS else 'flag_watch'
                    else:
                        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                    item_id = tree.insert('', 'end', values=values, tags=(tag,))
                    self._remember_item_row(item_id, row)
                tree.update_idletasks()
            self._apply_saved_or_measured_column_widths(tree, save_missing=True)
        except Exception as e:
            messagebox.showerror(text['error_title'], text['treeview_error'].format(str(e)))

    def refresh_treeview(self) -> None:
        """Refresh the Treeview with updated data."""
        text = ENGLISH_TEXT
        self.loading_label = tk.Label(self.root, text=text['loading_message'], fg='blue')
        self.loading_label.pack(pady=10)
        self.root.update()
        for widget in self.search_frame.winfo_children():
            if isinstance(widget, (tk.Button, ttk.Button)):
                widget.config(state='disabled')
        threading.Thread(target=self.threaded_refresh, daemon=True).start()

    def threaded_refresh(self) -> None:
        """Thread function for refreshing data."""
        headers, data = self.load_db_data()
        self.root.after(0, lambda: self.finish_refresh(headers, data))

    def finish_refresh(self, headers, data) -> None:
        """Complete refresh by updating UI in the main thread."""
        self.headers = headers
        self.display_headers = self._make_display_headers(headers)
        self.data = data
        if self.current_phone_history:
            self.show_phone_history(self.current_phone_history, update_entry=True, show_if_empty=True)
        else:
            self.show_all_records(clear_phone_entry=False)
        if hasattr(self, 'loading_label') and self.loading_label.winfo_exists():
            self.loading_label.destroy()
        for widget in self.search_frame.winfo_children():
            if isinstance(widget, (tk.Button, ttk.Button)):
                widget.config(state='normal')

    def _set_bike_list_buttons_state(self, state: str) -> None:
        """Enable/disable every button inside the bike-list toolbar, including nested groups."""
        try:
            if not hasattr(self, 'search_frame') or not self.search_frame.winfo_exists():
                return

            def walk(widget):
                for child in widget.winfo_children():
                    if isinstance(child, (tk.Button, ttk.Button)):
                        try:
                            child.config(state=state)
                        except Exception:
                            pass
                    walk(child)
            walk(self.search_frame)
        except Exception:
            pass

    def _copy_record_row(self, row: Dict) -> Dict:
        """Make a safe copy of one database row, including the hidden database id."""
        copied = {'_db_id': row.get('_db_id')}
        for col in self.headers:
            copied[col] = row.get(col, '')
        return copied

    def _row_match_params(self, row: Dict) -> Tuple[str, Tuple]:
        """Build a strict fallback WHERE clause for records without a database id."""
        issues = [row.get(f'issue{i}', '') or '' for i in range(1, 11)]
        where = (
            "datetime=? AND phone=? AND model=? "
            "AND COALESCE(issue1,'')=? AND COALESCE(issue2,'')=? AND COALESCE(issue3,'')=? "
            "AND COALESCE(issue4,'')=? AND COALESCE(issue5,'')=? AND COALESCE(issue6,'')=? "
            "AND COALESCE(issue7,'')=? AND COALESCE(issue8,'')=? AND COALESCE(issue9,'')=? "
            "AND COALESCE(issue10,'')=?"
        )
        params = (row.get('datetime', ''), row.get('phone', ''), row.get('model', ''), *issues[:10])
        return where, params

    def _clear_bike_cache(self) -> None:
        """Force the next bike-list reload to read fresh database data."""
        try:
            _app_cache.pop('bike_data', None)
            _app_cache.pop('bike_headers', None)
            _app_cache.pop('bike_data_time', None)
        except Exception:
            pass

    def _delete_row_from_database(self, row: Dict) -> Tuple[bool, str]:
        """Delete exactly one row from the database without rewriting the whole table."""
        text = ENGLISH_TEXT
        deadline = time.monotonic() + DB_LOCK_MAX_WAIT_SECONDS
        last_error = None
        while True:
            conn = None
            try:
                self.backup_db()
                conn = get_db_connection()
                conn.execute(f'PRAGMA busy_timeout={DB_LOCK_RETRY_BUSY_TIMEOUT_MS}')
                cursor = conn.cursor()
                db_id = row.get('_db_id')
                if db_id not in (None, ''):
                    cursor.execute('DELETE FROM customer_bikes WHERE id=?', (db_id,))
                else:
                    where, params = self._row_match_params(row)
                    cursor.execute(f'DELETE FROM customer_bikes WHERE rowid IN (SELECT rowid FROM customer_bikes WHERE {where} LIMIT 1)', params)
                if cursor.rowcount <= 0:
                    conn.rollback()
                    return False, 'The selected record was not found. Refresh Bike List and try again.'
                conn.commit()
                self._clear_bike_cache()
                return True, ''
            except sqlite3.OperationalError as e:
                last_error = e
                try:
                    if conn:
                        conn.rollback()
                except Exception:
                    pass
                if is_database_locked_error(e) and time.monotonic() < deadline:
                    time.sleep(DB_LOCK_RETRY_DELAY_SECONDS)
                    continue
                return False, text['db_error_message'].format(str(last_error))
            except Exception as e:
                try:
                    if conn:
                        conn.rollback()
                except Exception:
                    pass
                return False, str(e)
            finally:
                if conn:
                    try:
                        conn.execute(f'PRAGMA busy_timeout={DB_DEFAULT_BUSY_TIMEOUT_MS}')
                    except Exception:
                        pass
                    release_db_connection(conn)

    def _insert_deleted_row_to_database(self, row: Dict) -> Tuple[bool, str]:
        """Restore one deleted row for Undo."""
        text = ENGLISH_TEXT
        deadline = time.monotonic() + DB_LOCK_MAX_WAIT_SECONDS
        last_error = None
        base_columns = ['datetime', 'phone', 'model'] + [f'issue{i}' for i in range(1, 11)] + ['shop_note']
        base_values = [row.get(col, '') or '' for col in base_columns]
        while True:
            conn = None
            try:
                conn = get_db_connection()
                conn.execute(f'PRAGMA busy_timeout={DB_LOCK_RETRY_BUSY_TIMEOUT_MS}')
                cursor = conn.cursor()
                db_id = row.get('_db_id')
                if db_id not in (None, ''):
                    try:
                        placeholders = ','.join(['?'] * (len(base_columns) + 1))
                        cursor.execute(
                            f"INSERT INTO customer_bikes (id,{','.join(base_columns)}) VALUES ({placeholders})",
                            [db_id] + base_values
                        )
                    except sqlite3.IntegrityError:
                        placeholders = ','.join(['?'] * len(base_columns))
                        cursor.execute(
                            f"INSERT INTO customer_bikes ({','.join(base_columns)}) VALUES ({placeholders})",
                            base_values
                        )
                else:
                    placeholders = ','.join(['?'] * len(base_columns))
                    cursor.execute(
                        f"INSERT INTO customer_bikes ({','.join(base_columns)}) VALUES ({placeholders})",
                        base_values
                    )
                conn.commit()
                self._clear_bike_cache()
                return True, ''
            except sqlite3.OperationalError as e:
                last_error = e
                try:
                    if conn:
                        conn.rollback()
                except Exception:
                    pass
                if is_database_locked_error(e) and time.monotonic() < deadline:
                    time.sleep(DB_LOCK_RETRY_DELAY_SECONDS)
                    continue
                return False, text['db_error_message'].format(str(last_error))
            except Exception as e:
                try:
                    if conn:
                        conn.rollback()
                except Exception:
                    pass
                return False, str(e)
            finally:
                if conn:
                    try:
                        conn.execute(f'PRAGMA busy_timeout={DB_DEFAULT_BUSY_TIMEOUT_MS}')
                    except Exception:
                        pass
                    release_db_connection(conn)

    def _restore_edited_row_to_database(self, before_row: Dict, after_row: Dict) -> Tuple[bool, str]:
        """Undo a no-print edit by restoring the old values."""
        text = ENGLISH_TEXT
        deadline = time.monotonic() + DB_LOCK_MAX_WAIT_SECONDS
        last_error = None
        set_columns = ['phone', 'model'] + [f'issue{i}' for i in range(1, 11)] + ['shop_note']
        set_values = [before_row.get(col, '') or '' for col in set_columns]
        while True:
            conn = None
            try:
                self.backup_db()
                conn = get_db_connection()
                conn.execute(f'PRAGMA busy_timeout={DB_LOCK_RETRY_BUSY_TIMEOUT_MS}')
                cursor = conn.cursor()
                db_id = before_row.get('_db_id') or after_row.get('_db_id')
                if db_id not in (None, ''):
                    cursor.execute(
                        "UPDATE customer_bikes SET phone=?, model=?, issue1=?, issue2=?, issue3=?, issue4=?, issue5=?, "
                        "issue6=?, issue7=?, issue8=?, issue9=?, issue10=?, shop_note=? WHERE id=?",
                        tuple(set_values) + (db_id,)
                    )
                else:
                    where, params = self._row_match_params(after_row)
                    cursor.execute(
                        "UPDATE customer_bikes SET phone=?, model=?, issue1=?, issue2=?, issue3=?, issue4=?, issue5=?, "
                        f"issue6=?, issue7=?, issue8=?, issue9=?, issue10=?, shop_note=? WHERE rowid IN (SELECT rowid FROM customer_bikes WHERE {where} LIMIT 1)",
                        tuple(set_values) + tuple(params)
                    )
                if cursor.rowcount <= 0:
                    conn.rollback()
                    return False, 'The edited record was not found. Refresh Bike List and try again.'
                conn.commit()
                self._clear_bike_cache()
                return True, ''
            except sqlite3.OperationalError as e:
                last_error = e
                try:
                    if conn:
                        conn.rollback()
                except Exception:
                    pass
                if is_database_locked_error(e) and time.monotonic() < deadline:
                    time.sleep(DB_LOCK_RETRY_DELAY_SECONDS)
                    continue
                return False, text['db_error_message'].format(str(last_error))
            except Exception as e:
                try:
                    if conn:
                        conn.rollback()
                except Exception:
                    pass
                return False, str(e)
            finally:
                if conn:
                    try:
                        conn.execute(f'PRAGMA busy_timeout={DB_DEFAULT_BUSY_TIMEOUT_MS}')
                    except Exception:
                        pass
                    release_db_connection(conn)

    def _reload_data_after_change(self) -> Tuple[List[str], List[Dict]]:
        """Reload bike data after a delete/edit/undo operation."""
        self._clear_bike_cache()
        return self.load_db_data()

    def _apply_reloaded_data(self, headers: List[str], data: List[Dict]) -> None:
        """Apply freshly loaded data to the current bike-list view."""
        if headers:
            self.headers = headers
            self.display_headers = self._make_display_headers(headers)
        self.data = data or []
        if getattr(self, 'current_phone_history', ''):
            self.show_phone_history(self.current_phone_history, update_entry=True, show_if_empty=True)
        else:
            self.show_all_records(clear_phone_entry=False)

    def delete_record(self) -> None:
        """Delete a selected record and keep an undo snapshot."""
        text = ENGLISH_TEXT
        selected_items = self.tree.selection()
        if not selected_items or len(selected_items) > 1:
            app_info('Info', text['delete_info'], self.root)
            return
        if not messagebox.askyesno(text['confirm_delete_title'], text['confirm_delete_message']):
            return
        row = self._get_row_for_item(selected_items[0])
        if not row:
            app_info('Info', text['select_record_info'], self.root)
            return
        row_snapshot = self._copy_record_row(row)
        self.delete_loading_label = tk.Label(self.root, text='Deleting...', fg='blue')
        self.delete_loading_label.pack(pady=10)
        self.root.update()
        self._set_bike_list_buttons_state(tk.DISABLED)
        threading.Thread(target=lambda: self.threaded_delete_record(row_snapshot), daemon=True).start()

    def threaded_delete_record(self, row_snapshot: Dict) -> None:
        """Worker for deleting one record."""
        success, error = self._delete_row_from_database(row_snapshot)
        headers, data = self._reload_data_after_change() if success else (self.headers, self.data)
        self.root.after(0, lambda: self.finish_delete_record(success, error, row_snapshot, headers, data))

    def finish_delete_record(self, success: bool, error: str, row_snapshot: Dict, headers: List[str], data: List[Dict]) -> None:
        """Finish a delete operation in the UI thread."""
        if hasattr(self, 'delete_loading_label') and self.delete_loading_label.winfo_exists():
            self.delete_loading_label.destroy()
        self._set_bike_list_buttons_state(tk.NORMAL)
        if not success:
            messagebox.showerror(ENGLISH_TEXT['error_title'], error or 'Delete failed.')
            return
        self._apply_reloaded_data(headers, data)
        self.undo_buffer = {'action': 'delete', 'row': row_snapshot}
        try:
            self.undo_button.config(state=tk.NORMAL)
        except Exception:
            pass
        try:
            show_toast(self.root, ENGLISH_TEXT['delete_success'], 2000)
        except Exception:
            pass

    def register_edit_undo_from_main(self, before_state: Dict[str, Any], phone: str, model: str, issues: List[str], shop_note: str='') -> None:
        """Register a main-window edit so the Bike List Undo button can restore it."""
        try:
            before_row = {'_db_id': before_state.get('db_id'), 'datetime': before_state.get('datetime', '')}
            before_row['phone'] = before_state.get('phone', '')
            before_row['model'] = before_state.get('model', '')
            before_issues = list(before_state.get('issues') or [])
            before_issues.extend([''] * (10 - len(before_issues)))
            for i in range(1, 11):
                before_row[f'issue{i}'] = before_issues[i - 1]
            before_row['shop_note'] = before_state.get('shop_note', '') or ''
            after_row = {'_db_id': before_state.get('db_id'), 'datetime': before_state.get('datetime', ''), 'phone': phone, 'model': model, 'shop_note': shop_note}
            after_issues = list(issues or [])
            after_issues.extend([''] * (10 - len(after_issues)))
            for i in range(1, 11):
                after_row[f'issue{i}'] = after_issues[i - 1]
            self.undo_buffer = {'action': 'edit', 'before': before_row, 'after': after_row}
            if hasattr(self, 'undo_button'):
                self.undo_button.config(state=tk.NORMAL)
        except Exception:
            pass

    def undo_last_action(self) -> None:
        """Undo the last delete or no-print edit."""
        if not self.undo_buffer:
            return
        undo_snapshot = dict(self.undo_buffer)
        self.undo_loading_label = tk.Label(self.root, text='Undoing...', fg='blue')
        self.undo_loading_label.pack(pady=10)
        self.root.update()
        self._set_bike_list_buttons_state(tk.DISABLED)
        threading.Thread(target=lambda: self.threaded_undo(undo_snapshot), daemon=True).start()

    def threaded_undo(self, undo_snapshot: Dict) -> None:
        """Restore the last deleted or edited record in a worker thread."""
        action = undo_snapshot.get('action')
        if action == 'delete':
            success, error = self._insert_deleted_row_to_database(undo_snapshot.get('row') or {})
        elif action == 'edit':
            success, error = self._restore_edited_row_to_database(undo_snapshot.get('before') or {}, undo_snapshot.get('after') or {})
        else:
            success, error = False, 'Nothing to undo.'
        headers, data = self._reload_data_after_change() if success else (self.headers, self.data)
        self.root.after(0, lambda: self.finish_undo(success, error, headers, data))

    def finish_undo(self, success: bool, error: str, headers: List[str], data: List[Dict]) -> None:
        """Clean up after undo operation."""
        if hasattr(self, 'undo_loading_label') and self.undo_loading_label.winfo_exists():
            self.undo_loading_label.destroy()
        self._set_bike_list_buttons_state(tk.NORMAL)
        if not success:
            messagebox.showerror(ENGLISH_TEXT['error_title'], error or 'Undo failed.')
            return
        self._apply_reloaded_data(headers, data)
        app_info(ENGLISH_TEXT['undo_title'], ENGLISH_TEXT['undo_message'], self.root)
        try:
            self.undo_button.config(state=tk.DISABLED)
        except Exception:
            pass
        self.undo_buffer = None

    def search_similar_phone_numbers(self, target_phone: str, threshold: int=2) -> List[Dict]:
        """Search for phone numbers similar to the target (prefix pre-filter for speed)."""
        prefix = target_phone[:3] if target_phone else ''
        if prefix:
            candidates = [row for row in self.data if str(row.get('phone', '')).startswith(prefix)]
        else:
            candidates = self.data[:1000]
        return [row for row in candidates if Levenshtein.distance(str(row.get('phone', '')), target_phone) <= threshold]

    def on_search(self) -> None:
        """Search for similar phone numbers."""
        text = ENGLISH_TEXT
        target_phone = self.phone_entry.get().strip()
        if not target_phone:
            app_info('Info', text['search_info'])
            return
        self.loading_label = tk.Label(self.root, text=text['searching_message'], fg='blue')
        self.loading_label.pack(pady=10)
        self.root.update()
        for widget in self.search_frame.winfo_children():
            if isinstance(widget, (tk.Button, ttk.Button)):
                widget.config(state='disabled')
        threading.Thread(target=lambda: self.threaded_search(target_phone), daemon=True).start()

    def threaded_search(self, target_phone) -> None:
        """Perform search in a separate thread."""
        similar_phone_data = self.search_similar_phone_numbers(target_phone)
        self.root.after(0, lambda: self.finish_search(similar_phone_data))

    def finish_search(self, similar_phone_data) -> None:
        """Complete search by updating UI in the main thread with highlighted results."""
        self.history_mode = False
        self.current_phone_history = ''
        self.current_view_data = list(similar_phone_data)
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree_item_rows = {}
        search_phone = self.phone_entry.get().strip()
        flag_map = get_all_customer_flags()
        for key, info in CUSTOMER_FLAG_LEVELS.items():
            self.tree.tag_configure(f'flag_{key}', background=info['color'])
        self.tree.tag_configure('evenrow', background='#f7f9fc')
        self.tree.tag_configure('oddrow', background='white')
        self.tree.tag_configure('hoverrow', background='#eaf3ff')
        for index, row in enumerate(similar_phone_data):
            values = self._tree_values_for_row(row)
            base_tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            item_id = self.tree.insert('', 'end', values=values, tags=(base_tag,))
            self._remember_item_row(item_id, row)
            phone_value = str(row.get('phone', ''))
            phone_key = normalize_flag_phone(phone_value)
            flag = flag_map.get(phone_key)
            if flag:
                level = str(flag['flag_level'] or 'watch')
                tag = f'flag_{level}' if level in CUSTOMER_FLAG_LEVELS else 'flag_watch'
                self.tree.item(item_id, tags=(tag,))
            elif phone_value == search_phone:
                self.tree.tag_configure('exact_match', background='#90EE90')
                self.tree.item(item_id, tags=('exact_match',))
            elif Levenshtein.distance(phone_value, search_phone) <= 2:
                self.tree.tag_configure('similar_match', background='#FFFF99')
                self.tree.item(item_id, tags=('similar_match',))
        self.auto_adjust_columns(self.tree)
        try:
            self.history_status_var.set(f'Search result: {len(similar_phone_data)} records')
        except Exception:
            pass
        if hasattr(self, 'loading_label') and self.loading_label.winfo_exists():
            self.loading_label.destroy()
        for widget in self.search_frame.winfo_children():
            if isinstance(widget, (tk.Button, ttk.Button)):
                widget.config(state='normal')

    def search_from_context(self) -> None:
        """Search using the selected phone number."""
        text = ENGLISH_TEXT
        selected_item = self.tree.selection()
        if not selected_item:
            app_info('Info', text['select_record_info'])
            return
        phone_number = self.get_selected_phone_from_tree()
        self.phone_entry.delete(0, tk.END)
        self.phone_entry.insert(0, phone_number)
        self.on_search()

    def copy_selected_bike_model(self) -> None:
        """Copy the selected row's bike/model text to the clipboard."""
        try:
            row = self._get_selected_row()
            value = str(row.get('model', '')).strip() if row else ''
            if not value:
                return
            self.root.clipboard_clear()
            self.root.clipboard_append(value)
            show_toast(self.root, 'Bike copied.', 1800)
        except Exception:
            pass

    def get_selected_phone_from_tree(self) -> str:
        """Return the phone number from the selected row."""
        selected_item = self.tree.selection()
        if not selected_item:
            return ''
        row = self._get_selected_row()
        if not row:
            return ''
        return str(row.get('phone', '')).strip()

    def refresh_current_tree_flag_colors(self) -> None:
        """Refresh only the visible Treeview row colors after mark changes."""
        try:
            for key, info in CUSTOMER_FLAG_LEVELS.items():
                self.tree.tag_configure(f'flag_{key}', background=info['color'])
            self.tree.tag_configure('evenrow', background='#f7f9fc')
            self.tree.tag_configure('oddrow', background='white')
            self.tree.tag_configure('hoverrow', background='#eaf3ff')
            flag_map = get_all_customer_flags()
            for index, item in enumerate(self.tree.get_children()):
                row = self._get_row_for_item(item)
                phone = str(row.get('phone', '')).strip() if row else ''
                flag = flag_map.get(normalize_flag_phone(phone))
                if flag:
                    level = str(flag['flag_level'] or 'watch')
                    tag = f'flag_{level}' if level in CUSTOMER_FLAG_LEVELS else 'flag_watch'
                else:
                    tag = 'evenrow' if index % 2 == 0 else 'oddrow'
                self.tree.item(item, tags=(tag,))
        except Exception:
            pass

    def mark_selected_customer(self, flag_level: str) -> None:
        """Mark the selected row's phone number from the right-click menu."""
        phone = self.get_selected_phone_from_tree()
        if not normalize_flag_phone(phone):
            return
        existing = get_customer_flag(phone)
        note = str(existing['note'] or '') if existing else ''
        if save_customer_flag(phone, flag_level, note):
            self.refresh_current_tree_flag_colors()
            try:
                check_customer_flag_warning()
            except Exception:
                pass

    def manage_selected_customer_mark(self) -> None:
        """Open the marked-number manager with the selected phone prefilled."""
        phone = self.get_selected_phone_from_tree()
        if not normalize_flag_phone(phone):
            return
        open_customer_flags_list_dialog(self.root, initial_phone=phone)

    def delete_selected_customer_mark(self) -> None:
        """Remove the mark for the selected row's phone number from the right-click menu."""
        phone = self.get_selected_phone_from_tree()
        if not normalize_flag_phone(phone):
            return
        if delete_customer_flag(phone):
            self.refresh_current_tree_flag_colors()
            try:
                check_customer_flag_warning()
            except Exception:
                pass

    def on_refresh(self) -> None:
        """Refresh the Treeview and clear search."""
        text = ENGLISH_TEXT
        self.loading_label = tk.Label(self.root, text=text['loading_message'], fg='blue')
        self.loading_label.pack(pady=10)
        self.root.update()
        for widget in self.search_frame.winfo_children():
            if isinstance(widget, (tk.Button, ttk.Button)):
                widget.config(state='disabled')
        threading.Thread(target=self.threaded_on_refresh, daemon=True).start()

    def threaded_on_refresh(self) -> None:
        """Thread function for full refresh with search clearing."""
        headers, data = self.load_db_data()
        self.root.after(0, lambda: self.finish_on_refresh(headers, data))

    def finish_on_refresh(self, headers, data) -> None:
        """Complete refresh and clear search in the main thread."""
        self.headers = headers
        self.display_headers = self._make_display_headers(headers)
        self.data = data
        self.show_all_records(clear_phone_entry=True)
        if hasattr(self, 'loading_label') and self.loading_label.winfo_exists():
            self.loading_label.destroy()
        for widget in self.search_frame.winfo_children():
            if isinstance(widget, (tk.Button, ttk.Button)):
                widget.config(state='normal')

    def generate_sms_qr(self, phone_number: str, sms_text: str) -> Optional[Image.Image]:
        """Generate an SMS QR code."""
        text = ENGLISH_TEXT
        sms_uri = f'sms:{phone_number}?body={sms_text}'
        qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=10, border=4)
        qr.add_data(sms_uri)
        qr.make(fit=True)
        try:
            return qr.make_image(fill='black', back_color='white')
        except Exception as e:
            messagebox.showerror(text['error_title'], text['qr_error'].format(str(e)))
            return None

    def display_qr_code(self, qr_image: Image.Image) -> None:
        """Display the QR code in a new window."""
        text = ENGLISH_TEXT
        if self.qr_window and self.qr_window.winfo_exists():
            self.qr_window.destroy()
        self.qr_window = tk.Toplevel(self.root)
        self.qr_window.title(text['qr_title'])
        self.qr_window.geometry(f'{self.QR_SIZE[0] + 40}x{self.QR_SIZE[1] + 100}')
        self.qr_window.resizable(False, False)
        set_window_icon(self.qr_window)
        selected_item = self.tree.selection()
        if selected_item:
            row = self._get_row_for_item(selected_item[0])
            registration_time = parse_visit_datetime(row.get('datetime', '')) if row else None
            if not registration_time:
                registration_time = datetime.now()
            current_time = datetime.now()
            time_elapsed = current_time - registration_time
            days = time_elapsed.days
            hours = time_elapsed.seconds // 3600
            if days > 0:
                time_text = f"{days} {'days'} "
                if hours > 0:
                    time_text += f"{hours} {'hours'}"
            else:
                time_text = f"{hours} {'hours'}"
            time_label = tk.Label(self.qr_window, text=f"{'Time since registration:'} {time_text}", font=('Arial', 10))
            time_label.pack(pady=5)
        img = qr_image.resize(self.QR_SIZE, Image.Resampling.LANCZOS)
        qr_img = ImageTk.PhotoImage(img)
        label = tk.Label(self.qr_window, image=qr_img)
        label.image = qr_img
        label.pack(expand=True)
        tk.Button(self.qr_window, text=text['print_qr_button'], command=lambda: self.print_qr_code(qr_image)).pack(pady=10)

    def image_to_escpos(self, image: Image.Image, values: Tuple, target_width: int) -> bytes:
        """Convert image and text to ESC/POS for printing."""
        im = image.convert('1')
        if im.width != target_width:
            ratio = target_width / im.width
            new_height = int(im.height * ratio)
            im = im.resize((target_width, new_height), Image.Resampling.LANCZOS)
        width, height = im.size
        width_bytes = (width + 7) // 8
        buffer = io.BytesIO()
        buffer.write(ESC_INIT + ESC_CODEPAGE_CP850)
        buffer.write(ALIGN_CENTER)
        buffer.write(b'\x1dv0\x00')
        buffer.write(width_bytes.to_bytes(2, 'little'))
        buffer.write(height.to_bytes(2, 'little'))
        for y in range(height):
            for x in range(0, width, 8):
                byte = 0
                for bit in range(8):
                    if x + bit < width and im.getpixel((x + bit, y)) == 0:
                        byte |= 1 << 7 - bit
                buffer.write(bytes([byte]))
        buffer.write(b'\n')
        buffer.write(b'\n')
        for i, header in enumerate(self.headers):
            if i < len(values):
                label = escpos_text(f'{header}: ')
                value = escpos_text(values[i] if values[i] else 'Not specified')
                buffer.write(label)
                buffer.write(value)
                buffer.write(b'\n')
        settings = load_settings()
        empty_lines = settings.get('empty_lines', DEFAULT_SETTINGS['empty_lines'])
        for _ in range(empty_lines):
            buffer.write(b'\n')
        buffer.write(ALIGN_LEFT)
        buffer.write(b'\x1dV\x01')
        return buffer.getvalue()

    def print_qr_code(self, qr_image: Image.Image) -> None:
        """Print the QR code and associated data through the Windows printer driver/GDI."""
        text = ENGLISH_TEXT
        settings = load_settings()
        printer_name = settings.get('shop_printer') or settings.get('printer') or settings.get('customer_printer')
        if not printer_name:
            messagebox.showwarning(text['printer_warning_title'], text['printer_warning_message'])
            return
        selected_item = self.tree.selection()
        if not selected_item:
            app_info('Info', text['select_record_info'])
            return
        row = self._get_selected_row()
        if not row:
            app_info('Info', text['select_record_info'])
            return
        values = self._db_values_for_row(row)
        try:
            qr_print_image = build_qr_record_image_gdi(qr_image, self.headers, values, settings)
            print_images_with_windows_driver(printer_name, [qr_print_image], doc_name='BikeTrack QR Code')
            app_info(text['print_success_title'], text['print_success_message'])
        except Exception as e:
            messagebox.showerror(text['error_title'], text['print_error'].format(str(e)))

    def on_treeview_double_click(self, event: tk.Event) -> None:
        """Double-click header/separator = autofit; double-click row = normal row action."""
        region = self.tree.identify_region(event.x, event.y)
        if region in ('heading', 'separator'):
            col = self._column_from_event_x(event.x)
            if col:
                self.auto_fit_column(col)
            return 'break'
        text = ENGLISH_TEXT
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
        selected_item = self.tree.selection()
        if not selected_item:
            app_info('Info', text['select_phone_info'])
            return
        if getattr(self, 'history_mode', False):
            self.use_selected_bike_in_main_form()
            return
        phone_number = self.get_selected_phone_from_tree()
        sms_text = self.load_sms_message()
        qr_image = self.generate_sms_qr(phone_number, sms_text)
        if qr_image:
            self.display_qr_code(qr_image)

    def generate_qr_from_selection(self) -> None:
        """Generate QR code from the selected record."""
        text = ENGLISH_TEXT
        selected_item = self.tree.selection()
        if not selected_item:
            app_info('Info', text['select_record_info'])
            return
        phone_number = self.get_selected_phone_from_tree()
        sms_text = self.load_sms_message()
        qr_image = self.generate_sms_qr(phone_number, sms_text)
        if qr_image:
            self.display_qr_code(qr_image)

    def copy_phone_number(self) -> None:
        """Copy the selected phone number to clipboard."""
        text = ENGLISH_TEXT
        selected_item = self.tree.selection()
        if not selected_item:
            app_info('Info', text['select_record_info'])
            return
        phone_number = self.get_selected_phone_from_tree()
        self.root.clipboard_clear()
        self.root.clipboard_append(phone_number)
        app_info(text['print_success_title'], text['copy_success'])

    def open_settings(self) -> None:
        """Open settings window for BikeShopApp."""
        text = ENGLISH_TEXT

        def save_settings_changes():
            message = message_text.get('1.0', tk.END).strip()
            printer = printer_combo.get()
            try:
                lines = int(empty_lines_entry.get())
                if lines < 0:
                    raise ValueError(text['lines_error'])
                font_size = int(font_size_entry.get())
                if font_size < 6 or font_size > 20:
                    raise ValueError('Font size must be between 6 and 20')
            except ValueError as e:
                messagebox.showerror(text['error_title'], str(e) if str(e) != text['lines_error'] else text['number_error'])
                return
            settings = load_settings()
            settings['sms_message'] = message
            local_settings = load_local_settings()
            local_settings['shop_printer'] = printer if printer else None
            local_settings['printer'] = printer if printer else None
            if not local_settings.get('customer_printer'):
                local_settings['customer_printer'] = printer if printer else None
            save_local_settings(local_settings)
            settings['shop_printer'] = printer if printer else None
            settings['printer'] = printer if printer else None
            settings['empty_lines'] = lines
            settings['treeview_font_size'] = font_size
            save_settings(settings)
            self.treeview_font_size = font_size
            style = ttk.Style()
            style.configure('Treeview', font=('Arial', font_size))
            style.configure('Treeview.Heading', font=('Arial', font_size, 'bold'))
            self.refresh_treeview()
            save_window_geometry(settings_window, 'bike_list_settings')
            settings_window.destroy()
            app_info(text['bike_saved_title'], text['bike_saved_message'])
        settings_window = tk.Toplevel(self.root)
        settings_window.title(text['bike_settings_title'])
        settings_window.geometry(get_saved_window_geometry('bike_list_settings', '400x400'))
        set_window_icon(settings_window)
        settings = load_settings()
        tk.Label(settings_window, text=text['sms_label']).pack(pady=5)
        message_text = tk.Text(settings_window, wrap='word', height=5)
        message_text.pack(padx=20, pady=5, fill='both', expand=True)
        message_text.insert(tk.END, settings.get('sms_message', DEFAULT_SETTINGS['sms_message']))
        tk.Label(settings_window, text=text['bike_printer_label']).pack(pady=5)
        printer_combo = ttk.Combobox(settings_window, values=list_printers())
        printer_combo.pack(padx=20, pady=5)
        selected_qr_printer = settings.get('shop_printer') or settings.get('printer') or settings.get('customer_printer')
        if selected_qr_printer in printer_combo['values']:
            printer_combo.set(selected_qr_printer)
        tk.Label(settings_window, text=text['empty_lines_label']).pack(pady=5)
        empty_lines_entry = tk.Entry(settings_window)
        empty_lines_entry.pack(padx=20, pady=5)
        empty_lines_entry.insert(0, str(settings.get('empty_lines', DEFAULT_SETTINGS['empty_lines'])))
        tk.Label(settings_window, text='Treeview Font Size:').pack(pady=5)
        font_size_entry = tk.Entry(settings_window)
        font_size_entry.pack(padx=20, pady=5)
        font_size_entry.insert(0, str(settings.get('treeview_font_size', DEFAULT_SETTINGS['treeview_font_size'])))
        tk.Button(settings_window, text=text['save_button'], command=save_settings_changes).pack(pady=10)
        settings_window.protocol('WM_DELETE_WINDOW', lambda: (save_window_geometry(settings_window, 'bike_list_settings'), settings_window.destroy()))

    def _restore_hover_row(self) -> None:
        """Restore tags for the row that previously had hover styling."""
        try:
            item = getattr(self, '_hover_item', None)
            if item and self.tree.exists(item):
                original = self._hover_original_tags.get(item)
                if original is not None:
                    self.tree.item(item, tags=original)
            self._hover_item = None
            self._hover_original_tags = {}
        except Exception:
            pass

    def on_tree_motion(self, event: tk.Event) -> None:
        """Light hover effect for a more modern list feel."""
        try:
            if self.tree.identify_region(event.x, event.y) != 'cell':
                self._restore_hover_row()
                return
            item = self.tree.identify_row(event.y)
            if item == getattr(self, '_hover_item', None):
                return
            self._restore_hover_row()
            if item and self.tree.exists(item):
                original = tuple(self.tree.item(item, 'tags'))
                self._hover_original_tags[item] = original
                self.tree.item(item, tags=original + ('hoverrow',))
                self._hover_item = item
        except Exception:
            pass

    def on_tree_leave(self, event: Optional[tk.Event]=None) -> None:
        """Remove hover styling when the pointer leaves the Treeview."""
        self._restore_hover_row()

    def show_context_menu(self, event: tk.Event) -> None:
        """Show row or header context menu on right-click."""
        region = self.tree.identify_region(event.x, event.y)
        if region in ('heading', 'separator'):
            self.show_header_menu(event)
            return
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def setup_ui(self) -> None:
        """Set up the BikeShopApp UI."""
        self.setup_search_frame()
        self.setup_treeview()
        self.tree.bind('<Double-1>', self.on_treeview_double_click)
        self.tree.bind('<ButtonRelease-1>', self.schedule_save_column_widths, add='+')
        self.tree.bind('<Button-3>', self.show_context_menu)
        self.tree.bind('<Motion>', self.on_tree_motion)
        self.tree.bind('<Leave>', self.on_tree_leave)

    def setup_search_frame(self) -> None:
        """Set up a clean, professional toolbar for search and bike-list actions."""
        text = ENGLISH_TEXT
        self.search_frame = tk.Frame(self.root, bg='#eef2f5', bd=0, highlightthickness=0)
        self.search_frame.pack(fill='x', padx=8, pady=(6, 3))

        find_text = 'Find Phone'
        all_text = 'Show All'
        use_text = 'Use Bike'
        edit_main_text = 'Edit Records'
        marked_text = 'Marked Numbers'
        columns_text = 'Columns'
        autofit_text = 'Autofit'

        toolbar_bg = '#eef2f5'
        label_bg = toolbar_bg
        label_fg = '#1f2933'

        def _darker(hex_color: str, amount: int=18) -> str:
            """Return a slightly darker shade for button hover colors."""
            try:
                hex_color = hex_color.lstrip('#')
                r = max(0, int(hex_color[0:2], 16) - amount)
                g = max(0, int(hex_color[2:4], 16) - amount)
                b = max(0, int(hex_color[4:6], 16) - amount)
                return f'#{r:02x}{g:02x}{b:02x}'
            except Exception:
                return '#333333'

        def pro_button(parent: tk.Widget, label: str, command, bg: str, fg: str='white',
                       width: Optional[int]=None, state: Optional[str]=None) -> tk.Button:
            """Create a modern colored toolbar button."""
            kwargs = {
                'text': label,
                'command': command,
                'bg': bg,
                'fg': fg,
                'activebackground': _darker(bg),
                'activeforeground': fg,
                'disabledforeground': '#8a8f98',
                'relief': 'flat',
                'bd': 0,
                'highlightthickness': 0,
                'padx': 10,
                'pady': 4,
                'takefocus': 0,
                'cursor': 'hand2',
                'font': ('Segoe UI', 9, 'bold')
            }
            if width is not None:
                kwargs['width'] = width
            if state is not None:
                kwargs['state'] = state
            button = tk.Button(parent, **kwargs)
            normal_bg = bg
            hover_bg = _darker(bg, 25)

            def on_enter(_event: tk.Event) -> None:
                try:
                    if str(button.cget('state')) == str(tk.NORMAL):
                        button.configure(bg=hover_bg)
                except Exception:
                    pass

            def on_leave(_event: tk.Event) -> None:
                try:
                    button.configure(bg=normal_bg)
                except Exception:
                    pass

            button.bind('<Enter>', on_enter)
            button.bind('<Leave>', on_leave)
            return button

        def make_group(parent: tk.Widget, title: str) -> tk.Frame:
            """Create a small labelled group inside the toolbar."""
            group = tk.Frame(parent, bg=toolbar_bg)
            group.pack(side='left', padx=(0, 14), pady=0)
            tk.Label(group, text=title, bg=label_bg, fg='#5f6b7a', font=('Segoe UI', 8, 'bold')).pack(side='left', padx=(0, 5))
            inner = tk.Frame(group, bg=toolbar_bg)
            inner.pack(side='left')
            return inner

        # Row 1: search controls
        search_row = tk.Frame(self.search_frame, bg=toolbar_bg)
        search_row.pack(fill='x', anchor='w')

        tk.Label(search_row, text=text['phone_label'], bg=label_bg, fg=label_fg,
                 font=('Segoe UI', 9, 'bold')).pack(side='left', padx=(0, 5), pady=2)
        self.phone_entry = tk.Entry(search_row, width=18, relief='solid', bd=1, font=('Segoe UI', 9))
        self.phone_entry.pack(side='left', padx=(0, 6), pady=2, ipady=3)
        self.phone_entry.bind('<Return>', lambda event: self.find_phone_records())

        pro_button(search_row, find_text, self.find_phone_records, '#2563eb', width=11).pack(side='left', padx=(0, 5), pady=2)
        pro_button(search_row, all_text, self.show_all_records, '#64748b', width=10).pack(side='left', padx=(0, 5), pady=2)

        # Row 2: record actions and tools
        actions_row = tk.Frame(self.search_frame, bg=toolbar_bg)
        actions_row.pack(fill='x', anchor='w', pady=(6, 0))

        record_group = make_group(actions_row, 'Record')
        pro_button(record_group, use_text, self.use_selected_bike_in_main_form, '#16a34a', width=10).pack(side='left', padx=(0, 5), pady=1)
        pro_button(record_group, edit_main_text, self.edit_selected_record_in_main_form, '#ea580c', width=12).pack(side='left', padx=(0, 5), pady=1)
        pro_button(record_group, text['delete_button'], self.delete_record, '#dc2626', width=8).pack(side='left', padx=(0, 5), pady=1)
        self.undo_button = pro_button(record_group, text['undo_button'], self.undo_last_action, '#9ca3af', fg='#111827', width=8, state=tk.DISABLED)
        self.undo_button.pack(side='left', padx=(0, 5), pady=1)

        tools_group = make_group(actions_row, 'Tools')
        pro_button(tools_group, marked_text, lambda: open_customer_flags_list_dialog(self.root), '#7c3aed', width=15).pack(side='left', padx=(0, 5), pady=1)
        pro_button(tools_group, autofit_text, lambda: self.auto_adjust_columns(self.tree), '#0891b2', width=9).pack(side='left', padx=(0, 5), pady=1)
        pro_button(tools_group, columns_text, self.open_column_visibility_dialog, '#0f766e', width=9).pack(side='left', padx=(0, 5), pady=1)
        pro_button(tools_group, text['settings_button'], self.open_settings, '#334155', width=9).pack(side='left', padx=(0, 5), pady=1)

        self.history_status_var = tk.StringVar(value='Showing all bikes')
        self.history_status_label = tk.Label(self.search_frame, textvariable=self.history_status_var,
                                             bg=toolbar_bg, fg='#345', anchor='w', font=('Segoe UI', 9))
        self.history_status_label.pack(fill='x', padx=0, pady=(6, 0), anchor='w')

    def open_column_visibility_dialog(self) -> None:
        """Open a small dialog to hide/show bike-list columns."""
        title = 'Choose columns'
        save_label = 'Save'
        all_label = 'Show all'
        default_label = 'Default'
        warning = 'At least one column must be visible.'
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry(get_saved_window_geometry('bike_list_columns', '320x520'))
        set_window_icon(dialog)
        dialog.transient(self.root)
        outer = tk.Frame(dialog)
        outer.pack(fill='both', expand=True, padx=10, pady=10)
        canvas = tk.Canvas(outer, highlightthickness=0)
        scroll = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas)
        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')
        current_visible = set(self._load_visible_columns())
        vars_by_col: Dict[str, tk.BooleanVar] = {}
        for col in self.display_headers:
            var = tk.BooleanVar(value=col in current_visible)
            vars_by_col[col] = var
            cb = tk.Checkbutton(inner, text=self._column_label(col), variable=var, anchor='w')
            cb.pack(fill='x', anchor='w', pady=1)
        buttons = tk.Frame(dialog)
        buttons.pack(fill='x', padx=10, pady=(0, 10))

        def set_all(value: bool) -> None:
            for var in vars_by_col.values():
                var.set(value)

        def set_default() -> None:
            default_cols = set(self._default_visible_columns())
            for col, var in vars_by_col.items():
                var.set(col in default_cols)

        def save_columns() -> None:
            selected = [col for col in self.display_headers if vars_by_col[col].get()]
            if not selected:
                messagebox.showwarning(title, warning, parent=dialog)
                return
            self.save_current_column_widths()
            settings = load_settings()
            settings['bike_list_visible_columns'] = selected
            save_settings(settings)
            self.visible_columns = selected
            self.tree['displaycolumns'] = selected
            self._refresh_heading_labels()
            self._apply_saved_or_measured_column_widths(self.tree, save_missing=True)
            save_window_geometry(dialog, 'bike_list_columns')
            dialog.destroy()
        tk.Button(buttons, text=all_label, command=lambda: set_all(True)).pack(side='left', padx=(0, 4))
        tk.Button(buttons, text=default_label, command=set_default).pack(side='left', padx=(0, 4))
        tk.Button(buttons, text=save_label, command=save_columns).pack(side='right')
        dialog.protocol('WM_DELETE_WINDOW', lambda: (save_window_geometry(dialog, 'bike_list_columns'), dialog.destroy()))

    def _row_phone_digits(self, row: Dict) -> str:
        """Return normalized phone digits for a row."""
        return _phone_digits(row.get('phone', ''))

    def _filtered_history_rows(self, phone: str) -> List[Dict]:
        """Return exact/contains phone history rows for the normalized phone value."""
        clean = _phone_digits(phone)
        if not clean:
            return []
        return [row for row in self.data if clean in self._row_phone_digits(row)]

    def find_phone_records(self) -> None:
        """One clear phone action: show exact history first, otherwise show similar phone numbers."""
        text = ENGLISH_TEXT
        clean = _phone_digits(self.phone_entry.get() if hasattr(self, 'phone_entry') else '')
        if hasattr(self, 'phone_entry'):
            self.phone_entry.delete(0, tk.END)
            self.phone_entry.insert(0, clean)
        if not clean:
            try:
                self.history_status_var.set('Enter a phone number')
            except Exception:
                pass
            app_info('Info', text['search_info'], self.root)
            return
        exact_rows = self._filtered_history_rows(clean)
        if exact_rows:
            self.show_phone_history(clean, update_entry=True, show_if_empty=True)
            return
        self.loading_label = tk.Label(self.root, text=text['searching_message'], fg='blue')
        self.loading_label.pack(pady=10)
        self.root.update()
        for widget in self.search_frame.winfo_children():
            if isinstance(widget, (tk.Button, ttk.Button)):
                widget.config(state='disabled')
        threading.Thread(target=lambda: self.threaded_find_similar_phone(clean), daemon=True).start()

    def threaded_find_similar_phone(self, clean_phone: str) -> None:
        """Search similar phones in a worker thread for the combined Find Phone action."""
        similar_phone_data = self.search_similar_phone_numbers(clean_phone)
        self.root.after(0, lambda: self.finish_find_phone_records(clean_phone, similar_phone_data))

    def finish_find_phone_records(self, clean_phone: str, similar_phone_data: List[Dict]) -> None:
        """Finish the combined Find Phone action in the main thread."""
        self.finish_search(similar_phone_data)
        try:
            if similar_phone_data:
                self.history_status_var.set(f'No exact history for {clean_phone}. Showing {len(similar_phone_data)} similar phone records.')
            else:
                self.history_status_var.set(f'No history or similar phones for {clean_phone}')
        except Exception:
            pass

    def show_phone_history(self, phone: str, update_entry: bool=True, show_if_empty: bool=True) -> None:
        """Show previous visits for one phone number inside this same bike-list window."""
        clean = _phone_digits(phone)
        if update_entry and hasattr(self, 'phone_entry'):
            self.phone_entry.delete(0, tk.END)
            self.phone_entry.insert(0, clean)
        if not clean:
            if show_if_empty:
                self.history_status_var.set('Enter a phone number')
            return
        rows = self._filtered_history_rows(clean)
        self.current_phone_history = clean
        self.history_mode = True
        self.populate_treeview(self.tree, rows)
        count = len(rows)
        if count:
            msg = f'Showing {count} previous visits for {clean}. Double-click = use bike.'
        else:
            msg = f'No history for {clean}'
        self.history_status_var.set(msg)
        try:
            self.root.title(ENGLISH_TEXT['bike_list_title'] + ' – ' + 'Phone History')
        except Exception:
            pass

    def show_all_records(self, clear_phone_entry: bool=True) -> None:
        """Return the list window to normal all-records mode."""
        self.current_phone_history = ''
        self.history_mode = False
        if clear_phone_entry and hasattr(self, 'phone_entry'):
            self.phone_entry.delete(0, tk.END)
        self.populate_treeview(self.tree, self.data)
        self.history_status_var.set('Showing all bikes')
        try:
            self.root.title(ENGLISH_TEXT['bike_list_title'])
        except Exception:
            pass

    def use_selected_bike_in_main_form(self) -> None:
        """Copy the selected row's phone and bike model back to the main entry form."""
        text = ENGLISH_TEXT
        selected_item = self.tree.selection()
        if not selected_item:
            app_info('Info', text['select_record_info'], self.root)
            return
        row = self._get_selected_row()
        if not row:
            return
        try:
            selected_phone = str(row.get('phone', '')).strip()
            selected_model = str(row.get('model', '')).strip()
            phone_entry.delete(0, tk.END)
            phone_entry.insert(0, selected_phone)
            model_entry.delete(0, tk.END)
            model_entry.insert(0, selected_model)
            clear_main_issue_table()
            reload_main_visit_history(selected_phone)
            quick_issue = getattr(root, '_main_issue_entry', None)
            if quick_issue is not None and quick_issue.winfo_exists():
                quick_issue.focus_set()
            root.lift()
            show_toast(root, 'Bike copied from history.', 2500)
        except Exception as e:
            messagebox.showerror(text['error_title'], str(e))

    def edit_selected_record_in_main_form(self) -> None:
        """Load the selected old record into the main input form for no-print updating."""
        text = ENGLISH_TEXT
        selected_item = self.tree.selection()
        if not selected_item:
            app_info('Info', text['select_record_info'], self.root)
            return
        row = self._get_selected_row()
        if not row:
            app_info('Info', text['select_record_info'], self.root)
            return
        enter_main_edit_mode(row)
        show_toast(root, 'Old record loaded for editing.', 2500)

    def setup_treeview(self) -> None:
        """Set up the Treeview widget."""
        text = ENGLISH_TEXT
        self.tree_frame = tk.Frame(self.root)
        self.tree_frame.pack(fill='both', expand=True)
        scrollbar_y = ttk.Scrollbar(self.tree_frame, orient='vertical')
        scrollbar_y.pack(side='right', fill='y')
        scrollbar_x = ttk.Scrollbar(self.tree_frame, orient='horizontal')
        scrollbar_x.pack(side='bottom', fill='x')
        style = ttk.Style()
        style.configure('Treeview', font=('Arial', self.treeview_font_size), rowheight=max(22, self.treeview_font_size + 12))
        style.configure('Treeview.Heading', font=('Arial', self.treeview_font_size, 'bold'), padding=(4, 3))
        try:
            style.map('Treeview', background=[('selected', '#2563eb')], foreground=[('selected', 'white')])
        except Exception:
            pass
        self.tree = ttk.Treeview(self.tree_frame, columns=getattr(self, 'display_headers', self.headers), show='headings', yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        self.tree.pack(fill='both', expand=True)
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)
        self.visible_columns = self._load_visible_columns()
        self.tree['displaycolumns'] = self.visible_columns
        for h in getattr(self, 'display_headers', self.headers):
            self.tree.heading(h, text=self._heading_text(h), command=lambda col=h: self.sort_tree_by_column(col))
            anchor = 'center' if h in ('age_since_visit', 'visit_date', 'visit_time', 'datetime', 'phone') else 'w'
            self.tree.column(h, anchor=anchor)
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label='Use Bike', command=self.use_selected_bike_in_main_form)
        self.context_menu.add_command(label='Edit Records', command=self.edit_selected_record_in_main_form)
        self.context_menu.add_separator()
        self.context_menu.add_command(label=text['context_delete'], command=self.delete_record)
        self.context_menu.add_command(label=text['context_generate'], command=self.generate_qr_from_selection)
        self.context_menu.add_command(label=text['context_search'], command=self.search_from_context)
        self.context_menu.add_command(label='Reprint Receipt', command=self.reprint_receipt_from_selection)
        self.context_menu.add_separator()
        mark_menu = tk.Menu(self.context_menu, tearoff=0)
        mark_menu.add_command(label='Watch', command=lambda: self.mark_selected_customer('watch'))
        mark_menu.add_command(label='Problem', command=lambda: self.mark_selected_customer('problem'))
        mark_menu.add_command(label='High Risk', command=lambda: self.mark_selected_customer('high_risk'))
        self.context_menu.add_cascade(label='Mark Customer', menu=mark_menu)
        self.context_menu.add_command(label='Manage This Mark', command=self.manage_selected_customer_mark)
        self.context_menu.add_command(label='Remove Mark', command=self.delete_selected_customer_mark)
        self.context_menu.add_separator()
        self.context_menu.add_command(label=text['context_copy'], command=self.copy_phone_number)
        self.context_menu.add_command(label='Copy Bike', command=self.copy_selected_bike_model)
        self.populate_treeview(self.tree, self.data)

    def reprint_receipt_from_selection(self) -> None:
        """Reprint receipt for the selected item."""
        text = ENGLISH_TEXT
        selected_item = self.tree.selection()
        if not selected_item:
            app_info('Info', text['select_record_info'])
            return
        status_message = tk.Label(self.root, text='Processing...', fg='blue')
        status_message.pack(pady=10)
        self.root.update()
        row = self._get_row_for_item(selected_item[0])
        if not row:
            status_message.destroy()
            messagebox.showerror(text['reprint_title'], 'Record data not found.')
            return
        record_data = [row.get('datetime', ''), row.get('phone', ''), row.get('model', '')]
        for i in range(1, 11):
            issue_key = f'issue{i}'
            record_data.append(row.get(issue_key, '') or '')
        record_data.append(row.get('shop_note', '') or '')
        threading.Thread(target=lambda: self.threaded_reprint_receipt(record_data, status_message), daemon=True).start()

    def threaded_reprint_receipt(self, record_data, status_message):
        """Print receipt in a background thread and handle UI updates."""
        text = ENGLISH_TEXT
        print_receipt(record_data)
        self.root.after(0, lambda: self.finish_reprint(status_message))

    def finish_reprint(self, status_message):
        """Clean up after reprinting receipt."""
        text = ENGLISH_TEXT
        if status_message and status_message.winfo_exists():
            status_message.destroy()
        success_message = tk.Label(self.root, text=text['reprint_success'], fg='green')
        success_message.pack(pady=10)
        self.root.after(3000, lambda: success_message.destroy() if success_message.winfo_exists() else None)

def open_bike_list(initial_phone_history: str='', lift: bool=True) -> None:
    """Open the bike list window, optionally filtered to one phone number's history."""
    global bike_list_window, bike_shop_app
    text = ENGLISH_TEXT
    if not os.path.exists(DATABASE_FILE):
        init_database()
    clean_history = _phone_digits(initial_phone_history)
    if bike_list_window and tk.Toplevel.winfo_exists(bike_list_window):
        if lift:
            bike_list_window.lift()
        if clean_history and bike_shop_app:
            try:
                bike_shop_app.show_phone_history(clean_history, update_entry=True, show_if_empty=True)
            except Exception:
                pass
        return
    try:
        bike_list_window = tk.Toplevel(root)
        set_window_icon(bike_list_window)
        bike_shop_app = BikeShopApp(bike_list_window, initial_phone_history=clean_history)

        def _close_bike_list_window():
            try:
                if bike_shop_app:
                    bike_shop_app.save_current_column_widths()
            except Exception:
                pass
            save_window_geometry(bike_list_window, 'bike_list')
            bike_list_window.destroy()
            globals().update(bike_list_window=None, bike_shop_app=None)
        bike_list_window.protocol('WM_DELETE_WINDOW', _close_bike_list_window)
        status_label.config(text=text['status_opened'], fg='green')
    except Exception as e:
        status_label.config(text=text['status_open_error'].format(str(e)), fg='red')

def has_visit_history(phone: str) -> bool:
    """Quickly check if any visit exists for this phone. Returns True/False."""
    try:
        clean = _phone_digits(phone)
        if not clean:
            return False
        conn = get_db_connection()
        try:
            c = conn.cursor()
            c.execute("\n                SELECT 1\n                FROM customer_bikes\n                WHERE REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(phone,' ',''),'-',''),'+',''),'(',''),')','') LIKE ?\n                LIMIT 1\n                ", (f'%{clean}%',))
            row = c.fetchone()
            return row is not None
        finally:
            release_db_connection(conn)
    except Exception:
        return False


def show_previous_visits_for_phone(phone: str, lift: bool=True) -> None:
    """Open or update a separate visit-history window for the entered phone number."""
    text = ENGLISH_TEXT
    clean = _phone_digits(phone)
    if not clean:
        app_info(text.get('visit_history_title', 'Visit History'), text.get('visit_history_empty', 'Phone number is empty.'))
        return

    existing = getattr(root, '_visit_history_dialog', None)
    try:
        if existing and existing.winfo_exists():
            reload_visit_history_dialog(existing, clean)
            if lift:
                existing.lift()
                existing.focus_force()
            return
    except Exception:
        pass

    try:
        dialog = tk.Toplevel(root)
        set_window_icon(dialog)
        dialog.title(text.get('visit_history_title', 'Visit History') + ' – ' + clean)
        try:
            dialog.geometry(get_saved_window_geometry('visit_history', '920x430'))
        except Exception:
            dialog.geometry('920x430')
        dialog.minsize(720, 280)
        dialog.resizable(True, True)
        root._visit_history_dialog = dialog
        dialog._vh_phone = clean

        def _close_dialog():
            try:
                save_window_geometry(dialog, 'visit_history')
            except Exception:
                pass
            try:
                if getattr(root, '_visit_history_dialog', None) is dialog:
                    root._visit_history_dialog = None
            except Exception:
                pass
            try:
                dialog.destroy()
            except Exception:
                pass

        def _clear_ref(event=None):
            try:
                if event is not None and event.widget is not dialog:
                    return
                if getattr(root, '_visit_history_dialog', None) is dialog:
                    root._visit_history_dialog = None
            except Exception:
                pass

        def _selected_values():
            tree = getattr(dialog, '_vh_tree', None)
            if not tree:
                return None
            selected = tree.selection()
            if not selected:
                return None
            values = tree.item(selected[0], 'values')
            if not values or len(values) < 5:
                return None
            return values

        def _use_selected_bike(event=None):
            values = _selected_values()
            if not values:
                app_info('Info', text.get('select_record_info', 'Please select a record.'), dialog)
                return 'break'
            _dt, _ago, row_phone, model, _issues = values[:5]
            try:
                if phone_entry:
                    phone_entry.delete(0, tk.END)
                    phone_entry.insert(0, str(row_phone).strip())
                if model_entry:
                    model_entry.delete(0, tk.END)
                    model_entry.insert(0, str(model).strip())
                    model_entry.focus_set()
                root.lift()
                show_toast(root, 'Bike model copied from visit history.', 2500)
            except Exception as exc:
                messagebox.showerror(text.get('error_title', 'Error'), str(exc))
            return 'break'

        def _copy_selected_phone():
            values = _selected_values()
            if not values:
                app_info('Info', text.get('select_record_info', 'Please select a record.'), dialog)
                return
            try:
                dialog.clipboard_clear()
                dialog.clipboard_append(str(values[2]))
                show_toast(dialog, text.get('copy_success', 'Phone number copied to clipboard.'), 2000)
            except Exception:
                pass

        top = ttk.Frame(dialog, padding=(10, 8, 10, 4))
        top.pack(fill='x')
        status_var = tk.StringVar(value='Loading visit history...')
        dialog._vh_status_var = status_var
        ttk.Label(top, textvariable=status_var, anchor='w').pack(side='left', fill='x', expand=True)

        tree_frame = ttk.Frame(dialog, padding=(10, 4, 10, 6))
        tree_frame.pack(fill='both', expand=True)
        scrollbar_y = ttk.Scrollbar(tree_frame, orient='vertical')
        scrollbar_y.pack(side='right', fill='y')
        scrollbar_x = ttk.Scrollbar(tree_frame, orient='horizontal')
        scrollbar_x.pack(side='bottom', fill='x')

        columns = ('datetime', 'ago', 'phone', 'model', 'issues')
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        tree.configure(displaycolumns=('datetime', 'ago', 'model', 'issues'))
        dialog._vh_tree = tree
        tree.heading('datetime', text='Date')
        tree.heading('ago', text=text.get('history_col_ago', 'Ago'))
        tree.heading('phone', text=text.get('phone_label', 'Phone Number:').rstrip(':'))
        tree.heading('model', text=text.get('history_col_model', 'Model'))
        tree.heading('issues', text=text.get('history_col_issues', 'Issues'))
        tree.column('datetime', width=110, minwidth=90, anchor='center')
        tree.column('ago', width=120, minwidth=90, anchor='center')
        tree.column('phone', width=130, minwidth=100, anchor='center')
        tree.column('model', width=210, minwidth=130, anchor='w')
        tree.column('issues', width=430, minwidth=180, anchor='w')
        tree.pack(fill='both', expand=True)
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        tree.bind('<Double-1>', _use_selected_bike)
        tree.bind('<Return>', _use_selected_bike)

        context_menu = tk.Menu(dialog, tearoff=0)
        context_menu.add_command(label='Use Bike Model', command=_use_selected_bike)
        context_menu.add_command(label=text.get('context_copy', 'Copy Phone Number'), command=_copy_selected_phone)

        def _show_context_menu(event):
            try:
                iid = tree.identify_row(event.y)
                if iid:
                    tree.selection_set(iid)
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                try:
                    context_menu.grab_release()
                except Exception:
                    pass
        tree.bind('<Button-3>', _show_context_menu)

        buttons = ttk.Frame(dialog, padding=(10, 0, 10, 10))
        buttons.pack(fill='x')
        ttk.Button(buttons, text='Use Bike Model', command=_use_selected_bike).pack(side='left')
        refresh_button = ttk.Button(buttons, text=text.get('refresh_button', 'Refresh'), command=lambda: reload_visit_history_dialog(dialog, getattr(dialog, '_vh_phone', clean)))
        refresh_button.pack(side='left', padx=(8, 0))
        dialog._vh_refresh_button = refresh_button
        ttk.Button(buttons, text=text.get('close_button', 'Close'), command=_close_dialog).pack(side='right')

        dialog.protocol('WM_DELETE_WINDOW', _close_dialog)
        dialog.bind('<Destroy>', _clear_ref, add='+')
        dialog.bind('<Escape>', lambda event=None: (_close_dialog(), 'break')[-1])

        try:
            if not get_saved_window_geometry('visit_history', ''):
                root.update_idletasks()
                nx = root.winfo_rootx() + root.winfo_width() + 20
                ny = root.winfo_rooty()
                dialog.geometry(f'+{nx}+{ny}')
        except Exception:
            pass

        reload_visit_history_dialog(dialog, clean)
        if lift:
            dialog.lift()
    except Exception as e:
        messagebox.showerror(text.get('visit_history_title', 'Visit History'), f"{text.get('visit_history_error', 'Error loading history')}: {e}")

def open_reprint_dialog() -> None:
    """Open dialog to select an entry for reprinting."""
    text = ENGLISH_TEXT
    dialog = tk.Toplevel(root)
    dialog.title(text['reprint_title'])
    dialog.geometry('600x400')
    dialog.resizable(True, True)
    set_window_icon(dialog)
    loading_label = tk.Label(dialog, text=text['loading_message'], fg='blue')
    loading_label.pack(pady=20)
    dialog.update()
    tree_frame = ttk.Frame(dialog)
    button_frame = ttk.Frame(dialog)
    columns = ('datetime', 'phone', 'model')
    tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
    tree_scroll_y = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
    tree_scroll_x = ttk.Scrollbar(tree_frame, orient='horizontal', command=tree.xview)
    tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
    tree.heading('datetime', text='Date and Time')
    tree.heading('phone', text='Phone')
    tree.heading('model', text='Model')
    tree.column('datetime', width=150, anchor='center')
    tree.column('phone', width=100, anchor='center')
    tree.column('model', width=200, anchor='center')
    threading.Thread(target=lambda: load_entries_for_reprint(dialog, tree, loading_label, tree_frame, button_frame), daemon=True).start()

def load_entries_for_reprint(dialog, tree, loading_label, tree_frame, button_frame):
    """Load database entries for reprint in a separate thread."""
    text = ENGLISH_TEXT
    try:
        with sqlite3.connect(DATABASE_FILE, timeout=30, check_same_thread=False) as conn:
            conn.execute('PRAGMA journal_mode=DELETE')
            conn.execute('PRAGMA busy_timeout=30000')
            conn.execute('PRAGMA synchronous=NORMAL')
            conn.execute('PRAGMA foreign_keys=ON')
            cursor = conn.cursor()
            cursor.execute('SELECT datetime, phone, model, issue1, issue2, issue3, issue4, issue5, issue6, issue7, issue8, issue9, issue10, shop_note FROM customer_bikes ORDER BY datetime DESC')
            entries = cursor.fetchall()
        dialog.after(0, lambda: finish_reprint_dialog_setup(dialog, tree, loading_label, tree_frame, button_frame, entries))
    except sqlite3.Error as e:
        dialog.after(0, lambda: messagebox.showerror(text['reprint_title'], text['reprint_error'].format(str(e))))
        dialog.after(0, dialog.destroy)

def finish_reprint_dialog_setup(dialog, tree, loading_label, tree_frame, button_frame, entries):
    """Complete the setup of reprint dialog in the main thread."""
    text = ENGLISH_TEXT
    loading_label.destroy()
    if not entries:
        app_info(text['reprint_title'], text['no_entries'])
        dialog.destroy()
        return
    tree_frame.pack(fill='both', expand=True, padx=10, pady=10)
    tree_scroll_y = tree.config('yscrollcommand')[4]
    tree_scroll_y.pack(side='right', fill='y')
    tree_scroll_x = tree.config('xscrollcommand')[4]
    tree_scroll_x.pack(side='bottom', fill='x')
    tree.pack(fill='both', expand=True)
    for entry in entries:
        tree.insert('', 'end', values=(entry[0], entry[1], entry[2]))
    button_frame.pack(pady=10)

    def reprint_selected():
        selected = tree.selection()
        if not selected:
            app_info(text['reprint_title'], text['select_record_info'])
            return
        selected_item = tree.item(selected[0], 'values')
        for entry in entries:
            if entry[0] == selected_item[0] and entry[1] == selected_item[1] and (entry[2] == selected_item[2]):
                threading.Thread(target=lambda: print_receipt(list(entry)), daemon=True).start()
                status_label.config(text=text['reprint_success'], fg='green')
                dialog.destroy()
                return
        messagebox.showerror(text['reprint_title'], 'Entry not found.')
    ttk.Button(button_frame, text=text['reprint_button'], command=reprint_selected).pack(side='left', padx=5)
    ttk.Button(button_frame, text='Cancel', command=dialog.destroy).pack(side='left', padx=5)
    dialog.protocol('WM_DELETE_WINDOW', dialog.destroy)
if __name__ == '__main__':
    start_main_application()
