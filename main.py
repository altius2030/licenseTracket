"""
License Tracker - Main Application
Windows app for tracking license expiration dates.
Built with CustomTkinter, SQLite backend, Excel import/export.
"""
import customtkinter as ctk
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk, simpledialog
from datetime import datetime, date, timedelta
import calendar
from typing import Optional, Dict, Any
import os
import re
import sys
import json
import threading
import tempfile
import subprocess
import urllib.request
import queue
import ctypes
import ctypes.wintypes

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from database import Database, import_excel_file, get_app_dir, parse_excel_date
from models import (
    LicenseStatus, ManualStatus, STATUS_COLORS, STATUS_LABELS,
    MANUAL_STATUS_LABELS, get_fields_for_sheet
)


# ===== Constants =====
APP_TITLE = "License Tracker"
APP_WIDTH = 1400
APP_HEIGHT = 800
APP_VERSION = "1.1.17"
GITHUB_REPO = "altius2030/licenseTracket"
GITHUB_RELEASES_LATEST = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
DB_PATH = os.path.join(get_app_dir(), "licenses.db")
EXCEL_SOURCE = r"C:\Users\user\Desktop\Лицензии\База Данных.xlsx"

# Program categories (nav bar items and card selector)
CATEGORIES = ["GBS-market", "Shopuchet", "МойСклад", "Отвалы"]

# ===== Theme palettes =====
THEMES = {
    'dark': {
        'BG_APP': '#0a0d14',       # app background (near black)
        'BG_PANEL': '#11161f',     # card/panel background
        'BG_PANEL_2': '#171e2a',   # inputs, rows
        'BG_NAV': '#0d111a',       # top bar
        'ACCENT': '#7c6cf6',
        'ACCENT_HOVER': '#9082ff',
        'TEXT_MAIN': '#eef0f7',
        'TEXT_MUTED': '#949cb0',
        'BORDER': '#2a3244',
        'TREE_BG': '#151b26',
        'TREE_HEAD': '#1c2431',
        'TREE_HEAD_HOVER': '#26304a',
        # Row tints (vivid)
        'TAG_BG': {
            'active':   '#12351f',
            'expiring': '#453107',
            'expired':  '#4a1220',
            'manual':   '#18315c',
            'critical': '#641016',
        },
        'TAG_FG': {
            'active':   '#4ade80',
            'expiring': '#fbbf24',
            'expired':  '#fb7185',
            'manual':   '#60a5fa',
            'critical': '#f87171',
        },
    },
    'light': {
        'BG_APP': '#f2f4f9',
        'BG_PANEL': '#ffffff',
        'BG_PANEL_2': '#e9edf4',
        'BG_NAV': '#ffffff',
        'ACCENT': '#6d5bf0',
        'ACCENT_HOVER': '#5946dd',
        'TEXT_MAIN': '#1b2230',
        'TEXT_MUTED': '#5f6876',
        'BORDER': '#d7dce6',
        'TREE_BG': '#ffffff',
        'TREE_HEAD': '#e9edf4',
        'TREE_HEAD_HOVER': '#dde3ee',
        # Row tints (light vivid)
        'TAG_BG': {
            'active':   '#d8f5e3',
            'expiring': '#fdeec9',
            'expired':  '#fce0e6',
            'manual':   '#dbe9fd',
            'critical': '#ffd4d9',
        },
        'TAG_FG': {
            'active':   '#0f7a42',
            'expiring': '#8a5a00',
            'expired':  '#b01835',
            'manual':   '#1055b5',
            'critical': '#c40f22',
        },
    },
}

# Current palette lives in module globals (BG_APP, BG_PANEL, ... are reassigned)
TAG_BG = {}   # filled by apply_theme
TAG_FG = {}

def apply_theme(name: str):
    """Populate module-level color constants from THEMES[name]."""
    pal = THEMES.get(name, THEMES['dark'])
    g = globals()
    for key, value in pal.items():
        if key.startswith('TAG_'):
            target = g[key]
            target.clear()
            target.update(value)
        else:
            g[key] = value
    g['CURRENT_THEME'] = name

apply_theme('dark')

# Month names for the date picker
MONTH_NAMES = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
               "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]


class DropdownOption(ctk.CTkButton):
    """Dropdown button: the list opens ABOVE the button (falls back to below
    when there is no room at the top). Selection happens only by explicit click,
    so a hovering/stuck mouse never picks an item accidentally."""

    def __init__(self, master, values, variable: ctk.StringVar, command=None,
                 width=170, height=30, **kw):
        super().__init__(
            master, text="  " + str(variable.get()), anchor="w",
            width=width, height=height, corner_radius=8,
            fg_color=kw.pop('fg_color', BG_PANEL_2),
            hover_color=kw.pop('hover_color', BORDER),
            text_color=kw.pop('text_color', TEXT_MAIN),
            font=kw.pop('font', ctk.CTkFont(size=12)),
            command=self._open_menu
        )
        self._values = list(values)
        self._var = variable
        # NOTE: do NOT store this in self._command - CTkButton uses that
        # attribute as the callback invoked on a real click. Overwriting it
        # made left clicks run the app callback and never open the list.
        self._on_select = command
        variable.trace_add("write", lambda *a: self._sync_text())

    def _sync_text(self):
        self.configure(text="  " + str(self._var.get()))

    def set_values(self, values):
        self._values = list(values)

    def _open_menu(self):
        # Toggle / close an already open popup
        if getattr(self, "_popup", None) is not None:
            try:
                if self._popup.winfo_exists():
                    self._close()
                    return
            except Exception:
                pass
            self._popup = None

        popup = ctk.CTkToplevel(self)
        popup.overrideredirect(True)      # borderless popup
        popup.configure(fg_color=BG_PANEL_2)
        # Keep a reference so the popup is not garbage-collected
        self._popup = popup

        item_h = 34
        max_h = 400
        n = len(self._values)
        scroll = n > 10

        if scroll:
            inner = ctk.CTkScrollableFrame(
                popup, height=max_h, width=max(self.winfo_width(), 200),
                fg_color=BG_PANEL_2,
                scrollbar_button_color=BORDER, scrollbar_button_hover_color=ACCENT,
                corner_radius=0
            )
            inner.pack(fill="both", expand=True)
        else:
            inner = popup

        for v in self._values:
            ctk.CTkButton(
                inner, text=v, anchor="w", height=item_h,
                corner_radius=0, border_spacing=12,
                fg_color="transparent", hover_color=ACCENT,
                text_color=TEXT_MAIN, font=ctk.CTkFont(size=12),
                command=lambda val=v: self._choose(val)
            ).pack(fill="x")

        # Geometry: above the button, width matches/surpasses button
        popup.update_idletasks()
        self.update_idletasks()
        popup_w = max(self.winfo_width(), 180) + 4
        popup_h = min(n * item_h + 8, max_h + 4)

        btn_x = self.winfo_rootx()
        btn_top = self.winfo_rooty()
        y_above = btn_top - popup_h - 4
        if y_above >= 40:  # don't cover the screen top / taskbar area
            popup.geometry(f"{popup_w}x{popup_h}+{btn_x}+{y_above}")   # ABOVE
        else:
            popup.geometry(
                f"{popup_w}x{popup_h}+{btn_x}+{btn_top + self.winfo_height() + 2}"
            )                                                          # below if no room

        popup.lift()
        # Close when the user clicks anywhere outside the popup.
        # No grab_set / focus_force here: borderless windows often refuse
        # focus on Windows, which made the popup close instantly (or never
        # receive clicks at all). Plain click-outside tracking is reliable.
        owner = self.winfo_toplevel()
        self._outside_click_bind = owner.bind(
            "<Button-1>", self._on_owner_click, add="+")
        self._escape_bind = owner.bind(
            "<Escape>", lambda e: self._close(), add="+")

    def _on_owner_click(self, event):
        """Close the popup when the click lands outside of it."""
        popup = getattr(self, "_popup", None)
        if popup is None:
            return
        try:
            if not popup.winfo_exists():
                return
        except Exception:
            return
        widget = event.widget
        # A click on the dropdown button itself is handled by the button
        # command, which toggles the popup closed.
        if widget is self or str(widget).startswith(str(self) + "."):
            return
        self._close()

    def _close(self):
        popup = getattr(self, "_popup", None)
        if popup is None:
            return
        self._popup = None

        try:
            owner = self.winfo_toplevel()
        except Exception:
            owner = None
        if owner is not None:
            for attr, sequence in (("_outside_click_bind", "<Button-1>"),
                                   ("_escape_bind", "<Escape>")):
                fid = getattr(self, attr, None)
                if fid:
                    try:
                        owner.unbind(sequence, fid)
                    except Exception:
                        pass
                setattr(self, attr, None)

        try:
            if popup.winfo_exists():
                popup.destroy()
        except Exception:
            pass

    def _choose(self, value):
        self._var.set(value)
        self._close()
        if self._on_select:
            try:
                self._on_select(value)
            except TypeError:
                self._on_select()


# ===== Date helpers =====

def parse_date_any(val) -> Optional[datetime]:
    """Parse a date value from DB/Excel into datetime. Returns None on failure."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return datetime.combine(val.date(), datetime.min.time())
    if isinstance(val, date):
        return datetime.combine(val, datetime.min.time())
    s = str(val).strip()
    if not s:
        return None
    # ISO (with or without time / 'Z' / fractional seconds)
    try:
        return datetime.fromisoformat(s.replace('Z', '').split('.')[0])
    except ValueError:
        pass
    try:
        return datetime.strptime(s[:10], "%Y-%m-%d")
    except ValueError:
        pass
    # DD.MM.YYYY variants (Excel import)
    for fmt in ("%d.%m.%Y", "%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S"):
        try:
            return datetime.strptime(s[:19], fmt)
        except ValueError:
            continue
    return None


def add_months(d: date, months: int) -> date:
    """Add months to a date, clamping the day (e.g. 31.01 + 1 month = 28.02)."""
    idx = d.month - 1 + months
    year = d.year + idx // 12
    month = idx % 12 + 1
    day = min(d.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


# Order of the lines produced by the "copy card" action
COPY_FIELD_ORDER = ("full_name", "email", "store_name", "phone")


def format_card_copy_text(data: Dict[str, Any]) -> str:
    """Build the contact block copied from a card, one value per line:

        Фамилия и имя
        Почта
        Магазин
        Телефон
        6:XXXX... (GBS ID / ключ лицензии)
    """
    if not data:
        return ""

    lines = []
    for key in COPY_FIELD_ORDER:
        value = data.get(key)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            lines.append(text)

    license_id = str(data.get('gbs_id') or '').strip()
    if not license_id:
        license_id = str(data.get('license_key') or '').strip()
    if license_id:
        lines.append(license_id)

    return "\n".join(lines)


def copy_to_clipboard(widget, text: str) -> bool:
    """Put text on the clipboard; returns False when there is nothing to copy.

    On Windows we write through the Win32 clipboard API so the data persists
    even after the owning widget / dialog is destroyed (Tk's clipboard_append
    keeps the data only while the selection owner exists, which made a second
    copy from a newly opened card silently fail).
    """
    if not text:
        return False

    # Windows: write directly to the OS clipboard (reliable overwrite).
    try:
        import ctypes
        CF_UNICODETEXT = 13
        GMEM_MOVEABLE = 0x0002
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        user32.OpenClipboard.argtypes = [ctypes.c_void_p]
        user32.OpenClipboard.restype = ctypes.c_bool
        user32.EmptyClipboard.restype = ctypes.c_bool
        user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
        user32.SetClipboardData.restype = ctypes.c_void_p
        user32.CloseClipboard.restype = ctypes.c_bool
        kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
        kernel32.GlobalAlloc.restype = ctypes.c_void_p
        kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalFree.argtypes = [ctypes.c_void_p]

        if not user32.OpenClipboard(None):
            raise RuntimeError("OpenClipboard failed")
        try:
            user32.EmptyClipboard()
            payload = text.encode("utf-16-le") + b"\x00\x00"
            handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(payload))
            if not handle:
                raise RuntimeError("GlobalAlloc failed")
            locked = kernel32.GlobalLock(handle)
            if not locked:
                kernel32.GlobalFree(handle)
                raise RuntimeError("GlobalLock failed")
            ctypes.memmove(locked, payload, len(payload))
            kernel32.GlobalUnlock(handle)
            if not user32.SetClipboardData(CF_UNICODETEXT, handle):
                kernel32.GlobalFree(handle)
                raise RuntimeError("SetClipboardData failed")
            # ownership of `handle` transfers to the clipboard on success
        finally:
            user32.CloseClipboard()
        return True
    except Exception:
        pass

    # Fallback (non-Windows / ctypes unavailable): Tk-based clipboard.
    try:
        widget.clipboard_clear()
        widget.clipboard_append(text)
        try:
            widget.update()
        except Exception:
            pass
        return True
    except Exception:
        return False


# ===== Auto-update helpers =====

def _parse_version(value) -> tuple:
    """Parse a version string like 'v1.2.3' into a comparable tuple."""
    text = str(value or "").strip().lstrip("vV")
    parts = []
    for chunk in text.split("."):
        num = ""
        for ch in chunk:
            if ch.isdigit():
                num += ch
            else:
                break
        parts.append(int(num) if num else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def _is_newer(latest, current) -> bool:
    return _parse_version(latest) > _parse_version(current)


def fetch_latest_release(timeout: int = 15):
    """Query GitHub for the latest release. Returns the parsed JSON dict."""
    req = urllib.request.Request(
        GITHUB_RELEASES_LATEST,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": f"LicenseTracker/{APP_VERSION}",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def find_update_asset_url(release) -> Optional[str]:
    """Find the downloadable .exe asset URL in a release payload."""
    if not isinstance(release, dict):
        return None
    assets = release.get("assets") or []
    # Prefer an asset whose name contains 'license' and ends with .exe
    for asset in assets:
        name = str(asset.get("name") or "").lower()
        if name.endswith(".exe") and "license" in name:
            return asset.get("browser_download_url")
    for asset in assets:
        name = str(asset.get("name") or "").lower()
        if name.endswith(".exe"):
            return asset.get("browser_download_url")
    return None


def download_update_file(url: str, dest: str, progress_cb=None, timeout: int = 120):
    """Download url to dest, calling progress_cb(downloaded, total) periodically."""
    req = urllib.request.Request(url, headers={"User-Agent": f"LicenseTracker/{APP_VERSION}"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        total = int(resp.headers.get("Content-Length") or 0)
        downloaded = 0
        with open(dest, "wb") as f:
            while True:
                chunk = resp.read(65536)
                if not chunk:
                    break
                f.write(chunk)
                downloaded += len(chunk)
                if progress_cb:
                    try:
                        progress_cb(downloaded, total)
                    except Exception:
                        pass


def apply_update_and_restart(new_exe_path: str):
    """Replace the current exe with new_exe_path via a helper batch, then exit.

    Works only for a frozen (PyInstaller) build. The batch waits for this
    process to finish, copies the new file over and relaunches the app.
    """
    if not getattr(sys, "frozen", False):
        return False
    current = sys.executable
    bat_path = os.path.join(tempfile.gettempdir(), "lt_update.bat")
    lines = [
        "@echo off",
        "chcp 65001 >nul",
        "timeout /t 2 /nobreak >nul",
        ":wait",
        f'tasklist /FI "IMAGENAME eq {os.path.basename(current)}" 2>nul | find /I "{os.path.basename(current)}" >nul',
        "if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto wait)",
        ":copy",
        f'copy /Y "{new_exe_path}" "{current}" >nul',
        "if errorlevel 1 (timeout /t 1 /nobreak >nul & goto copy)",
        f'start "" "{current}"',
        f'del "{new_exe_path}" >nul 2>&1',
        'del "%~f0" >nul 2>&1',
    ]
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    subprocess.Popen(
        ["cmd", "/c", bat_path],
        creationflags=subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS,
    )
    return True


class LicenseCardDialog(ctk.CTkToplevel):
    """Dialog for viewing/editing a license card."""
    
    def __init__(self, parent, db: Database, license_data: Optional[Dict] = None, sheet_name: str = "", category: str = "GBS-market"):
        super().__init__(parent)
        self.db = db
        self.license_data = license_data
        if license_data and license_data.get('sheet_name'):
            self.sheet_name = license_data['sheet_name']
        else:
            self.sheet_name = sheet_name or "GBS"
        if license_data and license_data.get('category'):
            self.category = license_data['category']
        else:
            self.category = category or "GBS-market"
        self.is_new = license_data is None
        self.result = None
        
        self.title("Новая лицензия" if self.is_new else f"Карточка: {license_data.get('full_name', '')}")
        self.geometry("700x800")
        self.resizable(False, True)
        self.transient(parent)
        self.grab_set()
        
        self.fields = get_fields_for_sheet(self.sheet_name)
        self.entries = {}
        
        self._build_ui()
        self._load_data()
        self._center_window()

        # Shortcut: Ctrl+Shift+C copies the card data
        self.bind("<Control-Shift-C>", lambda e: self._on_copy_card())
        # Ctrl+C: copy selected text in a field, otherwise the whole card
        self.bind("<Control-c>", self._on_ctrl_c_card)
        self.bind("<Control-C>", self._on_ctrl_c_card)
        self.bind("<KeyPress>", self._on_key_press_card, add="+")
        self.bind("<Control-Cyrillic_es>", self._on_ctrl_c_card)
        self.bind("<Control-Cyrillic_ES>", self._on_ctrl_c_card)
        self.bind("<<Copy>>", self._on_ctrl_c_card)
    
    def _center_window(self):
        self.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width() // 2) - (self.winfo_width() // 2)
        y = self.master.winfo_y() + (self.master.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")
    
    def _build_ui(self):
        # Header frame (fixed at top, not scrolling)
        header = ctk.CTkFrame(self, height=70, corner_radius=0, fg_color=BG_NAV)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        title_text = "Новая лицензия" if self.is_new else "Карточка лицензии"
        icon = "➕" if self.is_new else "👤"
        ctk.CTkLabel(
            header, text=f"{icon}  {title_text}",
            font=ctk.CTkFont(size=16, weight="bold"), text_color=TEXT_MAIN
        ).pack(side="left", padx=(18, 5), pady=20)

        # Category selector: lets the card be moved between categories
        category_box = ctk.CTkFrame(header, fg_color="transparent")
        category_box.pack(side="right", padx=16)

        ctk.CTkLabel(
            category_box, text="Категория:", font=ctk.CTkFont(size=11),
            text_color=TEXT_MUTED
        ).pack(side="left", padx=(0, 6))

        self.category_var = ctk.StringVar(value=self.category)
        self.category_menu = DropdownOption(
            category_box, CATEGORIES, self.category_var,
            width=160, height=30
        )
        self.category_menu.pack(side="left")

        self.sheet_label = ctk.CTkLabel(
            category_box, text=f" · {self.sheet_name}",
            font=ctk.CTkFont(size=11), text_color=TEXT_MUTED
        )
        self.sheet_label.pack(side="left")
        
        # Main scrollable frame
        self.scroll_frame = ctk.CTkScrollableFrame(self, corner_radius=0, fg_color="transparent")
        self.scroll_frame.pack(fill="both", expand=True)
        
        # Section headers for field grouping
        SECTION_MARKERS = {
            'install_date': '📋  Основная информация',
            'gbs_id': '🔑  Лицензия',
            'payment_date': '💰  Оплата',
            'sent_info': '📝  Прочее',
        }
        
        # Fields
        self._field_copy_buttons = {}
        for field_name, label, field_type in self.fields:
            if field_name in SECTION_MARKERS:
                ctk.CTkLabel(
                    self.scroll_frame, text=SECTION_MARKERS[field_name],
                    font=ctk.CTkFont(size=11, weight="bold"), anchor="w",
                    text_color=TEXT_MUTED
                ).pack(fill="x", pady=(16, 4))
            self._create_field(field_name, label, field_type)
        
        # Status section (always shown)
        self._create_status_section()

        # Comments section (always shown; requires saved card)
        self._create_comments_section()
        
        # Buttons (fixed at bottom)
        btn_frame = ctk.CTkFrame(self, height=62, corner_radius=0, fg_color=BG_NAV)
        btn_frame.pack(fill="x")
        btn_frame.pack_propagate(False)
        
        ctk.CTkButton(
            btn_frame, text="💾  Сохранить", command=self._on_save,
            fg_color=ACCENT, hover_color=ACCENT_HOVER, width=150, height=38,
            corner_radius=8, font=ctk.CTkFont(size=13, weight="bold")
        ).pack(side="left", padx=(16, 6), pady=12)
        
        ctk.CTkButton(
            btn_frame, text="Отмена", command=self._on_cancel,
            fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
            width=100, height=38, corner_radius=8
        ).pack(side="left", padx=6, pady=12)

        # Copy the card's contact block to the clipboard
        self.copy_btn = ctk.CTkButton(
            btn_frame, text="📋  Копировать", command=self._on_copy_card,
            fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
            width=140, height=38, corner_radius=8,
            font=ctk.CTkFont(size=12)
        )
        self.copy_btn.pack(side="left", padx=6, pady=12)
        
        if not self.is_new:
            ctk.CTkButton(
                btn_frame, text="🗑  Удалить", command=self._on_delete,
                fg_color="#e5484d", hover_color="#c73840",
                width=110, height=38, corner_radius=8
            ).pack(side="right", padx=(6, 16), pady=12)
    
    def _create_field(self, field_name: str, label: str, field_type: str):
        frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        frame.pack(fill="x", pady=3)

        # Copy button for the field
        copy_btn = ctk.CTkButton(
            frame, text="📋", width=32, height=32, corner_radius=6,
            fg_color=BG_PANEL_2, hover_color=ACCENT, text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=14),
            command=lambda fn=field_name: self._copy_field_value(fn)
        )
        copy_btn.pack(side="left", padx=(0, 4))
        self._field_copy_buttons[field_name] = copy_btn

        ctk.CTkLabel(frame, text=label, width=200, anchor="w",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MAIN).pack(side="left", padx=(0, 12))
        
        if field_type == 'date':
            entry = ctk.CTkEntry(frame, placeholder_text="ДД.ММ.ГГГГ", width=200, corner_radius=8,
                                 fg_color=BG_PANEL_2, border_color=BORDER)
            entry.pack(side="left")
            # Add calendar button
            ctk.CTkButton(
                frame, text="📅", width=38, height=30, corner_radius=8,
                fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
                command=lambda e=entry: self._pick_date(e)
            ).pack(side="left", padx=(6, 0))
        elif field_type == 'email':
            entry = ctk.CTkEntry(frame, placeholder_text="email@example.com", corner_radius=8,
                                 fg_color=BG_PANEL_2, border_color=BORDER)
            entry.pack(side="left", fill="x", expand=True)
        elif field_type == 'number':
            entry = ctk.CTkEntry(frame, placeholder_text="0", width=100, corner_radius=8,
                                 fg_color=BG_PANEL_2, border_color=BORDER)
            entry.pack(side="left")
        elif field_type == 'text':
            if field_name in ['comment', 'activation_info', 'payment_info', 'sent_info']:
                entry = ctk.CTkTextbox(frame, height=64, width=400, corner_radius=8,
                                       fg_color=BG_PANEL_2, border_color=BORDER, border_width=1)
                entry.pack(side="left", fill="x", expand=True)
            else:
                entry = ctk.CTkEntry(frame, corner_radius=8,
                                     fg_color=BG_PANEL_2, border_color=BORDER)
                entry.pack(side="left", fill="x", expand=True)
        else:
            entry = ctk.CTkEntry(frame, corner_radius=8,
                                 fg_color=BG_PANEL_2, border_color=BORDER)
            entry.pack(side="left", fill="x", expand=True)
        
        self.entries[field_name] = entry
        self._attach_field_copy_menu(entry)

    def _attach_field_copy_menu(self, entry):
        """Right-click context menu with 'Копировать' for any card input."""

        def on_right_click(event):
            menu = tk.Menu(self, tearoff=0, font=("Segoe UI", 10))

            def do_copy():
                text = ""
                if isinstance(entry, ctk.CTkTextbox):
                    try:
                        text = entry._textbox.selection_get()
                    except Exception:
                        text = entry.get("1.0", "end-1c")
                else:
                    try:
                        text = entry._entry.selection_get()
                    except Exception:
                        text = entry.get()
                if not text:
                    try:
                        text = entry.get() if not isinstance(entry, ctk.CTkTextbox) else entry.get("1.0", "end-1c")
                    except Exception:
                        text = ""
                copy_to_clipboard(self, text)

            menu.add_command(label="📋  Копировать", command=do_copy)
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
            return "break"

        targets = [entry]
        for attr in ("_entry", "_textbox"):
            inner = getattr(entry, attr, None)
            if inner is not None:
                targets.append(inner)
        for t in targets:
            try:
                t.bind("<Button-3>", on_right_click, add="+")
            except Exception:
                pass

    def _create_status_section(self):
        frame = ctk.CTkFrame(self.scroll_frame, fg_color=BG_PANEL, corner_radius=12)
        frame.pack(fill="x", pady=(18, 5))
        
        ctk.CTkLabel(
            frame, text="🏷  Статус лицензии", font=ctk.CTkFont(size=11, weight="bold"), anchor="w",
            text_color=TEXT_MUTED
        ).pack(anchor="w", padx=12, pady=(10, 2))
        
        # Auto status display
        auto_frame = ctk.CTkFrame(frame, fg_color="transparent")
        auto_frame.pack(fill="x", padx=12, pady=(2, 4))
        
        ctk.CTkLabel(auto_frame, text="Автоматически вычислен:", width=170, anchor="w",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED).pack(side="left")
        self.auto_status_label = ctk.CTkLabel(
            auto_frame, text="—", font=ctk.CTkFont(size=12, weight="bold"), text_color=TEXT_MAIN
        )
        self.auto_status_label.pack(side="left")

    # ===== Comments =====

    def _create_comments_section(self):
        self.comments_area = None
        self.new_comment_box = None

        self.comments_frame = ctk.CTkFrame(self.scroll_frame, fg_color=BG_PANEL, corner_radius=12)
        self.comments_frame.pack(fill="x", pady=(18, 5))

        ctk.CTkLabel(
            self.comments_frame, text="💬  Комментарии",
            font=ctk.CTkFont(size=11, weight="bold"), anchor="w",
            text_color=TEXT_MUTED
        ).pack(anchor="w", padx=12, pady=(10, 2))

        if self.is_new:
            ctk.CTkLabel(
                self.comments_frame,
                text="Комментарии будут доступны после сохранения карточки.",
                font=ctk.CTkFont(size=12), text_color=TEXT_MUTED, anchor="w", justify="left"
            ).pack(anchor="w", padx=12, pady=(2, 12))
            return

        # List of existing comments
        self.comments_area = ctk.CTkFrame(self.comments_frame, fg_color="transparent")
        self.comments_area.pack(fill="x", padx=12, pady=(2, 4))

        # Input for a new comment
        input_row = ctk.CTkFrame(self.comments_frame, fg_color="transparent")
        input_row.pack(fill="x", padx=12, pady=(4, 12))

        self.new_comment_box = ctk.CTkTextbox(
            input_row, height=56, corner_radius=8,
            fg_color=BG_PANEL_2, border_color=BORDER, border_width=1
        )
        self.new_comment_box.pack(side="left", fill="x", expand=True)
        self._attach_field_copy_menu(self.new_comment_box)

        ctk.CTkButton(
            input_row, text="＋ Добавить", width=110,
            corner_radius=8, fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=12), command=self._on_add_comment
        ).pack(side="left", padx=(8, 0), fill="y")

        self._reload_comments()

    @staticmethod
    def _format_comment_dt(iso_str):
        """Format stored ISO timestamp as ДД.ММ.ГГГГ ЧЧ:ММ."""
        if not iso_str:
            return ""
        try:
            return datetime.fromisoformat(iso_str).strftime("%d.%m.%Y %H:%M")
        except Exception:
            return str(iso_str)

    def _reload_comments(self):
        if self.comments_area is None:
            return
        for widget in self.comments_area.winfo_children():
            widget.destroy()

        comments = self.db.get_comments(self.license_data['id'])
        if not comments:
            ctk.CTkLabel(
                self.comments_area, text="Пока нет комментариев.",
                font=ctk.CTkFont(size=12), text_color=TEXT_MUTED
            ).pack(anchor="w", pady=(2, 6))
            return

        for comment in comments:
            self._create_comment_row(comment)

    def _create_comment_row(self, comment: Dict):
        row = ctk.CTkFrame(self.comments_area, fg_color=BG_PANEL_2, corner_radius=9)
        row.pack(fill="x", pady=4)

        body = ctk.CTkFrame(row, fg_color="transparent")
        body.pack(fill="x", padx=10, pady=(8, 2))
        ctk.CTkLabel(
            body, text=comment['text'], font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_MAIN, justify="left", wraplength=460, anchor="w"
        ).pack(anchor="w")

        footer = ctk.CTkFrame(row, fg_color="transparent")
        footer.pack(fill="x", padx=10, pady=(0, 8))

        stamp = self._format_comment_dt(comment.get('created_at'))
        updated = comment.get('updated_at')
        if updated and updated != comment.get('created_at'):
            stamp += f"   (изменён: {self._format_comment_dt(updated)})"
        ctk.CTkLabel(
            footer, text="🕒 " + stamp, font=ctk.CTkFont(size=10),
            text_color=TEXT_MUTED
        ).pack(side="left")

        ctk.CTkButton(
            footer, text="✏", width=32, height=24, corner_radius=6,
            fg_color="transparent", border_width=1, border_color=BORDER,
            hover_color=BG_PANEL, text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=11), command=lambda c=comment: self._on_edit_comment(c)
        ).pack(side="right")
        ctk.CTkButton(
            footer, text="🗑", width=32, height=24, corner_radius=6,
            fg_color="#e5484d", hover_color="#c73840", text_color="#ffffff",
            font=ctk.CTkFont(size=11), command=lambda c=comment: self._on_delete_comment(c)
        ).pack(side="right", padx=(0, 6))

    def _on_add_comment(self):
        text = self.new_comment_box.get("1.0", "end-1c").strip()
        if not text:
            messagebox.showwarning("Комментарий", "Введите текст комментария.", parent=self)
            return
        # created_at = current date and time, set on save
        self.db.add_comment(self.license_data['id'], text)
        self.new_comment_box.delete("1.0", "end")
        self._reload_comments()

    def _on_edit_comment(self, comment: Dict):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Редактировать комментарий")
        dlg.geometry("540x280")
        dlg.transient(self)
        dlg.configure(fg_color=BG_NAV)

        box = ctk.CTkTextbox(
            dlg, corner_radius=8, fg_color=BG_PANEL_2, border_color=BORDER, border_width=1,
            font=ctk.CTkFont(size=12)
        )
        box.pack(fill="both", expand=True, padx=14, pady=(14, 8))
        box.insert("1.0", comment['text'])

        def save():
            text = box.get("1.0", "end-1c").strip()
            if not text:
                messagebox.showwarning("Комментарий", "Текст не может быть пустым.", parent=dlg)
                return
            self.db.update_comment(comment['id'], text)  # updated_at = now
            dlg.destroy()
            self._reload_comments()

        btns = ctk.CTkFrame(dlg, fg_color="transparent")
        btns.pack(fill="x", padx=14, pady=(0, 14))
        ctk.CTkButton(
            btns, text="💾  Сохранить", command=save, width=130, height=34,
            corner_radius=8, fg_color=ACCENT, hover_color=ACCENT_HOVER
        ).pack(side="left")
        ctk.CTkButton(
            btns, text="Отмена", command=dlg.destroy, width=100, height=34,
            corner_radius=8, fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN
        ).pack(side="left", padx=8)

        dlg.update_idletasks()
        dlg.geometry(f"+{self.winfo_x() + 80}+{self.winfo_y() + 140}")
        dlg.after(100, dlg.grab_set)

    def _on_delete_comment(self, comment: Dict):
        if messagebox.askyesno("Подтверждение", "Удалить этот комментарий?", parent=self):
            self.db.delete_comment(comment['id'])
            self._reload_comments()
    
    def _pick_date(self, entry: ctk.CTkEntry):
        """Open an aligned date picker (День / Месяц / Год) with safe dropdowns."""
        picker = ctk.CTkToplevel(self)
        picker.title("Выбор даты")
        picker.geometry("360x230")
        picker.resizable(False, False)
        picker.transient(self)
        picker.configure(fg_color=BG_PANEL)

        # Current date or today
        current_text = entry.get().strip()
        try:
            if current_text:
                current_date = datetime.strptime(current_text, "%d.%m.%Y").date()
            else:
                current_date = date.today()
        except Exception:
            current_date = date.today()

        year_var = ctk.StringVar(value=str(current_date.year))
        month_var = ctk.StringVar(value=MONTH_NAMES[current_date.month - 1])
        day_var = ctk.StringVar(value=f"{current_date.day:02d}")

        HEADER_FONT = ctk.CTkFont(size=11, weight="bold")

        def make_col(parent, col, title, values, var, width):
            """One aligned column: header label + aligned dropdown."""
            cell = ctk.CTkFrame(parent, fg_color="transparent")
            cell.grid(row=0, column=col, padx=8, sticky="n")
            ctk.CTkLabel(cell, text=title, font=HEADER_FONT,
                         text_color=TEXT_MUTED, width=width).pack()
            menu = DropdownOption(
                cell, values, var, width=width, height=32
            )
            menu.pack(pady=(6, 0))
            return menu

        fields = ctk.CTkFrame(picker, fg_color="transparent")
        fields.pack(padx=14, pady=(16, 4))
        fields.grid_columnconfigure(0, weight=1)
        fields.grid_columnconfigure(1, weight=1)
        fields.grid_columnconfigure(2, weight=1)

        make_col(fields, 0, "ДЕНЬ",  [f"{d:02d}" for d in range(1, 32)], day_var, 64)
        make_col(fields, 1, "МЕСЯЦ", MONTH_NAMES, month_var, 108)
        make_col(fields, 2, "ГОД",   [str(y) for y in range(2020, 2040)], year_var, 80)

        def current_month_days():
            y, m = int(year_var.get()), MONTH_NAMES.index(month_var.get()) + 1
            return calendar.monthrange(y, m)[1]

        # Keep the selected day valid when month/year changes (e.g. 31.02 → 28.02)
        def clamp_day():
            try:
                day = int(day_var.get())
            except ValueError:
                return
            max_day = current_month_days()
            if day > max_day:
                day_var.set(f"{max_day:02d}")

        def on_ok():
            try:
                d = date(int(year_var.get()),
                         MONTH_NAMES.index(month_var.get()) + 1,
                         int(day_var.get()))
                entry.delete(0, "end")
                entry.insert(0, d.strftime("%d.%m.%Y"))
                picker.destroy()
            except ValueError:
                messagebox.showerror("Ошибка", "Неверная дата", parent=picker)

        year_var.trace_add("write", lambda *a: clamp_day())
        month_var.trace_add("write", lambda *a: clamp_day())

        btns = ctk.CTkFrame(picker, fg_color="transparent")
        btns.pack(pady=18)
        ctk.CTkButton(
            btns, text="✔ Выбрать", command=on_ok, width=130, height=36,
            corner_radius=8, fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            btns, text="Очистить", width=110, height=36, corner_radius=8,
            fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=12),
            command=lambda: (entry.delete(0, "end"), picker.destroy())
        ).pack(side="left", padx=5)

        # Center picker
        picker.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (picker.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (picker.winfo_height() // 2)
        picker.geometry(f"+{x}+{y}")
        picker.after(100, picker.grab_set)
    
    def _load_data(self):
        if not self.license_data:
            return
        
        for field_name, entry in self.entries.items():
            value = self.license_data.get(field_name)
            if value is None:
                continue
            
            if isinstance(entry, ctk.CTkTextbox):
                entry.delete("1.0", "end")
                entry.insert("1.0", str(value))
            elif isinstance(entry, ctk.CTkEntry):
                # Format dates to ДД.ММ.ГГГГ
                if field_name in ['install_date', 'id_expiry', 'payment_date', 'renewal_date', 'next_payment']:
                    if isinstance(value, str) and value:
                        try:
                            if 'T' in value:
                                value = datetime.fromisoformat(value.replace('Z', '')).strftime("%d.%m.%Y")
                            else:
                                value = datetime.strptime(value[:10], "%Y-%m-%d").strftime("%d.%m.%Y")
                        except:
                            pass
                entry.delete(0, "end")
                entry.insert(0, str(value))
        
        # Update auto status display
        auto_status = self.license_data.get('status', 'active')
        try:
            status_enum = LicenseStatus(auto_status)
        except ValueError:
            status_enum = LicenseStatus.ACTIVE
        self.auto_status_label.configure(
            text=STATUS_LABELS.get(status_enum, auto_status),
            text_color=STATUS_COLORS.get(status_enum, "white")
        )
    
    def _collect_data(self) -> Dict[str, Any]:
        data = {'sheet_name': self.sheet_name, 'category': self.category_var.get()}
        
        for field_name, entry in self.entries.items():
            if isinstance(entry, ctk.CTkTextbox):
                value = entry.get("1.0", "end-1c").strip()
            else:
                value = entry.get().strip()
            
            # Parse dates to ISO format for storage
            if field_name in ['install_date', 'id_expiry', 'payment_date', 'renewal_date', 'next_payment']:
                if value:
                    try:
                        parsed = datetime.strptime(value, "%d.%m.%Y").date()
                        value = parsed.isoformat()
                    except:
                        pass  # Keep as string if parsing fails
                else:
                    value = None  # Clear the date field
            
            # Always include the field (even empty) so clearing works
            data[field_name] = value if value else None
        
        return data
    
    def _on_key_press_card(self, event):
        """Layout-independent Ctrl+C (works on Russian keyboard too)."""
        try:
            ks = str(getattr(event, "keysym", "") or "")
            kc = getattr(event, "keycode", None)
            st = getattr(event, "state", 0) or 0
            try:
                import os as _os, tempfile as _tf
                log = _os.path.join(_tf.gettempdir(), "lt_keys.txt")
                with open(log, "a", encoding="utf-8") as _f:
                    _f.write(f"CARD keycode={kc} keysym={ks!r} state={st} (ctrl={bool(st & 0x4)})\n")
            except Exception:
                pass
            if st & 0x4:
                low = ks.lower()
                c_keysyms = ("c", "u0441", "cyrillic_es", "cyrillic c", "с")
                if kc == 67 or low in c_keysyms or "es" in low or low.endswith("c") or low.endswith("с"):
                    return self._on_ctrl_c_card(event)
        except Exception:
            pass
        return None

    def _on_ctrl_c_card(self, event=None):
        """Ctrl+C in the card: copy selected field text, else the whole card."""
        for entry in getattr(self, "entries", {}).values():
            inner = getattr(entry, "_entry", None)
            if inner is None:
                inner = getattr(entry, "_textbox", None)
            if inner is None:
                continue
            try:
                sel = inner.selection_get()
            except Exception:
                sel = ""
            if sel:
                copy_to_clipboard(self, sel)
                return "break"
        # No selection in any field -> copy the whole card data
        self._on_copy_card()
        return "break"

    def _on_copy_card(self):
        """Copy name / email / store / phone / license ID to the clipboard."""
        text = format_card_copy_text(self._collect_data())
        if not copy_to_clipboard(self, text):
            messagebox.showwarning(
                "Копирование",
                "В карточке нет данных для копирования",
                parent=self
            )
            return
        self._flash_copy_button()

    def _flash_copy_button(self):
        """Brief visual confirmation without a blocking dialog."""
        try:
            self.copy_btn.configure(text="✓  Скопировано")
        except Exception:
            return

        def restore():
            try:
                if self.copy_btn.winfo_exists():
                    self.copy_btn.configure(text="📋  Копировать")
            except Exception:
                pass

        self.after(1500, restore)

    def _copy_field_value(self, field_name):
        """Copy a single field's value to the clipboard."""
        entry = self.entries.get(field_name)
        btn = self._field_copy_buttons.get(field_name)
        if entry is None:
            return
        try:
            if isinstance(entry, ctk.CTkTextbox):
                text = entry.get("1.0", "end-1c").strip()
            else:
                text = entry.get().strip()
        except Exception:
            text = ""
        if not text:
            return
        copy_to_clipboard(self, text)
        # Brief visual feedback on the button
        if btn is not None:
            try:
                btn.configure(text="✓")
            except Exception:
                pass
            def restore(b=btn):
                try:
                    if b.winfo_exists():
                        b.configure(text="📋")
                except Exception:
                    pass
            self.after(1200, restore)

    def _on_save(self):
        data = self._collect_data()
        
        # Validate required fields
        if not data.get('full_name'):
            messagebox.showerror("Ошибка", "Заполните поле 'Фамилия и имя'", parent=self)
            return
        
        try:
            if self.is_new:
                license_id = self.db.create_license(data)
                self.result = ('create', license_id)
            else:
                self.db.update_license(self.license_data['id'], data)
                self.result = ('update', self.license_data['id'])
            self.destroy()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить: {e}", parent=self)
    
    def _on_cancel(self):
        self.result = None
        self.destroy()
    
    def _on_delete(self):
        if messagebox.askyesno("Подтверждение", "Удалить эту запись?", parent=self):
            try:
                self.db.delete_license(self.license_data['id'])
                self.result = ('delete', self.license_data['id'])
                self.destroy()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось удалить: {e}", parent=self)


class MoveCategoryDialog(ctk.CTkToplevel):
    """Small dialog for moving selected records to another category."""

    def __init__(self, parent, count, current_category, on_move):
        super().__init__(parent)
        self.on_move = on_move

        self.title("Переместить в категорию")
        self.geometry("540x320")
        self.resizable(False, False)
        self.transient(parent)
        self.configure(fg_color=BG_PANEL)
        self.after(100, self.grab_set)

        header = ctk.CTkFrame(self, corner_radius=0, fg_color=BG_NAV, height=64)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(
            header, text="📁  Перемещение в категорию",
            font=ctk.CTkFont(size=15, weight="bold"), text_color=TEXT_MAIN
        ).pack(side="left", padx=16)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=16)

        ctk.CTkLabel(
            body, text=f"Выбрано записей: {count}", anchor="w",
            font=ctk.CTkFont(size=13, weight="bold"), text_color=TEXT_MAIN
        ).pack(fill="x")

        ctk.CTkLabel(
            body, text="Выберите новую категорию:", anchor="w",
            font=ctk.CTkFont(size=12), text_color=TEXT_MUTED
        ).pack(fill="x", pady=(10, 6))

        self.category_var = ctk.StringVar(
            value=current_category if current_category in CATEGORIES else CATEGORIES[0]
        )
        self.category_seg = ctk.CTkSegmentedButton(
            body, values=CATEGORIES, variable=self.category_var, height=36,
            font=ctk.CTkFont(size=12, weight="bold"),
            selected_color=ACCENT, selected_hover_color=ACCENT_HOVER,
            unselected_color=BG_PANEL_2, unselected_hover_color=BORDER,
            text_color=TEXT_MAIN, fg_color=BG_PANEL_2, border_width=0,
            command=self._on_selected
        )
        self.category_seg.pack(fill="x", pady=4)

        self.hint_label = ctk.CTkLabel(
            body, text="", anchor="w",
            font=ctk.CTkFont(size=11), text_color=TEXT_MUTED
        )
        self.hint_label.pack(fill="x", pady=(8, 0))

        buttons = ctk.CTkFrame(body, fg_color="transparent")
        buttons.pack(fill="x", side="bottom", pady=(14, 0))

        self.move_btn = ctk.CTkButton(
            buttons, text="Переместить", width=150, height=36, corner_radius=8,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._confirm
        )
        self.move_btn.pack(side="left")

        ctk.CTkButton(
            buttons, text="Отмена", width=110, height=36, corner_radius=8,
            fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=12), command=self.destroy
        ).pack(side="left", padx=8)

        self._start_category = current_category
        self.bind("<Escape>", lambda e: self.destroy())
        self._on_selected(self.category_var.get())
        self._center(parent)

    def _on_selected(self, value=None):
        value = value or self.category_var.get()
        # Nothing to move if the target category is the current one
        same = value == self._start_category
        try:
            self.move_btn.configure(state="disabled" if same else "normal")
            self.hint_label.configure(
                text="Записи уже находятся в этой категории"
                if same else f"Записи будут перенесены в «{value}»"
            )
        except Exception:
            pass

    def _confirm(self):
        category = self.category_var.get()
        self.destroy()
        self.on_move(category)

    def _center(self, parent):
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")


# ===== Excel import with column mapping =====

NO_COLUMN = "— не заполнять —"

IMPORT_CATEGORIES = CATEGORIES

DATE_FIELDS = {"install_date", "id_expiry", "payment_date", "renewal_date", "next_payment"}

# (db field, label shown in UI, extra synonyms for auto-matching)
IMPORT_FIELD_DEFS = [
    ("full_name", "ФИО",
     ["фамилия и имя", "фио клиента", "имя", "клиент", "name"]),
    ("store_name", "Название магазина",
     ["магазин", "торговая точка", "название торговой точки", "точка", "store"]),
    ("phone", "Телефон",
     ["тел", "номер телефона", "телефон 1", "phone", "mobile"]),
    ("email", "Почта",
     ["email", "e-mail", "мейл", "электронная почта", "почта клиента"]),
    ("gbs_id", "GBS ID",
     ["gbsid", "gbs-id", "id лицензии", "license id"]),
    ("license_key", "Ключ",
     ["ключ лицензии", "лицензионный ключ", "license key", "key"]),
    ("install_date", "Дата установки",
     ["дата установки", "установка", "install date"]),
    ("id_expiry", "Срок ID",
     ["срок id", "срок действия", "действует до", "expiry", "expiration"]),
    ("payment_date", "Дата оплаты",
     ["дата оплаты", "срок когда оплатил", "оплатил", "payment date"]),
    ("renewal_date", "Дата продления",
     ["дата продления", "продление"]),
    ("next_payment", "Следующая оплата",
     ["следующая оплата", "след оплата", "след. оплата", "next payment"]),
    ("activation_period", "Срок активации",
     ["срок активации", "период активации"]),
    ("address", "Адрес", ["адрес"]),
    ("row_number", "№", ["номер", "п/п"]),
    ("sent_info", "Отправлено", ["отправлено"]),
    ("activation_info", "Активация и срок", ["активация и срок"]),
    ("payment_info", "Оплата (текст)", ["оплата", "payment info"]),
    ("receipt_number", "№ квитанции", ["квитанция", "номер квитанции"]),
    ("migrated_from", "Переход с другой программы",
     ["переход с другой программы", "переход"]),
    ("comment", "Комментарий", ["примечание", "коммент"]),
]


def _norm_header(text) -> str:
    """Normalise a column header for fuzzy matching."""
    value = str(text or "").strip().lower().replace("ё", "е")
    value = re.sub(r"[^0-9a-zа-я]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def auto_map_columns(headers):
    """Guess which file column corresponds to each card field."""
    cols = [(str(h).strip(), _norm_header(h)) for h in headers]
    mapping = {}
    used = set()

    def candidates(field_def):
        return [field_def[1]] + list(field_def[2])

    # Pass 1: exact header matches
    for field_def in IMPORT_FIELD_DEFS:
        field = field_def[0]
        for cand in candidates(field_def):
            norm_cand = _norm_header(cand)
            if not norm_cand:
                continue
            found = None
            for col, norm in cols:
                if col in used or not norm:
                    continue
                if norm == norm_cand:
                    found = col
                    break
            if found:
                mapping[field] = found
                used.add(found)
                break

    # Pass 2: partial matches for the remaining fields
    for field_def in IMPORT_FIELD_DEFS:
        field = field_def[0]
        if field in mapping:
            continue
        best = None
        best_score = 0
        for cand in candidates(field_def):
            norm_cand = _norm_header(cand)
            if len(norm_cand) < 4:
                continue
            for col, norm in cols:
                if col in used or not norm:
                    continue
                if norm_cand in norm or norm in norm_cand:
                    if len(norm_cand) > best_score:
                        best, best_score = col, len(norm_cand)
        if best:
            mapping[field] = best
            used.add(best)

    return mapping


def normalize_frame_columns(df):
    """Give every DataFrame column a non-empty, unique string name."""
    names = []
    seen = set()
    for idx, column in enumerate(df.columns):
        name = str(column).strip()
        if not name or name.lower().startswith("unnamed"):
            name = f"Колонка {idx + 1}"
        base = name
        counter = 2
        while name in seen:
            name = f"{base} ({counter})"
            counter += 1
        seen.add(name)
        names.append(name)
    df = df.copy()
    df.columns = names
    return df


def _is_blank(value) -> bool:
    if value is None:
        return True
    try:
        if value != value:  # NaN
            return True
    except Exception:
        pass
    text = str(value).strip()
    return text == "" or text.lower() in ("nan", "nat", "none")


def build_records_from_frame(df, field_to_column):
    """Turn DataFrame rows into DB records using the chosen column mapping."""
    records = []
    for _, row in df.iterrows():
        record = {}
        for field, column in field_to_column.items():
            if not column or column == NO_COLUMN:
                continue
            try:
                raw = row[column]
            except (KeyError, IndexError):
                continue
            if _is_blank(raw):
                continue

            if field in DATE_FIELDS:
                value = parse_excel_date(raw)
            elif field == "row_number":
                try:
                    value = int(float(raw))
                except (TypeError, ValueError):
                    value = None
            elif isinstance(raw, float) and raw.is_integer():
                value = str(int(raw))  # phones/IDs read as 8747.0
            else:
                value = str(raw).strip()

            if value is not None and str(value).strip() != "":
                record[field] = value

        # Skip fully empty rows (a lone row number is not a record)
        if set(record) - {"row_number"}:
            records.append(record)
    return records


class ImportDialog(ctk.CTkToplevel):
    """Excel import wizard: pick a file, preview it, map columns, choose category."""

    def __init__(self, parent, db: Database, default_category="GBS-market", on_done=None):
        super().__init__(parent)
        self.db = db
        self.on_done = on_done
        self.file_path = None
        self.sheets = {}
        self.headers = []
        self.field_vars = {}
        self.field_menus = {}

        self.title("Импорт клиентов из Excel")
        self.geometry("860x900")
        self.minsize(780, 640)
        self.transient(parent)
        self.configure(fg_color=BG_APP)
        self.after(100, self.grab_set)

        self._build_ui(default_category)
        self._center(parent)

    # ---------- UI ----------

    def _build_ui(self, default_category):
        header = ctk.CTkFrame(self, height=64, corner_radius=0, fg_color=BG_NAV)
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkLabel(
            header, text="📥  Импорт клиентов из Excel",
            font=ctk.CTkFont(size=16, weight="bold"), text_color=TEXT_MAIN
        ).pack(side="left", padx=18, pady=18)

        ctk.CTkButton(
            header, text="Скачать шаблон", width=150, height=32, corner_radius=8,
            fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=12), command=self._download_template
        ).pack(side="right", padx=16, pady=16)

        footer = ctk.CTkFrame(self, height=64, corner_radius=0, fg_color=BG_NAV)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        self.import_btn = ctk.CTkButton(
            footer, text="✅  Импортировать", width=170, height=38, corner_radius=8,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._do_import, state="disabled"
        )
        self.import_btn.pack(side="left", padx=16, pady=12)

        ctk.CTkButton(
            footer, text="Отмена", width=110, height=38, corner_radius=8,
            fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=12), command=self.destroy
        ).pack(side="left", padx=6, pady=12)

        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent", corner_radius=0)
        self.scroll.pack(fill="both", expand=True)

        # --- file card ---
        file_card = ctk.CTkFrame(self.scroll, corner_radius=14, fg_color=BG_PANEL)
        file_card.pack(fill="x", padx=20, pady=(16, 10))

        ctk.CTkLabel(
            file_card, text="Файл Excel / CSV", anchor="w",
            font=ctk.CTkFont(size=13, weight="bold"), text_color=TEXT_MAIN
        ).pack(fill="x", padx=16, pady=(14, 2))

        ctk.CTkLabel(
            file_card,
            text="Поддерживаются .xlsx и .csv. Первая строка файла должна содержать заголовки колонок.",
            anchor="w", font=ctk.CTkFont(size=11), text_color=TEXT_MUTED
        ).pack(fill="x", padx=16)

        file_row = ctk.CTkFrame(file_card, fg_color="transparent")
        file_row.pack(fill="x", padx=16, pady=(12, 6))

        ctk.CTkButton(
            file_row, text="Выбор файла", width=130, height=34, corner_radius=8,
            fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=12), command=self._choose_file
        ).pack(side="left")

        self.file_label = ctk.CTkLabel(
            file_row, text="Файл не выбран", anchor="w",
            font=ctk.CTkFont(size=12), text_color=TEXT_MUTED
        )
        self.file_label.pack(side="left", padx=12)

        self.load_btn = ctk.CTkButton(
            file_card, text="Загрузить и показать предпросмотр",
            height=38, corner_radius=8, fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._load_preview, state="disabled"
        )
        self.load_btn.pack(anchor="w", padx=16, pady=(6, 16))
        # --- preview card (hidden until a file is loaded) ---
        self.preview_card = ctk.CTkFrame(self.scroll, corner_radius=14, fg_color=BG_PANEL)

        preview_top = ctk.CTkFrame(self.preview_card, fg_color="transparent")
        preview_top.pack(fill="x", padx=16, pady=(14, 4))

        self.preview_title = ctk.CTkLabel(
            preview_top, text="Предпросмотр", anchor="w",
            font=ctk.CTkFont(size=14, weight="bold"), text_color=TEXT_MAIN
        )
        self.preview_title.pack(side="left")

        self.rows_label = ctk.CTkLabel(
            preview_top, text="", anchor="e",
            font=ctk.CTkFont(size=11), text_color=TEXT_MUTED
        )
        self.rows_label.pack(side="right")

        sheet_row = ctk.CTkFrame(self.preview_card, fg_color="transparent")
        sheet_row.pack(fill="x", padx=16, pady=(6, 2))

        ctk.CTkLabel(
            sheet_row, text="Лист Excel:", font=ctk.CTkFont(size=12), text_color=TEXT_MUTED
        ).pack(side="left", padx=(0, 8))

        self.sheet_var = ctk.StringVar(value="")
        self.sheet_menu = DropdownOption(
            sheet_row, [], self.sheet_var, command=self._on_sheet_change,
            width=240, height=30
        )
        self.sheet_menu.pack(side="left")

        ctk.CTkLabel(
            self.preview_card, text="СОПОСТАВЛЕНИЕ КОЛОНОК", anchor="w",
            font=ctk.CTkFont(size=11, weight="bold"), text_color=TEXT_MUTED
        ).pack(fill="x", padx=16, pady=(14, 4))

        head = ctk.CTkFrame(self.preview_card, fg_color="transparent")
        head.pack(fill="x", padx=16)
        ctk.CTkLabel(
            head, text="ПОЛЕ КАРТОЧКИ", width=240, anchor="w",
            font=ctk.CTkFont(size=10, weight="bold"), text_color=TEXT_MUTED
        ).pack(side="left")
        ctk.CTkLabel(
            head, text="КОЛОНКА В ФАЙЛЕ", width=320, anchor="w",
            font=ctk.CTkFont(size=10, weight="bold"), text_color=TEXT_MUTED
        ).pack(side="left", padx=8)

        mapping_frame = ctk.CTkFrame(self.preview_card, fg_color="transparent")
        mapping_frame.pack(fill="x", padx=16, pady=(4, 8))

        for field, label, _aliases in IMPORT_FIELD_DEFS:
            row = ctk.CTkFrame(mapping_frame, fg_color="transparent")
            row.pack(fill="x", pady=1)

            ctk.CTkLabel(
                row, text=label, width=240, anchor="w",
                font=ctk.CTkFont(size=12), text_color=TEXT_MAIN
            ).pack(side="left")

            var = ctk.StringVar(value=NO_COLUMN)
            menu = DropdownOption(row, [NO_COLUMN], var, width=320, height=30)
            menu.pack(side="left", padx=8)

            self.field_vars[field] = var
            self.field_menus[field] = menu

        ctk.CTkLabel(
            self.preview_card, text="КАТЕГОРИЯ", anchor="w",
            font=ctk.CTkFont(size=11, weight="bold"), text_color=TEXT_MUTED
        ).pack(fill="x", padx=16, pady=(10, 2))

        self.category_var = ctk.StringVar(value=default_category)
        self.category_seg = ctk.CTkSegmentedButton(
            self.preview_card, values=IMPORT_CATEGORIES,
            variable=self.category_var, height=34,
            font=ctk.CTkFont(size=12),
            selected_color=ACCENT, selected_hover_color=ACCENT_HOVER,
            unselected_color=BG_PANEL_2, unselected_hover_color=BORDER,
            text_color=TEXT_MAIN, fg_color=BG_PANEL_2, border_width=0
        )
        self.category_seg.pack(anchor="center", pady=(4, 16))

    def _center(self, parent):
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    # ---------- File loading ----------

    def _choose_file(self):
        path = filedialog.askopenfilename(
            title="Выберите файл Excel или CSV",
            filetypes=[
                ("Excel / CSV", "*.xlsx *.xlsm *.xls *.csv"),
                ("Excel", "*.xlsx *.xlsm *.xls"),
                ("CSV", "*.csv"),
                ("Все файлы", "*.*"),
            ]
        )
        if not path:
            return
        self.file_path = path
        self.file_label.configure(text=os.path.basename(path), text_color=TEXT_MAIN)
        self.load_btn.configure(state="normal")
        self.import_btn.configure(state="disabled")
        self.preview_card.pack_forget()
        self.sheets = {}

    def _load_preview(self):
        if not self.file_path:
            return
        try:
            import pandas as pd
        except ImportError:
            messagebox.showerror(
                "Ошибка", "Не установлена библиотека pandas.\nУстановите: pip install pandas openpyxl",
                parent=self
            )
            return

        try:
            if self.file_path.lower().endswith(".csv"):
                frame = pd.read_csv(self.file_path, dtype=object)
                sheets = {"CSV": frame}
            else:
                sheets = pd.read_excel(self.file_path, sheet_name=None, dtype=object)
        except Exception as e:
            messagebox.showerror(
                "Ошибка чтения", f"Не удалось прочитать файл:\n{e}", parent=self
            )
            return

        sheets = {
            name: normalize_frame_columns(frame)
            for name, frame in sheets.items()
            if len(frame.columns) > 0 and len(frame) > 0
        }
        if not sheets:
            messagebox.showwarning(
                "Предпросмотр", "В файле не найдено данных.", parent=self
            )
            return

        self.sheets = sheets
        self.preview_card.pack(fill="x", padx=20, pady=(4, 16))

        names = list(self.sheets.keys())
        self.sheet_menu.set_values(names)
        if self.sheet_var.get() not in names:
            self.sheet_var.set(names[0])
        self._apply_sheet(self.sheet_var.get())
        self.import_btn.configure(state="normal")

    def _on_sheet_change(self, sheet_name=None):
        sheet_name = sheet_name or self.sheet_var.get()
        self._apply_sheet(sheet_name)

    def _apply_sheet(self, sheet_name):
        frame = self.sheets.get(sheet_name)
        if frame is None:
            return

        self.headers = [str(c) for c in frame.columns]
        self.preview_title.configure(
            text=f"Предпросмотр: {os.path.basename(self.file_path)}"
        )
        self.rows_label.configure(text=f"строк в файле: {len(frame)}")

        values = [NO_COLUMN] + self.headers
        mapping = auto_map_columns(self.headers)

        for field, var in self.field_vars.items():
            self.field_menus[field].set_values(values)
            var.set(mapping.get(field, NO_COLUMN))

    # ---------- Import ----------

    def _do_import(self):
        if not self.sheets:
            return

        sheet = self.sheet_var.get()
        frame = self.sheets.get(sheet)
        if frame is None:
            return

        mapping = {
            field: var.get()
            for field, var in self.field_vars.items()
            if var.get() and var.get() != NO_COLUMN
        }
        if not mapping:
            messagebox.showwarning(
                "Импорт", "Сопоставьте хотя бы одну колонку с полем карточки.", parent=self
            )
            return

        category = self.category_var.get()
        records = build_records_from_frame(frame, mapping)
        if not records:
            messagebox.showwarning(
                "Импорт", "В выбранных колонках нет данных для импорта.", parent=self
            )
            return

        try:
            imported, updated, skipped = self.db.bulk_upsert(records, sheet, category)
            self.db.log_import(self.file_path, sheet, imported, updated, skipped)
        except Exception as e:
            messagebox.showerror("Ошибка импорта", str(e), parent=self)
            return

        messagebox.showinfo(
            "Импорт завершён",
            f"Категория: {category}\nЛист: {sheet}\n\n"
            f"Новых: {imported}\nОбновлено: {updated}\nПропущено: {skipped}",
            parent=self
        )

        if self.on_done:
            try:
                self.on_done()
            except Exception:
                pass
        self.destroy()

    def _download_template(self):
        path = filedialog.asksaveasfilename(
            title="Сохранить шаблон импорта",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="Шаблон_импорта.xlsx"
        )
        if not path:
            return
        try:
            import openpyxl
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Импорт"
            ws.append([label for _field, label, _aliases in IMPORT_FIELD_DEFS])
            wb.save(path)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e), parent=self)
            return
        messagebox.showinfo("Шаблон", f"Шаблон сохранён:\n{path}", parent=self)


# (Low-level keyboard hook removed — it blocked Ctrl+C globally and triggered AV).
# Ctrl+C on Russian layout is handled by _on_key_press below using keycode (physical key).


class LicenseTrackerApp(ctk.CTk):
    """Main application window."""
    
    def __init__(self):
        super().__init__()
        self.db = Database(DB_PATH)
        self._migrate_otvaly_category()
        try:
            install_keyboard_hook()
            self.after(500, lambda: set_app_hwnd(self.winfo_id()))
        except Exception:
            pass
        
        self.title(APP_TITLE)
        self.geometry(f"{APP_WIDTH}x{APP_HEIGHT}")
        self.minsize(1000, 600)
        
        # Set theme (loaded from settings)
        ctk.set_default_color_theme("blue")
        self.current_theme = self.db.get_setting("theme", "dark") or "dark"
        if self.current_theme not in THEMES:
            self.current_theme = "dark"
        apply_theme(self.current_theme)
        ctk.set_appearance_mode("dark" if self.current_theme == "dark" else "light")
        
        # State
        self.current_category = "GBS-market"
        self.current_sheet = ""
        self.sort_column = ""
        self.sort_reverse = False
        self.filter_text = ""
        self.status_filter = "Все"
        self.days_filter = "Все"
        
        self._build_ui()
        self._load_categories()
        self._load_sheets()
        self._refresh_data()
        self._update_stats()
        self._bind_shortcuts()
        # Auto-backup to Documents on startup
        self.after(1500, self._auto_backup_to_documents)
        # Background update check on launch (only for frozen builds)
        self.after(2500, self._auto_update_check)
    
    def _build_ui(self):
        self.configure(fg_color=BG_APP)
        
        # ===== Top navigation bar (стиль референса) =====
        nav = ctk.CTkFrame(self, height=68, corner_radius=0, fg_color=BG_NAV)
        nav.pack(fill="x")
        nav.pack_propagate(False)
        
        # Brand (left)
        ctk.CTkLabel(
            nav, text="🔑 База лицензий",
            font=ctk.CTkFont(size=14, weight="bold"), text_color=TEXT_MAIN
        ).pack(side="left", padx=(16, 10))

        # File menu button
        ctk.CTkButton(
            nav, text="📁  Файл", width=80, height=36, corner_radius=8,
            fg_color="transparent", hover_color=BG_PANEL_2,
            text_color=TEXT_MUTED, font=ctk.CTkFont(size=12),
            command=self._show_file_menu
        ).pack(side="left", padx=0, pady=14)

        # Columns visibility button
        ctk.CTkButton(
            nav, text="📊  Столбцы", width=100, height=36, corner_radius=8,
            fg_color="transparent", hover_color=BG_PANEL_2,
            text_color=TEXT_MUTED, font=ctk.CTkFont(size=12),
            command=self._show_columns_dialog
        ).pack(side="left", padx=0, pady=14)

        # Category nav items (centered)
        self.category_frame = ctk.CTkFrame(nav, fg_color="transparent")
        self.category_frame.pack(side="left", fill="both", expand=True, pady=14)
        self.category_inner = ctk.CTkFrame(self.category_frame, fg_color="transparent")
        self.category_inner.place(relx=0.5, rely=0.5, anchor="center")

        # Right side: version badge + refresh + theme toggle
        ctk.CTkLabel(
            nav, text=f"v{APP_VERSION}", height=24, corner_radius=12,
            fg_color=BG_NAV, border_width=1, border_color=BORDER,
            font=ctk.CTkFont(size=11), text_color=TEXT_MUTED
        ).pack(side="right", padx=(0, 14), pady=17)

        ctk.CTkButton(
            nav, text="⬆", width=36, height=32, corner_radius=8,
            fg_color=BG_PANEL_2, hover_color=BORDER,
            font=ctk.CTkFont(size=14), text_color=TEXT_MAIN,
            command=self._check_updates_manual
        ).pack(side="right", padx=(0, 6), pady=13)

        ctk.CTkButton(
            nav, text="↻", width=36, height=32, corner_radius=8,
            fg_color=BG_PANEL_2, hover_color=BORDER,
            font=ctk.CTkFont(size=14), text_color=TEXT_MAIN,
            command=self._refresh_all
        ).pack(side="right", padx=(0, 12), pady=13)

        self.theme_btn = ctk.CTkButton(
            nav, text="☀" if self.current_theme == "dark" else "🌙",
            width=36, height=32, corner_radius=8,
            fg_color=BG_PANEL_2, hover_color=BORDER,
            font=ctk.CTkFont(size=13), text_color=TEXT_MAIN,
            command=self._toggle_theme
        )
        self.theme_btn.pack(side="right", padx=(0, 6), pady=13)


        
        # ===== Content container =====
        page = ctk.CTkFrame(self, fg_color="transparent")
        page.pack(fill="both", expand=True)
        
        # ===== Toolbar card =====
        toolbar = ctk.CTkFrame(page, corner_radius=14, fg_color=BG_PANEL)
        toolbar.pack(fill="x", padx=24, pady=(18, 12))
        
        tf = ctk.CTkFrame(toolbar, fg_color="transparent", height=64)
        tf.pack(fill="x")
        tf.pack_propagate(False)
        
        # Right group (rightmost first): Удалить → Добавить → Импорт → Экспорт → Продлить
        ctk.CTkButton(
            tf, text="🗑 Удалить", command=self._on_delete_selected,
            width=100, height=36, corner_radius=8,
            fg_color="#e5484d", hover_color="#c73840", text_color="#ffffff",
            font=ctk.CTkFont(size=13)
        ).pack(side="right", padx=(6, 16), pady=14)

        ctk.CTkButton(
            tf, text="＋ Добавить", command=self._on_add,
            width=120, height=36, corner_radius=8,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=13)
        ).pack(side="right", padx=4, pady=14)

        ctk.CTkButton(
            tf, text="📥 Импорт", command=self._on_import,
            width=100, height=36, corner_radius=8,
            fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=13)
        ).pack(side="right", padx=4, pady=14)

        ctk.CTkButton(
            tf, text="📤 Экспорт", command=self._on_export,
            width=100, height=36, corner_radius=8,
            fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=13)
        ).pack(side="right", padx=4, pady=14)

        ctk.CTkButton(
            tf, text="⏳ Продлить", command=self._on_renew_selected,
            width=110, height=36, corner_radius=8,
            fg_color="#0e9f6e", hover_color="#0c8a5f", text_color="#ffffff",
            font=ctk.CTkFont(size=13)
        ).pack(side="right", padx=4, pady=14)

        ctk.CTkButton(
            tf, text="📁 Переместить", command=self._on_move_selected,
            width=130, height=36, corner_radius=8,
            fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=13)
        ).pack(side="right", padx=4, pady=14)

        # Search (left side)
        self.search_var = ctk.StringVar()
        self.search_entry = ctk.CTkEntry(
            tf, placeholder_text="🔍  Поиск: ФИО, магазин, телефон, ключ…",
            textvariable=self.search_var, width=280, height=36, corner_radius=8,
            fg_color=BG_PANEL_2, border_width=1, border_color=BORDER
        )
        self.search_entry.pack(side="left", padx=(16, 4), pady=14)
        self.search_var.trace_add("write", lambda *args: self._on_filter_change())
        
        # ===== Table card =====
        content = ctk.CTkFrame(page, corner_radius=14, fg_color=BG_PANEL)
        content.pack(fill="both", expand=True, padx=24, pady=(0, 12))
        
        # Style ttk.Treeview
        self._configure_tree_style()
        
        tree_frame = ctk.CTkFrame(content, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Define columns
        self.columns = [
            ("full_name", "ФИО", 190),
            ("store_name", "Магазин", 155),
            ("phone", "Телефон", 135),
            ("email", "Email", 185),
            ("gbs_id", "GBS ID", 205),
            ("license_key", "Ключ", 185),
            ("activation_period", "Срок активации", 105),
            ("id_expiry", "Срок ID", 105),
            ("next_payment", "След. оплата", 110),
            ("days_left", "Осталось", 95),
            ("status", "Статус", 125),
        ]

        # Load column visibility settings (default: all visible)
        self._column_visibility = {}
        for col_id, _label, _width in self.columns:
            val = self.db.get_setting(f"column_visible_{col_id}", "1")
            self._column_visibility[col_id] = val == "1"

        col_ids = [c[0] for c in self.columns]
        self.tree = ttk.Treeview(
            tree_frame, columns=col_ids, show="headings",
            selectmode="extended", style="Modern.Treeview"
        )

        for col_id, col_name, col_width in self.columns:
            self.tree.heading(col_id, text=col_name, command=lambda c=col_id: self._on_column_sort(c))
            if self._column_visibility.get(col_id, True):
                self.tree.column(col_id, width=col_width, minwidth=55, anchor="w")
            else:
                self.tree.column(col_id, width=0, minwidth=0, stretch=False)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Configure tags
        for tag_key, bg in TAG_BG.items():
            self.tree.tag_configure(tag_key, background=bg, foreground=TAG_FG[tag_key])
        
        # Empty state overlay
        self.empty_state = ctk.CTkLabel(
            tree_frame,
            text="📭  Нет данных для отображения\n\nДобавьте лицензию кнопкой «＋ Добавить»\nили загрузите Excel через «📥 Импорт»",
            font=ctk.CTkFont(size=15),
            text_color=TEXT_MUTED,
            justify="center"
        )
        self.empty_state.place(relx=0.5, rely=0.42, anchor="center")
        self.empty_state.lower()
        
        # Bind events
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<Delete>", lambda e: self._on_delete_selected())
        # Layout-independent Ctrl+C: check physical key (keycode 67 = C) with Control held
        self.tree.bind("<Control-KeyPress>", self._on_tree_ctrl_key)
        self.tree.bind("<KeyPress>", self._on_key_press, add="+")
        self.tree.bind("<<Copy>>", self._on_ctrl_c, add="+")
        self.tree.bind("<Control-Cyrillic_es>", self._on_ctrl_c, add="+")
        
        # Context menu
        self.context_menu = tk.Menu(self, tearoff=0, font=("Segoe UI", 10))
        self.context_menu.add_command(label="✏️  Редактировать", command=self._on_edit_selected)
        self.context_menu.add_command(label="📁  Переместить в категорию…", command=self._on_move_selected)
        self.context_menu.add_command(label="🗑  Удалить  (Delete)", command=self._on_delete_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="📋  Копировать данные", command=self._on_copy_card_data)
        self.context_menu.add_command(label="🔑  Копировать ID", command=self._on_copy_id)
        
        # ===== Bottom card: filters + sheet selector =====
        bottom = ctk.CTkFrame(page, corner_radius=14, fg_color=BG_PANEL)
        bottom.pack(fill="x", padx=24, pady=(0, 12))
        
        bar = ctk.CTkFrame(bottom, fg_color="transparent", height=54)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        
        ctk.CTkLabel(bar, text="Статус:", font=ctk.CTkFont(size=12),
                     text_color=TEXT_MUTED).pack(side="left", padx=(16, 6))
        self.status_filter_var = ctk.StringVar(value="Все")
        self.status_filter_menu = DropdownOption(
            bar, ["Все", "Активна", "Скоро истекает", "Истекла", "Ручная"],
            self.status_filter_var, command=self._on_filter_change,
            width=150, height=30
        )
        self.status_filter_menu.pack(side="left", pady=12)

        ctk.CTkLabel(bar, text="Срок:", font=ctk.CTkFont(size=12),
                     text_color=TEXT_MUTED).pack(side="left", padx=(14, 6))
        self.days_filter_var = ctk.StringVar(value="Все")
        self.days_filter_menu = DropdownOption(
            bar, ["Все", "0-3 дня", "5 дней", "Неделя (7 дней)", "10 дней", "30 дней", "Просроченные"],
            self.days_filter_var, command=self._on_filter_change,
            width=175, height=30
        )
        self.days_filter_menu.pack(side="left", pady=12)
        
        ctk.CTkButton(
            bar, text="Сбросить", width=90, height=30,
            corner_radius=8, fg_color="transparent", border_width=1,
            border_color=BORDER, text_color=TEXT_MUTED,
            hover_color=BG_PANEL_2, font=ctk.CTkFont(size=11),
            command=self._reset_filters
        ).pack(side="left", padx=14, pady=12)
        
        self.info_label = ctk.CTkLabel(bar, text="", font=ctk.CTkFont(size=11), text_color=TEXT_MUTED)
        self.info_label.pack(side="left", padx=8)
        
        ctk.CTkLabel(bar, text="Лист:", font=ctk.CTkFont(size=12),
                     text_color=TEXT_MUTED).pack(side="right", padx=(10, 6))
        self.sheet_seg = ctk.CTkSegmentedButton(
            bar, height=30, font=ctk.CTkFont(size=11),
            selected_color=ACCENT, selected_hover_color=ACCENT_HOVER,
            unselected_color=BG_PANEL_2, unselected_hover_color=BORDER,
            text_color=TEXT_MAIN,
            fg_color=BG_PANEL_2, border_width=0,
            command=self._on_sheet_change
        )
        self.sheet_seg.pack(side="right", padx=(0, 16), pady=12)
        
        # ===== Stat chips card =====
        stats_card = ctk.CTkFrame(page, corner_radius=14, fg_color=BG_PANEL)
        stats_card.pack(fill="x", padx=24, pady=(0, 16))
        
        stats_row = ctk.CTkFrame(stats_card, fg_color="transparent", height=46)
        stats_row.pack(fill="x")
        stats_row.pack_propagate(False)
        
        self.stat_chips = {}
        chip_defs = [
            ('total',    "Всего",          ACCENT),
            ('active',   "Активных",       "#3ddc97"),
            ('expiring', "Скоро истекает", "#f5a623"),
            ('critical', "0-3 дня",        "#ff5c5c"),
            ('expired',  "Просрочены",     "#c1556a"),
            ('manual',   "Ручные",         "#6fa8ff"),
        ]
        for key, label_text, color in chip_defs:
            chip = ctk.CTkFrame(stats_row, corner_radius=10, fg_color=BG_PANEL_2)
            chip.pack(side="left", padx=5, pady=8)
            ctk.CTkFrame(chip, width=12, height=8, corner_radius=2,
                         fg_color=color).pack(side="left", pady=(0, 12), padx=(10, 6))
            ctk.CTkLabel(chip, text=label_text, font=ctk.CTkFont(size=11),
                         text_color=TEXT_MUTED).pack(side="left")
            count_label = ctk.CTkLabel(
                chip, text="0", font=ctk.CTkFont(size=12, weight="bold"),
                text_color=color, width=30
            )
            count_label.pack(side="left", padx=(4, 10))
            self.stat_chips[key] = count_label
    
    def _configure_tree_style(self):
        """Apply modern styling to the ttk.Treeview widget (theme-aware)."""
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "Modern.Treeview",
            rowheight=36,
            font=("Segoe UI", 11),
            background=TREE_BG,
            fieldbackground=TREE_BG,
            foreground=TEXT_MAIN,
            borderwidth=0,
            relief="flat",
        )
        style.configure(
            "Modern.Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background=TREE_HEAD,
            foreground=TEXT_MUTED,
            borderwidth=1,
            bordercolor=TREE_HEAD,
            relief="flat",
            padding=(10, 8),
        )
        style.map("Modern.Treeview.Heading", background=[("active", TREE_HEAD_HOVER)])
        style.map(
            "Modern.Treeview",
            background=[("selected", ACCENT)],
            foreground=[("selected", "#ffffff")],
        )
    
    def _migrate_otvaly_category(self):
        """Move the legacy 'отвалы' sheet into the new «Отвалы» category once."""
        try:
            if self.db.get_all_licenses(category="Отвалы"):
                return
            self.db.move_sheet_records("отвалы", "GBS-market", "Отвалы")
        except Exception:
            pass

    def _refresh_all(self):
        """Refresh data and statuses from database."""
        self.db.refresh_all_statuses()
        self._load_categories()
        self._load_sheets()
        self._refresh_data()

    def _toggle_theme(self):
        """Switch between dark and light themes and rebuild the UI."""
        # Move focus off widgets that are about to be destroyed: CustomTkinter
        # restores the focused widget 1 ms after an appearance-mode change and
        # would hit an already destroyed treeview otherwise.
        try:
            self.focus_set()
            self.update_idletasks()
        except Exception:
            pass

        self.current_theme = "light" if self.current_theme == "dark" else "dark"
        apply_theme(self.current_theme)
        ctk.set_appearance_mode("dark" if self.current_theme == "dark" else "light")
        try:
            self.db.set_setting("theme", self.current_theme)
        except Exception:
            pass

        # Rebuild all widgets of the main window
        for widget in self.winfo_children():
            widget.destroy()

        self.title(APP_TITLE)
        self._build_ui()
        self._load_categories()
        self._load_sheets()
        self._refresh_data()
    
    # ===== Auto-update (GitHub Releases) =====

    def _post(self, cb):
        """Schedule cb on the UI thread, ignoring destroyed windows."""
        try:
            if self.winfo_exists():
                self.after(0, cb)
        except Exception:
            pass

    # ===== File menu (backup + updates) =====

    def _show_file_menu(self):
        """Show a popup menu with backup and update options."""
        menu = tk.Menu(self, tearoff=0, font=("Segoe UI", 11), bg=BG_PANEL,
                       fg=TEXT_MAIN, activebackground=ACCENT,
                       activeforeground="#ffffff", borderwidth=0)
        menu.add_command(label="💾  Резервная копия базы…",
                         command=lambda: (menu.destroy(), self._backup_database()))
        menu.add_command(label="📅  Резервная копия по расписанию…",
                         command=lambda: (menu.destroy(), self._schedule_backup_dialog()))
        menu.add_separator()
        menu.add_command(label="⬆  Проверить обновления",
                         command=lambda: (menu.destroy(), self._check_updates_manual()))
        try:
            menu.tk_popup(self.winfo_pointerx(), self.winfo_pointery())
        except Exception:
            try:
                menu.destroy()
            except Exception:
                pass

    def _show_columns_dialog(self):
        """Show a dialog with checkboxes to toggle column visibility."""
        dlg = ctk.CTkToplevel(self)
        dlg.title("Настройка столбцов")
        dlg.geometry("340x420")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.configure(fg_color=BG_PANEL)
        dlg.after(100, dlg.grab_set)

        ctk.CTkLabel(
            dlg, text="📊  Отображаемые столбцы",
            font=ctk.CTkFont(size=15, weight="bold"), text_color=TEXT_MAIN,
        ).pack(pady=(20, 12))

        # Scrollable frame for checkboxes
        scroll = ctk.CTkScrollableFrame(dlg, fg_color="transparent", width=300, height=300)
        scroll.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # Checkbox variables
        self._col_vars = {}
        for col_id, col_name, _col_width in self.columns:
            var = ctk.StringVar(value="1" if self._column_visibility.get(col_id, True) else "0")
            self._col_vars[col_id] = var
            cb = ctk.CTkCheckBox(
                scroll, text=col_name, variable=var, onvalue="1", offvalue="0",
                font=ctk.CTkFont(size=13), text_color=TEXT_MAIN,
                fg_color=ACCENT, hover_color=ACCENT_HOVER,
            )
            cb.pack(anchor="w", pady=4, padx=4)

        def apply():
            # Save settings and rebuild the UI
            for col_id, var in self._col_vars.items():
                self.db.set_setting(f"column_visible_{col_id}", var.get())
            dlg.destroy()
            # Rebuild the table with new visibility
            self._rebuild_table()

        buttons = ctk.CTkFrame(dlg, fg_color="transparent")
        buttons.pack(fill="x", padx=20, pady=(0, 16))
        ctk.CTkButton(
            buttons, text="Применить", width=120, height=36, corner_radius=8,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"), command=apply
        ).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ctk.CTkButton(
            buttons, text="Отмена", width=120, height=36, corner_radius=8,
            fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=13), command=dlg.destroy
        ).pack(side="right", expand=True, fill="x", padx=(6, 0))

    def _rebuild_table(self):
        """Rebuild the table to reflect column visibility changes."""
        # Save current state
        items = self.tree.get_children()
        item_values = {item: self.tree.item(item) for item in items}
        current_selection = self.tree.selection()
        current_scroll = self.tree.yview()

        # Re-read column visibility
        self._column_visibility = {}
        for col_id, _label, _width in self.columns:
            val = self.db.get_setting(f"column_visible_{col_id}", "1")
            self._column_visibility[col_id] = val == "1"

        # Reconfigure columns
        for col_id, col_name, col_width in self.columns:
            if self._column_visibility.get(col_id, True):
                self.tree.column(col_id, width=col_width, minwidth=55, anchor="w")
            else:
                self.tree.column(col_id, width=0, minwidth=0, stretch=False)

    def _backup_database(self):
        """Save a copy of the database to a user-chosen location."""
        initial = f"licenses_backup_{date.today().isoformat()}.db"
        path = filedialog.asksaveasfilename(
            title="Сохранить резервную копию базы",
            initialfile=initial,
            defaultextension=".db",
            filetypes=[("SQLite DB", "*.db"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            import shutil
            src = self.db.db_path
            # Write-ahead log handling: copy the main file
            shutil.copy2(src, path)
            # Also copy WAL/SHM if they exist
            for ext in ("-wal", "-shm"):
                wal = src + ext
                if os.path.exists(wal):
                    shutil.copy2(wal, path + ext)
            messagebox.showinfo(
                "Резервная копия",
                f"База данных сохранена в:\n{path}",
                parent=self,
            )
        except Exception as e:
            messagebox.showerror(
                "Ошибка",
                f"Не удалось создать резервную копию:\n{e}",
                parent=self,
            )

    def _schedule_backup_dialog(self):
        """Set up automatic daily backup at a chosen time."""
        dlg = ctk.CTkToplevel(self)
        dlg.title("Резервная копия по расписанию")
        dlg.geometry("420x280")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.configure(fg_color=BG_PANEL)
        dlg.after(100, dlg.grab_set)

        ctk.CTkLabel(
            dlg, text="📅  Автоматическое резервное копирование",
            font=ctk.CTkFont(size=15, weight="bold"), text_color=TEXT_MAIN,
        ).pack(pady=(20, 12))

        # State: on/off
        enabled_var = ctk.StringVar(
            value="on" if self.db.get_setting("backup_schedule_enabled") == "1" else "off"
        )
        ctk.CTkCheckBox(
            dlg, text="Включить ежедневное резервное копирование",
            variable=enabled_var, onvalue="on", offvalue="off",
            font=ctk.CTkFont(size=13), text_color=TEXT_MAIN,
        ).pack(pady=(0, 10))

        # Folder
        folder_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        folder_frame.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkLabel(folder_frame, text="Папка:", width=60, anchor="w",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED).pack(side="left")
        folder_var = ctk.StringVar(
            value=self.db.get_setting("backup_schedule_dir") or os.path.join(
                os.path.expanduser("~"), "Documents", "LicenseTracker_Backups"
            )
        )
        folder_entry = ctk.CTkEntry(folder_frame, textvariable=folder_var, width=240)
        folder_entry.pack(side="left", fill="x", expand=True, padx=(6, 4))
        ctk.CTkButton(
            folder_frame, text="…", width=28, height=28, corner_radius=6,
            fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
            command=lambda: (
                folder_var.set(filedialog.askdirectory(title="Выберите папку") or folder_var.get())
            )
        ).pack(side="left")

        # Time
        time_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        time_frame.pack(fill="x", padx=20, pady=(0, 12))
        ctk.CTkLabel(time_frame, text="Время:", width=60, anchor="w",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED).pack(side="left")
        time_var = ctk.StringVar(value=self.db.get_setting("backup_schedule_time") or "02:00")
        time_entry = ctk.CTkEntry(time_frame, textvariable=time_var, width=80)
        time_entry.pack(side="left", padx=(6, 0))
        ctk.CTkLabel(time_frame, text=" (ЧЧ:ММ)", anchor="w",
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED).pack(side="left")

        status_label = ctk.CTkLabel(dlg, text="", anchor="w",
                                    font=ctk.CTkFont(size=11), text_color=TEXT_MUTED)
        status_label.pack(fill="x", padx=20)

        def save():
            try:
                self.db.set_setting("backup_schedule_enabled", "1" if enabled_var.get() == "on" else "0")
                self.db.set_setting("backup_schedule_dir", folder_var.get())
                self.db.set_setting("backup_schedule_time", time_var.get())
                # Ensure the folder exists
                os.makedirs(folder_var.get(), exist_ok=True)
                # Restart the scheduler thread
                self._start_backup_scheduler()
                dlg.destroy()
                messagebox.showinfo("Расписание", "Настройки сохранены.", parent=self)
            except Exception as e:
                status_label.configure(text=f"Ошибка: {e}", text_color="#e5484d")

        ctk.CTkButton(
            dlg, text="Сохранить", width=120, height=36, corner_radius=8,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"), command=save
        ).pack(pady=(8, 16))

        self._start_backup_scheduler()

    def _start_backup_scheduler(self):
        """Start a background thread that performs scheduled backups."""
        if getattr(self, "_backup_scheduler_thread", None) and self._backup_scheduler_thread.is_alive():
            return

        def scheduler():
            while True:
                try:
                    if self.db.get_setting("backup_schedule_enabled") == "1":
                        time_str = self.db.get_setting("backup_schedule_time") or "02:00"
                        folder = self.db.get_setting("backup_schedule_dir")
                        now = datetime.now()
                        target = datetime.strptime(time_str, "%H:%M").replace(
                            year=now.year, month=now.month, day=now.day
                        )
                        # If already passed today, wait until tomorrow
                        if now >= target:
                            target += timedelta(days=1)
                        # Wait until the target time (check every 30s)
                        while datetime.now() < target:
                            import time as _time
                            _time.sleep(30)
                        # Perform backup
                        self._run_scheduled_backup(folder)
                    import time as _time
                    _time.sleep(60)
                except Exception:
                    import time as _time
                    _time.sleep(60)

        self._backup_scheduler_thread = threading.Thread(target=scheduler, daemon=True)
        self._backup_scheduler_thread.start()

    def _run_scheduled_backup(self, folder):
        """Actually copy the DB to the backup folder."""
        try:
            os.makedirs(folder, exist_ok=True)
            ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            dest = os.path.join(folder, f"licenses_backup_{ts}.db")
            import shutil
            shutil.copy2(self.db.db_path, dest)
            # Keep only the last 30 backups
            backups = sorted(
                f for f in os.listdir(folder) if f.startswith("licenses_backup_") and f.endswith(".db")
            )
            for old in backups[:-30]:
                try:
                    os.remove(os.path.join(folder, old))
                except Exception:
                    pass
        except Exception:
            pass

    def _auto_backup_to_documents(self):
        """Automatically back up the database to Documents/LicenseTracker_Backups once per day."""
        try:
            last = self.db.get_setting("auto_backup_date", "")
            today = date.today().isoformat()
            if last == today:
                return  # already backed up today

            docs = os.path.join(os.path.expanduser("~"), "Documents")
            folder = os.path.join(docs, "LicenseTracker_Backups")
            os.makedirs(folder, exist_ok=True)

            ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            dest = os.path.join(folder, f"licenses_backup_{ts}.db")
            import shutil
            shutil.copy2(self.db.db_path, dest)

            # Keep only the last 14 daily backups
            backups = sorted(
                f for f in os.listdir(folder)
                if f.startswith("licenses_backup_") and f.endswith(".db")
            )
            for old in backups[:-14]:
                try:
                    os.remove(os.path.join(folder, old))
                except Exception:
                    pass

            self.db.set_setting("auto_backup_date", today)
        except Exception:
            pass

    def _download_and_install_update(self, download_url, version, release_url):
        """Download the new exe and replace the current one."""
        try:
            # Download to temp
            tmp_dir = tempfile.gettempdir()
            new_exe = os.path.join(tmp_dir, "LicenseTracker_update.exe")
            req = urllib.request.Request(download_url, headers={"User-Agent": f"LicenseTracker/{APP_VERSION}"})
            with urllib.request.urlopen(req, timeout=120) as resp, open(new_exe, "wb") as f:
                while True:
                    chunk = resp.read(65536)
                    if not chunk:
                        break
                    f.write(chunk)

            # Create a helper batch to replace and restart
            current = os.path.join(os.path.dirname(sys.executable), "LicenseTracker.exe")
            bat = os.path.join(tmp_dir, "lt_update.bat")
            with open(bat, "w", encoding="utf-8") as f:
                f.write("@echo off\n")
                f.write("timeout /t 2 /nobreak >nul\n")
                f.write(f'copy /Y "{new_exe}" "{current}" >nul\n')
                f.write(f'start "" "{current}"\n')
                f.write('del "%~f0"\n')

            subprocess.Popen(["cmd", "/c", bat],
                             creationflags=subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS)
            self.after(0, self.destroy)
            self.after(100, os._exit, 0)
        except Exception as e:
            messagebox.showerror("Обновление", f"Ошибка при установке обновления:\n{e}", parent=self)

    def _auto_update_check(self):
        if not getattr(sys, "frozen", False):
            return
        if getattr(self, "_update_in_progress", False):
            return

        def worker():
            try:
                release = fetch_latest_release()
            except Exception:
                return
            if not isinstance(release, dict):
                return
            tag = release.get("tag_name", "")
            if not _is_newer(tag, APP_VERSION):
                return
            url = find_update_asset_url(release)
            if not url:
                return
            # Auto-install without asking
            self._update_in_progress = True
            self._post( lambda: messagebox.showinfo(
                "Обновление",
                f"Найдена новая версия {tag}. Приложение будет обновлено и перезапущено.",
                parent=self))
            self._download_and_install_update(url, tag, release.get("html_url", ""))

        threading.Thread(target=worker, daemon=True).start()

    def _check_updates_manual(self):
        if getattr(self, "_update_in_progress", False):
            return

        def worker():
            try:
                release = fetch_latest_release()
            except Exception as e:
                self._post( lambda: messagebox.showerror(
                    "Обновление", f"Не удалось проверить обновления:\n{e}", parent=self))
                return
            if not isinstance(release, dict):
                self._post( lambda: messagebox.showerror(
                    "Обновление", "Некорректный ответ сервера.", parent=self))
                return
            tag = release.get("tag_name", "")
            if not _is_newer(tag, APP_VERSION):
                self._post( lambda: messagebox.showinfo(
                    "Обновление", "У вас установлена последняя версия.", parent=self))
                return

            url = find_update_asset_url(release)
            if not url:
                self._post( lambda: messagebox.showwarning(
                    "Обновление", "В релизе нет файла для загрузки.", parent=self))
                return

            def ask():
                if messagebox.askyesno(
                    "Обновление",
                    f"Доступна новая версия {tag}.\nЗагрузить и установить?",
                    parent=self,
                ):
                    self._update_in_progress = True
                    self._download_and_install_update(url, tag, release.get("html_url", ""))

            self._post( ask)

        threading.Thread(target=worker, daemon=True).start()

    def _start_update_download(self, release, tag, url, silent=False):
        if getattr(self, "_update_in_progress", False):
            return
        self._update_in_progress = True
        self._show_update_progress(tag)
        tmp = os.path.join(tempfile.gettempdir(), "LicenseTracker_update.exe")

        def worker():
            try:
                download_update_file(
                    url, tmp,
                    progress_cb=lambda d, t: self._post( lambda dd=d, tt=t: self._update_progress(dd, tt)),
                )
            except Exception as e:
                self._post( lambda: self._update_failed(str(e)))
                return
            self._post( lambda: self._apply_update_and_restart_ui(tmp, tag))

        threading.Thread(target=worker, daemon=True).start()

    def _show_update_progress(self, tag):
        try:
            if getattr(self, "_update_dlg", None) and self._update_dlg.winfo_exists():
                self._update_dlg.destroy()
        except Exception:
            pass
        dlg = ctk.CTkToplevel(self)
        dlg.title("Обновление")
        dlg.geometry("420x170")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()
        dlg.configure(fg_color=BG_PANEL)

        ctk.CTkLabel(
            dlg, text=f"Загрузка обновления {tag}…",
            font=ctk.CTkFont(size=14, weight="bold"), text_color=TEXT_MAIN,
        ).pack(pady=(20, 8))

        bar = ctk.CTkProgressBar(dlg, width=360, height=14)
        bar.pack(pady=6)
        bar.set(0)

        label = ctk.CTkLabel(
            dlg, text="Подготовка…", font=ctk.CTkFont(size=11), text_color=TEXT_MUTED,
        )
        label.pack(pady=4)

        self._update_dlg = dlg
        self._update_bar = bar
        self._update_label = label

    def _update_progress(self, downloaded, total):
        try:
            if total and total > 0:
                self._update_bar.set(downloaded / total)
                mb_d = downloaded / (1024 * 1024)
                mb_t = total / (1024 * 1024)
                self._update_label.configure(text=f"{mb_d:.1f} / {mb_t:.1f} МБ")
            else:
                self._update_bar.set(0)
                self._update_label.configure(text=f"{downloaded // 1024} КБ загружено")
        except Exception:
            pass

    def _destroy_update_progress(self):
        try:
            if getattr(self, "_update_dlg", None) and self._update_dlg.winfo_exists():
                self._update_dlg.destroy()
        except Exception:
            pass
        self._update_dlg = None

    def _update_failed(self, error):
        self._destroy_update_progress()
        self._update_in_progress = False
        try:
            messagebox.showerror("Обновление", f"Ошибка загрузки обновления:\\n{error}", parent=self)
        except Exception:
            pass

    def _apply_update_and_restart_ui(self, tmp_path, tag):
        self._destroy_update_progress()
        if not getattr(sys, "frozen", False):
            self._update_in_progress = False
            messagebox.showinfo(
                "Обновление",
                f"Новая версия {tag} загружена в:\\n{tmp_path}\\n"
                "Автоустановка доступна только в собранной версии.",
                parent=self,
            )
            return
        try:
            messagebox.showinfo(
                "Обновление",
                f"Обновление {tag} загружено. Приложение будет перезапущено для установки.",
                parent=self,
            )
        except Exception:
            pass
        ok = apply_update_and_restart(tmp_path)
        self._update_in_progress = False
        if ok:
            try:
                self.destroy()
            except Exception:
                pass
            os._exit(0)

    def _on_ctrl_c(self, event=None):
        """Ctrl+C: copy selected rows (table) or selected text (entries)."""
        try:
            widget = self.focus_get()
        except Exception:
            widget = None
        tree = getattr(self, "tree", None)
        # Be lenient about focus: if the table has a selection and focus is
        # on the table or the main window, treat it as a table copy request.
        if tree is not None and tree.selection() and (widget is tree or widget is None or widget is self):
            self._on_copy_card_data()
            return "break"
        # Entries / text widgets: copy the current selection explicitly
        if isinstance(widget, (tk.Entry, tk.Text, tk.Spinbox)):
            try:
                sel = widget.selection_get()
            except Exception:
                sel = ""
            if sel:
                copy_to_clipboard(self, sel)
                return "break"
        return None

    def _on_key_press(self, event):
        """Layout-independent Ctrl+C (works on Russian keyboard too).
        Uses keycode (physical key 67 = C) — ignores keysym which varies by layout."""
        try:
            ks = str(getattr(event, "keysym", "") or "")
            kc = getattr(event, "keycode", None)
            st = getattr(event, "state", 0) or 0
            if st & 0x4 and kc == 67:  # Ctrl + physical C key
                return self._on_ctrl_c(event)
        except Exception:
            pass
        return None

    def _on_tree_ctrl_key(self, event):
        """Tree focused + Ctrl+key: copy if it's the C key (any layout)."""
        if getattr(event, "keycode", None) == 67:
            if self.tree.selection():
                self._on_copy_card_data()
                return "break"
        return None

    def _poll_ctrl_c(self):
        """No-op: hook removed. Kept for compatibility."""
        pass

    def _bind_shortcuts(self):
        """Global keyboard shortcuts."""
        self.bind("<Control-f>", lambda e: self.search_entry.focus_set())
        self.bind("<F5>", lambda e: self._refresh_data())
        self.bind("<Control-c>", self._on_ctrl_c, add="+")
        self.bind("<Control-C>", self._on_ctrl_c, add="+")
        self.bind("<KeyPress>", self._on_key_press, add="+")
        self.bind("<Control-Cyrillic_es>", self._on_ctrl_c, add="+")
        self.bind("<Control-Cyrillic_ES>", self._on_ctrl_c, add="+")
        self.bind("<<Copy>>", self._on_ctrl_c, add="+")
        self.bind("<Control-n>", lambda e: self._on_add())
        self.bind("<Insert>", lambda e: self._on_add())
    
    def _reset_filters(self):
        """Reset all filters to defaults."""
        self.search_var.set("")
        self.status_filter_var.set("Все")
        self.days_filter_var.set("Все")
        self.filter_text = ""
        self.status_filter = "Все"
        self.days_filter = "Все"
        self._refresh_data()
    
    def _load_categories(self):
        """Load category nav items (like menu links in the reference)."""
        CATEGORIES_LIST = list(CATEGORIES)
        
        counts = {cat: len(self.db.get_all_licenses(category=cat)) for cat in CATEGORIES_LIST}
        
        # Rebuild nav buttons (centered container)
        for widget in self.category_inner.winfo_children():
            widget.destroy()
        
        self.category_buttons = {}
        for cat in CATEGORIES_LIST:
            count = counts[cat]
            label = f"{cat}  {count}" if count else cat
            btn = ctk.CTkButton(
                self.category_inner,
                text=label,
                height=36,
                corner_radius=8,
                border_spacing=10,
                fg_color="transparent",
                hover_color=BG_PANEL_2,
                text_color=TEXT_MUTED,
                font=("Segoe UI", 12),
                command=lambda c=cat: self._on_category_change(c)
            )
            btn.pack(side="left", padx=3)
            self.category_buttons[cat] = btn
        
        if self.current_category not in CATEGORIES:
            self.current_category = CATEGORIES[0]
        self._update_category_buttons()
    
    def _update_category_buttons(self):
        """Highlight the active category nav item."""
        for cat, btn in self.category_buttons.items():
            if cat == self.current_category:
                btn.configure(
                    fg_color=ACCENT,
                    hover_color=ACCENT_HOVER,
                    text_color="#ffffff",
                    font=("Segoe UI", 12, "bold")
                )
            else:
                btn.configure(
                    fg_color="transparent",
                    hover_color=BG_PANEL_2,
                    text_color=TEXT_MUTED,
                    font=("Segoe UI", 12)
                )
    
    def _on_category_change(self, category: str):
        """Handle category nav click."""
        if category == self.current_category:
            return
        self.current_category = category
        self.current_sheet = ""  # Reset sheet filter
        self._update_category_buttons()
        self._load_sheets()
        self._refresh_data()
    
    def _load_sheets(self):
        """Load sheets for the current category into segmented buttons."""
        sheets = self.db.get_sheets(self.current_category)
        values = ["Все"] + list(sheets)
        self.sheet_seg.configure(values=values)
        
        current = self.sheet_seg.get()
        if current in values:
            self.sheet_seg.set(current)
            self.current_sheet = "" if current == "Все" else current
        else:
            self.sheet_seg.set("Все")
            self.current_sheet = ""
    
    def _on_sheet_change(self, value=None):
        selected = self.sheet_seg.get()
        self.current_sheet = "" if selected == "Все" else selected
        self._refresh_data()
    
    def _on_filter_change(self, *args):
        self.filter_text = self.search_var.get().lower()
        self.status_filter = self.status_filter_var.get()
        self.days_filter = self.days_filter_var.get()
        self._refresh_data()
    
    def _calc_days_left(self, lic: Dict) -> Optional[int]:
        """Calculate days until license expires. Primary source: Срок ID (id_expiry)."""
        today = date.today()
        
        for field in ['id_expiry', 'next_payment']:  # id_expiry first
            val = lic.get(field)
            if val:
                try:
                    if isinstance(val, str):
                        if 'T' in val:
                            exp_date = datetime.fromisoformat(val.replace('Z', '')).date()
                        else:
                            exp_date = datetime.strptime(val[:10], "%Y-%m-%d").date()
                    elif isinstance(val, date):
                        exp_date = val
                    else:
                        continue
                    return (exp_date - today).days
                except:
                    continue
        return None
    
    def _matches_days_filter(self, days_left: Optional[int], filter_val: str) -> bool:
        """Check if days_left matches the selected filter."""
        if filter_val == "Все":
            return True
        
        if days_left is None:
            return False
        
        if filter_val == "0-3 дня":
            return 0 <= days_left <= 3
        elif filter_val == "5 дней":
            return 0 <= days_left <= 5
        elif filter_val == "Неделя (7 дней)":
            return 0 <= days_left <= 7
        elif filter_val == "10 дней":
            return 0 <= days_left <= 10
        elif filter_val == "30 дней":
            return 0 <= days_left <= 30
        elif filter_val == "Просроченные":
            return days_left < 0
        
        return True
    
    def _on_column_sort(self, column):
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        self._update_heading_arrows()
        self._refresh_data()

    def _update_heading_arrows(self):
        """Show sort direction marker in the active column heading."""
        for col_id, col_name, _ in self.columns:
            if col_id == self.sort_column:
                arrow = "  ▼" if self.sort_reverse else "  ▲"
                self.tree.heading(col_id, text=col_name + arrow)
            else:
                self.tree.heading(col_id, text=col_name)
    
    def _refresh_data(self):
        # Get data from database
        licenses = self.db.get_all_licenses(
            category=self.current_category,
            sheet_name=self.current_sheet if self.current_sheet else None
        )
        
        # Calculate days_left for each license
        for lic in licenses:
            lic['days_left'] = self._calc_days_left(lic)
        
        # Apply filters
        filtered = []
        for lic in licenses:
            # Text search
            if self.filter_text:
                searchable = " ".join([
                    str(lic.get('full_name', '')),
                    str(lic.get('store_name', '')),
                    str(lic.get('phone', '')),
                    str(lic.get('email', '')),
                    str(lic.get('gbs_id', '')),
                    str(lic.get('license_key', '')),
                ]).lower()
                if self.filter_text not in searchable:
                    continue
            
            # Status filter
            if self.status_filter != "Все":
                status_map = {
                    "Активна": LicenseStatus.ACTIVE,
                    "Скоро истекает": LicenseStatus.EXPIRING,
                    "Истекла": LicenseStatus.EXPIRED,
                    "Ручная": LicenseStatus.MANUAL,
                }
                if lic.get('status') != status_map.get(self.status_filter):
                    continue
            
            # Days filter
            if hasattr(self, 'days_filter') and not self._matches_days_filter(lic.get('days_left'), self.days_filter):
                continue
            
            filtered.append(lic)
        
        # Sort (dates parsed properly; missing values always go to the bottom)
        if self.sort_column:
            col = self.sort_column
            DATE_COLS = {'install_date', 'id_expiry', 'payment_date', 'renewal_date', 'next_payment'}
            NUM_COLS = {'days_left'}

            def sort_key(item):
                val = item.get(col)
                missing_rank = 2  # missing values always last
                if col in NUM_COLS:
                    if val is None:
                        return (missing_rank, 0)
                    try:
                        return (0, float(val))
                    except (TypeError, ValueError):
                        return (missing_rank, 0)
                if col in DATE_COLS and val is not None:
                    parsed = parse_date_any(val)
                    return (0, parsed) if parsed else (missing_rank, datetime.min)
                if val is None or str(val).strip() == "":
                    return (missing_rank, "")
                if col == 'status' or col == 'manual_status':
                    return (1, str(val).lower())
                return (1, str(val).lower())

            filtered.sort(key=sort_key, reverse=self.sort_reverse)
        
        # Update tree with visual grouping: expired first, sorted by days left
        self.tree.delete(*self.tree.get_children())
        
        for lic in filtered:
            values = []
            for col_id, _, _ in self.columns:
                val = lic.get(col_id, "")
                if col_id in ['install_date', 'id_expiry', 'payment_date', 'renewal_date', 'next_payment'] and val:
                    try:
                        val = datetime.fromisoformat(str(val).replace('Z', '')).strftime("%d.%m.%Y")
                    except:
                        pass
                elif col_id == 'days_left':
                    days = lic.get('days_left')
                    if days is not None:
                        if days < 0:
                            val = f"−{abs(days)} дн."
                        elif days == 0:
                            val = "Сегодня"
                        else:
                            val = f"{days} дн."
                    else:
                        val = "—"
                elif col_id == 'status':
                    manual = lic.get('manual_status')
                    if manual:
                        try:
                            val = MANUAL_STATUS_LABELS.get(ManualStatus(manual), manual)
                        except ValueError:
                            val = manual
                    else:
                        try:
                            val = STATUS_LABELS.get(LicenseStatus(val), val)
                        except ValueError:
                            pass
                values.append(val if val is not None else "")
            
            # Determine tag for coloring (based on computed urgency)
            days = lic.get('days_left')
            if days is not None and days < 0:
                tag = 'expired'
            elif days == 0:
                tag = 'critical'
            elif days is not None and days <= 3:
                tag = 'critical'
            elif days is not None and days <= 7:
                tag = 'expiring'
            elif lic.get('manual_status'):
                tag = 'manual'
            else:
                tag = 'active'
            
            self.tree.insert("", "end", values=values, tags=(tag,), iid=str(lic['id']))
        
        # Show/hide empty state
        if not filtered:
            self.empty_state.lift()
        else:
            self.empty_state.lower()
        
        self._update_stats(len(filtered), len(licenses))
    
    def _update_stats(self, filtered_count: int = None, total_count: int = None):
        all_licenses = self.db.get_all_licenses(
            category=self.current_category,
            sheet_name=self.current_sheet if self.current_sheet else None
        )
        
        # Calculate days for all
        for lic in all_licenses:
            lic['days_left'] = self._calc_days_left(lic)
        
        if filtered_count is None or total_count is None:
            total_count = len(all_licenses)
            filtered_count = total_count
        
        # Count by computed urgency
        active = expiring = critical = expired = manual = 0
        for l in all_licenses:
            days = l.get('days_left')
            if l.get('manual_status'):
                manual += 1
                continue
            if days is None:
                active += 1
            elif days < 0:
                expired += 1
            elif days <= 3:
                critical += 1
            elif days <= 7:
                expiring += 1
            else:
                active += 1
        
        # Update stat chips
        self.stat_chips['total'].configure(text=str(total_count))
        self.stat_chips['active'].configure(text=str(active))
        self.stat_chips['expiring'].configure(text=str(expiring))
        self.stat_chips['critical'].configure(text=str(critical))
        self.stat_chips['expired'].configure(text=str(expired))
        self.stat_chips['manual'].configure(text=str(manual))
        
        sheet_info = f"{self.current_category} · Лист: {self.current_sheet or 'Все'} · Показано: {filtered_count}"
        self.info_label.configure(text=sheet_info)
    
    def _on_add(self):
        dialog = LicenseCardDialog(
            self, self.db, 
            sheet_name=self.current_sheet or "GBS", 
            category=self.current_category
        )
        self.wait_window(dialog)
        if dialog.result:
            self._load_categories()
            self._load_sheets()
            self._refresh_data()
    
    def _on_edit_selected(self):
        selection = self.tree.selection()
        if not selection:
            return
        # Edit first selected
        license_id = int(selection[0])
        self._edit_license(license_id)
    
    def _on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            license_id = int(item)
            self._edit_license(license_id)
    
    def _on_right_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.tk_popup(event.x_root, event.y_root)
    
    def _edit_license(self, license_id: int):
        license_data = self.db.get_license(license_id)
        if not license_data:
            return
        
        dialog = LicenseCardDialog(
            self, self.db, license_data,
            category=self.current_category
        )
        self.wait_window(dialog)
        if dialog.result:
            self._load_categories()
            self._load_sheets()
            self._refresh_data()
    
    def _on_move_selected(self):
        """Move the selected records to another category."""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning(
                "Перемещение",
                "Сначала выберите запись (или несколько) в таблице.",
                parent=self
            )
            return

        ids = []
        for item_id in selection:
            try:
                ids.append(int(item_id))
            except (TypeError, ValueError):
                pass
        if not ids:
            return

        def do_move(category):
            moved = self.db.move_licenses(ids, category)
            if moved:
                self._load_categories()
                self._load_sheets()
                self._refresh_data()

        MoveCategoryDialog(self, len(ids), self.current_category, do_move)

    def _on_delete_selected(self):
        selection = self.tree.selection()
        if not selection:
            return

        count = len(selection)
        if count == 1:
            msg = "Удалить выбранную запись?"
        else:
            msg = f"Удалить выбранные записи ({count})?\n\nБудут удалены все выделенные: {count} шт."
        if not messagebox.askyesno("Подтверждение", msg):
            return

        deleted = 0
        for item_id in selection:
            try:
                if self.db.delete_license(int(item_id)):
                    deleted += 1
            except Exception:
                pass
        if deleted:
            self._refresh_data()
            self._load_sheets()
            self._load_categories()

    # ===== Renewal (Продление срока ID) =====

    RENEW_OPTIONS = [
        ("1 месяц — +1 мес.", 1),
        ("Квартал — +3 мес.", 3),
        ("Полгода — +6 мес.", 6),
        ("Год — +12 мес.", 12),
    ]

    def _on_renew_selected(self):
        """Open the renewal dialog for selected row(s)."""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning(
                "Продление",
                "Сначала выберите запись (или несколько) в таблице.",
                parent=self
            )
            return

        licenses = []
        for item_id in selection:
            lic = self.db.get_license(int(item_id))
            if lic:
                licenses.append(lic)
        if not licenses:
            return

        self._show_renew_dialog(licenses)

    def _show_renew_dialog(self, licenses):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Продление срока ID")
        dlg.geometry("520x560")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.configure(fg_color=BG_PANEL)

        # Header
        header = ctk.CTkFrame(dlg, corner_radius=0, fg_color=BG_NAV, height=64)
        header.pack(fill="x")
        header.pack_propagate(False)
        title = "⏳  Продление срока ID"
        if len(licenses) > 1:
            title += f"  ·  выбрано: {len(licenses)}"
        ctk.CTkLabel(
            header, text=title,
            font=ctk.CTkFont(size=15, weight="bold"), text_color=TEXT_MAIN
        ).pack(side="left", padx=16)

        body = ctk.CTkFrame(dlg, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=(12, 6))

        # Current expiry list (up to 8 rows)
        list_frame = ctk.CTkFrame(body, fg_color=BG_PANEL_2, corner_radius=10)
        list_frame.pack(fill="x")

        max_show = 8
        for i, lic in enumerate(licenses):
            if i >= max_show:
                ctk.CTkLabel(
                    list_frame, text=f"… и ещё {len(licenses) - max_show}",
                    font=ctk.CTkFont(size=11), text_color=TEXT_MUTED, anchor="w"
                ).pack(anchor="w", padx=12, pady=(0, 6))
                break
            cur = lic.get('id_expiry')
            cur_txt = "— не указан"
            if cur:
                parsed = parse_date_any(cur)
                cur_txt = parsed.strftime("%d.%m.%Y") if parsed else str(cur)
            ctk.CTkLabel(
                list_frame,
                text=f"• {lic.get('full_name') or 'Без имени'}   —   текущий Срок ID: {cur_txt}",
                font=ctk.CTkFont(size=12), text_color=TEXT_MAIN, anchor="w"
            ).pack(anchor="w", padx=12, pady=(8, 0))
        pad = ctk.CTkFrame(list_frame, height=4, fg_color="transparent")
        pad.pack()

        # Rule hint
        ctk.CTkLabel(
            body,
            text="Правило: если срок уже прошёл или не указан — отсчёт с сегодняшнего дня,\nиначе — от текущей даты срока ID.",
            font=ctk.CTkFont(size=11), text_color=TEXT_MUTED, justify="left", anchor="w"
        ).pack(anchor="w", pady=(10, 6))

        # Quick options
        ctk.CTkLabel(
            body, text="Выберите период:",
            font=ctk.CTkFont(size=12, weight="bold"), text_color=TEXT_MAIN, anchor="w"
        ).pack(anchor="w", pady=(4, 6))

        btns = ctk.CTkFrame(body, fg_color="transparent")
        btns.pack(fill="x")
        for label_text, months in self.RENEW_OPTIONS:
            ctk.CTkButton(
                btns, text=label_text, height=38, corner_radius=8,
                fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
                font=ctk.CTkFont(size=12),
                command=lambda m=months: self._apply_renewal(licenses, m, 0, dlg)
            ).pack(fill="x", pady=3)

        # Custom period
        custom = ctk.CTkFrame(body, fg_color=BG_PANEL_2, corner_radius=10)
        custom.pack(fill="x", pady=(10, 0))
        ctk.CTkLabel(
            custom, text="Другой период:", font=ctk.CTkFont(size=12),
            text_color=TEXT_MAIN
        ).pack(side="left", padx=(12, 6), pady=12)
        months_entry = ctk.CTkEntry(custom, placeholder_text="мес.", width=70, corner_radius=8,
                                    fg_color=BG_PANEL, border_color=BORDER)
        months_entry.pack(side="left", pady=12)
        days_entry = ctk.CTkEntry(custom, placeholder_text="дней", width=70, corner_radius=8,
                                  fg_color=BG_PANEL, border_color=BORDER)
        days_entry.pack(side="left", padx=(6, 6), pady=12)
        ctk.CTkButton(
            custom, text="Продлить", width=100, height=32, corner_radius=8,
            fg_color="#0e9f6e", hover_color="#0c8a5f", text_color="#ffffff",
            font=ctk.CTkFont(size=12),
            command=lambda: self._apply_renewal_custom(licenses, months_entry, days_entry, dlg)
        ).pack(side="left", padx=(6, 12), pady=12)

        # Cancel
        ctk.CTkButton(
            body, text="Закрыть", width=100, height=32, corner_radius=8,
            fg_color=BG_PANEL_2, hover_color=BORDER, text_color=TEXT_MAIN,
            command=dlg.destroy
        ).pack(anchor="e", pady=(10, 4))

        dlg.update_idletasks()
        dlg.geometry(f"+{self.winfo_x() + 200}+{self.winfo_y() + 120}")
        dlg.after(100, dlg.grab_set)

    def _apply_renewal_custom(self, licenses, months_entry, days_entry, dlg):
        months, days = 0, 0
        try:
            months = int(months_entry.get().strip() or 0)
        except ValueError:
            months = 0
        try:
            days = int(days_entry.get().strip() or 0)
        except ValueError:
            days = 0
        if months == 0 and days == 0:
            messagebox.showwarning(
                "Продление", "Укажите количество месяцев или дней (больше 0).", parent=dlg
            )
            return
        self._apply_renewal(licenses, months, days, dlg)

    def _apply_renewal(self, licenses, months: int, days: int, dlg=None):
        """Extend id_expiry for every selected license and save immediately."""
        today = date.today()
        updated = 0
        try:
            for lic in licenses:
                current = parse_date_any(lic.get('id_expiry'))
                if current is None or current.date() < today:
                    base = today  # expired/missing -> from today
                else:
                    base = current.date()  # valid -> extend from current expiry
                new_date = add_months(base, months) + timedelta(days=days)
                if self.db.update_license(lic['id'], {'id_expiry': new_date.isoformat()}):
                    updated += 1
        except Exception as e:
            messagebox.showerror("Продление", f"Не удалось продлить: {e}", parent=self)
            return

        if dlg is not None:
            dlg.destroy()
        self._refresh_data()

        # Summary popup only when opened from the dialog (not in silent mode)
        if dlg is not None:
            if len(licenses) > 1:
                msg = f"Продлено записей: {updated}\nПериод: {self._renew_period_text(months, days)}"
            else:
                msg = f"Срок ID продлён.\nПериод: {self._renew_period_text(months, days)}"
            messagebox.showinfo("Продление", msg, parent=self)

    @staticmethod
    def _renew_period_text(months: int, days: int) -> str:
        parts = []
        if months:
            parts.append(f"{months} мес.")
        if days:
            parts.append(f"{days} дн.")
        return " + ".join(parts) if parts else "—"

    def _on_copy_id(self):
        selection = self.tree.selection()
        if selection:
            # Copy first selected ID
            copy_to_clipboard(self, str(selection[0]))

    def _on_copy_card_data(self):
        """Copy contact block(s) for the selected row(s)."""
        selection = self.tree.selection()
        if not selection:
            return

        blocks = []
        for iid in selection:
            try:
                lic = self.db.get_license(int(iid))
            except (TypeError, ValueError):
                lic = None
            if not lic:
                continue
            text = format_card_copy_text(lic)
            if text:
                blocks.append(text)

        if not blocks:
            return

        copy_to_clipboard(self, "\n\n".join(blocks))
    
    def _on_import(self):
        """Open the import wizard (preview + column mapping + category)."""
        ImportDialog(
            self, self.db,
            default_category=self.current_category,
            on_done=self._after_import
        )

    def _after_import(self):
        self._load_categories()
        self._load_sheets()
        self._refresh_data()

    def _on_export(self):
        file_path = filedialog.asksaveasfilename(
            title="Сохранить как Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"Licenses_{self.current_category}_{date.today().isoformat()}.xlsx"
        )
        if not file_path:
            return
        
        try:
            sheet = self.current_sheet if self.current_sheet else None
            self.db.export_to_excel(file_path, category=self.current_category, sheet_name=sheet)
            messagebox.showinfo(
                "Экспорт", 
                f"Данные категории '{self.current_category}' экспортированы в:\n{file_path}",
                parent=self
            )
        except Exception as e:
            messagebox.showerror("Ошибка экспорта", str(e), parent=self)


def main():
    app = LicenseTrackerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
