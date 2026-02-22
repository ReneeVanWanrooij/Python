import calendar
import hashlib
import json
import os
import shutil
import sys
import ctypes
import urllib.request
import urllib.error
from ctypes import wintypes
from dataclasses import dataclass
from datetime import date, datetime, timedelta

import holidays
try:
    import psutil
except Exception:
    psutil = None
try:
    import pygetwindow as gw
except Exception:
    gw = None
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from PySide6.QtCore import Qt, Signal, QEvent, QPoint, QObject, QTimer, QPropertyAnimation, QEasingCurve, QRect, QParallelAnimationGroup
from PySide6.QtGui import QAction, QColor, QFont, QBrush, QLinearGradient, QGradient, QPen, QPixmap, QPainter, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QColorDialog,
    QScrollArea,
    QSystemTrayIcon,
    QMenu,
    QStatusBar,
    QStyledItemDelegate,
    QSizePolicy,
    QTabBar,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QFrame,
    QHeaderView,
    QToolBar,
    QVBoxLayout,
    QWidget,
    QSpinBox,
    QCheckBox,
    QAbstractItemView,
)


WEEKDAYS = ["MA", "DI", "WO", "DO", "VR", "ZA", "ZO"]
MONTHS_NL = [
    "Januari",
    "Februari",
    "Maart",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Augustus",
    "September",
    "Oktober",
    "November",
    "December",
]

DEFAULT_EXTRA_INFO_OPTIONS = [
    "cursus",
    "bijzonder verlof",
    "zwangerschap",
]

COLOR_DEFAULTS = {
    "bg_today_dark": "#1f2a3a",
    "bg_vacation_dark": "#3b2e1a",
    "bg_holiday_dark": "#173D37",
    "bg_school_dark": "#1d4ed8",
    "bg_weekend_dark": "#232a34",
    "bg_default_dark": "#161b22",
    "daynum_today_dark": "#79c0ff",
    "daynum_holiday_dark": "#ff9492",
    "daynum_vacation_dark": "#ffb757",
    "daynum_school_dark": "#dbeafe",
    "daynum_weekend_dark": "#8b949e",
    "daynum_default_dark": "#e6edf3",
    "weeknum_bg_dark": "#334155",
    "weeknum_fg_dark": "#c9d1d9",
    "weektotal_bg_dark": "#223047",
    "weektotal_fg_dark": "#9ecbff",
    "empty_bg_dark": "#1b2230",
    "hours_fg_dark": "#c9d1d9",
    "planned_work_bg_dark": "#3f5f87",
    "planned_free_bg_dark": "#7a6a44",
    "timer_bg_dark": "#1f2835",
    "timer_text_dark": "#e6edf7",
    "timer_btn_dark": "#355071",
}

GLASS_OPACITY_MIN = 0.15
GLASS_OPACITY_MAX = 0.95


def normalize_hhmm(value: str | None, default: str = "00:00") -> str:
    if value is None:
        return default
    s = str(value).strip()
    if not s:
        return default
    if ":" in s:
        parts = s.split(":")
        if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
            h = min(23, max(0, int(parts[0])))
            m = min(59, max(0, int(parts[1])))
            return f"{h:02d}:{m:02d}"
    digits = "".join(ch for ch in s if ch.isdigit())
    if not digits:
        return default
    if len(digits) <= 2:
        h = 0
        m = int(digits)
    else:
        h = int(digits[:-2])
        m = int(digits[-2:])
    h = min(23, max(0, h))
    m = min(59, max(0, m))
    return f"{h:02d}:{m:02d}"


def parse_hhmm_strict(value: str | None) -> bool:
    if value is None:
        return False
    s = str(value).strip()
    if len(s) != 5 or s[2] != ":":
        return False
    hh, mm = s.split(":")
    if not (hh.isdigit() and mm.isdigit()):
        return False
    h, m = int(hh), int(mm)
    return 0 <= h <= 23 and 0 <= m <= 59


def hhmm_to_minutes(value: str | None) -> int:
    if not value:
        return 0
    hhmm = normalize_hhmm(value, "00:00")
    hh, mm = hhmm.split(":")
    return int(hh) * 60 + int(mm)


def minutes_to_hhmm(minutes: int) -> str:
    m = max(0, int(minutes))
    return f"{m // 60:02d}:{m % 60:02d}"


def seconds_to_hhmmss(seconds: int) -> str:
    s = max(0, int(seconds))
    hh = s // 3600
    mm = (s % 3600) // 60
    ss = s % 60
    return f"{hh:02d}:{mm:02d}:{ss:02d}"


def hhmmss_to_seconds(value) -> int:
    if value is None:
        return 0
    if isinstance(value, int):
        return max(0, value)
    txt = str(value).strip().replace(".", ":")
    if not txt:
        return 0
    parts = txt.split(":")
    try:
        parts_i = [int(p) for p in parts]
    except ValueError:
        return 0
    if len(parts_i) == 3:
        h, m, s = parts_i
    elif len(parts_i) == 2:
        h, m = parts_i
        s = 0
    else:
        return 0
    return max(0, h * 3600 + m * 60 + s)


def normalize_to_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%d-%m-%Y"):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                pass
    return None


def daterange(start_date: date, end_date: date):
    for n in range((end_date - start_date).days + 1):
        yield start_date + timedelta(n)


@dataclass
class DayData:
    w: str = "00:00"
    v: str = "00:00"
    z: str = "00:00"
    worked: str = ""


def split_reason_and_type(raw_text: str) -> tuple[str, str]:
    def _looks_like_free_prefix(value: str) -> bool:
        s = value.strip().lower()
        if s == "vrij":
            return True
        if s.endswith(" vrij") and len(s) >= 10 and s[2] == ":" and s[:2].isdigit() and s[3:5].isdigit():
            return True
        return False

    if not raw_text:
        return "", "Vrij"
    txt = str(raw_text).strip()
    if " - " in txt:
        left, right = txt.split(" - ", 1)
        if _looks_like_free_prefix(left):
            return "", right.strip()
        kind = left.split("(", 1)[0].strip() or "Vrij"
        return kind, right.strip()
    if _looks_like_free_prefix(txt):
        return "", ""
    kind = txt.split("(", 1)[0].strip() or "Vrij"
    return kind, ""


def parse_nl_date(value: str) -> date | None:
    s = str(value).strip()
    if not s:
        return None
    for fmt in ("%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def normalize_extra_info_options(options: list[str] | None) -> list[str]:
    out: list[str] = []
    seen = set()

    for d in DEFAULT_EXTRA_INFO_OPTIONS:
        k = d.casefold()
        if k not in seen:
            seen.add(k)
            out.append(d)

    for raw in options or []:
        v = str(raw).strip()
        if not v:
            continue
        if v.casefold() == "optioneel":
            continue
        k = v.casefold()
        if k in seen:
            continue
        seen.add(k)
        out.append(v)

    return out


def normalize_extra_info_enabled(options: list[str] | None, enabled_options: list[str] | None) -> list[str]:
    opts = normalize_extra_info_options(options)
    enabled = {str(v).strip().casefold() for v in (enabled_options or []) if str(v).strip()}
    if not enabled:
        return opts
    return [opt for opt in opts if opt.casefold() in enabled]


def default_extra_info_color(option: str) -> str:
    preset = {
        "cursus": "#5b6f8f",
        "bijzonder verlof": "#7b5f3e",
        "zwangerschap": "#7a4f79",
    }
    k = option.strip().casefold()
    if k in preset:
        return preset[k]
    h = hashlib.md5(k.encode("utf-8")).hexdigest()
    r = 72 + (int(h[0:2], 16) % 92)
    g = 72 + (int(h[2:4], 16) % 92)
    b = 72 + (int(h[4:6], 16) % 92)
    return f"#{r:02x}{g:02x}{b:02x}"


class HhmmEntryFilter(QObject):
    DIGIT_POS = (0, 1, 3, 4)

    def __init__(self, line_edit: QLineEdit, default: str = "00:00"):
        super().__init__(line_edit)
        self.line_edit = line_edit
        self.default = default

    def _ensure_mask(self):
        cur = self.line_edit.text()
        if len(cur) != 5 or (len(cur) > 2 and cur[2] != ":"):
            self.line_edit.setText(normalize_hhmm(cur if cur else self.default, self.default))

    def _set_char(self, idx: int, ch: str):
        s = self.line_edit.text()
        if len(s) != 5:
            s = self.default
        out = s[:idx] + ch + s[idx + 1 :]
        self.line_edit.setText(out)

    def _next_digit_pos(self, pos: int) -> int:
        for p in self.DIGIT_POS:
            if p >= pos:
                return p
        return self.DIGIT_POS[0]

    def _prev_digit_pos(self, pos: int) -> int:
        for p in reversed(self.DIGIT_POS):
            if p <= pos:
                return p
        return self.DIGIT_POS[0]

    def eventFilter(self, obj, event):
        if obj is not self.line_edit:
            return False

        if event.type() == QEvent.FocusIn:
            self._ensure_mask()
            self.line_edit.setCursorPosition(0)
            return False

        if event.type() == QEvent.FocusOut:
            self.line_edit.setText(normalize_hhmm(self.line_edit.text() or self.default, self.default))
            return False

        if event.type() != QEvent.KeyPress:
            return False

        if event.modifiers() & (Qt.ControlModifier | Qt.AltModifier | Qt.MetaModifier):
            return False

        self._ensure_mask()
        key = event.key()
        txt = event.text()
        pos = self.line_edit.cursorPosition()

        if txt and txt.isdigit():
            target = self._next_digit_pos(pos if pos <= 4 else 5)
            self._set_char(target, txt)
            next_pos = self.DIGIT_POS[0]
            for p in self.DIGIT_POS:
                if p > target:
                    next_pos = p
                    break
            if target == self.DIGIT_POS[-1]:
                next_pos = self.DIGIT_POS[0]
            self.line_edit.setCursorPosition(next_pos)
            return True

        if key == Qt.Key_Backspace:
            if pos > 0:
                target = self._prev_digit_pos(pos - 1)
                self._set_char(target, "0")
                self.line_edit.setCursorPosition(target)
            return True

        if key == Qt.Key_Delete:
            target = self._next_digit_pos(pos)
            self._set_char(target, "0")
            self.line_edit.setCursorPosition(target)
            return True

        if key == Qt.Key_Left:
            target = self._prev_digit_pos(max(0, pos - 1))
            self.line_edit.setCursorPosition(target)
            return True

        if key == Qt.Key_Right:
            target = self._next_digit_pos(min(4, pos + 1))
            self.line_edit.setCursorPosition(target)
            return True

        if key == Qt.Key_Home:
            self.line_edit.setCursorPosition(0)
            return True

        if key == Qt.Key_End:
            self.line_edit.setCursorPosition(4)
            return True

        if key in (Qt.Key_Tab, Qt.Key_Backtab, Qt.Key_Up, Qt.Key_Down, Qt.Key_Return, Qt.Key_Enter):
            return False

        if txt == ":":
            return True

        if txt and txt.isprintable():
            return True

        return False


def force_hhmm_line_edit(line_edit: QLineEdit, default: str = "00:00"):
    line_edit.setText(normalize_hhmm(line_edit.text() or default, default))
    filt = HhmmEntryFilter(line_edit, default=default)
    line_edit.installEventFilter(filt)
    line_edit._hhmm_filter = filt


class FramelessDialog(QDialog):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowFlag(Qt.FramelessWindowHint, True)
        self._drag_active = False
        self._drag_offset = QPoint()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_active = True
            self._drag_offset = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._drag_active:
            self.move(event.globalPosition().toPoint() - self._drag_offset)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_active = False
        super().mouseReleaseEvent(event)


class CalendarCellDelegate(QStyledItemDelegate):
    def __init__(self, owner):
        super().__init__(owner.table)
        self.owner = owner

    def paint(self, painter, option, index):
        super().paint(painter, option, index)
        if bool(index.data(Qt.UserRole)):
            painter.save()
            painter.setPen(QPen(QColor("#2e3948"), 1))
            painter.drawRect(option.rect.adjusted(0, 0, -1, -1))
            painter.restore()


class ExcelStore:
    """Persistente opslaglaag (Excel).

    Alle workbook-migraties, normalisatie en IO lopen via deze klasse.
    UI mag hierop lezen/schrijven, maar niet direct aan openpyxl-sheets zitten.
    """
    def __init__(self, year: int, base_dir: str):
        self.year = year
        self.base_dir = base_dir
        self.path = os.path.join(base_dir, f"Time_tabel_{year}.xlsx")
        self.nl_holidays = holidays.NL(years=[year])
        self.planned_data: dict[str, dict[str, str]] = {}
        self.worked_data: dict[str, str] = {}
        self.vakantie_dagen: dict[str, str] = {}
        self.school_vakanties: dict[str, str] = {}
        self.extra_info_data: dict[str, str] = {}
        self.weekday_pattern: dict[int, dict[str, str]] = {}
        self.extra_info_options: list[str] = list(DEFAULT_EXTRA_INFO_OPTIONS)
        self.extra_info_enabled: list[str] = list(DEFAULT_EXTRA_INFO_OPTIONS)
        self.data_log: dict[str, dict[str, int]] = {}
        self.day_max_minutes = 480
        self.school_region = "zuid"
        self._load_or_init()
        self.load_all()

    def _load_or_init(self):
        created_new = False
        if not os.path.exists(self.path):
            created_new = True
            wb = Workbook()
            wb.remove(wb.active)
            for month_name in MONTHS_NL:
                ws = wb.create_sheet(month_name)
                ws.append(["WK"] + WEEKDAYS + ["Wk totaal"])
            ws_vrij = wb.create_sheet("vrije dagen")
            ws_vrij.append(["datum", "type"])
            ws_plan = wb.create_sheet("Planning")
            ws_plan.append(["datum", "W"])
            ws_sv = wb.create_sheet("Schoolvakanties")
            ws_sv.append(["datum", "vakantie"])
            ws_wp = wb.create_sheet("Werkpatroon")
            ws_wp.append(["weekdag", "W", "MAX"])
            defaults = {
                "MA": ("08:00", "00:00", "00:00", "08:00"),
                "DI": ("08:00", "00:00", "00:00", "08:00"),
                "WO": ("08:00", "00:00", "00:00", "08:00"),
                "DO": ("08:00", "00:00", "00:00", "08:00"),
                "VR": ("08:00", "00:00", "00:00", "08:00"),
                "ZA": ("00:00", "00:00", "00:00", "00:00"),
                "ZO": ("00:00", "00:00", "00:00", "00:00"),
            }
            for wd, vals in defaults.items():
                ws_wp.append([wd, vals[0], vals[3]])
            ws_set = wb.create_sheet("Instellingen")
            ws_set.append(["key", "value"])
            ws_set.append(["dag_max", "08:00"])
            ws_ai = wb.create_sheet("AanvullendeInfo")
            ws_ai.append(["optie", "actief"])
            for opt in DEFAULT_EXTRA_INFO_OPTIONS:
                ws_ai.append([opt, 1])
            ws_aid = wb.create_sheet("AanvullendeInfoData")
            ws_aid.append(["datum", "info"])
            ws_log = wb.create_sheet("Data_log")
            ws_log.append(["Datum", "Tijd", "Werk_sec", "Idle_sec", "Call_sec"])
            wb.save(self.path)
        self.wb = load_workbook(self.path)
        self.ensure_sheets()
        if created_new:
            self.seed_public_holidays_to_free_days()
            self.seed_school_holidays_from_api(region_filter=None)
            self.wb.save(self.path)

    def ensure_sheets(self):
        if "Planning" not in self.wb.sheetnames:
            ws = self.wb.create_sheet("Planning")
            ws.append(["datum", "W"])
        if "vrije dagen" not in self.wb.sheetnames:
            ws = self.wb.create_sheet("vrije dagen")
            ws.append(["datum", "type"])
        if "Schoolvakanties" not in self.wb.sheetnames:
            ws = self.wb.create_sheet("Schoolvakanties")
            ws.append(["datum", "vakantie", "regio"])
        else:
            ws = self.wb["Schoolvakanties"]
            h1 = str(ws.cell(1, 1).value or "").strip().casefold()
            h2 = str(ws.cell(1, 2).value or "").strip().casefold()
            h3 = str(ws.cell(1, 3).value or "").strip().casefold()
            if h1 != "datum" or h2 != "vakantie":
                ws.cell(1, 1, "datum")
                ws.cell(1, 2, "vakantie")
            if h3 != "regio":
                ws.cell(1, 3, "regio")
        if "Werkpatroon" not in self.wb.sheetnames:
            ws = self.wb.create_sheet("Werkpatroon")
            ws.append(["weekdag", "W", "MAX"])
        if "Instellingen" not in self.wb.sheetnames:
            ws = self.wb.create_sheet("Instellingen")
            ws.append(["key", "value"])
            ws.append(["dag_max", "08:00"])
        if "AanvullendeInfo" not in self.wb.sheetnames:
            ws = self.wb.create_sheet("AanvullendeInfo")
            ws.append(["optie", "actief"])
            for opt in DEFAULT_EXTRA_INFO_OPTIONS:
                ws.append([opt, 1])
        if "AanvullendeInfoData" not in self.wb.sheetnames:
            ws = self.wb.create_sheet("AanvullendeInfoData")
            ws.append(["datum", "info"])
        if "Data_log" not in self.wb.sheetnames:
            ws = self.wb.create_sheet("Data_log")
            ws.append(["Datum", "Tijd", "Werk_sec", "Idle_sec", "Call_sec"])
        for m in MONTHS_NL:
            if m not in self.wb.sheetnames:
                ws = self.wb.create_sheet(m)
                ws.append(["WK"] + WEEKDAYS + ["Wk totaal"])
        self._normalize_planning_sheet()
        self._normalize_workpattern_sheet()
        self.wb.save(self.path)

    def seed_public_holidays_to_free_days(self):
        ws = self.wb["vrije dagen"]
        existing = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[0]:
                d = normalize_to_date(row[0])
                if d:
                    existing.add(d)

        for d in sorted(self.nl_holidays.keys()):
            if d.year != self.year:
                continue
            if d in existing:
                continue
            ws.append([d, str(self.nl_holidays.get(d) or "Feestdag")])
            existing.add(d)

    def fetch_school_holidays_from_api(self, region_filter: str | None = None) -> list[tuple[str, str, str]]:
        url = "https://opendata.rijksoverheid.nl/v1/infotypes/schoolholidays?output=json"
        req = urllib.request.Request(url, headers={"User-Agent": "TijdplannerPro/1.0"})
        out: dict[tuple[str, str], str] = {}

        try:
            with urllib.request.urlopen(req, timeout=12) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
        except Exception:
            return []

        items = payload if isinstance(payload, list) else payload.get("result", []) if isinstance(payload, dict) else []
        year_str = str(self.year)
        rf = (region_filter or "").strip().casefold()

        for item in items:
            content_list = item.get("content", []) if isinstance(item, dict) else []
            for content in content_list:
                schoolyear = str(content.get("schoolyear", ""))
                if year_str not in schoolyear:
                    continue
                vacations = content.get("vacations", [])
                for vac in vacations:
                    vac_type = str(vac.get("type", "")).strip() or "Schoolvakantie"
                    for reg in vac.get("regions", []):
                        region_name = str(reg.get("region", "")).strip().casefold()
                        if not region_name:
                            continue
                        if rf and rf not in region_name:
                            continue
                        if "noord" in region_name:
                            region_key = "noord"
                        elif "midden" in region_name:
                            region_key = "midden"
                        elif "zuid" in region_name:
                            region_key = "zuid"
                        else:
                            continue
                        try:
                            start = date.fromisoformat(str(reg.get("startdate", ""))[:10])
                            end = date.fromisoformat(str(reg.get("enddate", ""))[:10])
                        except Exception:
                            continue
                        if end < start:
                            continue
                        cur = start
                        while cur <= end:
                            if cur.year == self.year:
                                out[(cur.strftime("%Y-%m-%d"), region_key)] = vac_type
                            cur += timedelta(days=1)

        rows = [(d, t, r) for (d, r), t in out.items()]
        rows.sort(key=lambda x: (x[0], x[2]))
        return rows

    def seed_school_holidays_from_api(self, region_filter: str | None = None) -> int:
        ws = self.wb["Schoolvakanties"]
        existing = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[0] and row[1]:
                d = normalize_to_date(row[0])
                if d:
                    reg = str(row[2] if len(row) > 2 and row[2] else "").strip().casefold()
                    if reg not in {"noord", "midden", "zuid"}:
                        reg = ""
                    existing.add((d.strftime("%Y-%m-%d"), reg))

        school_rows = self.fetch_school_holidays_from_api(region_filter=region_filter)
        added = 0
        for d_key, vac_type, reg in school_rows:
            pair = (d_key, reg)
            if pair in existing:
                continue
            d = normalize_to_date(d_key)
            if not d:
                continue
            ws.append([d, vac_type, reg])
            existing.add(pair)
            added += 1
        return added

    def _normalize_planning_sheet(self):
        ws = self.wb["Planning"]
        rows: list[tuple[date, str]] = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            d = normalize_to_date(row[0])
            if not d:
                continue
            w = normalize_hhmm((row[1] if len(row) > 1 else "00:00") or "00:00")
            rows.append((d, w))
        rows.sort(key=lambda x: x[0])
        if ws.max_row > 0:
            ws.delete_rows(1, ws.max_row)
        ws.append(["datum", "W"])
        for d, w in rows:
            ws.append([d, w])

    def _normalize_workpattern_sheet(self):
        ws = self.wb["Werkpatroon"]
        idx = {"MA": 0, "DI": 1, "WO": 2, "DO": 3, "VR": 4, "ZA": 5, "ZO": 6}
        vals: dict[str, tuple[str, str]] = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            wd = str(row[0]).strip().upper()
            if wd not in idx:
                continue
            w = normalize_hhmm((row[1] if len(row) > 1 else "00:00") or "00:00")
            if len(row) > 4:
                m_raw = row[4]
            else:
                m_raw = row[2] if len(row) > 2 else ("08:00" if idx[wd] < 5 else "00:00")
            m = normalize_hhmm(m_raw or ("08:00" if idx[wd] < 5 else "00:00"))
            vals[wd] = (w, m)
        if ws.max_row > 0:
            ws.delete_rows(1, ws.max_row)
        ws.append(["weekdag", "W", "MAX"])
        for wd in ("MA", "DI", "WO", "DO", "VR", "ZA", "ZO"):
            i = idx[wd]
            default_w = "08:00" if i < 5 else "00:00"
            default_m = "08:00" if i < 5 else "00:00"
            w, m = vals.get(wd, (default_w, default_m))
            ws.append([wd, w, m])

    def load_all(self):
        self.load_patterns()
        self.load_settings()
        self.load_free_days()
        self.load_school_holidays()
        if not self.school_vakanties:
            added = self.seed_school_holidays_from_api(region_filter=None)
            if added:
                self.wb.save(self.path)
                self.load_school_holidays()
        self.load_extra_info_options()
        self.load_extra_info_data()
        self.load_data_log()
        self.load_planning()
        self.load_worked()
        self.highlight_today_excel()

    def load_extra_info_options(self):
        ws = self.wb["AanvullendeInfo"]
        raw_options = []
        raw_enabled = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            opt = str(row[0]).strip()
            if not opt:
                continue
            raw_options.append(opt)
            is_active = True
            if len(row) > 1:
                marker = str(row[1]).strip().casefold()
                if marker in {"0", "false", "nee", "no", "n"}:
                    is_active = False
            if is_active:
                raw_enabled.append(opt)

        opts = normalize_extra_info_options(raw_options)
        if not raw_options:
            self.save_extra_info_options(opts, opts)
            return

        enabled = normalize_extra_info_enabled(opts, raw_enabled)
        self.extra_info_options = opts
        self.extra_info_enabled = enabled

    def save_extra_info_options(self, options: list[str], enabled_options: list[str] | None = None):
        opts = normalize_extra_info_options(options)
        enabled = normalize_extra_info_enabled(opts, enabled_options if enabled_options is not None else opts)
        enabled_set = {o.casefold() for o in enabled}

        ws = self.wb["AanvullendeInfo"]
        if ws.max_row > 0:
            ws.delete_rows(1, ws.max_row)
        ws.append(["optie", "actief"])
        for opt in opts:
            ws.append([opt, 1 if opt.casefold() in enabled_set else 0])

        self.extra_info_options = opts
        self.extra_info_enabled = enabled
        self.wb.save(self.path)

    def load_extra_info_data(self):
        self.extra_info_data.clear()
        ws = self.wb["AanvullendeInfoData"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[0] and row[1]:
                d = normalize_to_date(row[0])
                if d:
                    self.extra_info_data[d.strftime("%Y-%m-%d")] = str(row[1]).strip()

    def load_data_log(self):
        self.data_log.clear()
        ws = self.wb["Data_log"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            d = normalize_to_date(row[0])
            if not d:
                continue
            key = d.strftime("%Y-%m-%d")
            self.data_log[key] = {
                "work": hhmmss_to_seconds(row[2] if len(row) > 2 else 0),
                "idle": hhmmss_to_seconds(row[3] if len(row) > 3 else 0),
                "call": hhmmss_to_seconds(row[4] if len(row) > 4 else 0),
            }

    def _save_extra_info_cell(self, dt: date, info: str):
        ws = self.wb["AanvullendeInfoData"]
        found = None
        for row in ws.iter_rows(min_row=2):
            if normalize_to_date(row[0].value) == dt:
                found = row
                break
        if found:
            found[1].value = info
        else:
            ws.append([dt, info])

    def _delete_extra_info_cell(self, dt: date):
        ws = self.wb["AanvullendeInfoData"]
        rows = []
        for row in ws.iter_rows(min_row=2):
            if normalize_to_date(row[0].value) == dt:
                rows.append(row[0].row)
        for r in sorted(rows, reverse=True):
            ws.delete_rows(r, 1)

    def get_extra_info(self, dt: date) -> str:
        return self.extra_info_data.get(dt.strftime("%Y-%m-%d"), "")

    def load_patterns(self):
        ws = self.wb["Werkpatroon"]
        idx = {"MA": 0, "DI": 1, "WO": 2, "DO": 3, "VR": 4, "ZA": 5, "ZO": 6}
        self.weekday_pattern.clear()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            wd = str(row[0]).strip().upper()
            if wd in idx:
                i = idx[wd]
                max_raw = row[4] if len(row) > 4 else (row[2] if len(row) > 2 else ("08:00" if i < 5 else "00:00"))
                self.weekday_pattern[i] = {
                    "W": normalize_hhmm((row[1] if len(row) > 1 else "00:00") or "00:00"),
                    "V": "00:00",
                    "Z": "00:00",
                    "M": normalize_hhmm(max_raw or ("08:00" if i < 5 else "00:00")),
                }
        for i in range(7):
            if i not in self.weekday_pattern:
                self.weekday_pattern[i] = {"W": "00:00", "V": "00:00", "Z": "00:00", "M": "08:00" if i < 5 else "00:00"}

    def load_settings(self):
        ws = self.wb["Instellingen"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[0] == "dag_max":
                self.day_max_minutes = hhmm_to_minutes(normalize_hhmm(row[1] or "08:00")) or 480
                return

    def load_free_days(self):
        self.vakantie_dagen.clear()
        ws = self.wb["vrije dagen"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[0] and row[1]:
                d = normalize_to_date(row[0])
                if d:
                    self.vakantie_dagen[d.strftime("%Y-%m-%d")] = str(row[1])

    def load_school_holidays(self):
        self.school_vakanties.clear()
        ws = self.wb["Schoolvakanties"]
        region = (self.school_region or "zuid").strip().casefold()
        if region not in {"noord", "midden", "zuid"}:
            region = "zuid"
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0] or not row[1]:
                continue
            d = normalize_to_date(row[0])
            if not d:
                continue
            row_region = str(row[2] if len(row) > 2 and row[2] else "").strip().casefold()
            if row_region and row_region != region:
                continue
            self.school_vakanties[d.strftime("%Y-%m-%d")] = str(row[1])

    def load_planning(self):
        self.planned_data.clear()
        ws = self.wb["Planning"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            d = normalize_to_date(row[0])
            if not d:
                continue
            key = d.strftime("%Y-%m-%d")
            free_txt = self.vakantie_dagen.get(key, "")
            free_h = "00:00"
            if free_txt:
                left = free_txt.split(" - ", 1)[0].strip().lower()
                if left.endswith(" vrij") and len(left) >= 10 and left[2] == ":":
                    free_h = normalize_hhmm(left[:5], "00:00")
            self.planned_data[key] = {
                "W": normalize_hhmm(row[1] or "00:00"),
                "V": free_h,
                "Z": "00:00",
            }

    def load_worked(self):
        self.worked_data.clear()
        for month_name in MONTHS_NL:
            ws = self.wb[month_name]
            month = MONTHS_NL.index(month_name) + 1
            for row in ws.iter_rows(min_row=2, values_only=True):
                wk = row[0] if row else None
                if not isinstance(wk, str) or not wk.startswith("wk"):
                    continue
                try:
                    week = int(wk[2:])
                except ValueError:
                    continue
                for wd in range(7):
                    try:
                        dt = date.fromisocalendar(self.year, week, wd + 1)
                    except ValueError:
                        continue
                    if dt.month != month:
                        continue
                    val = row[wd + 1] if wd + 1 < len(row) else ""
                    self.worked_data[dt.strftime("%Y-%m-%d")] = normalize_hhmm(val or "", "")

    def highlight_today_excel(self):
        today = date.today()
        if today.year != self.year:
            return
        ws = self.wb[MONTHS_NL[today.month - 1]]
        for row in ws.iter_rows(min_row=2):
            wk = row[0].value
            if not isinstance(wk, str) or not wk.startswith("wk"):
                continue
            try:
                week = int(wk[2:])
            except ValueError:
                continue
            for c in range(7):
                try:
                    dt = date.fromisocalendar(self.year, week, c + 1)
                except ValueError:
                    continue
                if dt == today:
                    row[c + 1].fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        ws.sheet_properties.tabColor = "ADD8E6"
        self.wb.save(self.path)

    def get_day(self, dt: date) -> DayData:
        key = dt.strftime("%Y-%m-%d")
        p = self.planned_data.get(key, {"W": "00:00", "V": "00:00", "Z": "00:00"})
        return DayData(
            w=normalize_hhmm(p.get("W", "00:00")),
            v=normalize_hhmm(p.get("V", "00:00")),
            z=normalize_hhmm(p.get("Z", "00:00")),
            worked=normalize_hhmm(self.worked_data.get(key, ""), ""),
        )

    def get_timer_log(self, dt: date) -> dict[str, int]:
        key = dt.strftime("%Y-%m-%d")
        row = self.data_log.get(key)
        if not row:
            return {"work": 0, "idle": 0, "call": 0}
        return {
            "work": int(row.get("work", 0)),
            "idle": int(row.get("idle", 0)),
            "call": int(row.get("call", 0)),
        }

    def set_worked_seconds(self, dt: date, seconds: int):
        key = dt.strftime("%Y-%m-%d")
        worked = seconds_to_hhmmss(seconds)
        self.worked_data[key] = worked
        self._save_worked_cell(dt, worked)
        self.wb.save(self.path)

    def save_timer_log(self, dt: date, work_s: int, idle_s: int, call_s: int):
        key = dt.strftime("%Y-%m-%d")
        self.data_log[key] = {"work": max(0, int(work_s)), "idle": max(0, int(idle_s)), "call": max(0, int(call_s))}
        ws = self.wb["Data_log"]
        found = None
        for row in ws.iter_rows(min_row=2):
            if normalize_to_date(row[0].value) == dt:
                found = row
                break
        now_txt = datetime.now().strftime("%H:%M:%S")
        if found:
            found[1].value = now_txt
            found[2].value = seconds_to_hhmmss(work_s)
            found[3].value = seconds_to_hhmmss(idle_s)
            found[4].value = seconds_to_hhmmss(call_s)
        else:
            ws.append([dt, now_txt, seconds_to_hhmmss(work_s), seconds_to_hhmmss(idle_s), seconds_to_hhmmss(call_s)])
        self.set_worked_seconds(dt, work_s)

    def set_day(self, dt: date, data: DayData, reason: str = "", extra_info: str = ""):
        key = dt.strftime("%Y-%m-%d")
        self.planned_data[key] = {"W": data.w, "V": data.v, "Z": data.z}
        self.worked_data[key] = data.worked
        self._save_planning_row(dt, data)
        self._save_worked_cell(dt, data.worked)
        info = (extra_info or "").strip()
        if info:
            self.extra_info_data[key] = info
            self._save_extra_info_cell(dt, info)
        else:
            self.extra_info_data.pop(key, None)
            self._delete_extra_info_cell(dt)
        if hhmm_to_minutes(data.v) > 0 or hhmm_to_minutes(data.z) > 0:
            abs_val = data.v if hhmm_to_minutes(data.v) > 0 else data.z
            text = f"{abs_val} vrij"
            details = (reason or "").strip()
            if details:
                text = f"{text} - {details}"
            self._save_free_day(dt, text)
        else:
            self._delete_free_day(dt)
        self.wb.save(self.path)

    def _save_planning_row(self, dt: date, d: DayData):
        ws = self.wb["Planning"]
        found = None
        for row in ws.iter_rows(min_row=2, max_col=2):
            if normalize_to_date(row[0].value) == dt:
                found = row
                break
        if found:
            found[1].value = d.w
        else:
            ws.append([dt, d.w])
        rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0]]
        rows.sort(key=lambda x: normalize_to_date(x[0]) or date.min)
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        for i, row in enumerate(rows, start=2):
            ws.cell(i, 1, row[0]).number_format = "DD-MM-YYYY"
            ws.cell(i, 2, normalize_hhmm(row[1] or "00:00"))

    def _save_worked_cell(self, dt: date, worked: str):
        ws = self.wb[MONTHS_NL[dt.month - 1]]
        week_key = f"wk{dt.isocalendar()[1]}"
        row_obj = None
        for row in ws.iter_rows(min_row=2):
            if row[0].value == week_key:
                row_obj = row
                break
        if row_obj is None:
            ws.append([week_key, "", "", "", "", "", "", "", ""])
            row_obj = next(ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row))
        row_obj[dt.weekday() + 1].value = worked

    def _save_free_day(self, dt: date, text: str):
        ws = self.wb["vrije dagen"]
        exists = False
        for row in ws.iter_rows(min_row=2):
            if normalize_to_date(row[0].value) == dt:
                row[1].value = text
                exists = True
                break
        if not exists:
            ws.append([dt, text])
        self.vakantie_dagen[dt.strftime("%Y-%m-%d")] = text

    def _delete_free_day(self, dt: date):
        ws = self.wb["vrije dagen"]
        rows = []
        for row in ws.iter_rows(min_row=2):
            if normalize_to_date(row[0].value) == dt:
                rows.append(row[0].row)
        for r in sorted(rows, reverse=True):
            ws.delete_rows(r, 1)
        self.vakantie_dagen.pop(dt.strftime("%Y-%m-%d"), None)

    def get_day_limit(self, dt: date) -> int:
        raw = self.weekday_pattern.get(dt.weekday(), {}).get("M", "08:00")
        return hhmm_to_minutes(normalize_hhmm(raw, "08:00"))

    def day_reason(self, dt: date) -> str:
        key = dt.strftime("%Y-%m-%d")
        if key in self.vakantie_dagen:
            return self.vakantie_dagen[key]
        if key in self.school_vakanties:
            return self.school_vakanties[key]
        if dt in self.nl_holidays:
            return str(self.nl_holidays.get(dt) or "")
        return ""

    def get_saldo_text(self) -> str:
        norm = 0
        free = 0
        worked = 0
        daily = int((38 * 60) / 5)
        start = date(self.year, 1, 1)
        end = min(date.today(), date(self.year, 12, 31))
        for dt in daterange(start, end):
            if dt.weekday() < 5 and dt not in self.nl_holidays:
                norm += daily
            d = self.get_day(dt)
            free += hhmm_to_minutes(d.v)
            worked += hhmm_to_minutes(d.worked)
        expected = max(0, norm - free)
        diff = worked - expected
        sign = "+" if diff >= 0 else "-"
        return f"Saldo {sign}{minutes_to_hhmm(abs(diff))} (gewerkt {minutes_to_hhmm(worked)} / norm {minutes_to_hhmm(expected)})"

    def export_copy(self) -> str:
        self.wb.save(self.path)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dst = os.path.join(self.base_dir, f"Time_tabel_{self.year}_export_{stamp}.xlsx")
        shutil.copy2(self.path, dst)
        return dst


class DayEditDialog(FramelessDialog):
    def __init__(
        self,
        parent: QWidget,
        dt: date,
        day_data: DayData,
        day_limit: int,
        current_reason: str = "",
        current_extra_info: str = "",
        extra_info_options: list[str] | None = None,
        extra_info_colors: dict[str, str] | None = None,
    ):
        super().__init__(parent)
        self.setWindowTitle(f"Dag bewerken - {dt.strftime('%d-%m-%Y')}")
        self.setModal(True)
        self.data_out: DayData | None = None
        self.reason_out = ""
        self.extra_info_out = ""
        self.day_limit = day_limit
        self.existing_worked = day_data.worked

        self.e_w = QLineEdit(day_data.w)
        self.e_v = QLineEdit(day_data.v)
        self.e_reason = QLineEdit()

        self.extra_options = normalize_extra_info_options(extra_info_options)
        self.extra_info_colors = dict(extra_info_colors or {})
        existing_type, existing_reason = split_reason_and_type(current_reason)
        selected_extra = current_extra_info.strip()
        if not selected_extra and existing_type and existing_type != "Vrij":
            selected_extra = existing_type
        if selected_extra and selected_extra.casefold() not in {o.casefold() for o in self.extra_options}:
            self.extra_options.append(selected_extra)

        self.extra_buttons: dict[str, QPushButton] = {}
        self.selected_extra: str | None = None
        self.extra_host = QWidget()
        chip_grid = QGridLayout(self.extra_host)
        chip_grid.setContentsMargins(0, 0, 0, 0)
        chip_grid.setHorizontalSpacing(6)
        chip_grid.setVerticalSpacing(6)
        cols = 3
        for idx, opt in enumerate(self.extra_options):
            btn = QPushButton(opt)
            btn.setCheckable(True)
            btn.clicked.connect(lambda checked, o=opt: self._on_extra_clicked(o, checked))
            btn.setStyleSheet(self._extra_button_style(self.extra_info_colors.get(opt, default_extra_info_color(opt))))
            self.extra_buttons[opt] = btn
            chip_grid.addWidget(btn, idx // cols, idx % cols)

        self.e_reason.setText(existing_reason)
        for w in (self.e_w, self.e_v):
            w.setPlaceholderText("hh:mm")
            force_hhmm_line_edit(w)
        self._syncing = False

        if selected_extra:
            self._set_extra_selection(selected_extra)
        else:
            self._set_extra_selection(None)

        form = QFormLayout()
        form.addRow("Werk (W)", self.e_w)
        form.addRow("Vrij (V)", self.e_v)
        form.addRow("Aanvullende informatie", self.extra_host)
        form.addRow("Reden vrij", self.e_reason)

        btn_cancel = QPushButton("Annuleren")
        btn_save = QPushButton("Opslaan")
        btn_cancel.clicked.connect(self.reject)
        btn_save.clicked.connect(self.on_save)

        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(btn_cancel)
        row.addWidget(btn_save)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)
        root.addLayout(form)
        root.addWidget(QLabel(f"Daglimiet: {minutes_to_hhmm(day_limit)}"))
        self.lbl_total = QLabel("")
        root.addWidget(self.lbl_total)
        root.addLayout(row)
        self.resize(560, 320)
        self.e_w.editingFinished.connect(lambda: self._rebalance("w"))
        self.e_v.editingFinished.connect(lambda: self._rebalance("v"))
        self._rebalance(None)

    def _extra_button_style(self, color: str) -> str:
        return (
            f"QPushButton {{ background:{color}; color:#f3f8ff; border:1px solid #5b6b80; border-radius:7px; padding:6px 10px; }}"
            "QPushButton:hover { border:1px solid #8aa2bf; }"
            "QPushButton:checked { border:2px solid #e6edf3; }"
        )

    def _on_extra_clicked(self, option: str, checked: bool):
        if checked:
            self._set_extra_selection(option)
        else:
            self._set_extra_selection(None)

    def _set_extra_selection(self, option: str | None):
        found = None
        if option:
            for key in self.extra_buttons.keys():
                if key.casefold() == option.strip().casefold():
                    found = key
                    break
        self.selected_extra = found
        for key, btn in self.extra_buttons.items():
            btn.blockSignals(True)
            btn.setChecked(key == found)
            btn.blockSignals(False)

    def _set_minutes(self, edit: QLineEdit, minutes: int):
        edit.setText(minutes_to_hhmm(max(0, minutes)))

    def _rebalance(self, changed: str | None):
        if self._syncing:
            return
        self._syncing = True
        w_min = min(hhmm_to_minutes(normalize_hhmm(self.e_w.text())), self.day_limit)
        v_min = min(hhmm_to_minutes(normalize_hhmm(self.e_v.text())), self.day_limit)
        total = w_min + v_min
        if total > self.day_limit:
            overflow = total - self.day_limit
            if changed == "w":
                reduce_v = min(v_min, overflow)
                v_min -= reduce_v
                overflow -= reduce_v
                if overflow > 0:
                    w_min = max(0, w_min - overflow)
            elif changed == "v":
                reduce_w = min(w_min, overflow)
                w_min -= reduce_w
                overflow -= reduce_w
                if overflow > 0:
                    v_min = max(0, v_min - overflow)
            else:
                reduce_v = min(v_min, overflow)
                v_min -= reduce_v
                overflow -= reduce_v
                if overflow > 0:
                    w_min = max(0, w_min - overflow)
        self._set_minutes(self.e_w, w_min)
        self._set_minutes(self.e_v, v_min)
        total = w_min + v_min
        self.lbl_total.setText(f"Totaal: {minutes_to_hhmm(total)} / {minutes_to_hhmm(self.day_limit)}")
        self._syncing = False

    def on_save(self):
        self._rebalance(None)
        w = normalize_hhmm(self.e_w.text())
        v = normalize_hhmm(self.e_v.text())
        z = "00:00"
        worked = normalize_hhmm(self.existing_worked, "")
        for x in (w, v):
            if not parse_hhmm_strict(x):
                QMessageBox.warning(self, "Fout", "Gebruik hh:mm (00:00 t/m 23:59).")
                return
        if hhmm_to_minutes(w) + hhmm_to_minutes(v) + hhmm_to_minutes(z) > self.day_limit:
            QMessageBox.warning(self, "Fout", f"Planning mag niet boven {minutes_to_hhmm(self.day_limit)}.")
            return
        self.data_out = DayData(w=w, v=v, z=z, worked=worked)
        self.reason_out = self.e_reason.text().strip()
        picked = (self.selected_extra or "").strip()
        self.extra_info_out = picked
        self.accept()


class MonthCard(QGroupBox):
    day_double_clicked = Signal(date)

    def __init__(
        self,
        store: ExcelStore,
        year: int,
        month: int,
        mode: str,
        focus_mode: bool = False,
        dark_mode: bool = False,
        colors: dict[str, str] | None = None,
        extra_info_colors: dict[str, str] | None = None,
    ):
        q = (month - 1) // 3 + 1
        super().__init__(f"{MONTHS_NL[month - 1]}  (Q{q})")
        self.store = store
        self.year = year
        self.month = month
        self.mode = mode
        self.focus_mode = focus_mode
        self.dark_mode = dark_mode
        self.colors = colors or dict(COLOR_DEFAULTS)
        self.extra_info_colors = dict(extra_info_colors or {})
        self.cell_map: dict[tuple[int, int], date] = {}
        self.table = QTableWidget(self)
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels(["WK"] + WEEKDAYS + ["Wk totaal"])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionMode(QTableWidget.NoSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setMouseTracking(True)
        self.table.setShowGrid(False)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setItemDelegate(CalendarCellDelegate(self))
        self.table.cellDoubleClicked.connect(self._on_double)
        self.table.setMinimumHeight(560 if focus_mode else 220)
        self.table.horizontalHeader().setMinimumSectionSize(70)
        if not self.focus_mode:
            self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            self.setMinimumWidth(680)
            self.setMaximumWidth(680)
        else:
            self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self.setMinimumWidth(0)
            self.setMaximumWidth(16777215)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(8, 10, 8, 8)
        lay.addWidget(self.table)
        self.refresh()

    def _weeks(self):
        first = date(self.year, self.month, 1)
        days = calendar.monthrange(self.year, self.month)[1]
        week_rows = {}
        for offset in range(days):
            dt = first + timedelta(days=offset)
            wn = dt.isocalendar()[1]
            wd = dt.weekday()
            if wn not in week_rows:
                week_rows[wn] = [None] * 7
            week_rows[wn][wd] = dt
        return sorted(week_rows.items())

    def _bg(self, dt: date) -> QColor:
        key = dt.strftime("%Y-%m-%d")
        if dt == date.today():
            return QColor(self.colors["bg_today_dark"])
        if key in self.store.vakantie_dagen:
            return QColor(self.colors["bg_vacation_dark"])
        if dt in self.store.nl_holidays:
            return QColor(self.colors["bg_holiday_dark"])
        if key in self.store.school_vakanties:
            return QColor(self.colors["bg_school_dark"])
        if dt.weekday() >= 5:
            return QColor(self.colors["bg_weekend_dark"])
        return QColor(self.colors["bg_default_dark"])

    def _extra_info_color(self, info: str) -> str | None:
        if not info:
            return None
        info_key = info.strip().casefold()
        for k, v in self.extra_info_colors.items():
            if k.strip().casefold() == info_key:
                return v
        return None

    def _planned_hours_brush(self, dt: date, fallback: QColor) -> QBrush:
        info = self.store.get_extra_info(dt).strip()
        override = self._extra_info_color(info)
        if override:
            return QBrush(QColor(override))
        d = self.store.get_day(dt)
        w_m = hhmm_to_minutes(d.w)
        v_m = hhmm_to_minutes(d.v)
        total = w_m + v_m + hhmm_to_minutes(d.z)
        if total <= 0:
            return QBrush(fallback)
        work_clr = QColor(self.colors["planned_work_bg_dark"])
        free_clr = QColor(self.colors["planned_free_bg_dark"])
        if w_m > 0 and v_m > 0:
            ratio = w_m / max(1, (w_m + v_m))
            grad = QLinearGradient(0.0, 0.0, 1.0, 0.0)
            grad.setCoordinateMode(QGradient.ObjectBoundingMode)
            stop = max(0.05, min(0.95, ratio))
            grad.setColorAt(0.0, work_clr)
            grad.setColorAt(stop, work_clr)
            grad.setColorAt(min(1.0, stop + 0.001), free_clr)
            grad.setColorAt(1.0, free_clr)
            return QBrush(grad)
        if w_m > 0:
            return QBrush(work_clr)
        if v_m > 0:
            return QBrush(free_clr)
        return QBrush(fallback)

    def _hours_text(self, dt: date) -> str:
        d = self.store.get_day(dt)
        if self.mode == "planned":
            total = hhmm_to_minutes(d.w) + hhmm_to_minutes(d.v) + hhmm_to_minutes(d.z)
            if total == 0:
                reason = self.store.day_reason(dt)
                return reason[:18] if reason else "00:00"
            return minutes_to_hhmm(total)
        if hhmm_to_minutes(d.w) == 0 and hhmm_to_minutes(d.v) > 0:
            reason = self.store.day_reason(dt)
            return f"{d.v}\n{reason[:14]}" if reason else d.v
        return d.worked or "00:00"

    def _day_number_color(self, dt: date) -> QColor:
        key = dt.strftime("%Y-%m-%d")
        if dt == date.today():
            return QColor(self.colors["daynum_today_dark"])
        if self.mode == "planned":
            d = self.store.get_day(dt)
            has_planning = (hhmm_to_minutes(d.w) + hhmm_to_minutes(d.v) + hhmm_to_minutes(d.z)) > 0
            if has_planning:
                if dt in self.store.nl_holidays:
                    return QColor(self.colors["daynum_holiday_dark"])
                if key in self.store.school_vakanties:
                    return QColor(self.colors["daynum_school_dark"])
                if key in self.store.vakantie_dagen:
                    return QColor(self.colors["daynum_vacation_dark"])
                return QColor(self.colors["daynum_default_dark"])
            if dt.weekday() >= 5:
                return QColor(self.colors["daynum_weekend_dark"])
            return QColor(self.colors["daynum_default_dark"])
        if dt in self.store.nl_holidays:
            return QColor(self.colors["daynum_holiday_dark"])
        if key in self.store.vakantie_dagen:
            return QColor(self.colors["daynum_vacation_dark"])
        if key in self.store.school_vakanties:
            return QColor(self.colors["daynum_school_dark"])
        if dt.weekday() >= 5:
            return QColor(self.colors["daynum_weekend_dark"])
        return QColor(self.colors["daynum_default_dark"])

    def refresh(self):
        total_label = "Gepland" if self.mode == "planned" else "Gewerkt"
        self.table.setHorizontalHeaderLabels(["WK"] + WEEKDAYS + [total_label])
        weeks = self._weeks()
        self.table.clearSpans()
        self.table.setRowCount(len(weeks) * 2)
        self.cell_map.clear()
        day_h = 36 if self.focus_mode else 18
        hour_h = 64 if self.focus_mode else 24
        for week_i, (week_no, week_days) in enumerate(weeks):
            day_row = week_i * 2
            hour_row = day_row + 1
            self.table.setRowHeight(day_row, day_h)
            self.table.setRowHeight(hour_row, hour_h)

            wk = QTableWidgetItem(f"wk{week_no}" if week_no is not None else "")
            wk.setTextAlignment(Qt.AlignCenter)
            wk.setBackground(QColor(self.colors["weeknum_bg_dark"]))
            wk.setForeground(QColor(self.colors["weeknum_fg_dark"]))
            wk.setData(Qt.UserRole, True)
            self.table.setItem(day_row, 0, wk)
            self.table.setSpan(day_row, 0, 2, 1)

            total_min = 0
            iso_year = None
            iso_week = None
            seed_dt = next((x for x in week_days if x is not None), None)
            if seed_dt:
                iso = seed_dt.isocalendar()
                iso_year, iso_week = iso[0], iso[1]
            for wd, dt in enumerate(week_days):
                c = wd + 1
                day_item = QTableWidgetItem("")
                day_item.setTextAlignment(Qt.AlignCenter)
                hours_item = QTableWidgetItem("")
                hours_item.setTextAlignment(Qt.AlignCenter)
                if dt:
                    bg = self._bg(dt)
                    day_item.setText(str(dt.day))
                    day_item.setBackground(bg)
                    day_item.setForeground(self._day_number_color(dt))
                    hours_item.setText(self._hours_text(dt))
                    hours_item.setBackground(self._planned_hours_brush(dt, bg))
                    hours_item.setForeground(QColor(self.colors["hours_fg_dark"]))
                    day_item.setData(Qt.UserRole, True)
                    hours_item.setData(Qt.UserRole, True)
                    reason = self.store.day_reason(dt)
                    if not reason and dt.strftime("%Y-%m-%d") in self.store.school_vakanties:
                        reason = self.store.school_vakanties.get(dt.strftime("%Y-%m-%d"), "")
                    if reason:
                        day_item.setToolTip(reason)
                        hours_item.setToolTip(reason)
                    self.cell_map[(day_row, c)] = dt
                    self.cell_map[(hour_row, c)] = dt
                    day = self.store.get_day(dt)
                    if self.mode == "planned":
                        total_min += hhmm_to_minutes(day.w) + hhmm_to_minutes(day.v) + hhmm_to_minutes(day.z)
                    else:
                        total_min += hhmm_to_minutes(day.worked)
                else:
                    blank = QColor(self.colors["empty_bg_dark"])
                    day_item.setBackground(blank)
                    hours_item.setBackground(blank)
                    day_item.setData(Qt.UserRole, False)
                    hours_item.setData(Qt.UserRole, False)
                self.table.setItem(day_row, c, day_item)
                self.table.setItem(hour_row, c, hours_item)

            if iso_year is not None and iso_week is not None:
                week_total = 0
                for d in range(1, 8):
                    try:
                        wdt = date.fromisocalendar(iso_year, iso_week, d)
                    except ValueError:
                        continue
                    day = self.store.get_day(wdt)
                    if self.mode == "planned":
                        week_total += hhmm_to_minutes(day.w) + hhmm_to_minutes(day.v) + hhmm_to_minutes(day.z)
                    else:
                        week_total += hhmm_to_minutes(day.worked)
                tot_txt = minutes_to_hhmm(week_total)
            else:
                tot_txt = minutes_to_hhmm(total_min)
            tot = QTableWidgetItem(tot_txt)
            tot.setTextAlignment(Qt.AlignCenter)
            tot.setBackground(QColor(self.colors["weektotal_bg_dark"]))
            tot.setForeground(QColor(self.colors["weektotal_fg_dark"]))
            tot.setFont(QFont("Segoe UI", 9, QFont.Bold))
            tot.setData(Qt.UserRole, True)
            self.table.setItem(day_row, 8, tot)
            self.table.setSpan(day_row, 8, 2, 1)

        if self.focus_mode:
            self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
            self.table.setColumnWidth(0, 92)
            for c in range(1, 8):
                self.table.horizontalHeader().setSectionResizeMode(c, QHeaderView.Stretch)
            self.table.horizontalHeader().setSectionResizeMode(8, QHeaderView.Fixed)
            self.table.setColumnWidth(8, 132)
        else:
            self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
            self.table.setColumnWidth(0, 56)
            for c in range(1, 8):
                self.table.horizontalHeader().setSectionResizeMode(c, QHeaderView.Fixed)
                self.table.setColumnWidth(c, 66)
            self.table.horizontalHeader().setSectionResizeMode(8, QHeaderView.Fixed)
            self.table.setColumnWidth(8, 106)

        total_h = self.table.horizontalHeader().height()
        for r in range(self.table.rowCount()):
            total_h += self.table.rowHeight(r)
        total_h += 8
        self.table.setMinimumHeight(total_h)
        self.table.setMaximumHeight(total_h)
        if not self.focus_mode:
            # Gebruik vaste hoogte op basis van 6 weekrijen zonder extra lege rasterrijen.
            header_h = self.table.horizontalHeader().height()
            fixed_h = header_h + (6 * (18 + 24)) + 8
            self.table.setMinimumHeight(fixed_h)
            self.table.setMaximumHeight(fixed_h)
            self.setMinimumHeight(fixed_h + 56)
            self.setMaximumHeight(fixed_h + 56)

    def _on_double(self, row: int, col: int):
        dt = self.cell_map.get((row, col))
        if dt and self.mode == "planned":
            self.day_double_clicked.emit(dt)


class CalendarBoard(QWidget):
    day_double_clicked = Signal(date)

    def __init__(
        self,
        store: ExcelStore,
        year: int,
        mode: str,
        months: list[int],
        per_row: int,
        focus_mode: bool = False,
        dark_mode: bool = False,
        colors: dict[str, str] | None = None,
        extra_info_colors: dict[str, str] | None = None,
    ):
        super().__init__()
        self.store = store
        self.year = year
        self.mode = mode
        self.months = months
        self.per_row = per_row
        self._active_per_row = per_row
        self.focus_mode = focus_mode
        self.dark_mode = dark_mode
        self.colors = colors or dict(COLOR_DEFAULTS)
        self.extra_info_colors = dict(extra_info_colors or {})
        self.cards: list[MonthCard] = []
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)
        root.setContentsMargins(0, 0, 0, 0)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setObjectName("calendarScroll")
        self.host = QWidget()
        self.host.setObjectName("calendarHost")
        self.grid = QGridLayout(self.host)
        self.grid.setContentsMargins(4, 4, 4, 4)
        self.grid.setSpacing(4)
        self.scroll.setWidget(self.host)
        root.addWidget(self.scroll)
        self.rebuild()

    def rebuild(self):
        cols = 1 if self.focus_mode else self.per_row
        self._active_per_row = cols

        while self.grid.count():
            item = self.grid.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self.cards.clear()
        for i, month in enumerate(self.months):
            card = MonthCard(
                self.store,
                self.year,
                month,
                self.mode,
                focus_mode=self.focus_mode,
                dark_mode=self.dark_mode,
                colors=self.colors,
                extra_info_colors=self.extra_info_colors,
            )
            card.day_double_clicked.connect(self.day_double_clicked.emit)
            self.cards.append(card)
            self.grid.addWidget(card, i // cols, i % cols)
        for c in range(cols):
            self.grid.setColumnStretch(c, 1)

    def set_mode(self, mode: str):
        self.mode = mode
        for c in self.cards:
            c.mode = mode
            c.refresh()

    def set_month(self, month: int):
        self.months = [month]
        self.per_row = 1
        self.rebuild()

    def refresh(self):
        for c in self.cards:
            c.refresh()

    def set_dark_mode(self, enabled: bool):
        self.dark_mode = enabled
        for c in self.cards:
            c.dark_mode = enabled
            c.refresh()

    def set_colors(self, colors: dict[str, str]):
        self.colors = colors
        for c in self.cards:
            c.colors = colors
            c.refresh()

    def set_extra_info_colors(self, extra_info_colors: dict[str, str]):
        self.extra_info_colors = dict(extra_info_colors or {})
        for c in self.cards:
            c.extra_info_colors = dict(self.extra_info_colors)
            c.refresh()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        return


class WorkPatternDialog(FramelessDialog):
    def __init__(self, parent: QWidget, store: ExcelStore, year: int, apply_callback=None):
        super().__init__(parent)
        self.store = store
        self.year = year
        self.apply_callback = apply_callback
        self.setWindowTitle("Werkpatroon en Dagmax")
        self.setModal(True)
        self.resize(640, 440)

        self.inputs_w: dict[int, QLineEdit] = {}
        self.inputs_m: dict[int, QLineEdit] = {}
        day_names = ["MA", "DI", "WO", "DO", "VR", "ZA", "ZO"]

        form_box = QWidget()
        grid = QGridLayout(form_box)
        title = QLabel("Weekpatroon (hh:mm)")
        title.setStyleSheet("font: 11pt 'Segoe UI Semibold'; color:#e6edf7; padding-bottom:4px;")
        grid.addWidget(title, 0, 0, 1, 3)
        grid.addWidget(QLabel("Dag"), 1, 0)
        grid.addWidget(QLabel("Werk"), 1, 1)
        grid.addWidget(QLabel("Max dag"), 1, 2)

        for i, dn in enumerate(day_names):
            row = i + 2
            grid.addWidget(QLabel(dn), row, 0)
            e_w = QLineEdit(self.store.weekday_pattern[i]["W"])
            e_m = QLineEdit(self.store.weekday_pattern[i]["M"])
            e_w.setPlaceholderText("hh:mm")
            e_m.setPlaceholderText("hh:mm")
            force_hhmm_line_edit(e_w)
            force_hhmm_line_edit(e_m, "00:00" if i > 4 else "08:00")
            self.inputs_w[i] = e_w
            self.inputs_m[i] = e_m
            grid.addWidget(e_w, row, 1)
            grid.addWidget(e_m, row, 2)

        btn_cancel = QPushButton("Annuleren")
        btn_save = QPushButton("Opslaan")
        btn_apply = QPushButton("Toepassen")
        btn_cancel.clicked.connect(self.reject)
        btn_save.clicked.connect(self._save_only)
        btn_apply.clicked.connect(self._apply)

        self.e_from = QLineEdit(f"01-01-{self.year}")
        self.e_to = QLineEdit(f"31-12-{self.year}")
        self.e_from.setPlaceholderText("dd-mm-jjjj")
        self.e_to.setPlaceholderText("dd-mm-jjjj")

        range_row = QHBoxLayout()
        range_row.addWidget(QLabel("Van"))
        range_row.addWidget(self.e_from)
        range_row.addWidget(QLabel("Tot en met"))
        range_row.addWidget(self.e_to)
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_save)
        btn_row.addWidget(btn_apply)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)
        root.addWidget(form_box)
        root.addLayout(range_row)
        root.addLayout(btn_row)

    def _apply(self):
        if not self._save_silent():
            return
        start_date = parse_nl_date(self.e_from.text())
        end_date = parse_nl_date(self.e_to.text())
        if not start_date or not end_date:
            QMessageBox.warning(self, "Fout", "Gebruik datumformaat dd-mm-jjjj voor van en tot.")
            return
        if start_date > end_date:
            QMessageBox.warning(self, "Fout", "Van-datum mag niet na tot-datum liggen.")
            return
        if self.apply_callback:
            self.apply_callback(start_date, end_date)
        self.accept()

    def _save_only(self):
        if self._save_silent():
            self.accept()

    def _save_silent(self):
        ws = self.store.wb["Werkpatroon"]
        if ws.max_row > 0:
            ws.delete_rows(1, ws.max_row)
        ws.append(["weekdag", "W", "MAX"])
        idx_to_day = {0: "MA", 1: "DI", 2: "WO", 3: "DO", 4: "VR", 5: "ZA", 6: "ZO"}
        for i in range(7):
            w = normalize_hhmm(self.inputs_w[i].text())
            m = normalize_hhmm(self.inputs_m[i].text(), "00:00" if i > 4 else "08:00")
            if not (parse_hhmm_strict(w) and parse_hhmm_strict(m)):
                QMessageBox.warning(self, "Fout", "Gebruik hh:mm voor werk en max.")
                return False
            if hhmm_to_minutes(w) > hhmm_to_minutes(m):
                QMessageBox.warning(self, "Fout", f"{idx_to_day[i]}: werk hoger dan max.")
                return False
            self.store.weekday_pattern[i] = {"W": w, "V": "00:00", "Z": "00:00", "M": m}
            ws.append([idx_to_day[i], w, m])
        self.store.wb.save(self.store.path)
        return True


class ColorSettingsDialog(FramelessDialog):
    LABELS = {
        "bg_today_dark": "Vandaag achtergrond",
        "bg_vacation_dark": "Vrij-dag achtergrond",
        "bg_holiday_dark": "Feestdag achtergrond",
        "bg_school_dark": "Schoolvakantie achtergrond",
        "bg_weekend_dark": "Weekend achtergrond",
        "bg_default_dark": "Werkdag achtergrond",
        "daynum_today_dark": "Dagnummer vandaag",
        "daynum_holiday_dark": "Dagnummer feestdag",
        "daynum_vacation_dark": "Dagnummer vrij",
        "daynum_school_dark": "Dagnummer schoolvakantie",
        "daynum_weekend_dark": "Dagnummer weekend",
        "daynum_default_dark": "Dagnummer werkdag",
        "weeknum_bg_dark": "Weeknummer achtergrond",
        "weeknum_fg_dark": "Weeknummer tekst",
        "weektotal_bg_dark": "Weektotaal achtergrond",
        "weektotal_fg_dark": "Weektotaal tekst",
        "empty_bg_dark": "Lege cel achtergrond",
        "hours_fg_dark": "Uren tekst",
        "planned_work_bg_dark": "Planning werk achtergrond",
        "planned_free_bg_dark": "Planning vrij achtergrond",
        "timer_bg_dark": "Timer achtergrond",
        "timer_text_dark": "Timer tekst",
        "timer_btn_dark": "Timer knop",
    }

    def __init__(
        self,
        parent: QWidget,
        colors: dict[str, str],
        extra_info_options: list[str],
        extra_info_colors: dict[str, str],
    ):
        super().__init__(parent)
        self.setWindowTitle("Kleuren kalender")
        self.setModal(True)
        self.resize(700, 820)
        self.colors = dict(colors)
        self.extra_info_options = [o for o in normalize_extra_info_options(extra_info_options) if o.casefold() != "optioneel"]
        self.extra_info_colors = dict(extra_info_colors or {})
        self.preview_buttons: dict[str, QPushButton] = {}
        self.preview_extra_buttons: dict[str, QPushButton] = {}
        for opt in self.extra_info_options:
            if opt not in self.extra_info_colors:
                self.extra_info_colors[opt] = default_extra_info_color(opt)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)
        grid_box = QGroupBox("Kies kleuren")
        grid = QGridLayout(grid_box)

        row = 0
        for key, label in self.LABELS.items():
            grid.addWidget(QLabel(label), row, 0)
            btn = QPushButton("Kies kleur")
            btn.clicked.connect(lambda _, k=key: self.pick_color(k))
            self.preview_buttons[key] = btn
            grid.addWidget(btn, row, 1)
            row += 1

        extra_box = QGroupBox("Aanvullende info kleuren (override planning)")
        extra_grid = QGridLayout(extra_box)
        extra_row = 0
        for opt in self.extra_info_options:
            extra_grid.addWidget(QLabel(opt), extra_row, 0)
            btn = QPushButton("Kies kleur")
            btn.clicked.connect(lambda _, o=opt: self.pick_extra_color(o))
            self.preview_extra_buttons[opt] = btn
            extra_grid.addWidget(btn, extra_row, 1)
            extra_row += 1

        self._refresh_previews()

        btn_reset = QPushButton("Reset standaard")
        btn_reset.clicked.connect(self.reset_defaults)
        btn_cancel = QPushButton("Annuleren")
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("Opslaan")
        btn_ok.clicked.connect(self.accept)

        row_btn = QHBoxLayout()
        row_btn.addWidget(btn_reset)
        row_btn.addStretch(1)
        row_btn.addWidget(btn_cancel)
        row_btn.addWidget(btn_ok)

        root.addWidget(grid_box)
        root.addWidget(extra_box)
        root.addLayout(row_btn)

    def _refresh_previews(self):
        for key, btn in self.preview_buttons.items():
            clr = self.colors.get(key, COLOR_DEFAULTS[key])
            btn.setText(clr.upper())
            btn.setStyleSheet(
                f"QPushButton {{ background:{clr}; color:#ffffff; border:1px solid #334155; border-radius:6px; padding:4px 8px; }}"
            )
        for opt, btn in self.preview_extra_buttons.items():
            clr = self.extra_info_colors.get(opt, default_extra_info_color(opt))
            btn.setText(clr.upper())
            btn.setStyleSheet(
                f"QPushButton {{ background:{clr}; color:#ffffff; border:1px solid #334155; border-radius:6px; padding:4px 8px; }}"
            )

    def pick_color(self, key: str):
        current = QColor(self.colors.get(key, COLOR_DEFAULTS[key]))
        c = QColorDialog.getColor(current, self, f"Kies kleur: {self.LABELS.get(key, key)}")
        if c.isValid():
            self.colors[key] = c.name()
            self._refresh_previews()

    def pick_extra_color(self, option: str):
        current = QColor(self.extra_info_colors.get(option, default_extra_info_color(option)))
        c = QColorDialog.getColor(current, self, f"Kies kleur: {option}")
        if c.isValid():
            self.extra_info_colors[option] = c.name()
            self._refresh_previews()

    def reset_defaults(self):
        self.colors = dict(COLOR_DEFAULTS)
        self.extra_info_colors = {opt: default_extra_info_color(opt) for opt in self.extra_info_options}
        self._refresh_previews()


class TimerSettingsDialog(FramelessDialog):
    def __init__(self, parent: QWidget, idle_threshold_sec: int):
        super().__init__(parent)
        self.setWindowTitle("Timer idle instellingen")
        self.setModal(True)
        self.resize(340, 120)
        self.idle_threshold_sec = max(15, min(3600, int(idle_threshold_sec or 60)))

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(6)

        total = self.idle_threshold_sec
        mins = total // 60
        secs = total % 60

        form = QFormLayout()
        form.setHorizontalSpacing(12)
        self.spin_min = QSpinBox()
        self.spin_min.setRange(0, 60)
        self.spin_min.setValue(mins)
        self.spin_sec = QSpinBox()
        self.spin_sec.setRange(0, 59)
        self.spin_sec.setValue(secs)
        form.addRow("Minuten", self.spin_min)
        form.addRow("Seconden", self.spin_sec)
        root.addLayout(form)

        btn_cancel = QPushButton("Annuleren")
        btn_save = QPushButton("Opslaan")
        btn_cancel.clicked.connect(self.reject)
        btn_save.clicked.connect(self._save)

        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(btn_cancel)
        row.addWidget(btn_save)
        root.addLayout(row)

    def _save(self):
        total = int(self.spin_min.value()) * 60 + int(self.spin_sec.value())
        self.idle_threshold_sec = max(15, min(3600, total))
        self.accept()


class GlassSettingsDialog(FramelessDialog):
    def __init__(self, parent: QWidget, inactive_glass_opacity: float):
        super().__init__(parent)
        self.setWindowTitle("Glassmorphism")
        self.setModal(True)
        self.resize(390, 150)
        try:
            op = float(inactive_glass_opacity)
        except Exception:
            op = 0.72
        self.inactive_glass_opacity = max(GLASS_OPACITY_MIN, min(GLASS_OPACITY_MAX, op))

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(6)

        form = QFormLayout()
        form.setHorizontalSpacing(12)
        self.spin_glass = QSpinBox()
        self.spin_glass.setRange(int(GLASS_OPACITY_MIN * 100), int(GLASS_OPACITY_MAX * 100))
        self.spin_glass.setSuffix("%")
        self.spin_glass.setValue(int(round(self.inactive_glass_opacity * 100)))
        form.addRow("Transparantie bij focusverlies", self.spin_glass)
        root.addLayout(form)

        hint = QLabel(
            f"Wordt toegepast als de app niet actief is (Alt-Tab naar andere app).\n"
            f"Vaste grenzen: min {int(GLASS_OPACITY_MIN * 100)}% / max {int(GLASS_OPACITY_MAX * 100)}%."
        )
        hint.setWordWrap(True)
        root.addWidget(hint)

        btn_cancel = QPushButton("Annuleren")
        btn_save = QPushButton("Opslaan")
        btn_cancel.clicked.connect(self.reject)
        btn_save.clicked.connect(self._save)
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(btn_cancel)
        row.addWidget(btn_save)
        root.addLayout(row)

    def _save(self):
        self.inactive_glass_opacity = max(GLASS_OPACITY_MIN, min(GLASS_OPACITY_MAX, float(self.spin_glass.value()) / 100.0))
        self.accept()


class SchoolRegionDialog(FramelessDialog):
    def __init__(self, parent: QWidget, current_region: str):
        super().__init__(parent)
        self.setWindowTitle("Schoolvakantie regio")
        self.setModal(True)
        self.resize(340, 140)
        self.region = (current_region or "zuid").strip().casefold()
        if self.region not in {"noord", "midden", "zuid"}:
            self.region = "zuid"

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)

        form = QFormLayout()
        self.cb_region = QComboBox()
        self.cb_region.addItems(["noord", "midden", "zuid"])
        self.cb_region.setCurrentText(self.region)
        form.addRow("Regio", self.cb_region)
        root.addLayout(form)

        btn_cancel = QPushButton("Annuleren")
        btn_save = QPushButton("Opslaan")
        btn_cancel.clicked.connect(self.reject)
        btn_save.clicked.connect(self._save)
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(btn_cancel)
        row.addWidget(btn_save)
        root.addLayout(row)

    def _save(self):
        self.region = self.cb_region.currentText().strip().casefold()
        self.accept()


class ExtraInfoSettingsDialog(FramelessDialog):
    def __init__(self, parent: QWidget, options: list[str], enabled_options: list[str]):
        super().__init__(parent)
        self.setWindowTitle("Aanvullende informatie opties")
        self.setModal(True)
        self.resize(720, 520)
        self.options = normalize_extra_info_options(options)
        self.enabled_options = normalize_extra_info_enabled(self.options, enabled_options)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)
        root.addWidget(QLabel("Beheer opties als losse regels en zet per optie actief/inactief."))

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Optie", "Actief"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        root.addWidget(self.table, 1)

        enabled_set = {o.casefold() for o in self.enabled_options}
        for opt in self.options:
            self._add_row(opt, opt.casefold() in enabled_set)

        add_row = QHBoxLayout()
        self.e_new = QLineEdit()
        self.e_new.setPlaceholderText("Nieuwe optie (bijv. training)")
        btn_add = QPushButton("Toevoegen")
        btn_remove = QPushButton("Verwijder geselecteerde")
        btn_add.clicked.connect(self._add_from_input)
        btn_remove.clicked.connect(self._remove_selected)
        add_row.addWidget(self.e_new, 1)
        add_row.addWidget(btn_add)
        add_row.addWidget(btn_remove)
        root.addLayout(add_row)

        btn_cancel = QPushButton("Annuleren")
        btn_save = QPushButton("Opslaan")
        btn_cancel.clicked.connect(self.reject)
        btn_save.clicked.connect(self._save)
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(btn_cancel)
        row.addWidget(btn_save)
        root.addLayout(row)

    def _contains_option(self, option: str) -> bool:
        key = option.casefold()
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it and it.text().strip().casefold() == key:
                return True
        return False

    def _add_row(self, option: str, active: bool):
        r = self.table.rowCount()
        self.table.insertRow(r)

        it_name = QTableWidgetItem(option)
        it_name.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
        self.table.setItem(r, 0, it_name)

        it_active = QTableWidgetItem("")
        it_active.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
        it_active.setCheckState(Qt.Checked if active else Qt.Unchecked)
        self.table.setItem(r, 1, it_active)

    def _add_from_input(self):
        val = self.e_new.text().strip()
        if not val:
            return
        if self._contains_option(val):
            QMessageBox.warning(self, "Bestaat al", f"Optie '{val}' bestaat al.")
            return
        self._add_row(val, True)
        self.e_new.clear()

    def _remove_selected(self):
        r = self.table.currentRow()
        if r < 0:
            return
        self.table.removeRow(r)

    def _save(self):
        options = []
        enabled = []
        seen = set()
        for r in range(self.table.rowCount()):
            it_name = self.table.item(r, 0)
            if not it_name:
                continue
            name = it_name.text().strip()
            if not name:
                continue
            key = name.casefold()
            if key in seen:
                QMessageBox.warning(self, "Dubbele optie", f"Optie '{name}' staat dubbel in de lijst.")
                return
            seen.add(key)
            options.append(name)
            it_active = self.table.item(r, 1)
            if it_active and it_active.checkState() == Qt.Checked:
                enabled.append(name)

        self.options = normalize_extra_info_options(options)
        self.enabled_options = normalize_extra_info_enabled(self.options, enabled)
        self.accept()


class TimerPanel(QFrame):
    """Timer met oud `Log tijd.py` gedrag (auto work/idle/call detectie)."""

    height_changed = Signal(int)

    DEFAULT_IDLE_THRESHOLD_SEC = 60

    class LASTINPUTINFO(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.UINT), ("dwTime", wintypes.DWORD)]

    def __init__(self, store: ExcelStore, on_refresh, on_toggle_full, on_drag_move, idle_threshold_sec: int = DEFAULT_IDLE_THRESHOLD_SEC):
        super().__init__()
        self.store = store
        self.on_refresh = on_refresh
        self.on_toggle_full = on_toggle_full
        self.on_drag_move = on_drag_move
        self.running = True
        self.work_seconds = 0
        self.idle_seconds = 0
        self.call_seconds = 0
        self.current_day = date.today()
        self.save_tick = 0
        self.save_interval = 5
        self.warning_alpha = 1.0
        self.warning_direction = -0.1
        self.warning_active = False
        self._drag_start = None
        self.idle_threshold_sec = self.DEFAULT_IDLE_THRESHOLD_SEC
        self.inactive_glass = False
        self.inactive_glass_opacity = 0.72
        self.min_inactive_glass_opacity = GLASS_OPACITY_MIN
        self.max_inactive_glass_opacity = GLASS_OPACITY_MAX
        self._colors_cache: dict[str, str] = {}

        self.setObjectName("timerPanel")
        self.setMinimumWidth(372)
        self.setMaximumWidth(372)
        self.setMinimumHeight(56)
        self.setMaximumHeight(56)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)
        root.setContentsMargins(4, 4, 4, 4)
        root.setSpacing(2)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(4)
        self.lbl_line = QLabel("T: 00:00:00  I: 00:00:00  C: 00:00:00")
        self.lbl_line.setStyleSheet("font: 11pt 'Consolas'; font-weight:700;")
        self.lbl_line.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        self.lbl_line.setMinimumHeight(42)
        self.lbl_line.installEventFilter(self)
        row.addWidget(self.lbl_line, 1)

        self.btn_open = QPushButton("📅")
        self.btn_open.setToolTip("Planner tonen/verbergen")
        self.btn_open.setFixedSize(44, 42)
        self.btn_open.setCheckable(True)
        self.btn_open.clicked.connect(self.on_toggle_full)
        row.addWidget(self.btn_open)
        root.addLayout(row)

        self.idle_label = QLabel("⚠ IDLE binnenkort")
        self.idle_label.setVisible(False)
        self.idle_label.setAlignment(Qt.AlignCenter)
        self.idle_label.setStyleSheet("font: 9pt 'Consolas'; font-weight:700;")
        root.addWidget(self.idle_label)

        self.timer = QTimer(self)
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self._tick)
        self.warn_timer = QTimer(self)
        self.warn_timer.setInterval(100)
        self.warn_timer.timeout.connect(self._animate_warning)

        self.set_idle_threshold(idle_threshold_sec)
        self.load_today()
        self.update_ui()
        self.timer.start()

    def eventFilter(self, obj, event):
        if obj is self.lbl_line:
            if event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                self._drag_start = event.globalPosition().toPoint()
                return True
            if event.type() == QEvent.MouseMove and self._drag_start is not None:
                delta = event.globalPosition().toPoint() - self._drag_start
                self._drag_start = event.globalPosition().toPoint()
                self.on_drag_move(delta)
                return True
            if event.type() == QEvent.MouseButtonRelease:
                self._drag_start = None
                return True
        return super().eventFilter(obj, event)

    def apply_colors(self, colors: dict[str, str]):
        self._colors_cache = dict(colors or {})
        bg = self._colors_cache.get("timer_bg_dark", "#000000")
        fg = self._colors_cache.get("timer_text_dark", "#7CFC00")
        btn = self._colors_cache.get("timer_btn_dark", "#222222")

        bg_c = QColor(bg)
        fg_c = QColor(fg)
        btn_c = QColor(btn)
        if self.inactive_glass:
            alpha = int(255 * self.inactive_glass_opacity)
            alpha = max(48, min(235, alpha))
            bg_c.setAlpha(alpha)
            fg_c = fg_c.lighter(115)
            fg_c.setAlpha(min(240, alpha + 65))
            btn_c.setAlpha(min(235, alpha + 35))

        bg_css = f"rgba({bg_c.red()}, {bg_c.green()}, {bg_c.blue()}, {bg_c.alpha()})"
        fg_css = f"rgba({fg_c.red()}, {fg_c.green()}, {fg_c.blue()}, {fg_c.alpha()})"
        btn_css = f"rgba({btn_c.red()}, {btn_c.green()}, {btn_c.blue()}, {btn_c.alpha()})"
        self.setStyleSheet(
            f"""
            QFrame#timerPanel {{
                background:qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 rgba(255,255,255,{18 if self.inactive_glass else 8}),
                    stop:1 {bg_css}
                );
                border:1px solid rgba(220,235,255,{130 if self.inactive_glass else 90});
                border-radius:10px;
            }}
            QFrame#timerPanel QLabel {{ color:{fg_css}; }}
            QFrame#timerPanel QPushButton {{
                color:#f8fbff;
                border-radius:7px;
                border-top:1px solid #9cb3cf;
                border-left:1px solid #839ab8;
                border-right:1px solid #32465f;
                border-bottom:2px solid #203248;
                background:{btn_css};
                background:qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #5e7895, stop:1 #355071);
                font: 12pt 'Segoe UI Emoji';
                padding: 0px;
            }}
            QFrame#timerPanel QPushButton:hover {{
                border-top:1px solid #b8ccee;
                background:qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #6f88a3, stop:1 #3f6287);
            }}
            QFrame#timerPanel QPushButton:checked, QFrame#timerPanel QPushButton:pressed {{
                border-top:2px solid #1e2f45;
                border-left:2px solid #22364d;
                border-right:1px solid #7f99b8;
                border-bottom:1px solid #9ab4d3;
                background:qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #2c4666, stop:1 #1f3650);
                padding-top:2px;
                padding-left:1px;
            }}
            """
        )
        self.idle_label.setStyleSheet("color:#111111; background:#ffff66; font: 10pt 'Consolas'; font-weight:700; border-radius:4px;")

    def set_inactive_glass(self, enabled: bool):
        flag = bool(enabled)
        if self.inactive_glass == flag:
            return
        self.inactive_glass = flag
        self.apply_colors(self._colors_cache)

    def set_glass_opacity(self, opacity: float):
        try:
            v = float(opacity)
        except Exception:
            v = 0.72
        v = max(self.min_inactive_glass_opacity, min(self.max_inactive_glass_opacity, v))
        if abs(v - self.inactive_glass_opacity) < 0.0001:
            return
        self.inactive_glass_opacity = v
        if self.inactive_glass:
            self.apply_colors(self._colors_cache)

    def set_glass_limits(self, min_opacity: float, max_opacity: float):
        changed = (
            abs(GLASS_OPACITY_MIN - self.min_inactive_glass_opacity) > 0.0001
            or abs(GLASS_OPACITY_MAX - self.max_inactive_glass_opacity) > 0.0001
        )
        self.min_inactive_glass_opacity = GLASS_OPACITY_MIN
        self.max_inactive_glass_opacity = GLASS_OPACITY_MAX
        if self.inactive_glass_opacity < GLASS_OPACITY_MIN or self.inactive_glass_opacity > GLASS_OPACITY_MAX:
            self.inactive_glass_opacity = max(GLASS_OPACITY_MIN, min(GLASS_OPACITY_MAX, self.inactive_glass_opacity))
            changed = True
        if changed and self.inactive_glass:
            self.apply_colors(self._colors_cache)

    def get_idle_time(self) -> float:
        if sys.platform != "win32":
            return 0.0
        lii = self.LASTINPUTINFO()
        lii.cbSize = ctypes.sizeof(self.LASTINPUTINFO)
        if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)):
            ms = ctypes.windll.kernel32.GetTickCount() - lii.dwTime
            return ms / 1000.0
        return 0.0

    def detect_call(self) -> bool:
        if psutil is None or gw is None:
            return False
        try:
            for proc in psutil.process_iter(["name"]):
                name = (proc.info.get("name") or "").lower()
                if any(x in name for x in ("teams", "chrome", "edge", "zoom", "slack")):
                    windows = (
                        gw.getWindowsWithTitle("Meeting")
                        + gw.getWindowsWithTitle("Call")
                        + gw.getWindowsWithTitle("Vergadering")
                    )
                    if windows:
                        return True
        except Exception:
            return False
        return False

    def load_today(self):
        self.current_day = date.today()
        row = self.store.get_timer_log(self.current_day)
        self.work_seconds = row["work"]
        self.idle_seconds = row["idle"]
        self.call_seconds = row["call"]

    def pause(self):
        self.running = False
        self.timer.stop()
        self.warn_timer.stop()
        self.persist()

    def start(self):
        if self.running:
            return
        self.running = True
        self.timer.start()

    def set_idle_threshold(self, seconds: int):
        try:
            sec = int(seconds)
        except Exception:
            sec = self.DEFAULT_IDLE_THRESHOLD_SEC
        self.idle_threshold_sec = max(15, min(3600, sec))

    def _show_idle_warning(self):
        if self.warning_active:
            return
        self.warning_active = True
        self.idle_label.setVisible(True)
        self.setMinimumHeight(76)
        self.setMaximumHeight(76)
        self.height_changed.emit(76)
        self.warn_timer.start()

    def _hide_idle_warning(self):
        if not self.warning_active:
            return
        self.warning_active = False
        self.warn_timer.stop()
        self.idle_label.setVisible(False)
        self.setMinimumHeight(56)
        self.setMaximumHeight(56)
        self.height_changed.emit(56)

    def _animate_warning(self):
        self.warning_alpha += self.warning_direction
        if self.warning_alpha <= 0.35 or self.warning_alpha >= 1.0:
            self.warning_direction *= -1
        intensity = int(200 + (55 * self.warning_alpha))
        self.idle_label.setStyleSheet(
            f"color:#111111; background:rgb(255,255,{intensity}); font: 10pt 'Consolas'; font-weight:700; border-radius:4px;"
        )

    def _tick(self):
        if not self.running:
            return
        if date.today() != self.current_day:
            self.load_today()

        idle_sec = self.get_idle_time()
        warning_sec = max(5, min(30, self.idle_threshold_sec // 6))
        if self.idle_threshold_sec - warning_sec <= idle_sec < self.idle_threshold_sec:
            self._show_idle_warning()
        else:
            self._hide_idle_warning()

        if idle_sec >= self.idle_threshold_sec:
            self.idle_seconds += 1
        else:
            self.work_seconds += 1

        if self.detect_call():
            self.call_seconds += 1

        self.save_tick += 1
        if self.save_tick >= self.save_interval:
            self.persist()
            self.save_tick = 0
        self.update_ui()

    def persist(self):
        self.store.save_timer_log(self.current_day, self.work_seconds, self.idle_seconds, self.call_seconds)
        self.on_refresh()

    def update_ui(self):
        self.lbl_line.setText(
            f"T: {seconds_to_hhmmss(self.work_seconds)}  I: {seconds_to_hhmmss(self.idle_seconds)}  C: {seconds_to_hhmmss(self.call_seconds)}"
        )


class MainWindow(QMainWindow):
    """Hoofdvenster.

    Bevat:
    - kalenderweergaven (jaar + maand),
    - timerdrawer,
    - instellingen/menu.
    """
    def __init__(self):
        super().__init__()
        self.year = datetime.now().year
        base_dir = os.path.dirname(os.path.abspath(__file__))
        self.base_dir = base_dir
        self.store = ExcelStore(self.year, base_dir)
        self.mode = "worked"
        self.dark_mode = True
        (
            self.colors,
            self.extra_info_colors,
            self.idle_threshold_sec,
            self.extra_info_enabled,
            self.school_region,
            self.inactive_glass_opacity,
            self.min_inactive_glass_opacity,
            self.max_inactive_glass_opacity,
        ) = self.load_color_settings()
        self.min_inactive_glass_opacity = GLASS_OPACITY_MIN
        self.max_inactive_glass_opacity = GLASS_OPACITY_MAX
        self.inactive_glass_opacity = max(GLASS_OPACITY_MIN, min(GLASS_OPACITY_MAX, self.inactive_glass_opacity))
        self.ensure_extra_info_colors()
        self.ensure_extra_info_enabled()
        self.school_region = (self.school_region or "zuid").strip().casefold()
        if self.school_region not in {"noord", "midden", "zuid"}:
            self.school_region = "zuid"
        self.store.school_region = self.school_region
        self.timer_host = None
        self._drag_active = False
        self._drag_offset = QPoint()
        self._allow_close = False
        self._slide_anim = None
        self._sliding = False
        self._needs_year_refresh = False
        self._needs_month_refresh = False
        self._planner_target_geometry = QRect(120, 80, 2140, 1050)
        self.setWindowTitle(f"Tijdplanner Pro - {self.year}")
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        # Start met drie maanden waar mogelijk, maar schaal veilig terug op kleinere schermen.
        self._planner_target_geometry = self._fit_rect_to_screen(self._planner_target_geometry)
        self.resize(self._planner_target_geometry.width(), self._planner_target_geometry.height())
        self.move(self._planner_target_geometry.topLeft())
        self.apply_theme()
        self.setup_ui()
        self.menuBar().installEventFilter(self)
        self.refresh_all()

    def _fit_rect_to_screen(self, rect: QRect, screen=None, margin: int = 18) -> QRect:
        if screen is None:
            screen = QApplication.screenAt(rect.center())
            if screen is None:
                screen = QApplication.primaryScreen()
        avail = screen.availableGeometry() if screen else QRect(0, 0, 2560, 1440)

        max_w = max(360, avail.width() - (margin * 2))
        max_h = max(260, avail.height() - (margin * 2))

        w = min(rect.width(), max_w)
        h = min(rect.height(), max_h)

        x_min = avail.left() + margin
        y_min = avail.top() + margin
        x_max = avail.right() - margin - w + 1
        y_max = avail.bottom() - margin - h + 1

        x = max(x_min, min(rect.x(), x_max))
        y = max(y_min, min(rect.y(), y_max))
        return QRect(x, y, w, h)

    def apply_theme(self):
        self.setStyleSheet(
            """
            QMainWindow { background: #1c212b; }
            QMainWindow[modeTheme="worked"] { background: #1b2620; }
            QMainWindow[modeTheme="planned"] { background: #1c212b; }
            QMenuBar { background: #121720; color: #e8edf7; padding: 5px; font: 10pt "Segoe UI"; }
            QMenuBar::item:selected { background: #355071; color:#ffffff; border-radius: 4px; }
            QMenu { background: #1d2532; color: #e8edf7; border: 1px solid #344156; }
            QMenu::item { padding: 6px 20px; border-radius: 4px; }
            QMenu::item:selected { background: #355071; color: #ffffff; }
            QDialog, QMessageBox { background: #1c2430; color: #e6edf3; }
            QToolTip { color: #e6edf3; background-color: #10151d; border: 1px solid #3b4a60; padding: 6px; }
            QToolBar { background: #232c3a; border: 0; spacing: 8px; padding: 6px; }
            QPushButton { background: #3f6ea6; color: #fff; border: 0; border-radius: 8px; padding: 8px 14px; font: 10pt "Segoe UI Semibold"; }
            QPushButton:hover { background: #4e7fba; }
            QPushButton:pressed { background: #325987; }
            QLabel#modeLeft, QLabel#modeRight { color:#9fb0c5; font: 10pt "Segoe UI Semibold"; padding: 0 4px; }
            QLabel#modeLeft[active="true"] { color:#9be3be; }
            QLabel#modeRight[active="true"] { color:#9ecbff; }
            QCheckBox#modeSwitch { spacing:0px; }
            QCheckBox#modeSwitch::indicator {
                width:56px; height:28px; border-radius:14px;
                border:1px solid #32465f;
                background:qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #3f7e5c, stop:1 #2a5a41);
            }
            QCheckBox#modeSwitch::indicator:unchecked {
                border-top:1px solid #88c6a7;
                border-left:1px solid #79b796;
                border-right:1px solid #2f4f44;
                border-bottom:2px solid #1f342b;
            }
            QCheckBox#modeSwitch::indicator:checked {
                border-top:1px solid #a8c6e7;
                border-left:1px solid #9bb8d7;
                border-right:1px solid #324a62;
                border-bottom:2px solid #213344;
                background:qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4d6f96, stop:1 #2f4f75);
            }
            QFrame#timerPanel QPushButton[active="true"] { background:#2f6f4b; border:1px solid #47946a; }
            QTabWidget::pane { border: 1px solid #3a4658; background: #222b38; border-radius: 8px; }
            QTabBar::tab { background: #2a3443; color: #d6deea; padding: 10px 18px; margin-right: 4px; border-top-left-radius: 8px; border-top-right-radius: 8px; font: 10pt "Segoe UI Semibold"; }
            QTabBar::tab:selected { background: #222b38; }
            QTabBar#monthTabBar::tab { background: #283240; color: #d6deea; margin-right: 3px; border: 1px solid #3a4658; border-bottom: 0; }
            QTabBar#monthTabBar::tab:selected { background: #355071; color: #f3f8ff; }
            QScrollArea#calendarScroll { background: #222b38; border: 0; }
            QScrollArea#calendarScroll[modeTheme="worked"] { background: #223129; }
            QScrollArea#calendarScroll[modeTheme="planned"] { background: #222b38; }
            QWidget#calendarHost { background: #222b38; }
            QWidget#calendarHost[modeTheme="worked"] { background: #223129; }
            QWidget#calendarHost[modeTheme="planned"] { background: #222b38; }
            QGroupBox { border: 1px solid #3b4759; border-radius: 10px; margin-top: 18px; padding-top: 12px; background: #263141; font: 10pt "Segoe UI Semibold"; color: #e0e7f2; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; top: -2px; padding: 2px 8px; }
            QTableWidget { background: #1f2835; alternate-background-color: #222e3d; border: 1px solid #3a4658; gridline-color: #3a4658; color: #e6edf7; font: 9pt "Segoe UI"; }
            QHeaderView::section { background: #2c3849; color: #dce4f0; padding: 6px; border: 1px solid #425066; font: 9pt "Segoe UI Semibold"; }
            QStatusBar { background: #232c3a; color: #d8e1ef; border-top: 1px solid #3a4658; font: 9pt "Segoe UI"; }
            QLabel { color: #dce4f0; }
            QLineEdit, QComboBox { border: 1px solid #45556d; border-radius: 6px; background: #1f2835; color: #eaf1fb; padding: 6px; font: 9pt "Segoe UI"; }
            QComboBox::drop-down { border: 0; width: 24px; }
            QComboBox::down-arrow { width: 10px; height: 10px; }
            QComboBox QAbstractItemView { background: #1f2835; color: #eaf1fb; border: 1px solid #45556d; selection-background-color: #355071; selection-color: #ffffff; outline: 0; }
            QScrollBar:vertical { background: #1a2230; width: 12px; margin: 2px; }
            QScrollBar::handle:vertical { background: #40516b; border-radius: 5px; min-height: 28px; }
            QScrollBar:horizontal { background: #1a2230; height: 12px; margin: 2px; }
            QScrollBar::handle:horizontal { background: #40516b; border-radius: 5px; min-width: 28px; }
            """
        )

    def setup_ui(self):
        self.build_menu()
        self.add_window_controls()
        self.build_toolbar()
        self.build_statusbar()

        root = QWidget()
        self.setCentralWidget(root)
        content_lay = QVBoxLayout(root)
        content_lay.setContentsMargins(6, 6, 6, 6)

        self.tabs = QTabWidget()
        self.tabs.tabBar().hide()
        content_lay.addWidget(self.tabs)
        self.tab_year = QWidget()
        self.tab_month = QWidget()
        self.tabs.addTab(self.tab_year, "Jaarweergave")
        self.tabs.addTab(self.tab_month, "Maandweergave")

        year_layout = QVBoxLayout(self.tab_year)
        self.year_board = CalendarBoard(
            self.store,
            self.year,
            self.mode,
            list(range(1, 13)),
            3,
            focus_mode=False,
            dark_mode=self.dark_mode,
            colors=self.colors,
            extra_info_colors=self.extra_info_colors,
        )
        self.year_board.day_double_clicked.connect(self.edit_day)
        year_layout.addWidget(self.year_board)

        month_layout = QVBoxLayout(self.tab_month)
        self.month_tabbar = QTabBar()
        self.month_tabbar.setObjectName("monthTabBar")
        self.month_tabbar.setDrawBase(False)
        self.month_tabbar.setExpanding(False)
        self.month_tabbar.setUsesScrollButtons(True)
        self.month_tabbar.setElideMode(Qt.ElideNone)
        for m in MONTHS_NL:
            self.month_tabbar.addTab(m)
        self.month_tabbar.setCurrentIndex(date.today().month - 1)
        self.month_tabbar.setMaximumHeight(36)
        self.month_tabbar.installEventFilter(self)
        self.month_tabbar.currentChanged.connect(self.on_month_tab_changed)
        month_layout.addWidget(self.month_tabbar)
        self.month_board = CalendarBoard(
            self.store,
            self.year,
            self.mode,
            [self.month_tabbar.currentIndex() + 1],
            1,
            focus_mode=True,
            dark_mode=self.dark_mode,
            colors=self.colors,
            extra_info_colors=self.extra_info_colors,
        )
        self.month_board.installEventFilter(self)
        self.month_board.scroll.installEventFilter(self)
        self.month_board.host.installEventFilter(self)
        self.month_board.day_double_clicked.connect(self.edit_day)
        self.month_board.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        month_layout.addWidget(self.month_board, 1)
        self.tabs.currentChanged.connect(self.on_tab_change)

    def slide_in_from(self, timer_geo: QRect):
        target = QRect(self._planner_target_geometry)
        if target.width() < 600 or target.height() < 400:
            target = QRect(120, 80, 2140, 1050)
        target = self._fit_rect_to_screen(target, QApplication.screenAt(timer_geo.center()))

        start = QRect(timer_geo.x(), timer_geo.y(), target.width(), target.height())
        self.setGeometry(start)
        self.setWindowOpacity(0.0)
        self.show()

        geo_anim = QPropertyAnimation(self, b"geometry")
        geo_anim.setDuration(320)
        geo_anim.setStartValue(start)
        geo_anim.setEndValue(target)
        geo_anim.setEasingCurve(QEasingCurve.OutCubic)

        opacity_anim = QPropertyAnimation(self, b"windowOpacity")
        opacity_anim.setDuration(220)
        opacity_anim.setStartValue(0.0)
        opacity_anim.setEndValue(1.0)
        opacity_anim.setEasingCurve(QEasingCurve.OutQuad)

        group = QParallelAnimationGroup(self)
        group.addAnimation(geo_anim)
        group.addAnimation(opacity_anim)
        self._slide_anim = group
        self._sliding = True

        def done():
            self._sliding = False
            self.setGeometry(target)
            self.setWindowOpacity(1.0)
            self._planner_target_geometry = QRect(target)

        group.finished.connect(done)
        group.start()

    def slide_out_to(self, timer_geo: QRect):
        start = self.geometry()
        if start.width() > 300 and start.height() > 200:
            self._planner_target_geometry = QRect(start)

        end = QRect(timer_geo.x(), timer_geo.y(), start.width(), start.height())

        geo_anim = QPropertyAnimation(self, b"geometry")
        geo_anim.setDuration(280)
        geo_anim.setStartValue(start)
        geo_anim.setEndValue(end)
        geo_anim.setEasingCurve(QEasingCurve.InOutCubic)

        opacity_anim = QPropertyAnimation(self, b"windowOpacity")
        opacity_anim.setDuration(210)
        opacity_anim.setStartValue(1.0)
        opacity_anim.setEndValue(0.0)
        opacity_anim.setEasingCurve(QEasingCurve.InQuad)

        group = QParallelAnimationGroup(self)
        group.addAnimation(geo_anim)
        group.addAnimation(opacity_anim)
        self._slide_anim = group
        self._sliding = True

        def done():
            self._sliding = False
            self.hide()
            self._planner_target_geometry = self._fit_rect_to_screen(self._planner_target_geometry)
            self.setGeometry(self._planner_target_geometry)
            self.setWindowOpacity(1.0)

        group.finished.connect(done)
        group.start()

    def build_menu(self):
        m = self.menuBar()
        menu_file = m.addMenu("Bestand")
        act_save = QAction("Opslaan", self)
        act_export = QAction("Exporteer kopie", self)
        act_exit = QAction("Afsluiten", self)
        act_save.triggered.connect(self.save)
        act_export.triggered.connect(self.export_copy)
        act_exit.triggered.connect(self.quit_app)
        menu_file.addAction(act_save)
        menu_file.addAction(act_export)
        menu_file.addSeparator()
        menu_file.addAction(act_exit)

        menu_view = m.addMenu("Beeld")
        menu_view.addAction("Jaarweergave", lambda: self.tabs.setCurrentIndex(0))
        menu_view.addAction("Maandweergave", lambda: self.tabs.setCurrentIndex(1))
        menu_view.addSeparator()
        menu_view.addAction("Ga naar huidige maand", self.goto_today)

        menu_plan = m.addMenu("Planning")
        menu_plan.addAction("Werkpatroon bewerken", self.open_work_pattern_editor)
        menu_settings = m.addMenu("Instellingen")
        menu_settings.addAction("Kleuren", self.open_color_settings)
        menu_settings.addAction("Aanvullende info opties", self.open_extra_info_settings)
        menu_settings.addAction("Schoolvakantie regio", self.open_school_region_settings)
        menu_settings.addAction("Timer idle detectie", self.open_timer_settings)
        menu_settings.addAction("Glassmorphism", self.open_glass_settings)
        menu_help = m.addMenu("Help")
        menu_help.addAction("Over", self.about)

    def add_window_controls(self):
        host = QWidget()
        row = QHBoxLayout(host)
        row.setContentsMargins(0, 0, 4, 0)
        row.setSpacing(4)

        btn_min = QPushButton("−")
        btn_close = QPushButton("×")
        for b in (btn_min, btn_close):
            b.setFixedSize(28, 24)
            b.setStyleSheet(
                "QPushButton { background:#2a3443; border:1px solid #3a4658; border-radius:4px; color:#dce4f0; padding:0px; }"
                "QPushButton:hover { background:#3a475b; }"
            )
        btn_close.setStyleSheet(
            "QPushButton { background:#4a2a30; border:1px solid #6a3d44; border-radius:4px; color:#ffdce0; padding:0px; }"
            "QPushButton:hover { background:#6a303a; }"
        )

        btn_min.clicked.connect(self.showMinimized)
        btn_close.clicked.connect(self.close)

        row.addWidget(btn_min)
        row.addWidget(btn_close)
        self.menuBar().setCornerWidget(host, Qt.TopRightCorner)

    def build_toolbar(self):
        tb = QToolBar("Main")
        tb.setMovable(False)
        self.addToolBar(Qt.TopToolBarArea, tb)

        mode_host = QWidget()
        mode_row = QHBoxLayout(mode_host)
        mode_row.setContentsMargins(0, 0, 0, 0)
        mode_row.setSpacing(6)
        self.lbl_mode_left = QLabel("Uren")
        self.lbl_mode_left.setObjectName("modeLeft")
        self.lbl_mode_right = QLabel("Planning")
        self.lbl_mode_right.setObjectName("modeRight")
        self.mode_switch = QCheckBox()
        self.mode_switch.setObjectName("modeSwitch")
        self.mode_switch.setTristate(False)
        self.mode_switch.toggled.connect(lambda checked: self.set_mode("planned" if checked else "worked"))
        mode_row.addWidget(self.lbl_mode_left)
        mode_row.addWidget(self.mode_switch)
        mode_row.addWidget(self.lbl_mode_right)
        tb.addWidget(mode_host)

        self.btn_pattern = QPushButton("Werkpatroon toepassen")
        self.btn_pattern.clicked.connect(self.open_work_pattern_editor)
        self.btn_pattern.setVisible(False)
        tb.addWidget(self.btn_pattern)
        self.btn_pattern_edit = QPushButton("Werkpatroon")
        self.btn_pattern_edit.clicked.connect(self.open_work_pattern_editor)
        self.btn_pattern_edit.setVisible(False)
        tb.addWidget(self.btn_pattern_edit)
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        tb.addWidget(spacer)
        self._refresh_mode_toggle_style()

    def build_statusbar(self):
        sb = QStatusBar()
        self.setStatusBar(sb)
        self.lbl_status = QLabel("Gereed")
        self.lbl_saldo = QLabel("")
        sb.addPermanentWidget(self.lbl_status, 1)
        sb.addPermanentWidget(self.lbl_saldo)

    def refresh_all(self):
        self.year_board.set_dark_mode(self.dark_mode)
        self.month_board.set_dark_mode(self.dark_mode)
        self.year_board.set_colors(self.colors)
        self.month_board.set_colors(self.colors)
        self.year_board.set_extra_info_colors(self.extra_info_colors)
        self.month_board.set_extra_info_colors(self.extra_info_colors)
        self.year_board.set_mode(self.mode)
        self.month_board.set_mode(self.mode)
        self.year_board.refresh()
        self.month_board.refresh()
        self._needs_year_refresh = False
        self._needs_month_refresh = False
        if self.timer_host:
            self.timer_host.panel.apply_colors(self.colors)
        self.lbl_saldo.setText(self.store.get_saldo_text())
        self._refresh_mode_toggle_style()
        self._apply_mode_theme()

    def _refresh_mode_toggle_style(self):
        checked = self.mode == "planned"
        if self.mode_switch.isChecked() != checked:
            self.mode_switch.blockSignals(True)
            self.mode_switch.setChecked(checked)
            self.mode_switch.blockSignals(False)
        self.lbl_mode_left.setProperty("active", "true" if self.mode == "worked" else "false")
        self.lbl_mode_right.setProperty("active", "true" if self.mode == "planned" else "false")
        for w in (self.mode_switch, self.lbl_mode_left, self.lbl_mode_right):
            w.style().unpolish(w)
            w.style().polish(w)
            w.update()

    def _apply_mode_theme(self):
        mode_theme = "worked" if self.mode == "worked" else "planned"
        self.setProperty("modeTheme", mode_theme)
        for w in (
            self,
            self.year_board.scroll,
            self.year_board.host,
            self.month_board.scroll,
            self.month_board.host,
        ):
            w.setProperty("modeTheme", mode_theme)
            w.style().unpolish(w)
            w.style().polish(w)
            w.update()

    def set_mode(self, mode: str):
        if mode not in ("worked", "planned"):
            return
        if self.mode == mode:
            return
        self.mode = mode
        is_planned = self.mode == "planned"
        self.btn_pattern.setVisible(is_planned)
        self.btn_pattern_edit.setVisible(is_planned)
        self.year_board.set_mode(self.mode)
        self.month_board.set_mode(self.mode)
        self._needs_year_refresh = True
        self._needs_month_refresh = True
        if self.tabs.currentIndex() == 0:
            self.year_board.refresh()
            self._needs_year_refresh = False
        else:
            self.month_board.refresh()
            self._needs_month_refresh = False
        self._refresh_mode_toggle_style()
        self._apply_mode_theme()
        self.lbl_status.setText("Actieve modus: Planning" if is_planned else "Actieve modus: Uren")

    def toggle_mode(self):
        self.set_mode("planned" if self.mode == "worked" else "worked")

    def on_month_tab_changed(self, idx: int):
        self.month_board.set_month(idx + 1)
        self.month_board.set_mode(self.mode)
        self.month_board.refresh()
        self.lbl_status.setText(f"Maand: {MONTHS_NL[idx]}")

    def on_tab_change(self, idx: int):
        if idx == 0 and self._needs_year_refresh:
            self.year_board.refresh()
            self._needs_year_refresh = False
        elif idx == 1 and self._needs_month_refresh:
            self.month_board.refresh()
            self._needs_month_refresh = False
        self.lbl_status.setText("Tab: Jaarweergave" if idx == 0 else "Tab: Maandweergave")

    def edit_day(self, dt: date):
        if self.mode != "planned":
            return
        existing_reason = self.store.day_reason(dt)
        dlg = DayEditDialog(
            self,
            dt,
            self.store.get_day(dt),
            self.store.get_day_limit(dt),
            existing_reason,
            self.store.get_extra_info(dt),
            self.extra_info_enabled,
            self.extra_info_colors,
        )
        if dlg.exec() != QDialog.Accepted or not dlg.data_out:
            return
        self.store.set_day(dt, dlg.data_out, dlg.reason_out, dlg.extra_info_out)
        self.refresh_all()
        self.lbl_status.setText(f"Dag opgeslagen: {dt.strftime('%d-%m-%Y')}")

    def open_work_pattern_editor(self):
        dlg = WorkPatternDialog(self, self.store, self.year, apply_callback=self.apply_pattern_year)
        if dlg.exec() == QDialog.Accepted:
            self.store.load_patterns()
            self.refresh_all()
            self.lbl_status.setText("Werkpatroon opgeslagen")

    def open_extra_info_settings(self):
        dlg = ExtraInfoSettingsDialog(self, self.store.extra_info_options, self.extra_info_enabled)
        if dlg.exec() != QDialog.Accepted:
            return
        self.store.save_extra_info_options(dlg.options, dlg.enabled_options)
        self.extra_info_enabled = list(dlg.enabled_options)
        self.ensure_extra_info_enabled()
        self.ensure_extra_info_colors()
        self.save_color_settings()
        self.refresh_all()
        self.lbl_status.setText("Aanvullende info opties opgeslagen")

    def open_school_region_settings(self):
        dlg = SchoolRegionDialog(self, self.school_region)
        if dlg.exec() != QDialog.Accepted:
            return
        self.school_region = dlg.region
        self.store.school_region = self.school_region
        self.store.load_school_holidays()
        if not self.store.school_vakanties:
            added = self.store.seed_school_holidays_from_api(region_filter=None)
            if added:
                self.store.wb.save(self.store.path)
                self.store.load_school_holidays()
        self.save_color_settings()
        self.refresh_all()
        self.lbl_status.setText(f"Schoolvakantie regio: {self.school_region}")

    def open_timer_settings(self):
        dlg = TimerSettingsDialog(self, self.idle_threshold_sec)
        if dlg.exec() != QDialog.Accepted:
            return
        self.idle_threshold_sec = dlg.idle_threshold_sec
        if self.timer_host and hasattr(self.timer_host, "panel"):
            self.timer_host.panel.set_idle_threshold(self.idle_threshold_sec)
        self.save_color_settings()
        self.lbl_status.setText(f"Idle detectie: {self.idle_threshold_sec // 60:02d}:{self.idle_threshold_sec % 60:02d}")

    def open_glass_settings(self):
        dlg = GlassSettingsDialog(self, self.inactive_glass_opacity)
        if dlg.exec() != QDialog.Accepted:
            return
        self.inactive_glass_opacity = dlg.inactive_glass_opacity
        if self.timer_host and hasattr(self.timer_host, "panel"):
            self.timer_host.panel.set_glass_limits(self.min_inactive_glass_opacity, self.max_inactive_glass_opacity)
            self.timer_host.panel.set_glass_opacity(self.inactive_glass_opacity)
            self.timer_host.update_focus_glass()
        self.save_color_settings()
        self.lbl_status.setText(
            f"Glassmorphism: {int(round(self.inactive_glass_opacity * 100)):02d}% (vast min {int(GLASS_OPACITY_MIN * 100)}% / max {int(GLASS_OPACITY_MAX * 100)}%)"
        )

    def save(self):
        self.store.wb.save(self.store.path)
        self.lbl_status.setText("Opgeslagen")
        self.lbl_saldo.setText(self.store.get_saldo_text())

    def export_copy(self):
        dst = self.store.export_copy()
        QMessageBox.information(self, "Export", f"Export gemaakt:\n{dst}")
        self.lbl_status.setText("Export gereed")

    def goto_today(self):
        t = date.today()
        if t.year != self.year:
            return
        self.tabs.setCurrentIndex(1)
        self.month_tabbar.setCurrentIndex(t.month - 1)

    def apply_pattern_year(self, start_date: date | None = None, end_date: date | None = None):
        ws = self.store.wb["Planning"]
        rows_by_date = {}
        updated_dates = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            d = normalize_to_date(row[0] if row else None)
            if d:
                rows_by_date[d] = normalize_hhmm((row[1] if len(row) > 1 else "00:00") or "00:00")

        for m in range(1, 13):
            days = calendar.monthrange(self.year, m)[1]
            for d in range(1, days + 1):
                dt = date(self.year, m, d)
                if start_date and dt < start_date:
                    continue
                if end_date and dt > end_date:
                    continue
                if dt in self.store.nl_holidays:
                    continue
                pat = self.store.weekday_pattern[dt.weekday()]
                w = normalize_hhmm(pat["W"])
                if hhmm_to_minutes(w) > self.store.get_day_limit(dt):
                    continue
                rows_by_date[dt] = w
                updated_dates.add(dt)
                key = dt.strftime("%Y-%m-%d")
                self.store.planned_data[key] = {"W": w, "V": "00:00", "Z": "00:00"}
                self.store.vakantie_dagen.pop(key, None)

        sorted_rows = sorted(rows_by_date.items(), key=lambda x: x[0])
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        for i, (dt, vals) in enumerate(sorted_rows, start=2):
            ws.cell(i, 1, dt).number_format = "DD-MM-YYYY"
            ws.cell(i, 2, vals)

        ws_vrij = self.store.wb["vrije dagen"]
        rows_to_delete = []
        for row in ws_vrij.iter_rows(min_row=2):
            r_dt = normalize_to_date(row[0].value)
            if r_dt in updated_dates:
                rows_to_delete.append(row[0].row)
        for r in sorted(rows_to_delete, reverse=True):
            ws_vrij.delete_rows(r, 1)

        self.store.load_free_days()
        self.store.wb.save(self.store.path)
        self.refresh_all()
        self.lbl_status.setText("Werkpatroon toegepast")
        if start_date and end_date:
            QMessageBox.information(self, "Klaar", f"Werkpatroon toegepast van {start_date.strftime('%d-%m-%Y')} t/m {end_date.strftime('%d-%m-%Y')}.")
        else:
            QMessageBox.information(self, "Klaar", "Werkpatroon toegepast op het jaar.")

    def colors_path(self) -> str:
        return os.path.join(self.base_dir, "calendar_colors.json")

    def ensure_extra_info_enabled(self):
        self.extra_info_enabled = normalize_extra_info_enabled(self.store.extra_info_options, self.extra_info_enabled)

    def ensure_extra_info_colors(self):
        normalized = {}
        for k, v in self.extra_info_colors.items():
            if isinstance(k, str) and isinstance(v, str) and v.startswith("#"):
                normalized[k.strip()] = v
        self.extra_info_colors = normalized
        for opt in self.store.extra_info_options:
            if opt not in self.extra_info_colors:
                self.extra_info_colors[opt] = default_extra_info_color(opt)

    def load_color_settings(self) -> tuple[dict[str, str], dict[str, str], int, list[str], str, float, float, float]:
        cfg = dict(COLOR_DEFAULTS)
        extra_cfg: dict[str, str] = {}
        idle_threshold_sec = 60
        extra_enabled: list[str] = []
        school_region = "zuid"
        inactive_glass_opacity = 0.72
        min_inactive_glass_opacity = GLASS_OPACITY_MIN
        max_inactive_glass_opacity = GLASS_OPACITY_MAX
        p = self.colors_path()
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    source_colors = data.get("colors") if isinstance(data.get("colors"), dict) else data
                    for k, v in source_colors.items():
                        if k in cfg and isinstance(v, str) and v.startswith("#"):
                            cfg[k] = v
                    source_extra = data.get("extra_info_colors") if isinstance(data.get("extra_info_colors"), dict) else {}
                    for k, v in source_extra.items():
                        if isinstance(k, str) and isinstance(v, str) and v.startswith("#"):
                            extra_cfg[k.strip()] = v
                    if isinstance(data.get("idle_threshold_sec"), int):
                        idle_threshold_sec = max(15, min(3600, int(data.get("idle_threshold_sec"))))
                    src_enabled = data.get("extra_info_enabled")
                    if isinstance(src_enabled, list):
                        extra_enabled = [str(x).strip() for x in src_enabled if str(x).strip()]
                    sr = str(data.get("school_region", "zuid")).strip().casefold()
                    if sr in {"noord", "midden", "zuid"}:
                        school_region = sr
                    if isinstance(data.get("inactive_glass_opacity"), (int, float)):
                        inactive_glass_opacity = max(GLASS_OPACITY_MIN, min(GLASS_OPACITY_MAX, float(data.get("inactive_glass_opacity"))))
            except:
                pass
        return (
            cfg,
            extra_cfg,
            idle_threshold_sec,
            extra_enabled,
            school_region,
            inactive_glass_opacity,
            min_inactive_glass_opacity,
            max_inactive_glass_opacity,
        )

    def save_color_settings(self):
        try:
            with open(self.colors_path(), "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "colors": self.colors,
                        "extra_info_colors": self.extra_info_colors,
                        "idle_threshold_sec": int(self.idle_threshold_sec),
                        "extra_info_enabled": list(self.extra_info_enabled),
                        "school_region": self.school_region,
                        "inactive_glass_opacity": float(self.inactive_glass_opacity),
                    },
                    f,
                    indent=2,
                )
        except Exception as exc:
            QMessageBox.warning(self, "Fout", f"Kleuren opslaan mislukt:\n{exc}")

    def open_color_settings(self):
        self.ensure_extra_info_colors()
        dlg = ColorSettingsDialog(
            self,
            self.colors,
            self.store.extra_info_options,
            self.extra_info_colors,
        )
        if dlg.exec() == QDialog.Accepted:
            self.colors = dict(dlg.colors)
            self.extra_info_colors = dict(dlg.extra_info_colors)
            self.ensure_extra_info_colors()
            self.save_color_settings()
            self.refresh_all()
            self.lbl_status.setText("Kleuren bijgewerkt")

    def reset_color_settings(self):
        self.colors = dict(COLOR_DEFAULTS)
        self.extra_info_colors = {}
        self.ensure_extra_info_colors()
        self.save_color_settings()
        self.refresh_all()
        self.lbl_status.setText("Kleuren hersteld")

    def about(self):
        QMessageBox.information(self, "Over", "Tijdplanner Pro (PySide6)\nLuxe desktop UI met tabs, menu, toolbar en snelle tabellen.")

    def toggle_maximize_restore(self):
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()

    def eventFilter(self, obj, event):
        if hasattr(self, "month_tabbar") and obj in {
            self.month_tabbar,
            getattr(self, "month_board", None),
            getattr(getattr(self, "month_board", None), "scroll", None),
            getattr(getattr(self, "month_board", None), "host", None),
        } and event.type() == QEvent.Wheel:
            delta = event.angleDelta().y()
            idx = self.month_tabbar.currentIndex()
            if delta < 0 and idx < self.month_tabbar.count() - 1:
                self.month_tabbar.setCurrentIndex(idx + 1)
            elif delta > 0 and idx > 0:
                self.month_tabbar.setCurrentIndex(idx - 1)
            return True

        if obj == self.menuBar():
            menu_popup_open = any(m.isVisible() for m in self.menuBar().findChildren(QMenu))
            if event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                # Alleen slepen vanaf lege menubalkruimte; menu-items blijven klikbaar.
                if (
                    not menu_popup_open
                    and self.menuBar().activeAction() is None
                    and self.menuBar().actionAt(event.position().toPoint()) is None
                    and not self.isMaximized()
                    and not self.isFullScreen()
                ):
                    self._drag_active = True
                    self._drag_offset = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
                    return True
            elif event.type() == QEvent.MouseMove and self._drag_active:
                if menu_popup_open:
                    self._drag_active = False
                elif not self.isMaximized() and not self.isFullScreen():
                    self.move(event.globalPosition().toPoint() - self._drag_offset)
                    return True
            elif event.type() == QEvent.MouseButtonRelease and self._drag_active:
                self._drag_active = False
                return True
        return super().eventFilter(obj, event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.isVisible() and not self._sliding and self.width() > 300 and self.height() > 200:
            self._planner_target_geometry = QRect(self.geometry())

    def moveEvent(self, event):
        super().moveEvent(event)
        if self.isVisible() and not self._sliding and self.width() > 300 and self.height() > 200:
            self._planner_target_geometry = QRect(self.geometry())

    def closeEvent(self, event):
        if self._allow_close:
            event.accept()
            return
        self.hide()
        event.ignore()

    def quit_app(self):
        if self.timer_host:
            self.timer_host.quit_all()
            return
        self._allow_close = True
        self.close()


class TimerWindow(QWidget):
    """Los timer-venster dat de planner in/uit laat schuiven."""

    def __init__(self, planner: MainWindow):
        super().__init__()
        self.planner = planner
        self._force_quit = False
        self.setObjectName("timerWindow")
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setWindowTitle("Tijdplanner Timer")
        self.setStyleSheet("QWidget#timerWindow { background: transparent; }")

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        self.panel = TimerPanel(
            self.planner.store,
            on_refresh=self.planner.refresh_all,
            on_toggle_full=self.toggle_planner,
            on_drag_move=self.move_window_by,
            idle_threshold_sec=self.planner.idle_threshold_sec,
        )
        self.panel.set_glass_limits(self.planner.min_inactive_glass_opacity, self.planner.max_inactive_glass_opacity)
        self.panel.set_glass_opacity(self.planner.inactive_glass_opacity)
        self.panel.apply_colors(self.planner.colors)
        self.panel.height_changed.connect(self.on_panel_height_changed)
        self.panel.timer.timeout.connect(self.update_tray_visual)
        root.addWidget(self.panel)
        self.on_panel_height_changed(self.panel.maximumHeight())
        self.setup_tray()
        app = QApplication.instance()
        if app is not None and hasattr(app, "applicationStateChanged"):
            app.applicationStateChanged.connect(lambda _state: self.update_focus_glass())
        self.update_focus_glass()

    def on_panel_height_changed(self, height: int):
        self.setFixedSize(self.panel.maximumWidth(), max(56, height))

    def move_window_by(self, delta: QPoint):
        self.move(self.pos() + delta)

    def update_focus_glass(self):
        app = QApplication.instance()
        app_active = bool(app is not None and app.applicationState() == Qt.ApplicationActive)
        self.panel.set_inactive_glass(not app_active)
        self.setWindowOpacity(1.0 if app_active else self.planner.inactive_glass_opacity)
        if self.planner.isVisible():
            planner_opacity = min(0.98, self.planner.inactive_glass_opacity + 0.08)
            self.planner.setWindowOpacity(1.0 if app_active else planner_opacity)

    def _planner_target_near_timer(self, timer_geo: QRect) -> QRect:
        target = QRect(self.planner._planner_target_geometry)

        screen = QApplication.screenAt(timer_geo.center())
        if screen is None:
            screen = QApplication.primaryScreen()
        avail = screen.availableGeometry() if screen else QRect(0, 0, 2560, 1440)

        gap = 8
        margin = 18
        max_w = max(520, avail.width() - (margin * 2))
        max_h = max(420, avail.height() - (margin * 2))
        desired_w = max(1200, target.width())
        desired_h = max(760, target.height())
        width = min(desired_w, max_w)
        height = min(desired_h, max_h)

        def clamp(v: int, lo: int, hi: int) -> int:
            if hi < lo:
                return lo
            return max(lo, min(v, hi))

        candidates = []

        x = timer_geo.right() + 1 + gap
        y = clamp(timer_geo.y() - 12, avail.top(), avail.bottom() + 1 - height)
        candidates.append(QRect(x, y, width, height))  # rechts

        x = timer_geo.x() - gap - width
        y = clamp(timer_geo.y() - 12, avail.top(), avail.bottom() + 1 - height)
        candidates.append(QRect(x, y, width, height))  # links

        x = clamp(timer_geo.x() - 40, avail.left(), avail.right() + 1 - width)
        y = timer_geo.bottom() + 1 + gap
        candidates.append(QRect(x, y, width, height))  # beneden

        x = clamp(timer_geo.x() - 40, avail.left(), avail.right() + 1 - width)
        y = timer_geo.y() - gap - height
        candidates.append(QRect(x, y, width, height))  # boven

        timer_box = QRect(timer_geo)
        timer_box.adjust(-gap, -gap, gap, gap)

        for c in candidates:
            if avail.contains(c) and not c.intersects(timer_box):
                return c

        best = candidates[0]
        best_score = -1
        for c in candidates:
            i = c.intersected(avail)
            visible_area = max(0, i.width()) * max(0, i.height())
            overlap = c.intersected(timer_box)
            overlap_area = max(0, overlap.width()) * max(0, overlap.height())
            score = visible_area - (overlap_area * 8)
            if score > best_score:
                best_score = score
                best = c

        final_rect = QRect(
            clamp(best.x(), avail.left(), avail.right() + 1 - width),
            clamp(best.y(), avail.top(), avail.bottom() + 1 - height),
            width,
            height,
        )

        if final_rect.intersects(timer_box):
            shifts = [
                QPoint((timer_box.right() + gap + 1) - final_rect.left(), 0),
                QPoint((timer_box.left() - gap - 1) - final_rect.right(), 0),
                QPoint(0, (timer_box.bottom() + gap + 1) - final_rect.top()),
                QPoint(0, (timer_box.top() - gap - 1) - final_rect.bottom()),
            ]
            fixed = None
            for d in shifts:
                c = final_rect.translated(d)
                c = QRect(
                    clamp(c.x(), avail.left(), avail.right() + 1 - width),
                    clamp(c.y(), avail.top(), avail.bottom() + 1 - height),
                    width,
                    height,
                )
                if not c.intersects(timer_box):
                    fixed = c
                    break
            if fixed is not None:
                final_rect = fixed

        return self.planner._fit_rect_to_screen(final_rect, screen, margin)

    def toggle_planner(self):
        geo = self.geometry()
        if self.planner.isVisible():
            self.planner.slide_out_to(geo)
            self.panel.btn_open.setChecked(False)
            self.panel.btn_open.setToolTip("Planner tonen")
        else:
            self.planner._planner_target_geometry = self._planner_target_near_timer(geo)
            self.planner.slide_in_from(geo)
            self.panel.btn_open.setChecked(True)
            self.panel.btn_open.setToolTip("Planner verbergen")

    def setup_tray(self):
        if not QSystemTrayIcon.isSystemTrayAvailable():
            self.tray = None
            return
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.make_tray_icon(self.panel.running, self.panel.work_seconds, self.panel.idle_seconds))
        self.tray.setToolTip("Tijdplanner Pro")
        menu = QMenu(self)
        act_open = QAction("Open timer", self)
        act_toggle = QAction("Toon/verberg planner", self)
        self.act_start = QAction("Start timer", self)
        self.act_pause = QAction("Pauze timer", self)
        act_exit = QAction("Afsluiten", self)
        act_open.triggered.connect(self.show_from_tray)
        act_toggle.triggered.connect(self.toggle_planner)
        self.act_start.triggered.connect(self.tray_start_timer)
        self.act_pause.triggered.connect(self.tray_pause_timer)
        act_exit.triggered.connect(self.quit_all)
        menu.addAction(act_open)
        menu.addAction(act_toggle)
        menu.addSeparator()
        menu.addAction(self.act_start)
        menu.addAction(self.act_pause)
        menu.addSeparator()
        menu.addAction(act_exit)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(self.on_tray_activated)
        self.tray.show()
        self.update_tray_visual()

    def make_tray_icon(self, running: bool, work_seconds: int, idle_seconds: int) -> QIcon:
        pm = QPixmap(64, 64)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing, True)

        grad = QLinearGradient(0, 0, 64, 64)
        if running:
            grad.setColorAt(0.0, QColor("#1ea672"))
            grad.setColorAt(1.0, QColor("#2563eb"))
        else:
            grad.setColorAt(0.0, QColor("#6b7280"))
            grad.setColorAt(1.0, QColor("#374151"))
        p.setBrush(grad)
        p.setPen(QPen(QColor("#dff7ff"), 2))
        p.drawRoundedRect(6, 6, 52, 52, 14, 14)

        # Buitenring voor "actief/idle"-gevoel.
        p.setBrush(Qt.NoBrush)
        ring_color = QColor("#facc15") if idle_seconds > 0 and running else QColor("#ecfeff")
        p.setPen(QPen(ring_color, 3))
        p.drawEllipse(12, 12, 40, 40)

        # Seconde-progressie binnen de minuut.
        sec = int(work_seconds) % 60
        arc_span = int((sec / 60.0) * 360 * 16)
        p.setPen(QPen(QColor("#ffffff"), 3))
        p.drawArc(12, 12, 40, 40, 90 * 16, -arc_span)

        # Mini timer-tekst (minuten, capped).
        mins = min(999, int(work_seconds) // 60)
        txt = f"{mins:02d}" if mins < 100 else "99+"
        p.setPen(QPen(QColor("#f8fafc"), 1))
        p.setFont(QFont("Segoe UI", 14, QFont.Bold))
        p.drawText(pm.rect(), Qt.AlignCenter, txt)

        # Run/pause-indicator onderaan.
        p.setBrush(QColor("#f8fafc"))
        p.setPen(QPen(QColor("#f8fafc"), 1))
        if running:
            p.setFont(QFont("Segoe UI Symbol", 8, QFont.Bold))
            p.drawText(QRect(44, 49, 14, 11), Qt.AlignCenter, "▶")
        else:
            p.drawRect(45, 50, 2, 8)
            p.drawRect(50, 50, 2, 8)
        p.end()
        return QIcon(pm)

    def update_tray_visual(self):
        running = bool(self.panel.running)
        w = int(self.panel.work_seconds)
        i = int(self.panel.idle_seconds)
        c = int(self.panel.call_seconds)

        icon = self.make_tray_icon(running, w, i)
        self.setWindowIcon(icon)
        self.planner.setWindowIcon(icon)
        app = QApplication.instance()
        if app is not None:
            app.setWindowIcon(icon)

        if not self.tray:
            return
        self.tray.setIcon(icon)
        state = "actief" if running else "gepauzeerd"
        self.tray.setToolTip(
            f"Tijdplanner Pro ({state})\n"
            f"Werk: {seconds_to_hhmmss(w)}\n"
            f"Idle: {seconds_to_hhmmss(i)}\n"
            f"Call: {seconds_to_hhmmss(c)}"
        )
        if hasattr(self, "act_start"):
            self.act_start.setEnabled(not running)
        if hasattr(self, "act_pause"):
            self.act_pause.setEnabled(running)

    def tray_start_timer(self):
        self.panel.start()
        self.update_tray_visual()

    def tray_pause_timer(self):
        self.panel.pause()
        self.update_tray_visual()

    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_from_tray()

    def show_from_tray(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()

    def quit_all(self):
        self._force_quit = True
        self.panel.pause()
        if self.tray:
            self.tray.hide()
        self.planner._allow_close = True
        self.planner.close()
        self.close()

    def closeEvent(self, event):
        if self._force_quit:
            event.accept()
            return
        if self.tray and self.tray.isVisible():
            self.hide()
            event.ignore()
            return
        event.accept()

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Tijdplanner Pro")
    planner = MainWindow()
    planner.hide()
    timer = TimerWindow(planner)
    planner.timer_host = timer
    timer.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
















