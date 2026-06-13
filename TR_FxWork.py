# -*- coding: utf-8 -*-
"""
KAGAMI 臺鐵工程本本 V0.1.7
- Python 標準函式庫版本：tkinter + sqlite3
- 關閉前自動儲存
- 可建立多個工程
- 開啟時自動載入上次編輯工程
- 基本資料、假期表、晴雨表、鐵路疏運表、週曆總表、計價資料、工程執行紀錄表
- V0.1.7：新增歷史區規則，工作區只保留最新版號檔案。
"""

import os
import hashlib
import shutil
import sqlite3
import tempfile
import zipfile
import calendar
import re
import sys
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog


APP_VERSION = "V0.1.7"
APP_RELEASE_SUMMARY = "新增歷史區規則，工作區只保留最新版號檔案。"
APP_TITLE = f"KAGAMI 臺鐵工程本本 {APP_VERSION}"


def get_app_dir():
    exe_path = os.path.abspath(sys.argv[0] or "")
    if exe_path.lower().endswith(".exe"):
        return os.path.dirname(exe_path)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()
DB_FILE_NAME = "TRFxWork_db"
LEGACY_DB_FILE_NAME = "TR_FxWork.db"
DB_FILE = os.path.join(APP_DIR, DB_FILE_NAME)
LEGACY_DB_FILE = os.path.join(APP_DIR, LEGACY_DB_FILE_NAME)

WEEKDAY_NAMES = ["一", "二", "三", "四", "五", "六", "日"]
PASSWORD_SALT = "1981"

PROJECT_EXTRA_FIELDS = [
    "contractor", "company_address", "responsible_person", "contact_person",
    "phone", "fax", "tax_id", "purchase_contract_no", "contract_date", "project_description",
    "contract_budget_net", "contract_award_net", "contract_budget_tax", "contract_award_tax",
    "contract_budget_total", "contract_award_total",
    "labor_budget", "labor_award", "deposit_difference", "deposit_performance", "deposit_total",
    "final_contract_amount", "warranty_rate", "warranty_deposit",
]

MONEY_FIELDS = {
    "contract_budget_net", "contract_award_net", "contract_budget_tax", "contract_award_tax",
    "contract_budget_total", "contract_award_total",
    "labor_budget", "labor_award", "deposit_difference", "deposit_performance", "deposit_total",
    "final_contract_amount", "warranty_deposit",
}


def today_str():
    return date.today().strftime("%Y-%m-%d")


def parse_date(text):
    text = (text or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def fmt_date(d):
    return d.strftime("%Y-%m-%d") if d else ""


def hash_password(password):
    return hashlib.sha256((PASSWORD_SALT + (password or "")).encode("utf-8")).hexdigest()


def excel_col_name(index):
    name = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def xml_escape(value):
    return (
        str(value or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def write_simple_xlsx(path, sheet_name, rows):
    sheet_title = sheet_name[:31] or "Sheet1"
    data = [[sheet_name]] + rows
    sheet_rows = []
    for r_idx, row in enumerate(data, start=1):
        cells = []
        for c_idx, value in enumerate(row, start=1):
            ref = f"{excel_col_name(c_idx - 1)}{r_idx}"
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{xml_escape(value)}</t></is></c>')
        sheet_rows.append(f'<row r="{r_idx}">{"".join(cells)}</row>')
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(sheet_rows)}</sheetData></worksheet>'
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{xml_escape(sheet_title)}" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    )
    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '</styleSheet>'
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", root_rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr("xl/styles.xml", styles_xml)


def read_simple_xlsx(path):
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path, "r") as zf:
        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("m:si", ns):
                shared.append("".join(t.text or "" for t in si.findall(".//m:t", ns)))
        root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
    rows = []
    for row in root.findall(".//m:row", ns):
        vals = []
        last_col = 0
        for cell in row.findall("m:c", ns):
            ref = cell.attrib.get("r", "")
            match = re.match(r"([A-Z]+)", ref)
            col = 0
            if match:
                for ch in match.group(1):
                    col = col * 26 + ord(ch) - 64
            while last_col + 1 < col:
                vals.append("")
                last_col += 1
            ctype = cell.attrib.get("t", "")
            if ctype == "inlineStr":
                vals.append("".join(t.text or "" for t in cell.findall(".//m:t", ns)))
            elif ctype == "s":
                v = cell.find("m:v", ns)
                vals.append(shared[int(v.text)] if v is not None and v.text and int(v.text) < len(shared) else "")
            else:
                v = cell.find("m:v", ns)
                vals.append(v.text if v is not None and v.text is not None else "")
            last_col = col
        rows.append(vals)
    return rows


def add_calendar_days(start, days):
    if not start or not days:
        return None
    return start + timedelta(days=int(days) - 1)


def add_work_days(start, days, exclude_dates):
    """從開工日當天開始累計第 1 工作日；週六週日與 exclude_dates 不計。"""
    if not start or not days:
        return None
    days = int(days)
    count = 0
    cur = start
    while True:
        is_weekend = cur.weekday() >= 5
        is_excluded = cur in exclude_dates
        if not is_weekend and not is_excluded:
            count += 1
            if count >= days:
                return cur
        cur += timedelta(days=1)


def count_work_days_until(start, end, exclude_dates):
    if not start or not end or end < start:
        return 0
    count = 0
    cur = start
    while cur <= end:
        if cur.weekday() < 5 and cur not in exclude_dates:
            count += 1
        cur += timedelta(days=1)
    return count


def resolve_database_path():
    if os.path.exists(DB_FILE):
        return DB_FILE
    if os.path.exists(LEGACY_DB_FILE):
        shutil.copy2(LEGACY_DB_FILE, DB_FILE)
    return DB_FILE


class DB:
    def __init__(self, path):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.init_schema()

    def init_schema(self):
        c = self.conn.cursor()
        c.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            exec_no TEXT,
            budget_no TEXT,
            award_date TEXT,
            planned_start TEXT,
            actual_start TEXT,
            contract_days INTEGER DEFAULT 0,
            day_type TEXT DEFAULT '工作日',
            planned_finish_holiday TEXT,
            planned_finish_transport TEXT,
            actual_finish TEXT,
            updated_at TEXT,
            password_hash TEXT DEFAULT ''
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS bids (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            round_no INTEGER,
            online_date TEXT,
            open_date TEXT
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS holidays (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            day TEXT,
            name TEXT
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS holiday_project_excludes (
            project_id INTEGER,
            day TEXT,
            excluded INTEGER DEFAULT 0,
            PRIMARY KEY(project_id, day)
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS workdays (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            day TEXT,
            name TEXT
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS workday_project_excludes (
            project_id INTEGER,
            day TEXT,
            excluded INTEGER DEFAULT 0,
            PRIMARY KEY(project_id, day)
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS weather (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            day TEXT,
            morning REAL DEFAULT 0,
            afternoon REAL DEFAULT 0,
            typhoon REAL DEFAULT 0,
            site REAL DEFAULT 0,
            note TEXT
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS railway (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            day TEXT,
            note TEXT
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS railway_project_excludes (
            project_id INTEGER,
            day TEXT,
            excluded INTEGER DEFAULT 0,
            PRIMARY KEY(project_id, day)
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS payment_contract (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            day TEXT,
            item TEXT,
            voucher_no TEXT,
            amount REAL DEFAULT 0,
            note TEXT
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS payment_other (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            day TEXT,
            item TEXT,
            voucher_no TEXT,
            amount REAL DEFAULT 0,
            note TEXT
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS payment_admin (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            day TEXT,
            item TEXT,
            voucher_no TEXT,
            amount REAL DEFAULT 0,
            note TEXT
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS execution_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            day TEXT,
            record_type TEXT,
            subject TEXT,
            content TEXT,
            note TEXT
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS project_milestones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            item_no TEXT,
            contract_item TEXT,
            start_date TEXT,
            deadline_days REAL DEFAULT 0,
            deadline_date TEXT,
            received_date TEXT,
            overdue TEXT,
            received_no TEXT,
            note TEXT,
            day_adjust REAL DEFAULT 0
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS execution_status (
            project_id INTEGER PRIMARY KEY,
            status TEXT DEFAULT '',
            updated_at TEXT
        )
        """)
        try:
            c.execute("ALTER TABLE projects ADD COLUMN password_hash TEXT DEFAULT ''")
        except sqlite3.OperationalError:
            pass
        for field in PROJECT_EXTRA_FIELDS:
            try:
                c.execute(f"ALTER TABLE projects ADD COLUMN {field} TEXT DEFAULT ''")
            except sqlite3.OperationalError:
                pass
        self.conn.commit()

    def get_setting(self, key, default=""):
        row = self.conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default

    def set_setting(self, key, value):
        self.conn.execute(
            "INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
            (key, str(value))
        )
        self.conn.commit()

    def projects(self):
        return self.conn.execute("SELECT * FROM projects ORDER BY updated_at DESC, id DESC").fetchall()

    def create_project(self, name):
        now = datetime.now().isoformat(timespec="seconds")
        cur = self.conn.execute(
            "INSERT INTO projects(name, updated_at) VALUES(?,?)",
            (name, now)
        )
        self.conn.commit()
        return cur.lastrowid

    def get_project(self, pid):
        return self.conn.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()

    def get_password_hash(self, pid):
        row = self.conn.execute("SELECT password_hash FROM projects WHERE id=?", (pid,)).fetchone()
        return (row["password_hash"] or "") if row and "password_hash" in row.keys() else ""

    def save_password_hash(self, pid, password_hash):
        self.conn.execute("UPDATE projects SET password_hash=?, updated_at=? WHERE id=?", (password_hash or "", datetime.now().isoformat(timespec="seconds"), pid))
        self.conn.commit()

    def get_project_by_name(self, name):
        return self.conn.execute("SELECT * FROM projects WHERE name=? ORDER BY id LIMIT 1", (name,)).fetchone()

    def delete_project(self, pid):
        for table in (
            "bids", "weather",
            "payment_contract", "payment_other", "payment_admin",
            "execution_records", "execution_status", "project_milestones",
            "holiday_project_excludes", "workday_project_excludes", "railway_project_excludes"
        ):
            self.conn.execute(f"DELETE FROM {table} WHERE project_id=?", (pid,))
        self.conn.execute("DELETE FROM projects WHERE id=?", (pid,))
        self.conn.commit()

    def get_status(self, pid):
        row = self.conn.execute("SELECT status FROM execution_status WHERE project_id=?", (pid,)).fetchone()
        return row["status"] if row else ""

    def save_status(self, pid, status):
        self.conn.execute(
            "INSERT OR REPLACE INTO execution_status(project_id, status, updated_at) VALUES(?,?,?)",
            (pid, status, datetime.now().isoformat(timespec="seconds"))
        )
        self.conn.commit()

    def save_project(self, pid, data):
        data["updated_at"] = datetime.now().isoformat(timespec="seconds")
        fields = [
            "name", "exec_no", "budget_no", "award_date", "planned_start",
            "actual_start", "contract_days", "day_type",
            "planned_finish_holiday", "planned_finish_transport", "actual_finish",
            "updated_at"
        ] + PROJECT_EXTRA_FIELDS
        sets = ", ".join(f"{f}=?" for f in fields)
        vals = [data.get(f, "") for f in fields] + [pid]
        self.conn.execute(f"UPDATE projects SET {sets} WHERE id=?", vals)
        self.conn.commit()

    def rows(self, table, pid):
        if table == "holidays":
            return self.shared_holidays(pid)
        if table == "workdays":
            return self.shared_workdays(pid)
        if table == "railway":
            return self.shared_railway(pid)
        if table == "project_milestones":
            return self.conn.execute("SELECT * FROM project_milestones WHERE project_id=? ORDER BY start_date, id", (pid,)).fetchall()
        return self.conn.execute(f"SELECT * FROM {table} WHERE project_id=? ORDER BY day, id", (pid,)).fetchall()

    def bids(self, pid):
        return self.conn.execute("SELECT * FROM bids WHERE project_id=? ORDER BY round_no, id", (pid,)).fetchall()

    def shared_holidays(self, pid):
        return self.conn.execute("""
            SELECT h.day, h.name, COALESCE(e.excluded, 0) AS excluded
            FROM (
                SELECT day, COALESCE(NULLIF(MAX(name), ''), '假日') AS name
                FROM holidays
                WHERE day IS NOT NULL AND TRIM(day) <> ''
                GROUP BY day
            ) h
            LEFT JOIN holiday_project_excludes e
                ON e.day = h.day AND e.project_id = ?
            ORDER BY h.day
        """, (pid,)).fetchall()

    def shared_workdays(self, pid):
        return self.conn.execute("""
            SELECT h.day, h.name, COALESCE(e.excluded, 0) AS excluded
            FROM (
                SELECT day, COALESCE(NULLIF(MAX(name), ''), '補班') AS name
                FROM workdays
                WHERE day IS NOT NULL AND TRIM(day) <> ''
                GROUP BY day
            ) h
            LEFT JOIN workday_project_excludes e
                ON e.day = h.day AND e.project_id = ?
            ORDER BY h.day
        """, (pid,)).fetchall()

    def shared_railway(self, pid):
        return self.conn.execute("""
            SELECT r.day, r.note, COALESCE(e.excluded, 0) AS excluded
            FROM (
                SELECT day, COALESCE(NULLIF(MAX(note), ''), '疏運') AS note
                FROM railway
                WHERE day IS NOT NULL AND TRIM(day) <> ''
                GROUP BY day
            ) r
            LEFT JOIN railway_project_excludes e
                ON e.day = r.day AND e.project_id = ?
            ORDER BY r.day
        """, (pid,)).fetchall()

    def replace_rows(self, table, pid, rows):
        if table not in ("holidays", "workdays", "railway"):
            self.conn.execute(f"DELETE FROM {table} WHERE project_id=?", (pid,))
        if table == "bids":
            for r in rows:
                self.conn.execute(
                    "INSERT INTO bids(project_id, round_no, online_date, open_date) VALUES(?,?,?,?)",
                    (pid, r.get("round_no", 1), r.get("online_date", ""), r.get("open_date", ""))
                )
        elif table == "holidays":
            self.conn.execute("DELETE FROM holidays")
            self.conn.execute("DELETE FROM holiday_project_excludes WHERE project_id=?", (pid,))
            for r in rows:
                day = r.get("day", "")
                self.conn.execute(
                    "INSERT INTO holidays(project_id, day, name) VALUES(?,?,?)",
                    (0, day, r.get("name", ""))
                )
                if r.get("excluded", 0):
                    self.conn.execute(
                        "INSERT OR REPLACE INTO holiday_project_excludes(project_id, day, excluded) VALUES(?,?,1)",
                        (pid, day)
                    )
        elif table == "workdays":
            self.conn.execute("DELETE FROM workdays")
            self.conn.execute("DELETE FROM workday_project_excludes WHERE project_id=?", (pid,))
            for r in rows:
                day = r.get("day", "")
                self.conn.execute(
                    "INSERT INTO workdays(project_id, day, name) VALUES(?,?,?)",
                    (0, day, r.get("name", ""))
                )
                if r.get("excluded", 0):
                    self.conn.execute(
                        "INSERT OR REPLACE INTO workday_project_excludes(project_id, day, excluded) VALUES(?,?,1)",
                        (pid, day)
                    )
        elif table == "weather":
            for r in rows:
                self.conn.execute(
                    "INSERT INTO weather(project_id, day, morning, afternoon, typhoon, site, note) VALUES(?,?,?,?,?,?,?)",
                    (
                        pid, r.get("day", ""), r.get("morning", 0), r.get("afternoon", 0),
                        r.get("typhoon", 0), r.get("site", 0), r.get("note", "")
                    )
                )
        elif table == "railway":
            self.conn.execute("DELETE FROM railway")
            self.conn.execute("DELETE FROM railway_project_excludes WHERE project_id=?", (pid,))
            for r in rows:
                day = r.get("day", "")
                self.conn.execute(
                    "INSERT INTO railway(project_id, day, note) VALUES(?,?,?)",
                    (0, day, r.get("note", ""))
                )
                if r.get("excluded", 0):
                    self.conn.execute(
                        "INSERT OR REPLACE INTO railway_project_excludes(project_id, day, excluded) VALUES(?,?,1)",
                        (pid, day)
                    )
        elif table in ("payment_contract", "payment_other", "payment_admin"):
            for r in rows:
                self.conn.execute(
                    f"INSERT INTO {table}(project_id, day, item, voucher_no, amount, note) VALUES(?,?,?,?,?,?)",
                    (pid, r.get("day", ""), r.get("item", ""), r.get("voucher_no", ""), r.get("amount", 0), r.get("note", ""))
                )
        elif table == "execution_records":
            for r in rows:
                self.conn.execute(
                    "INSERT INTO execution_records(project_id, day, record_type, subject, content, note) VALUES(?,?,?,?,?,?)",
                    (pid, r.get("day", ""), r.get("record_type", ""), r.get("subject", ""), r.get("content", ""), r.get("note", ""))
                )
        elif table == "project_milestones":
            for r in rows:
                self.conn.execute(
                    "INSERT INTO project_milestones(project_id, item_no, contract_item, start_date, deadline_days, deadline_date, received_date, overdue, received_no, note, day_adjust) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                    (
                        pid, r.get("item_no", ""), r.get("contract_item", ""), r.get("start_date", ""),
                        r.get("deadline_days", 0), r.get("deadline_date", ""), r.get("received_date", ""),
                        r.get("overdue", ""), r.get("received_no", ""), r.get("note", ""), r.get("day_adjust", 0)
                    )
                )
        self.conn.commit()


class EditableTree(ttk.Frame):
    def __init__(self, master, columns, headings, widths, on_changed=None, can_edit=None, add_command=None):
        super().__init__(master)
        self.columns = columns
        self.headings = headings
        self.on_changed = on_changed
        self.can_edit = can_edit or (lambda: True)
        self.add_command = add_command or self.add_row
        self.tree = ttk.Treeview(self, columns=columns, show="headings", height=12, style="Grid.Treeview")
        vs = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.tag_configure("pink", background="#f4cccc")
        self.tree.tag_configure("red", foreground="#cc0000")
        self.tree.tag_configure("year_sep", background="#000000", foreground="#ffffff")

        for col, head, width in zip(columns, headings, widths):
            self.tree.heading(col, text=head, command=lambda c=col: self.sort_by_column(c))
            fixed_cols = {"exclude", "day", "name", "note", "morning", "afternoon", "typhoon", "site"}
            self.tree.column(col, width=width, minwidth=width, anchor="center", stretch=False if col in fixed_cols else True)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        btns = ttk.Frame(self)
        btns.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
        ttk.Button(btns, text="新增一列", command=self.add_command).pack(side="left", padx=3)
        ttk.Button(btns, text="編輯選取列", command=self.edit_row).pack(side="left", padx=3)
        ttk.Button(btns, text="刪除選取列", command=self.delete_row).pack(side="left", padx=3)
        self.tree.bind("<Double-1>", lambda e: self.edit_row())
        self.tree.bind("<Button-1>", self.on_tree_click)

    def on_tree_click(self, event):
        if not self.columns or self.columns[0] != "exclude":
            return
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        if col != "#1":
            return
        item = self.tree.identify_row(event.y)
        if not item:
            return
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return "break"
        values = list(self.tree.item(item, "values"))
        values[0] = "" if values and values[0] == "✓" else "✓"
        self.tree.item(item, values=values)
        self.changed()
        return "break"

    def add_row(self, values=None):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        values = values or [""] * len(self.columns)
        self.tree.insert("", "end", values=values)
        self.changed()

    def add_row_after_selection(self, values=None):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        values = values or [""] * len(self.columns)
        selected = self.tree.selection()
        if selected:
            index = self.tree.index(selected[-1]) + 1
            item = self.tree.insert("", index, values=values)
        else:
            item = self.tree.insert("", "end", values=values)
        self.tree.selection_set(item)
        self.tree.focus(item)
        self.tree.see(item)
        self.changed()

    def edit_row(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        item = self.tree.focus()
        if not item:
            return
        old = list(self.tree.item(item, "values"))
        win = tk.Toplevel(self)
        win.title("編輯資料")
        win.grab_set()
        entries = []
        for i, col in enumerate(self.columns):
            label = self.headings[i] if i < len(self.headings) else col
            ttk.Label(win, text=label).grid(row=i, column=0, sticky="e", padx=8, pady=4)
            e = ttk.Entry(win, width=32)
            e.grid(row=i, column=1, sticky="ew", padx=8, pady=4)
            e.insert(0, old[i] if i < len(old) else "")
            entries.append(e)

        weather_edit = "morning" in self.columns and "afternoon" in self.columns
        if weather_edit:
            morning_index = self.columns.index("morning")
            afternoon_index = self.columns.index("afternoon")

            def morning_is_full_day():
                try:
                    return float(entries[morning_index].get().strip() or 0) >= 1.0
                except ValueError:
                    return False

            def update_afternoon_state(event=None):
                if morning_is_full_day():
                    entries[afternoon_index].delete(0, "end")
                    entries[afternoon_index].configure(state="disabled")
                else:
                    entries[afternoon_index].configure(state="normal")

            entries[morning_index].bind("<KeyRelease>", update_afternoon_state)
            entries[morning_index].bind("<FocusOut>", update_afternoon_state)
            update_afternoon_state()

        def ok():
            vals = [e.get().strip() for e in entries]
            if weather_edit:
                try:
                    if float(vals[morning_index] or 0) >= 1.0:
                        vals[afternoon_index] = ""
                except ValueError:
                    pass
            self.tree.item(item, values=vals)
            win.destroy()
            self.changed()

        ttk.Button(win, text="確定", command=ok).grid(row=len(self.columns), column=0, columnspan=2, pady=10)
        win.grid_columnconfigure(1, weight=1)

    def delete_row(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        for item in self.tree.selection():
            self.tree.delete(item)
        self.changed()

    def set_rows(self, rows):
        self.tree.delete(*self.tree.get_children())
        for row in rows:
            self.tree.insert("", "end", values=row)

    def apply_row_tags(self, tag_func):
        for item in self.tree.get_children():
            values = list(self.tree.item(item, "values"))
            tag = tag_func(values)
            self.tree.item(item, tags=(tag,) if tag else ())

    def get_rows(self):
        return [list(self.tree.item(i, "values")) for i in self.tree.get_children()]

    def changed(self):
        if self.on_changed:
            self.on_changed()

    def sort_by_column(self, col):
        rows = [(self.tree.set(item, col), item) for item in self.tree.get_children("")]
        rows.sort(key=lambda x: x[0])
        for index, (_, item) in enumerate(rows):
            self.tree.move(item, "", index)
        self.changed()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1250x820")
        self.minsize(1100, 720)
        self.resizable(True, True)

        self.db = DB(resolve_database_path())
        self.current_project_id = None
        self.loading = False
        self.dirty = False
        self.project_password_hash = ""
        self.edit_unlocked = True
        self.edit_widgets = []
        self.undo_snapshot = None
        self.last_state = None
        self.restoring = False
        self.recalculating = False
        self.save_after_id = None

        self.style = ttk.Style()
        self.style.configure("Top.TLabelframe.Label", font=("Microsoft JhengHei UI", 11, "bold"))
        self.style.configure("TLabel", font=("Microsoft JhengHei UI", 10))
        self.style.configure("TButton", font=("Microsoft JhengHei UI", 10))
        self.style.configure("Treeview", rowheight=26, font=("Microsoft JhengHei UI", 10))
        self.style.configure("Treeview.Heading", font=("Microsoft JhengHei UI", 10, "bold"))
        self.style.configure("Grid.Treeview", rowheight=26, font=("Microsoft JhengHei UI", 10), borderwidth=1, relief="solid")
        self.style.configure("Grid.Treeview.Heading", font=("Microsoft JhengHei UI", 10, "bold"), borderwidth=1, relief="solid")
        self.style.map("TEntry", foreground=[("disabled", "#1f4e79")])
        self.style.map("TCombobox", foreground=[("disabled", "#1f4e79")])

        self.build_ui()
        self.load_projects()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(3000, self.auto_save_loop)

    def build_ui(self):
        top_select = ttk.Frame(self, padding=8)
        top_select.pack(fill="x")

        ttk.Label(top_select, text="工程名稱：").pack(side="left")
        self.project_combo = ttk.Combobox(top_select, state="readonly", width=55)
        self.project_combo.pack(side="left", padx=5)
        self.project_combo.bind("<<ComboboxSelected>>", self.on_project_selected)
        ttk.Button(top_select, text="新增工程", command=self.new_project).pack(side="left", padx=5)
        ttk.Button(top_select, text="立即儲存", command=self.save_current).pack(side="left", padx=5)
        ttk.Button(top_select, text="回復上一個動作", command=self.undo_last_action).pack(side="left", padx=5)

        self.edit_password_var = tk.StringVar()
        self.lock_state_var = tk.StringVar(value="未鎖定")
        ttk.Label(top_select, textvariable=self.lock_state_var, foreground="#a64d00").pack(side="left", padx=3)

        ttk.Button(top_select, text="▶", width=3, command=self.toggle_function_panel).pack(side="right", padx=(6, 0))
        self.status_var = tk.StringVar(value="")
        ttk.Label(top_select, textvariable=self.status_var).pack(side="right")
        self.function_panel = None

        self.summary = ttk.LabelFrame(self, text="工程基本資料顯示區", padding=8, style="Top.TLabelframe")
        self.summary.pack(fill="x", padx=8, pady=(0, 8))
        self.summary_vars = {}
        labels = [
            ("開工時間", "start"),
            ("預定完工時間", "finish1"),
            ("修正後預訂完工時間", "finish2"),
            ("已經過多少施工日數", "elapsed"),
            ("到今天日期是第幾工作日", "workday_no"),
            ("發包總核銷金額", "contract_total"),
            ("發包以外核銷總金額", "other_total"),
            ("管理費總核銷金額", "admin_total"),
            ("工程執行狀態", "execution_status"),
        ]
        for i, (text, key) in enumerate(labels):
            row = 0 if i < 5 else 1
            col = i if i < 5 else i - 5
            ttk.Label(self.summary, text=text + "：").grid(row=row, column=col*2, sticky="e", padx=(5, 2), pady=3)
            v = tk.StringVar()
            self.summary_vars[key] = v
            ttk.Label(self.summary, textvariable=v, foreground="#1f4e79", font=("Microsoft JhengHei UI", 11, "bold")).grid(
                row=row, column=col*2+1, sticky="w", padx=(0, 12), pady=3
            )

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)

        self.tab_basic = ttk.Frame(self.nb, padding=8)
        self.tab_holiday = ttk.Frame(self.nb, padding=8)
        self.tab_weather = ttk.Frame(self.nb, padding=8)
        self.tab_railway = ttk.Frame(self.nb, padding=8)
        self.tab_calendar = ttk.Frame(self.nb, padding=8)
        self.tab_payment_contract = ttk.Frame(self.nb, padding=8)
        self.tab_payment_other = ttk.Frame(self.nb, padding=8)
        self.tab_payment_admin = ttk.Frame(self.nb, padding=8)
        self.tab_execution = ttk.Frame(self.nb, padding=8)
        self.tab_milestone = ttk.Frame(self.nb, padding=8)

        self.nb.add(self.tab_basic, text="工程基本資料")
        self.nb.add(self.tab_holiday, text="假期表")
        self.nb.add(self.tab_weather, text="晴雨表")
        self.nb.add(self.tab_railway, text="鐵路疏運表")
        self.nb.add(self.tab_calendar, text="施工日曆")
        self.nb.add(self.tab_payment_contract, text="發包工程費計價")
        self.nb.add(self.tab_payment_other, text="發包以外計價")
        self.nb.add(self.tab_payment_admin, text="管理費計價")
        self.nb.add(self.tab_execution, text="工程執行紀錄表")
        self.nb.add(self.tab_milestone, text="工程大事記")

        self.build_basic_tab()
        self.build_holiday_tab()
        self.build_weather_tab()
        self.build_railway_tab()
        self.build_calendar_tab()
        self.build_payment_tabs()
        self.build_execution_tab()
        self.build_milestone_tab()
        self.assign_tree_edit_guards()

    def can_edit(self):
        manual_var = getattr(self, "data_edit_enabled_var", None)
        manual_unlocked = True if manual_var is None else bool(manual_var.get())
        password_unlocked = self.edit_unlocked or not self.project_password_hash
        return manual_unlocked and password_unlocked

    def add_page_edit_toggle(self, parent):
        holder = ttk.Frame(parent)
        holder.pack(fill="x", pady=(0, 4))
        ttk.Checkbutton(
            holder,
            text="資料編輯鎖定解除（勾選才可編輯）",
            variable=self.data_edit_enabled_var,
            command=self.apply_edit_lock_state
        ).pack(anchor="w")
        return holder

    def mark_dirty(self, *_):
        if not self.loading and not self.restoring:
            if not self.can_edit():
                self.status_var.set("編輯已鎖定，請先輸入正確編輯密碼")
                return
            if self.recalculating:
                self.dirty = True
                self.schedule_auto_save()
                return
            if not self.recalculating:
                current_state = self.capture_state()
                if self.last_state is not None and current_state != self.last_state:
                    self.undo_snapshot = self.last_state
                self.last_state = current_state
            self.dirty = True
            self.recalculate()
            if not self.recalculating:
                self.last_state = self.capture_state()
            self.schedule_auto_save()

    def capture_state(self):
        state = {
            "basic": {},
            "day_type": self.day_type_var.get() if hasattr(self, "day_type_var") else "",
            "execution_status": self.execution_status_var.get() if hasattr(self, "execution_status_var") else "",
            "trees": {},
        }
        if hasattr(self, "basic_vars"):
            state["basic"] = {k: v.get() for k, v in self.basic_vars.items()}
        if hasattr(self, "project_description_text"):
            state["basic"]["project_description"] = self.project_description_text.get("1.0", "end-1c")
        for name in [
            "bid_tree", "holiday_tree", "workday_tree", "weather_tree", "railway_tree",
            "payment_contract_tree", "payment_other_tree", "payment_admin_tree",
            "execution_tree", "milestone_tree"
        ]:
            if hasattr(self, name):
                state["trees"][name] = [list(row) for row in getattr(self, name).get_rows()]
        return state

    def restore_state(self, state):
        if not state:
            return
        self.restoring = True
        try:
            for key, value in state.get("basic", {}).items():
                if key in self.basic_vars:
                    self.basic_vars[key].set(value)
            if hasattr(self, "project_description_text"):
                self.project_description_text.delete("1.0", "end")
                self.project_description_text.insert("1.0", state.get("basic", {}).get("project_description", ""))
            if hasattr(self, "day_type_var"):
                self.day_type_var.set(state.get("day_type", "工作日") or "工作日")
            if hasattr(self, "execution_status_var"):
                self.execution_status_var.set(state.get("execution_status", "規劃中") or "規劃中")
            for name, rows in state.get("trees", {}).items():
                if hasattr(self, name):
                    getattr(self, name).set_rows(rows)
        finally:
            self.restoring = False
        self.dirty = True
        self.recalculate()
        self.last_state = self.capture_state()

    def undo_last_action(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        if not self.undo_snapshot:
            self.status_var.set("目前沒有可回復的上一個動作")
            return
        snapshot = self.undo_snapshot
        self.undo_snapshot = None
        self.restore_state(snapshot)
        self.save_current()
        self.status_var.set("已回復上一個動作並自動儲存")

    def schedule_auto_save(self):
        if self.loading or self.restoring:
            return
        if self.save_after_id:
            try:
                self.after_cancel(self.save_after_id)
            except tk.TclError:
                pass
        self.save_after_id = self.after(800, self.run_scheduled_save)

    def run_scheduled_save(self):
        self.save_after_id = None
        if self.dirty:
            self.save_current()

    def format_money_value(self, value):
        amount = self.safe_amount(value)
        if amount == 0 and not str(value or "").strip().replace("元", "").replace(",", ""):
            return ""
        return f"{amount:,.0f}元"

    def format_money_field(self, key):
        if key not in self.basic_vars or self.loading or self.restoring:
            return
        var = self.basic_vars[key]
        raw = var.get().strip()
        if not raw:
            return
        formatted = self.format_money_value(raw)
        if formatted and formatted != raw:
            var.set(formatted)

    def toggle_function_panel(self):
        if self.function_panel and self.function_panel.winfo_exists():
            self.function_panel.destroy()
            self.function_panel = None
            return
        panel = tk.Toplevel(self)
        self.function_panel = panel
        panel.title("功能區")
        panel.transient(self)
        panel.resizable(False, True)
        x = self.winfo_rootx() + max(self.winfo_width() - 280, 0)
        y = self.winfo_rooty() + 60
        panel.geometry(f"280x520+{x}+{y}")
        box = ttk.Frame(panel, padding=12)
        box.pack(fill="both", expand=True)
        ttk.Label(box, text="編輯密碼").pack(anchor="w", pady=(0, 2))
        ttk.Entry(box, textvariable=self.edit_password_var, show="*", width=22).pack(fill="x", pady=(0, 4))
        pwd_buttons = ttk.Frame(box)
        pwd_buttons.pack(fill="x", pady=(0, 8))
        ttk.Button(pwd_buttons, text="鎖定", command=self.lock_project).pack(side="left", expand=True, fill="x", padx=(0, 2))
        ttk.Button(pwd_buttons, text="解鎖", command=self.unlock_project).pack(side="left", expand=True, fill="x", padx=2)
        ttk.Button(pwd_buttons, text="取消", command=self.clear_project_password).pack(side="left", expand=True, fill="x", padx=(2, 0))
        ttk.Separator(box).pack(fill="x", pady=8)
        ttk.Button(box, text="修改工程名稱", command=self.rename_current_project).pack(fill="x", pady=4)
        ttk.Button(box, text="資料庫打包備份", command=self.backup_database).pack(fill="x", pady=4)
        ttk.Button(box, text="異地備份資料庫", command=self.backup_database_offsite).pack(fill="x", pady=4)
        ttk.Button(box, text="匯入備份", command=self.import_database).pack(fill="x", pady=4)
        ttk.Button(box, text="匯出分頁檔案", command=self.export_page_excel).pack(fill="x", pady=4)
        ttk.Button(box, text="匯入 Excel 到分頁", command=self.import_page_excel).pack(fill="x", pady=4)
        ttk.Button(box, text="刪除工程", command=self.delete_current_project).pack(fill="x", pady=4)
        ttk.Button(box, text="關閉功能區", command=panel.destroy).pack(fill="x", pady=(18, 4))

    def open_date_picker(self, target_var, title="選擇日期"):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        base_date = parse_date(target_var.get()) or parse_date(self.basic_vars.get("planned_start", tk.StringVar()).get()) or date.today()
        win = tk.Toplevel(self)
        win.title(title)
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)

        year_var = tk.IntVar(value=base_date.year)
        month_var = tk.IntVar(value=base_date.month)
        cal_frame = ttk.Frame(win, padding=10)

        top = ttk.Frame(win, padding=10)
        top.pack(fill="x")

        def shift_month(delta):
            year = year_var.get()
            month = month_var.get() + delta
            if month < 1:
                month = 12
                year -= 1
            elif month > 12:
                month = 1
                year += 1
            year_var.set(year)
            month_var.set(month)
            render_calendar()

        ttk.Button(top, text="上一月", command=lambda: shift_month(-1)).pack(side="left", padx=(0, 8))
        ttk.Label(top, text="年").pack(side="left")
        year_spin = ttk.Spinbox(top, from_=base_date.year - 30, to=base_date.year + 30, textvariable=year_var, width=8, command=lambda: render_calendar())
        year_spin.pack(side="left", padx=(2, 8))
        ttk.Label(top, text="月").pack(side="left")
        month_spin = ttk.Spinbox(top, from_=1, to=12, textvariable=month_var, width=5, command=lambda: render_calendar())
        month_spin.pack(side="left", padx=(2, 8))
        ttk.Button(top, text="下一月", command=lambda: shift_month(1)).pack(side="left")
        ttk.Button(top, text="今天", command=lambda: (year_var.set(date.today().year), month_var.set(date.today().month), render_calendar())).pack(side="left", padx=8)
        cal_frame.pack()

        def choose(day):
            target_var.set(day.strftime("%Y-%m-%d"))
            win.destroy()

        def render_calendar():
            for child in cal_frame.winfo_children():
                child.destroy()
            try:
                year = int(year_var.get())
                month = int(month_var.get())
                if month < 1 or month > 12:
                    raise ValueError
            except (tk.TclError, ValueError):
                return
            for cidx, name in enumerate(WEEKDAY_NAMES):
                ttk.Label(cal_frame, text="星期" + name, width=9, anchor="center").grid(row=0, column=cidx, padx=1, pady=1)
            cal = calendar.Calendar(firstweekday=0)
            for ridx, week in enumerate(cal.monthdatescalendar(year, month), start=1):
                for cidx, day in enumerate(week):
                    if day.month != month:
                        ttk.Label(cal_frame, text="", width=9).grid(row=ridx, column=cidx, padx=1, pady=1)
                    else:
                        ttk.Button(cal_frame, text=str(day.day), width=8, command=lambda d=day: choose(d)).grid(row=ridx, column=cidx, padx=1, pady=1)

        year_spin.bind("<Return>", lambda e: render_calendar())
        month_spin.bind("<Return>", lambda e: render_calendar())
        year_spin.bind("<FocusOut>", lambda e: render_calendar())
        month_spin.bind("<FocusOut>", lambda e: render_calendar())
        render_calendar()

    def entry(self, parent, row, col, label, key, width=28, date_picker=False):
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="e", padx=3, pady=2)
        var = tk.StringVar()
        var.trace_add("write", self.mark_dirty)
        if date_picker:
            holder = ttk.Frame(parent)
            holder.grid(row=row, column=col+1, sticky="ew", padx=3, pady=2)
            ent = ttk.Entry(holder, textvariable=var, width=max(width - 4, 8))
            ent.pack(side="left", fill="x", expand=True)
            btn = ttk.Button(holder, text="▼", width=3, command=lambda v=var, t=label: self.open_date_picker(v, t))
            btn.pack(side="left", padx=(3, 0))
            self.edit_widgets.append(btn)
        else:
            ent = ttk.Entry(parent, textvariable=var, width=width)
            ent.grid(row=row, column=col+1, sticky="ew", padx=3, pady=2)
        self.edit_widgets.append(ent)
        self.basic_vars[key] = var
        if key in MONEY_FIELDS:
            ent.bind("<FocusOut>", lambda e, k=key: self.format_money_field(k))
            ent.bind("<Return>", lambda e, k=key: self.format_money_field(k))
        return ent

    def section_title(self, parent, text, row, columnspan=8):
        ttk.Label(parent, text=text, anchor="center", font=("Microsoft JhengHei UI", 12, "bold")).grid(
            row=row, column=0, columnspan=columnspan, sticky="ew", padx=6, pady=(10, 6)
        )

    def multiline_entry(self, parent, row, col, label, key, height=3):
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="ne", padx=3, pady=2)
        txt = tk.Text(parent, height=height, width=60, wrap="word", font=("Microsoft JhengHei UI", 10))
        txt.configure(borderwidth=1, relief="solid", highlightthickness=1, highlightbackground="#888888")
        txt.grid(row=row, column=col+1, columnspan=7, sticky="ew", padx=3, pady=2)
        txt.bind("<KeyRelease>", lambda e: self.mark_dirty())
        txt.bind("<FocusOut>", lambda e: self.mark_dirty())
        self.edit_widgets.append(txt)
        setattr(self, key + "_text", txt)
        return txt

    def build_basic_tab(self):
        self.basic_vars = {}
        canvas = tk.Canvas(self.tab_basic, highlightthickness=0)
        vs = ttk.Scrollbar(self.tab_basic, orient="vertical", command=canvas.yview)
        hs = ttk.Scrollbar(self.tab_basic, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        self.tab_basic.grid_rowconfigure(0, weight=1)
        self.tab_basic.grid_columnconfigure(0, weight=1)
        content = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=content, anchor="nw")
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=max(e.width, content.winfo_reqwidth())))
        def on_mousewheel(event):
            delta = -1 if event.delta > 0 else 1
            canvas.yview_scroll(delta * 3, "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        form = ttk.Frame(content, padding=8)
        form.pack(fill="x")
        self.section_title(form, "工程基本資料", 0)

        self.data_edit_enabled_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            form,
            text="資料編輯鎖定解除（勾選才可編輯）",
            variable=self.data_edit_enabled_var,
            command=self.apply_edit_lock_state
        ).grid(row=1, column=0, columnspan=8, sticky="w", padx=3, pady=2)

        self.entry(form, 2, 0, "工程名稱", "name", 18)
        self.entry(form, 2, 2, "工程執行號", "exec_no", 14)
        self.entry(form, 2, 4, "動支請示單號", "budget_no", 14)
        self.entry(form, 2, 6, "採購契約號碼", "purchase_contract_no", 14)
        self.entry(form, 3, 0, "決標日期", "award_date", 14, date_picker=True)
        self.entry(form, 3, 2, "簽約日期", "contract_date", 14, date_picker=True)
        self.entry(form, 3, 4, "預訂開工日", "planned_start", 14, date_picker=True)
        self.entry(form, 3, 6, "實際開工日", "actual_start", 14, date_picker=True)
        self.entry(form, 4, 0, "契約工期", "contract_days", 10)
        ttk.Label(form, text="工期類型").grid(row=4, column=2, sticky="e", padx=3, pady=2)
        self.day_type_var = tk.StringVar(value="工作日")
        self.day_type_var.trace_add("write", self.mark_dirty)
        self.day_type = ttk.Combobox(form, textvariable=self.day_type_var, state="readonly", values=["工作日", "日曆天"], width=10)
        self.day_type.grid(row=4, column=3, sticky="ew", padx=3, pady=2)
        self.edit_widgets.append(self.day_type)

        self.entry(form, 4, 4, "預訂竣工日（例假表）", "planned_finish_holiday", 14, date_picker=True)
        self.entry(form, 4, 6, "預訂竣工日（疏運表）", "planned_finish_transport", 14, date_picker=True)
        self.entry(form, 5, 0, "實際竣工日", "actual_finish", 14, date_picker=True)
        self.entry(form, 5, 2, "承攬商", "contractor", 14)
        self.entry(form, 5, 4, "公司地址", "company_address", 14)
        self.entry(form, 5, 6, "負責人", "responsible_person", 14)
        self.entry(form, 6, 0, "聯絡人", "contact_person", 14)
        self.entry(form, 6, 2, "電話", "phone", 14)
        self.entry(form, 6, 4, "傳真電話", "fax", 14)
        self.entry(form, 6, 6, "統一編號", "tax_id", 14)

        ttk.Label(form, text="工程執行狀態").grid(row=7, column=0, sticky="e", padx=3, pady=2)
        self.execution_status_var = tk.StringVar(value="規劃中")
        self.execution_status_var.trace_add("write", self.mark_dirty)
        self.execution_status_combo = ttk.Combobox(
            form,
            textvariable=self.execution_status_var,
            values=["規劃中", "招標中", "決標完成", "開工中", "施工中", "停工中", "復工中", "竣工中", "已竣工", "驗收中", "驗收完成", "結案"],
            width=14
        )
        self.execution_status_combo.grid(row=7, column=1, sticky="ew", padx=3, pady=2)
        self.edit_widgets.append(self.execution_status_combo)

        self.multiline_entry(form, 8, 0, "工程說明", "project_description", height=3)

        self.section_title(form, "發包工程費", 9)
        self.entry(form, 10, 0, "預算(未稅)", "contract_budget_net", 12)
        self.entry(form, 10, 2, "決標(未稅)", "contract_award_net", 12)
        self.entry(form, 10, 4, "稅金(預算)", "contract_budget_tax", 12)
        self.entry(form, 10, 6, "稅金(決標)", "contract_award_tax", 12)
        self.entry(form, 11, 0, "預算(含稅)", "contract_budget_total", 12)
        self.entry(form, 11, 2, "決標(契約金額含稅)", "contract_award_total", 12)

        self.section_title(form, "包工費", 12)
        self.entry(form, 13, 0, "預算", "labor_budget", 12)
        self.entry(form, 13, 2, "決標", "labor_award", 12)
        self.entry(form, 13, 4, "差額保證金", "deposit_difference", 12)
        self.entry(form, 13, 6, "履約保證金", "deposit_performance", 12)
        self.entry(form, 14, 0, "保證金總額(差額+履約)", "deposit_total", 12)

        self.section_title(form, "竣工發包工程費", 15)
        self.entry(form, 16, 0, "竣工發包工程費", "final_contract_amount", 12)
        self.entry(form, 16, 2, "保固金比例", "warranty_rate", 12)
        self.entry(form, 16, 4, "保固保證金", "warranty_deposit", 12)

        for i in range(8):
            form.grid_columnconfigure(i, weight=1)

        bid_box = ttk.LabelFrame(content, text="招標上網日 / 開標日期（可建立多次開標）", padding=8)
        bid_box.pack(fill="both", expand=True, pady=8)
        self.bid_tree = EditableTree(
            bid_box,
            ["round_no", "online_date", "open_date"],
            ["第幾次", "招標上網日", "開標日期"],
            [100, 180, 180],
            self.mark_dirty,
            add_command=self.open_bid_calendar_dialog
        )
        self.bid_tree.pack(fill="both", expand=True)

        ttk.Button(content, text="重新計算預訂竣工日", command=self.recalculate).pack(anchor="e", pady=4)

    def open_bid_calendar_dialog(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return

        base_date = parse_date(self.basic_vars.get("award_date", tk.StringVar()).get()) or date.today()
        win = tk.Toplevel(self)
        win.title("新增招標日期")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)

        year_var = tk.IntVar(value=base_date.year)
        month_var = tk.IntVar(value=base_date.month)
        selected_day_var = tk.StringVar(value="")
        round_var = tk.StringVar(value=str(len(self.bid_tree.get_rows()) + 1))
        date_type_var = tk.StringVar(value="招標上網日")

        top = ttk.Frame(win, padding=10)
        top.pack(fill="x")
        cal_frame = ttk.Frame(win, padding=(10, 0, 10, 10))

        def shift_month(delta):
            year = year_var.get()
            month = month_var.get() + delta
            if month < 1:
                month = 12
                year -= 1
            elif month > 12:
                month = 1
                year += 1
            year_var.set(year)
            month_var.set(month)
            selected_day_var.set("")
            render_calendar()

        ttk.Button(top, text="上一月", command=lambda: shift_month(-1)).pack(side="left", padx=(0, 8))
        ttk.Label(top, text="年").pack(side="left")
        year_spin = ttk.Spinbox(top, from_=base_date.year - 30, to=base_date.year + 30, textvariable=year_var, width=8, command=lambda: render_calendar())
        year_spin.pack(side="left", padx=(2, 8))
        ttk.Label(top, text="月").pack(side="left")
        month_spin = ttk.Spinbox(top, from_=1, to=12, textvariable=month_var, width=5, command=lambda: render_calendar())
        month_spin.pack(side="left", padx=(2, 8))
        ttk.Button(top, text="下一月", command=lambda: shift_month(1)).pack(side="left")
        ttk.Button(top, text="今天", command=lambda: (year_var.set(date.today().year), month_var.set(date.today().month), render_calendar())).pack(side="left", padx=8)

        cal_frame.pack()
        ttk.Label(win, textvariable=selected_day_var, foreground="#1f4e79").pack(anchor="w", padx=12)

        form = ttk.Frame(win, padding=10)
        form.pack(fill="x")
        ttk.Label(form, text="第幾次").grid(row=0, column=0, sticky="e", padx=(0, 6), pady=4)
        ttk.Entry(form, textvariable=round_var, width=10).grid(row=0, column=1, sticky="w", pady=4)
        ttk.Label(form, text="輸入狀況").grid(row=1, column=0, sticky="e", padx=(0, 6), pady=4)
        ttk.Combobox(form, textvariable=date_type_var, values=["招標上網日", "開標日期"], state="readonly", width=16).grid(row=1, column=1, sticky="w", pady=4)

        def select_day(day):
            selected_day_var.set(day.strftime("%Y-%m-%d"))

        def render_calendar():
            for child in cal_frame.winfo_children():
                child.destroy()
            try:
                year = int(year_var.get())
                month = int(month_var.get())
                if month < 1 or month > 12:
                    raise ValueError
            except (tk.TclError, ValueError):
                return
            for cidx, name in enumerate(WEEKDAY_NAMES):
                ttk.Label(cal_frame, text="星期" + name, width=9, anchor="center").grid(row=0, column=cidx, padx=1, pady=1)
            cal = calendar.Calendar(firstweekday=0)
            for ridx, week in enumerate(cal.monthdatescalendar(year, month), start=1):
                for cidx, day in enumerate(week):
                    if day.month != month:
                        ttk.Label(cal_frame, text="", width=9).grid(row=ridx, column=cidx, padx=1, pady=1)
                    else:
                        ttk.Button(cal_frame, text=str(day.day), width=8, command=lambda d=day: select_day(d)).grid(row=ridx, column=cidx, padx=1, pady=1)

        def ok():
            day_text = selected_day_var.get().strip()
            if not day_text:
                messagebox.showwarning("尚未選日期", "請先在日曆表點選招標日期。", parent=win)
                return
            round_text = round_var.get().strip() or str(len(self.bid_tree.get_rows()) + 1)
            if date_type_var.get() == "招標上網日":
                values = [round_text, day_text, ""]
            else:
                values = [round_text, "", day_text]
            self.bid_tree.add_row_after_selection(values)
            selected_day_var.set("")
            self.status_var.set(f"已新增招標資訊：{date_type_var.get()} {day_text}")

        btns = ttk.Frame(win, padding=(10, 0, 10, 10))
        btns.pack(fill="x")
        ttk.Button(btns, text="關閉", command=win.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="增加招標日期", command=ok).pack(side="right", padx=4)
        year_spin.bind("<Return>", lambda e: render_calendar())
        month_spin.bind("<Return>", lambda e: render_calendar())
        year_spin.bind("<FocusOut>", lambda e: render_calendar())
        month_spin.bind("<FocusOut>", lambda e: render_calendar())
        render_calendar()

    def build_holiday_tab(self):
        self.add_page_edit_toggle(self.tab_holiday)
        toolbar = ttk.Frame(self.tab_holiday)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Label(toolbar, text="假期表與補班日表：按「新增一列」用日曆新增。").pack(side="left")
        ttk.Button(toolbar, text="複製前一年度假期", command=self.copy_previous_year_holidays).pack(side="right", padx=4)
        ttk.Button(toolbar, text="確認假期", command=self.confirm_holidays).pack(side="right", padx=4)

        tables = ttk.Frame(self.tab_holiday)
        tables.pack(fill="both", expand=True)
        left = ttk.LabelFrame(tables, text="假期表", padding=6)
        left.pack(side="left", fill="both", expand=True, padx=(0, 4))
        right = ttk.LabelFrame(tables, text="補班日表", padding=6)
        right.pack(side="left", fill="both", expand=True, padx=(4, 0))

        self.holiday_tree = EditableTree(
            left,
            ["exclude", "day", "name"],
            ["排除", "日期", "假日名稱"],
            [45, 160, 160],
            self.mark_dirty,
            add_command=self.open_holiday_calendar_dialog
        )
        self.holiday_tree.pack(fill="both", expand=True, pady=6)
        self.workday_tree = EditableTree(
            right,
            ["exclude", "day", "name"],
            ["排除", "日期", "補班名稱"],
            [45, 160, 160],
            self.mark_dirty,
            add_command=self.open_workday_calendar_dialog
        )
        self.workday_tree.pack(fill="both", expand=True, pady=6)

    def copy_previous_year_holidays(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "請先解除鎖定。")
            return
        year = simpledialog.askinteger("複製前一年度假期", "請輸入要新增的年份：", parent=self, initialvalue=date.today().year)
        if not year:
            return
        prev_year = year - 1
        existing = {r[1] for r in self.holiday_tree.get_rows() if len(r) > 1}
        copied = 0
        for row in self.holiday_tree.get_rows():
            if len(row) < 3:
                continue
            d = parse_date(row[1])
            if not d or d.year != prev_year:
                continue
            try:
                new_day = d.replace(year=year)
            except ValueError:
                continue
            day_text = fmt_date(new_day)
            if day_text in existing:
                continue
            item = self.holiday_tree.tree.insert("", "end", values=["", day_text, row[2]], tags=("red",))
            self.holiday_tree.tree.see(item)
            existing.add(day_text)
            copied += 1
        self.mark_dirty()
        self.status_var.set(f"已複製 {copied} 筆 {year} 年假期，請確認假期")

    def confirm_holidays(self):
        for item in self.holiday_tree.tree.get_children():
            self.holiday_tree.tree.item(item, tags=())
        self.mark_dirty()
        self.status_var.set("假期已確認")

    def apply_year_separators(self, tree):
        last_year = None
        for item in tree.tree.get_children():
            vals = list(tree.tree.item(item, "values"))
            d = parse_date(vals[1] if len(vals) > 1 else "")
            if not d:
                continue
            tags = list(tree.tree.item(item, "tags"))
            if last_year is not None and d.year != last_year and "red" not in tags:
                tags.append("year_sep")
            last_year = d.year
            tree.tree.item(item, tags=tuple(tags))

    def open_holiday_calendar_dialog(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return

        start_date = parse_date(self.basic_vars.get("planned_start", tk.StringVar()).get())
        base_date = start_date or date.today()

        win = tk.Toplevel(self)
        win.title("新增假期")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)

        year_var = tk.IntVar(value=base_date.year)
        month_var = tk.IntVar(value=base_date.month)
        selected_day_var = tk.StringVar(value="")
        holiday_name_var = tk.StringVar(value="")

        top = ttk.Frame(win, padding=10)
        top.pack(fill="x")

        def shift_month(delta):
            year = year_var.get()
            month = month_var.get() + delta
            if month < 1:
                month = 12
                year -= 1
            elif month > 12:
                month = 1
                year += 1
            year_var.set(year)
            month_var.set(month)
            selected_day_var.set("")
            render_calendar()

        ttk.Button(top, text="上一月", command=lambda: shift_month(-1)).pack(side="left", padx=(0, 8))
        ttk.Label(top, text="年").pack(side="left")
        year_spin = ttk.Spinbox(top, from_=base_date.year - 30, to=base_date.year + 30, textvariable=year_var, width=8, command=lambda: render_calendar())
        year_spin.pack(side="left", padx=(2, 8))
        ttk.Label(top, text="月").pack(side="left")
        month_spin = ttk.Spinbox(top, from_=1, to=12, textvariable=month_var, width=5, command=lambda: render_calendar())
        month_spin.pack(side="left", padx=(2, 8))
        ttk.Button(top, text="下一月", command=lambda: shift_month(1)).pack(side="left")
        ttk.Button(top, text="今天", command=lambda: (year_var.set(date.today().year), month_var.set(date.today().month), render_calendar())).pack(side="left", padx=8)

        cal_frame = ttk.Frame(win, padding=(10, 0, 10, 10))
        cal_frame.pack()
        ttk.Label(win, textvariable=selected_day_var, foreground="#1f4e79").pack(anchor="w", padx=12)

        form = ttk.Frame(win, padding=10)
        form.pack(fill="x")
        ttk.Label(form, text="假期名稱").grid(row=0, column=0, sticky="e", padx=(0, 6), pady=4)
        name_entry = ttk.Entry(form, textvariable=holiday_name_var, width=34)
        name_entry.grid(row=0, column=1, sticky="ew", pady=4)
        form.grid_columnconfigure(1, weight=1)

        def select_day(day):
            selected_day_var.set(day.strftime("%Y-%m-%d"))
            name_entry.focus_set()

        def render_calendar():
            for child in cal_frame.winfo_children():
                child.destroy()

            try:
                year = int(year_var.get())
                month = int(month_var.get())
                if month < 1 or month > 12:
                    raise ValueError
            except (tk.TclError, ValueError):
                return

            for col, name in enumerate(WEEKDAY_NAMES):
                ttk.Label(cal_frame, text="星期" + name, width=9, anchor="center").grid(row=0, column=col, padx=1, pady=1)

            cal = calendar.Calendar(firstweekday=0)
            for row, week in enumerate(cal.monthdatescalendar(year, month), start=1):
                for col, day in enumerate(week):
                    if day.month != month:
                        ttk.Label(cal_frame, text="", width=9).grid(row=row, column=col, padx=1, pady=1)
                        continue
                    btn = ttk.Button(cal_frame, text=str(day.day), width=8, command=lambda d=day: select_day(d))
                    btn.grid(row=row, column=col, padx=1, pady=1)

        def ok():
            day_text = selected_day_var.get().strip()
            if not day_text:
                messagebox.showwarning("尚未選日期", "請先在日曆表點選假期日期。", parent=win)
                return
            name_text = holiday_name_var.get().strip()
            if not name_text:
                messagebox.showwarning("尚未輸入假期名稱", "請在日曆表下方輸入假期名稱。", parent=win)
                return
            self.holiday_tree.add_row(["", day_text, name_text])
            holiday_name_var.set("")
            selected_day_var.set("")
            self.status_var.set(f"已新增假期：{day_text} {name_text}")
            name_entry.focus_set()

        btns = ttk.Frame(win, padding=(10, 0, 10, 10))
        btns.pack(fill="x")
        ttk.Button(btns, text="關閉", command=win.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="增加假期", command=ok).pack(side="right", padx=4)

        year_spin.bind("<Return>", lambda e: render_calendar())
        month_spin.bind("<Return>", lambda e: render_calendar())
        year_spin.bind("<FocusOut>", lambda e: render_calendar())
        month_spin.bind("<FocusOut>", lambda e: render_calendar())
        render_calendar()
        name_entry.focus_set()

    def open_workday_calendar_dialog(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "請先解除鎖定。")
            return
        base_date = parse_date(self.basic_vars.get("planned_start", tk.StringVar()).get()) or date.today()
        win = tk.Toplevel(self)
        win.title("新增補班日")
        win.transient(self)
        win.grab_set()
        year_var = tk.IntVar(value=base_date.year)
        month_var = tk.IntVar(value=base_date.month)
        selected_day_var = tk.StringVar(value="")
        name_var = tk.StringVar(value="")
        top = ttk.Frame(win, padding=10)
        top.pack(fill="x")
        cal_frame = ttk.Frame(win, padding=10)
        cal_frame.pack()
        ttk.Label(win, textvariable=selected_day_var, foreground="#1f4e79").pack(anchor="w", padx=12)
        form = ttk.Frame(win, padding=10)
        form.pack(fill="x")
        ttk.Label(form, text="補班名稱").grid(row=0, column=0, sticky="e", padx=(0, 6))
        ttk.Entry(form, textvariable=name_var, width=32).grid(row=0, column=1, sticky="ew")

        def render():
            for child in cal_frame.winfo_children():
                child.destroy()
            y, m = int(year_var.get()), int(month_var.get())
            for c, name in enumerate(WEEKDAY_NAMES):
                ttk.Label(cal_frame, text="星期" + name, width=9, anchor="center").grid(row=0, column=c)
            for r, week in enumerate(calendar.Calendar(firstweekday=0).monthdatescalendar(y, m), start=1):
                for c, day in enumerate(week):
                    if day.month != m:
                        ttk.Label(cal_frame, text="", width=9).grid(row=r, column=c)
                    else:
                        ttk.Button(cal_frame, text=str(day.day), width=8, command=lambda d=day: selected_day_var.set(fmt_date(d))).grid(row=r, column=c)

        def shift(delta):
            y, m = year_var.get(), month_var.get() + delta
            if m < 1:
                y, m = y - 1, 12
            elif m > 12:
                y, m = y + 1, 1
            year_var.set(y)
            month_var.set(m)
            render()

        ttk.Button(top, text="上一月", command=lambda: shift(-1)).pack(side="left")
        ttk.Spinbox(top, from_=base_date.year - 30, to=base_date.year + 30, textvariable=year_var, width=8, command=render).pack(side="left", padx=4)
        ttk.Spinbox(top, from_=1, to=12, textvariable=month_var, width=5, command=render).pack(side="left", padx=4)
        ttk.Button(top, text="下一月", command=lambda: shift(1)).pack(side="left")
        ttk.Button(top, text="今天", command=lambda: (year_var.set(date.today().year), month_var.set(date.today().month), render())).pack(side="left", padx=8)

        def ok():
            day_text = selected_day_var.get().strip()
            if not day_text:
                messagebox.showwarning("尚未選日期", "請先點選補班日期。", parent=win)
                return
            name = name_var.get().strip() or "補班"
            self.workday_tree.add_row(["", day_text, name])
            selected_day_var.set("")
            name_var.set("")

        btns = ttk.Frame(win, padding=10)
        btns.pack(fill="x")
        ttk.Button(btns, text="關閉", command=win.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="增加補班日", command=ok).pack(side="right", padx=4)
        render()

    def build_weather_tab(self):
        self.add_page_edit_toggle(self.tab_weather)
        ttk.Label(
            self.tab_weather,
            text="晴雨表：按「新增一列」會開啟年、月日曆表；下方可分別輸入上午、下午、天氣、場地與備註。"
        ).pack(anchor="w")
        self.weather_tree = EditableTree(
            self.tab_weather,
            ["day", "morning", "afternoon", "typhoon", "site", "note"],
            ["日期", "上午", "下午", "天氣", "場地", "備註"],
            [130, 130, 130, 130, 130, 130],
            self.mark_dirty,
            add_command=self.open_weather_calendar_dialog
        )
        self.weather_tree.pack(fill="both", expand=True, pady=6)

    def open_weather_calendar_dialog(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return

        start_date = parse_date(self.basic_vars.get("planned_start", tk.StringVar()).get())
        base_date = start_date or date.today()

        win = tk.Toplevel(self)
        win.title("新增晴雨紀錄")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)

        year_var = tk.IntVar(value=base_date.year)
        month_var = tk.IntVar(value=base_date.month)
        selected_day_var = tk.StringVar(value="")
        morning_var = tk.StringVar(value="")
        afternoon_var = tk.StringVar(value="")
        typhoon_var = tk.StringVar(value="")
        site_var = tk.StringVar(value="")
        note_var = tk.StringVar(value="")

        top = ttk.Frame(win, padding=10)
        top.pack(fill="x")

        def shift_month(delta):
            year = year_var.get()
            month = month_var.get() + delta
            if month < 1:
                month = 12
                year -= 1
            elif month > 12:
                month = 1
                year += 1
            year_var.set(year)
            month_var.set(month)
            selected_day_var.set("")
            render_calendar()

        ttk.Button(top, text="上一月", command=lambda: shift_month(-1)).pack(side="left", padx=(0, 8))
        ttk.Label(top, text="年").pack(side="left")
        year_spin = ttk.Spinbox(top, from_=base_date.year - 30, to=base_date.year + 30, textvariable=year_var, width=8, command=lambda: render_calendar())
        year_spin.pack(side="left", padx=(2, 8))
        ttk.Label(top, text="月").pack(side="left")
        month_spin = ttk.Spinbox(top, from_=1, to=12, textvariable=month_var, width=5, command=lambda: render_calendar())
        month_spin.pack(side="left", padx=(2, 8))
        ttk.Button(top, text="下一月", command=lambda: shift_month(1)).pack(side="left")
        ttk.Button(top, text="今天", command=lambda: (year_var.set(date.today().year), month_var.set(date.today().month), render_calendar())).pack(side="left", padx=8)

        cal_frame = ttk.Frame(win, padding=(10, 0, 10, 10))
        cal_frame.pack()
        ttk.Label(win, textvariable=selected_day_var, foreground="#1f4e79").pack(anchor="w", padx=12)

        form = ttk.Frame(win, padding=10)
        form.pack(fill="x")
        fields = [
            ("上午", morning_var),
            ("下午", afternoon_var),
            ("天氣", typhoon_var),
            ("場地", site_var),
            ("備註", note_var),
        ]
        entries = []
        for i, (label, var) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=i, column=0, sticky="e", padx=(0, 6), pady=4)
            entry_width = 34 if label == "備註" else 12
            ent = ttk.Entry(form, textvariable=var, width=entry_width)
            ent.grid(row=i, column=1, sticky="ew", pady=4)
            entries.append(ent)
        form.grid_columnconfigure(1, weight=1)

        def morning_is_full_day():
            try:
                return float(morning_var.get().strip() or 0) >= 1.0
            except ValueError:
                return False

        def update_afternoon_state(*args):
            if morning_is_full_day():
                afternoon_var.set("")
                entries[1].configure(state="disabled")
            else:
                entries[1].configure(state="normal")

        morning_var.trace_add("write", update_afternoon_state)
        update_afternoon_state()

        def select_day(day):
            selected_day_var.set(day.strftime("%Y-%m-%d"))
            entries[0].focus_set()

        def render_calendar():
            for child in cal_frame.winfo_children():
                child.destroy()

            try:
                year = int(year_var.get())
                month = int(month_var.get())
                if month < 1 or month > 12:
                    raise ValueError
            except (tk.TclError, ValueError):
                return

            for col, name in enumerate(WEEKDAY_NAMES):
                ttk.Label(cal_frame, text="星期" + name, width=9, anchor="center").grid(row=0, column=col, padx=1, pady=1)

            cal = calendar.Calendar(firstweekday=0)
            for row, week in enumerate(cal.monthdatescalendar(year, month), start=1):
                for col, day in enumerate(week):
                    if day.month != month:
                        ttk.Label(cal_frame, text="", width=9).grid(row=row, column=col, padx=1, pady=1)
                        continue
                    btn = ttk.Button(cal_frame, text=str(day.day), width=8, command=lambda d=day: select_day(d))
                    btn.grid(row=row, column=col, padx=1, pady=1)

        def ok():
            day_text = selected_day_var.get().strip()
            if not day_text:
                messagebox.showwarning("尚未選日期", "請先在日曆表點選晴雨紀錄日期。", parent=win)
                return
            values = [
                day_text,
                morning_var.get().strip(),
                "" if morning_is_full_day() else afternoon_var.get().strip(),
                typhoon_var.get().strip(),
                site_var.get().strip(),
                note_var.get().strip(),
            ]
            self.weather_tree.add_row_after_selection(values)
            selected_day_var.set("")
            for var in (morning_var, afternoon_var, typhoon_var, site_var, note_var):
                var.set("")
            self.status_var.set(f"已新增晴雨紀錄：{day_text}")
            entries[0].focus_set()

        btns = ttk.Frame(win, padding=(10, 0, 10, 10))
        btns.pack(fill="x")
        ttk.Button(btns, text="關閉", command=win.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="增加晴雨紀錄", command=ok).pack(side="right", padx=4)

        year_spin.bind("<Return>", lambda e: render_calendar())
        month_spin.bind("<Return>", lambda e: render_calendar())
        year_spin.bind("<FocusOut>", lambda e: render_calendar())
        month_spin.bind("<FocusOut>", lambda e: render_calendar())
        render_calendar()
        entries[0].focus_set()

    def build_railway_tab(self):
        self.add_page_edit_toggle(self.tab_railway)
        top = ttk.Frame(self.tab_railway)
        top.pack(fill="x")
        ttk.Label(top, text="鐵路疏運停工日期：可讀入第二分頁假期表；新增時會插在目前選取列下方。").pack(side="left")
        ttk.Button(top, text="讀入第二分頁假期表", command=self.import_holidays_to_railway).pack(side="right")
        self.railway_tree = EditableTree(
            self.tab_railway,
            ["exclude", "day", "note"],
            ["排除", "日期", "疏運名稱"],
            [45, 160, 160],
            self.mark_dirty,
            add_command=self.open_railway_calendar_dialog
        )
        self.railway_tree.pack(fill="both", expand=True, pady=6)

    def import_holidays_to_railway(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        existing_days = {r[1] for r in self.railway_tree.get_rows() if len(r) > 1 and r[1]}
        added = 0
        for row in self.holiday_tree.get_rows():
            if len(row) < 3:
                continue
            if row[0] == "✓":
                continue
            day_text = (row[1] or "").strip()
            name_text = (row[2] or "").strip()
            if not parse_date(day_text) or day_text in existing_days:
                continue
            self.railway_tree.add_row(["", day_text, name_text])
            existing_days.add(day_text)
            added += 1
        self.status_var.set(f"已從假期表讀入 {added} 筆資料")

    def open_railway_calendar_dialog(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return

        start_date = parse_date(self.basic_vars.get("planned_start", tk.StringVar()).get())
        base_date = start_date or date.today()

        win = tk.Toplevel(self)
        win.title("新增鐵路疏運停工日")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)

        year_var = tk.IntVar(value=base_date.year)
        month_var = tk.IntVar(value=base_date.month)
        selected_day_var = tk.StringVar(value="")
        note_var = tk.StringVar(value="")

        top = ttk.Frame(win, padding=10)
        top.pack(fill="x")

        def shift_month(delta):
            year = year_var.get()
            month = month_var.get() + delta
            if month < 1:
                month = 12
                year -= 1
            elif month > 12:
                month = 1
                year += 1
            year_var.set(year)
            month_var.set(month)
            selected_day_var.set("")
            render_calendar()

        ttk.Button(top, text="上一月", command=lambda: shift_month(-1)).pack(side="left", padx=(0, 8))
        ttk.Label(top, text="年").pack(side="left")
        year_spin = ttk.Spinbox(top, from_=base_date.year - 30, to=base_date.year + 30, textvariable=year_var, width=8, command=lambda: render_calendar())
        year_spin.pack(side="left", padx=(2, 8))
        ttk.Label(top, text="月").pack(side="left")
        month_spin = ttk.Spinbox(top, from_=1, to=12, textvariable=month_var, width=5, command=lambda: render_calendar())
        month_spin.pack(side="left", padx=(2, 8))
        ttk.Button(top, text="下一月", command=lambda: shift_month(1)).pack(side="left")
        ttk.Button(top, text="今天", command=lambda: (year_var.set(date.today().year), month_var.set(date.today().month), render_calendar())).pack(side="left", padx=8)

        cal_frame = ttk.Frame(win, padding=(10, 0, 10, 10))
        cal_frame.pack()
        ttk.Label(win, textvariable=selected_day_var, foreground="#1f4e79").pack(anchor="w", padx=12)

        form = ttk.Frame(win, padding=10)
        form.pack(fill="x")
        ttk.Label(form, text="疏運名稱").grid(row=0, column=0, sticky="e", padx=(0, 6), pady=4)
        note_entry = ttk.Entry(form, textvariable=note_var, width=34)
        note_entry.grid(row=0, column=1, sticky="ew", pady=4)
        form.grid_columnconfigure(1, weight=1)

        def select_day(day):
            selected_day_var.set(day.strftime("%Y-%m-%d"))
            note_entry.focus_set()

        def render_calendar():
            for child in cal_frame.winfo_children():
                child.destroy()

            try:
                year = int(year_var.get())
                month = int(month_var.get())
                if month < 1 or month > 12:
                    raise ValueError
            except (tk.TclError, ValueError):
                return

            for col, name in enumerate(WEEKDAY_NAMES):
                ttk.Label(cal_frame, text="星期" + name, width=9, anchor="center").grid(row=0, column=col, padx=1, pady=1)

            cal = calendar.Calendar(firstweekday=0)
            for row, week in enumerate(cal.monthdatescalendar(year, month), start=1):
                for col, day in enumerate(week):
                    if day.month != month:
                        ttk.Label(cal_frame, text="", width=9).grid(row=row, column=col, padx=1, pady=1)
                        continue
                    btn = ttk.Button(cal_frame, text=str(day.day), width=8, command=lambda d=day: select_day(d))
                    btn.grid(row=row, column=col, padx=1, pady=1)

        def ok():
            day_text = selected_day_var.get().strip()
            if not day_text:
                messagebox.showwarning("尚未選日期", "請先在日曆表點選疏運停工日期。", parent=win)
                return
            note_text = note_var.get().strip() or "疏運"
            self.railway_tree.add_row_after_selection(["", day_text, note_text])
            note_var.set("")
            selected_day_var.set("")
            self.status_var.set(f"已新增疏運停工日：{day_text} {note_text}")
            note_entry.focus_set()

        btns = ttk.Frame(win, padding=(10, 0, 10, 10))
        btns.pack(fill="x")
        ttk.Button(btns, text="關閉", command=win.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="增加疏運日", command=ok).pack(side="right", padx=4)

        year_spin.bind("<Return>", lambda e: render_calendar())
        month_spin.bind("<Return>", lambda e: render_calendar())
        year_spin.bind("<FocusOut>", lambda e: render_calendar())
        month_spin.bind("<FocusOut>", lambda e: render_calendar())
        render_calendar()
        note_entry.focus_set()

    def build_calendar_tab(self):
        self.add_page_edit_toggle(self.tab_calendar)
        top = ttk.Frame(self.tab_calendar)
        top.pack(fill="x")
        ttk.Label(top, text="顯示月份：").pack(side="left")
        self.cal_month_var = tk.StringVar(value=date.today().strftime("%Y-%m"))
        ttk.Entry(top, textvariable=self.cal_month_var, width=12).pack(side="left", padx=4)
        ttk.Button(top, text="產生施工日曆", command=self.render_calendar).pack(side="left", padx=4)
        ttk.Button(top, text="上個月", command=lambda: self.shift_month(-1)).pack(side="left", padx=4)
        ttk.Button(top, text="下個月", command=lambda: self.shift_month(1)).pack(side="left", padx=4)

        cal_area = ttk.Frame(self.tab_calendar)
        cal_area.pack(fill="both", expand=True, pady=8)
        self.cal_canvas = tk.Canvas(cal_area, background="white")
        self.cal_canvas.grid(row=0, column=0, sticky="nsew")
        cal_area.grid_rowconfigure(0, weight=1)
        cal_area.grid_columnconfigure(0, weight=1)
        self.cal_canvas.bind("<Configure>", lambda e: self.render_calendar())
        self.cal_canvas.bind("<MouseWheel>", lambda e: self.shift_month(-1 if e.delta > 0 else 1))

    def build_payment_tabs(self):
        self.add_page_edit_toggle(self.tab_payment_contract)
        self.payment_contract_tree = self.make_payment_tree(
            self.tab_payment_contract,
            "發包工程費計價：先建立資料欄位，後續再依需求增加計價公式與報表。"
        )
        self.add_page_edit_toggle(self.tab_payment_other)
        self.payment_other_tree = self.make_payment_tree(
            self.tab_payment_other,
            "發包以外計價：先建立資料欄位，後續再依需求增加分類與核銷流程。"
        )
        self.add_page_edit_toggle(self.tab_payment_admin)
        self.payment_admin_tree = self.make_payment_tree(
            self.tab_payment_admin,
            "管理費計價：先建立資料欄位，後續再依需求增加管理費計算規則。"
        )

    def make_payment_tree(self, parent, hint):
        ttk.Label(parent, text=hint).pack(anchor="w")
        tree = EditableTree(
            parent,
            ["day", "item", "voucher_no", "amount", "note"],
            ["日期", "項目", "核銷/憑證號", "核銷金額", "備註"],
            [120, 260, 160, 140, 320],
            self.mark_dirty
        )
        tree.pack(fill="both", expand=True, pady=6)
        return tree

    def build_execution_tab(self):
        self.add_page_edit_toggle(self.tab_execution)
        ttk.Label(self.tab_execution, text="工程執行紀錄表：按「新增一列」用行事曆新增；可點選標題列排序。").pack(anchor="w")
        self.execution_tree = EditableTree(
            self.tab_execution,
            ["day", "record_type", "content", "note"],
            ["日期", "資料類型", "內容", "備註"],
            [120, 160, 520, 260],
            self.mark_dirty,
            add_command=self.open_execution_calendar_dialog
        )
        self.execution_tree.pack(fill="both", expand=True, pady=6)

    def open_execution_calendar_dialog(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return

        base_date = parse_date(self.basic_vars.get("planned_start", tk.StringVar()).get()) or date.today()
        win = tk.Toplevel(self)
        win.title("新增工程執行紀錄")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)

        year_var = tk.IntVar(value=base_date.year)
        month_var = tk.IntVar(value=base_date.month)
        selected_day_var = tk.StringVar(value="")
        type_var = tk.StringVar(value="工作會議")
        note_var = tk.StringVar(value="")

        top = ttk.Frame(win, padding=10)
        top.pack(fill="x")
        cal_frame = ttk.Frame(win, padding=(10, 0, 10, 10))

        def shift_month(delta):
            year = year_var.get()
            month = month_var.get() + delta
            if month < 1:
                month = 12
                year -= 1
            elif month > 12:
                month = 1
                year += 1
            year_var.set(year)
            month_var.set(month)
            selected_day_var.set("")
            render_calendar()

        ttk.Button(top, text="上一月", command=lambda: shift_month(-1)).pack(side="left", padx=(0, 8))
        ttk.Label(top, text="年").pack(side="left")
        year_spin = ttk.Spinbox(top, from_=base_date.year - 30, to=base_date.year + 30, textvariable=year_var, width=8, command=lambda: render_calendar())
        year_spin.pack(side="left", padx=(2, 8))
        ttk.Label(top, text="月").pack(side="left")
        month_spin = ttk.Spinbox(top, from_=1, to=12, textvariable=month_var, width=5, command=lambda: render_calendar())
        month_spin.pack(side="left", padx=(2, 8))
        ttk.Button(top, text="下一月", command=lambda: shift_month(1)).pack(side="left")
        ttk.Button(top, text="今天", command=lambda: (year_var.set(date.today().year), month_var.set(date.today().month), render_calendar())).pack(side="left", padx=8)
        cal_frame.pack()
        ttk.Label(win, textvariable=selected_day_var, foreground="#1f4e79").pack(anchor="w", padx=12)

        form = ttk.Frame(win, padding=10)
        form.pack(fill="x")
        ttk.Label(form, text="資料類型").grid(row=0, column=0, sticky="e", padx=(0, 6), pady=4)
        ttk.Combobox(
            form,
            textvariable=type_var,
            values=["工作會議", "會勘", "變更需求會議", "變更確認會議", "其他"],
            state="readonly",
            width=20
        ).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(form, text="內容").grid(row=1, column=0, sticky="ne", padx=(0, 6), pady=4)
        content_text = tk.Text(form, height=5, width=48, wrap="word", font=("Microsoft JhengHei UI", 10))
        content_text.grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Label(form, text="備註").grid(row=2, column=0, sticky="e", padx=(0, 6), pady=4)
        note_entry = ttk.Entry(form, textvariable=note_var, width=48)
        note_entry.grid(row=2, column=1, sticky="ew", pady=4)
        form.grid_columnconfigure(1, weight=1)

        def select_day(day):
            selected_day_var.set(day.strftime("%Y-%m-%d"))
            content_text.focus_set()

        def render_calendar():
            for child in cal_frame.winfo_children():
                child.destroy()
            try:
                year = int(year_var.get())
                month = int(month_var.get())
                if month < 1 or month > 12:
                    raise ValueError
            except (tk.TclError, ValueError):
                return
            for col, name in enumerate(WEEKDAY_NAMES):
                ttk.Label(cal_frame, text="星期" + name, width=9, anchor="center").grid(row=0, column=col, padx=1, pady=1)
            cal = calendar.Calendar(firstweekday=0)
            for row, week in enumerate(cal.monthdatescalendar(year, month), start=1):
                for col, day in enumerate(week):
                    if day.month != month:
                        ttk.Label(cal_frame, text="", width=9).grid(row=row, column=col, padx=1, pady=1)
                        continue
                    ttk.Button(cal_frame, text=str(day.day), width=8, command=lambda d=day: select_day(d)).grid(row=row, column=col, padx=1, pady=1)

        def ok():
            day_text = selected_day_var.get().strip()
            if not day_text:
                messagebox.showwarning("尚未選日期", "請先在日曆表點選日期。", parent=win)
                return
            content = content_text.get("1.0", "end-1c").strip()
            if not content:
                messagebox.showwarning("尚未輸入內容", "請輸入內容。", parent=win)
                return
            self.execution_tree.add_row_after_selection([day_text, type_var.get(), content, note_var.get().strip()])
            selected_day_var.set("")
            content_text.delete("1.0", "end")
            note_var.set("")
            content_text.focus_set()

        btns = ttk.Frame(win, padding=(10, 0, 10, 10))
        btns.pack(fill="x")
        ttk.Button(btns, text="關閉", command=win.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="增加資料", command=ok).pack(side="right", padx=4)
        render_calendar()

    def build_milestone_tab(self):
        self.add_page_edit_toggle(self.tab_milestone)
        ttk.Label(self.tab_milestone, text="工程大事記：起算日期與收文日期用行事曆點選；收文日期空白以粉紅色標示。").pack(anchor="w")
        self.milestone_tree = EditableTree(
            self.tab_milestone,
            ["item_no", "contract_item", "start_date", "deadline_days", "deadline_date", "received_date", "overdue", "received_no", "note", "day_adjust"],
            ["項次", "履約項目", "起算日期", "履約期限日數", "履約期限", "收文日期", "逾期", "收文文號", "註記", "日數調整"],
            [70, 200, 120, 120, 120, 120, 80, 160, 240, 100],
            self.mark_dirty,
            add_command=self.open_milestone_dialog
        )
        self.milestone_tree.pack(fill="both", expand=True, pady=6)

    def calc_milestone_row(self, values):
        vals = (values + [""] * 10)[:10]
        start = parse_date(vals[2])
        try:
            deadline_days = int(float(vals[3] or 0))
        except ValueError:
            deadline_days = 0
        try:
            adjust = int(float(vals[9] or 0))
        except ValueError:
            adjust = 0
        if start and deadline_days:
            vals[4] = fmt_date(start + timedelta(days=deadline_days + adjust))
        received = parse_date(vals[5])
        deadline = parse_date(vals[4])
        vals[6] = "逾期" if received and deadline and received > deadline else ""
        return vals

    def refresh_milestone_rows(self):
        if not hasattr(self, "milestone_tree"):
            return
        rows = [self.calc_milestone_row(r) for r in self.milestone_tree.get_rows()]
        self.milestone_tree.set_rows(rows)
        self.milestone_tree.apply_row_tags(lambda r: "pink" if len(r) > 5 and not r[5] else "")

    def pick_date_for_var(self, target_var, parent):
        base = parse_date(target_var.get()) or date.today()
        win = tk.Toplevel(parent)
        win.title("選擇日期")
        win.transient(parent)
        win.grab_set()
        year_var = tk.IntVar(value=base.year)
        month_var = tk.IntVar(value=base.month)
        top = ttk.Frame(win, padding=10)
        top.pack(fill="x")
        cal_frame = ttk.Frame(win, padding=10)
        cal_frame.pack()

        def render():
            for child in cal_frame.winfo_children():
                child.destroy()
            y, m = int(year_var.get()), int(month_var.get())
            for c, name in enumerate(WEEKDAY_NAMES):
                ttk.Label(cal_frame, text="星期" + name, width=9, anchor="center").grid(row=0, column=c)
            for r, week in enumerate(calendar.Calendar(firstweekday=0).monthdatescalendar(y, m), start=1):
                for c, day in enumerate(week):
                    if day.month != m:
                        ttk.Label(cal_frame, text="", width=9).grid(row=r, column=c)
                    else:
                        ttk.Button(cal_frame, text=str(day.day), width=8, command=lambda d=day: (target_var.set(fmt_date(d)), win.destroy())).grid(row=r, column=c)

        def shift(delta):
            y, m = year_var.get(), month_var.get() + delta
            if m < 1:
                y, m = y - 1, 12
            elif m > 12:
                y, m = y + 1, 1
            year_var.set(y)
            month_var.set(m)
            render()

        ttk.Button(top, text="上一月", command=lambda: shift(-1)).pack(side="left")
        ttk.Spinbox(top, from_=base.year - 30, to=base.year + 30, textvariable=year_var, width=8, command=render).pack(side="left", padx=4)
        ttk.Spinbox(top, from_=1, to=12, textvariable=month_var, width=5, command=render).pack(side="left", padx=4)
        ttk.Button(top, text="下一月", command=lambda: shift(1)).pack(side="left")
        ttk.Button(top, text="今天", command=lambda: (year_var.set(date.today().year), month_var.set(date.today().month), render())).pack(side="left", padx=8)
        render()

    def open_milestone_dialog(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先解除鎖定。")
            return
        win = tk.Toplevel(self)
        win.title("新增工程大事記")
        win.transient(self)
        win.grab_set()
        vars_ = {k: tk.StringVar() for k in ["item_no", "contract_item", "start_date", "deadline_days", "received_date", "received_no", "note", "day_adjust"]}
        labels = [("項次", "item_no"), ("履約項目", "contract_item"), ("起算日期", "start_date"), ("履約期限日數", "deadline_days"), ("收文日期", "received_date"), ("收文文號", "received_no"), ("註記", "note"), ("日數調整", "day_adjust")]
        form = ttk.Frame(win, padding=10)
        form.pack(fill="x")
        for i, (label, key) in enumerate(labels):
            ttk.Label(form, text=label).grid(row=i, column=0, sticky="e", padx=4, pady=3)
            ent = ttk.Entry(form, textvariable=vars_[key], width=32)
            ent.grid(row=i, column=1, sticky="ew", padx=4, pady=3)
            if key in ("start_date", "received_date"):
                ttk.Button(form, text="▼", width=3, command=lambda v=vars_[key]: self.pick_date_for_var(v, win)).grid(row=i, column=2, padx=2)
        form.grid_columnconfigure(1, weight=1)

        def ok():
            row = self.calc_milestone_row([
                vars_["item_no"].get(), vars_["contract_item"].get(), vars_["start_date"].get(), vars_["deadline_days"].get(),
                "", vars_["received_date"].get(), "", vars_["received_no"].get(), vars_["note"].get(), vars_["day_adjust"].get()
            ])
            self.milestone_tree.add_row_after_selection(row)
            self.refresh_milestone_rows()
            win.destroy()

        ttk.Button(win, text="新增", command=ok).pack(side="right", padx=10, pady=10)
        ttk.Button(win, text="取消", command=win.destroy).pack(side="right", pady=10)

    def build_status_tab(self):
        box = ttk.LabelFrame(self.tab_status, text="工程執行狀態", padding=12)
        box.pack(fill="x")
        ttk.Label(box, text="目前狀態：").grid(row=0, column=0, sticky="e", padx=6, pady=8)
        self.execution_status_var = tk.StringVar(value="規劃中")
        self.execution_status_var.trace_add("write", self.mark_dirty)
        self.execution_status_combo = ttk.Combobox(
            box,
            textvariable=self.execution_status_var,
            values=[
                "規劃中", "招標中", "決標完成", "開工中", "施工中",
                "停工中", "復工中", "竣工中", "已竣工", "驗收中",
                "驗收完成", "結案"
            ],
            width=28
        )
        self.execution_status_combo.grid(row=0, column=1, sticky="w", padx=6, pady=8)
        self.edit_widgets.append(self.execution_status_combo)
        ttk.Label(
            box,
            text="此欄位會即時顯示在上半部工程基本資料顯示區。後續可再擴充狀態日期、狀態說明與歷程紀錄。"
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=6, pady=8)
        box.grid_columnconfigure(1, weight=1)


    def build_settings_tab(self):
        box = ttk.LabelFrame(self.tab_settings, text="1、密碼設置和修改", padding=12)
        box.pack(fill="x")
        ttk.Label(box, text="新密碼：").grid(row=0, column=0, sticky="e", padx=6, pady=8)
        self.new_password_var = tk.StringVar()
        ttk.Entry(box, textvariable=self.new_password_var, show="*", width=24).grid(row=0, column=1, sticky="w", padx=6, pady=8)
        ttk.Label(box, text="確認新密碼：").grid(row=1, column=0, sticky="e", padx=6, pady=8)
        self.confirm_password_var = tk.StringVar()
        ttk.Entry(box, textvariable=self.confirm_password_var, show="*", width=24).grid(row=1, column=1, sticky="w", padx=6, pady=8)
        ttk.Button(box, text="儲存 / 修改密碼", command=self.set_project_password).grid(row=2, column=0, columnspan=2, sticky="w", padx=6, pady=8)
        ttk.Label(
            box,
            text="密碼至少 3 字元；使用 SHA256，公鑰/salt：1981。密碼雜湊會存在每一個工程資料裡。未設定密碼時不鎖定編輯。"
        ).grid(row=3, column=0, columnspan=2, sticky="w", padx=6, pady=8)

    def assign_tree_edit_guards(self):
        for name in [
            "bid_tree", "holiday_tree", "workday_tree", "weather_tree", "railway_tree",
            "payment_contract_tree", "payment_other_tree", "payment_admin_tree",
            "execution_tree", "milestone_tree"
        ]:
            if hasattr(self, name):
                getattr(self, name).can_edit = self.can_edit

    def apply_edit_lock_state(self):
        unlocked = self.can_edit()
        for w in getattr(self, "edit_widgets", []):
            try:
                w.configure(state="normal" if unlocked else "disabled")
                if isinstance(w, tk.Text):
                    w.configure(foreground="black" if unlocked else "#1f4e79")
            except tk.TclError:
                pass
        self.assign_tree_edit_guards()
        manual_var = getattr(self, "data_edit_enabled_var", None)
        if manual_var is not None and not manual_var.get():
            self.lock_state_var.set("資料編輯鎖定")
        elif not self.project_password_hash:
            self.lock_state_var.set("未設定密碼")
        elif unlocked:
            self.lock_state_var.set("已解鎖")
        else:
            self.lock_state_var.set("已鎖定")

    def unlock_project(self):
        if not self.current_project_id:
            return
        if not self.project_password_hash:
            messagebox.showwarning("尚未鎖定", "此工程尚未設定編輯密碼。請先輸入密碼並按「鎖定」。")
            return
        pwd = self.edit_password_var.get()
        if len(pwd) < 3:
            messagebox.showwarning("密碼長度不足", "編輯密碼至少需要 3 個字元。")
            return
        if hash_password(pwd) == self.project_password_hash:
            self.edit_unlocked = True
            self.apply_edit_lock_state()
            self.status_var.set("密碼正確，已解鎖編輯")
        else:
            self.edit_unlocked = False
            self.apply_edit_lock_state()
            messagebox.showerror("密碼錯誤", "編輯密碼不正確。")

    def lock_project(self):
        if not self.current_project_id:
            return
        pwd = self.edit_password_var.get()
        if len(pwd) < 3:
            messagebox.showwarning("密碼長度不足", "請先輸入至少 3 個字元的編輯密碼，再按「鎖定」。")
            return
        self.save_current()
        self.db.save_password_hash(self.current_project_id, hash_password(pwd))
        self.project_password_hash = self.db.get_password_hash(self.current_project_id)
        self.edit_unlocked = False
        self.edit_password_var.set("")
        self.apply_edit_lock_state()
        self.status_var.set("已使用 SHA256 儲存密碼並鎖定編輯")

    def clear_project_password(self):
        if not self.current_project_id:
            return
        if self.project_password_hash and not self.edit_unlocked:
            messagebox.showwarning("尚未解鎖", "請先輸入正確密碼並解鎖後，才能取消密碼。")
            return
        if not self.project_password_hash:
            self.status_var.set("此工程目前未設定密碼")
            return
        if not messagebox.askyesno("取消密碼", "確定要取消此工程的編輯密碼嗎？取消後重新開啟也不會自動鎖定。"):
            return
        self.db.save_password_hash(self.current_project_id, "")
        self.project_password_hash = ""
        self.edit_unlocked = True
        self.edit_password_var.set("")
        self.apply_edit_lock_state()
        self.status_var.set("已取消密碼，不再自動鎖定")

    def set_project_password(self):
        if not self.current_project_id:
            return
        if self.project_password_hash and not self.edit_unlocked:
            messagebox.showwarning("尚未解鎖", "此工程已設定密碼，請先在上半部輸入正確密碼並解鎖後，才能修改密碼。")
            return
        pwd = self.new_password_var.get()
        confirm = self.confirm_password_var.get()
        if len(pwd) < 3:
            messagebox.showwarning("密碼長度不足", "新密碼至少需要 3 個字元。")
            return
        if pwd != confirm:
            messagebox.showwarning("密碼不一致", "新密碼與確認新密碼不一致。")
            return
        self.db.save_password_hash(self.current_project_id, hash_password(pwd))
        self.project_password_hash = self.db.get_password_hash(self.current_project_id)
        self.edit_unlocked = True
        self.edit_password_var.set(pwd)
        self.new_password_var.set("")
        self.confirm_password_var.set("")
        self.apply_edit_lock_state()
        self.status_var.set("工程密碼已更新並儲存")
        messagebox.showinfo("密碼已更新", "此工程的編輯密碼已更新。")

    def safe_amount(self, value):
        try:
            return float(str(value).replace(",", "").replace("元", "").replace("%", "").strip() or 0)
        except ValueError:
            return 0.0

    def tree_amount_total(self, tree):
        total = 0.0
        if not hasattr(self, tree):
            return total
        for row in getattr(self, tree).get_rows():
            if len(row) >= 4:
                total += self.safe_amount(row[3])
        return total

    def money_text(self, value):
        return f"{value:,.0f}元"

    def shift_month(self, delta):
        ym = self.cal_month_var.get().strip()
        try:
            y, m = map(int, ym.split("-"))
            first = date(y, m, 1)
        except Exception:
            first = date.today().replace(day=1)
        m = first.month + delta
        y = first.year
        if m < 1:
            y -= 1
            m = 12
        elif m > 12:
            y += 1
            m = 1
        self.cal_month_var.set(f"{y:04d}-{m:02d}")
        self.render_calendar()

    def load_projects(self):
        projects = self.db.projects()
        if not projects:
            pid = self.db.create_project("新工程")
            self.db.set_setting("last_project_id", pid)
            projects = self.db.projects()

        self.project_map = {f"{p['name']}  #{p['id']}": p["id"] for p in projects}
        self.project_combo["values"] = list(self.project_map.keys())

        last = self.db.get_setting("last_project_id", "")
        selected_label = None
        for label, pid in self.project_map.items():
            if str(pid) == str(last):
                selected_label = label
                break
        if not selected_label:
            selected_label = list(self.project_map.keys())[0]

        self.project_combo.set(selected_label)
        self.load_project(self.project_map[selected_label])

    def new_project(self):
        self.save_current()
        name = simpledialog.askstring("新增工程", "請輸入工程名稱：", parent=self)
        if not name:
            return
        pid = self.db.create_project(name.strip())
        self.db.set_setting("last_project_id", pid)
        self.load_projects()

    def rename_current_project(self):
        if not self.current_project_id:
            return
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        current_name = self.basic_vars["name"].get().strip()
        new_name = simpledialog.askstring("修改工程名稱", "請輸入新的工程名稱：", initialvalue=current_name, parent=self)
        if not new_name:
            return
        self.basic_vars["name"].set(new_name.strip())
        self.save_current()
        self.load_projects()
        self.status_var.set("工程名稱已修改")

    def on_project_selected(self, *_):
        label = self.project_combo.get()
        pid = self.project_map.get(label)
        if pid and pid != self.current_project_id:
            self.save_current()
            self.load_project(pid)

    def load_project(self, pid):
        self.loading = True
        self.current_project_id = pid
        p = self.db.get_project(pid)

        for key in self.basic_vars:
            self.basic_vars[key].set(p[key] if key in p.keys() and p[key] is not None else "")
        if hasattr(self, "project_description_text"):
            self.project_description_text.delete("1.0", "end")
            self.project_description_text.insert("1.0", p["project_description"] if "project_description" in p.keys() and p["project_description"] is not None else "")
        self.day_type_var.set(p["day_type"] or "工作日")

        self.bid_tree.set_rows([[r["round_no"], r["online_date"], r["open_date"]] for r in self.db.bids(pid)])
        self.holiday_tree.set_rows([["✓" if r["excluded"] else "", r["day"], r["name"]] for r in self.db.rows("holidays", pid)])
        self.workday_tree.set_rows([["✓" if r["excluded"] else "", r["day"], r["name"]] for r in self.db.rows("workdays", pid)])
        self.apply_year_separators(self.holiday_tree)
        self.apply_year_separators(self.workday_tree)
        self.weather_tree.set_rows([[r["day"], r["morning"], r["afternoon"], r["typhoon"], r["site"], r["note"]] for r in self.db.rows("weather", pid)])
        self.railway_tree.set_rows([["✓" if r["excluded"] else "", r["day"], r["note"]] for r in self.db.rows("railway", pid)])
        self.payment_contract_tree.set_rows([[r["day"], r["item"], r["voucher_no"], r["amount"], r["note"]] for r in self.db.rows("payment_contract", pid)])
        self.payment_other_tree.set_rows([[r["day"], r["item"], r["voucher_no"], r["amount"], r["note"]] for r in self.db.rows("payment_other", pid)])
        self.payment_admin_tree.set_rows([[r["day"], r["item"], r["voucher_no"], r["amount"], r["note"]] for r in self.db.rows("payment_admin", pid)])
        self.execution_tree.set_rows([
            [r["day"], r["record_type"], "\n".join(x for x in [r["subject"], r["content"]] if x), r["note"]]
            for r in self.db.rows("execution_records", pid)
        ])
        self.milestone_tree.set_rows([
            self.calc_milestone_row([r["item_no"], r["contract_item"], r["start_date"], r["deadline_days"], r["deadline_date"], r["received_date"], r["overdue"], r["received_no"], r["note"], r["day_adjust"]])
            for r in self.db.rows("project_milestones", pid)
        ])
        self.refresh_milestone_rows()
        self.execution_status_var.set(self.db.get_status(pid) or "規劃中")
        self.project_password_hash = self.db.get_password_hash(pid)
        self.edit_unlocked = False if self.project_password_hash else True
        self.edit_password_var.set("")

        self.loading = False
        self.dirty = False
        self.db.set_setting("last_project_id", pid)
        self.recalculate()
        self.render_calendar()
        self.apply_edit_lock_state()
        self.undo_snapshot = None
        self.last_state = self.capture_state()
        self.status_var.set(f"已載入：{p['name']}")

    def collect_exclude_dates(self, include_railway=False):
        days = set()
        for row in self.holiday_tree.get_rows():
            if row and row[0] == "✓":
                continue
            d = parse_date(row[1] if len(row) > 1 else "")
            if d:
                days.add(d)
        if include_railway:
            for row in self.railway_tree.get_rows():
                if row and row[0] == "✓":
                    continue
                d = parse_date(row[1] if len(row) > 1 else "")
                if d:
                    days.add(d)
        return days

    def collect_railway_dates(self):
        days = set()
        for row in self.railway_tree.get_rows():
            if row and row[0] == "✓":
                continue
            d = parse_date(row[1] if len(row) > 1 else "")
            if d:
                days.add(d)
        return days

    def collect_workday_dates(self):
        days = set()
        for row in self.workday_tree.get_rows():
            if row and row[0] == "✓":
                continue
            d = parse_date(row[1] if len(row) > 1 else "")
            if d:
                days.add(d)
        return days

    def collect_weather_deductions(self):
        rows = {}
        for row in self.weather_tree.get_rows():
            d = parse_date(row[0] if row else "")
            if d:
                rows[d] = (
                    self.safe_amount(row[1] if len(row) > 1 else 0),
                    self.safe_amount(row[2] if len(row) > 2 else 0),
                )
        return rows

    def daily_work_increment(self, d, holiday_dates=None, railway_dates=None, workday_dates=None, weather_rows=None):
        if holiday_dates is None:
            holiday_dates = self.collect_exclude_dates(False)
        if railway_dates is None:
            railway_dates = self.collect_railway_dates()
        if workday_dates is None:
            workday_dates = self.collect_workday_dates()
        if weather_rows is None:
            weather_rows = self.collect_weather_deductions()

        if d in railway_dates:
            return 0
        base = 1 if self.day_type_var.get() == "日曆天" else (1 if d.weekday() < 5 else 0)
        if d in workday_dates:
            base = 1
        if d in holiday_dates:
            return 0
        if d in weather_rows:
            morning, afternoon = weather_rows[d]
            if morning >= 1.0:
                rain_deduct = 1.0
            else:
                rain_deduct = (1.0 if morning > 0 else 0) + (0.5 if afternoon > 0 else 0)
            return max(0, base - rain_deduct)
        return base

    def count_project_workdays_until(self, start, end):
        if not start or not end or end < start:
            return 0.0
        holiday_dates = self.collect_exclude_dates(False)
        railway_dates = self.collect_railway_dates()
        workday_dates = self.collect_workday_dates()
        weather_rows = self.collect_weather_deductions()

        total = 0.0
        cur = start
        while cur <= end:
            total += self.daily_work_increment(cur, holiday_dates, railway_dates, workday_dates, weather_rows)
            cur += timedelta(days=1)
        return total

    def recalculate(self):
        if not self.current_project_id:
            return
        self.recalculating = True
        try:
            start = parse_date(self.basic_vars["actual_start"].get()) or parse_date(self.basic_vars["planned_start"].get())
            planned_start = parse_date(self.basic_vars["planned_start"].get())
            try:
                days = int(float(self.basic_vars["contract_days"].get() or 0))
            except ValueError:
                days = 0

            day_type = self.day_type_var.get()
            holiday_ex = self.collect_exclude_dates(False)
            transport_ex = self.collect_exclude_dates(True)

            if day_type == "工作日":
                finish_holiday = add_work_days(planned_start, days, holiday_ex)
                finish_transport = add_work_days(planned_start, days, transport_ex)
            else:
                finish_holiday = add_calendar_days(planned_start, days)
                # 日曆天原則上不扣例假；疏運表仍視為停工順延。
                finish_transport = add_calendar_days(planned_start, days)
                if finish_transport:
                    extra = sum(1 for d in transport_ex if planned_start and planned_start <= d <= finish_transport)
                    finish_transport = finish_transport + timedelta(days=extra)

            finish_holiday_text = fmt_date(finish_holiday)
            finish_transport_text = fmt_date(finish_transport)
            if self.basic_vars["planned_finish_holiday"].get() != finish_holiday_text:
                self.basic_vars["planned_finish_holiday"].set(finish_holiday_text)
            if self.basic_vars["planned_finish_transport"].get() != finish_transport_text:
                self.basic_vars["planned_finish_transport"].set(finish_transport_text)

            auto_totals = {
                "contract_budget_total": self.safe_amount(self.basic_vars.get("contract_budget_net", tk.StringVar()).get()) + self.safe_amount(self.basic_vars.get("contract_budget_tax", tk.StringVar()).get()),
                "contract_award_total": self.safe_amount(self.basic_vars.get("contract_award_net", tk.StringVar()).get()) + self.safe_amount(self.basic_vars.get("contract_award_tax", tk.StringVar()).get()),
                "deposit_total": self.safe_amount(self.basic_vars.get("deposit_difference", tk.StringVar()).get()) + self.safe_amount(self.basic_vars.get("deposit_performance", tk.StringVar()).get()),
                "warranty_deposit": self.safe_amount(self.basic_vars.get("final_contract_amount", tk.StringVar()).get()) * (self.safe_amount(self.basic_vars.get("warranty_rate", tk.StringVar()).get()) / 100),
            }
            for key, value in auto_totals.items():
                if key in self.basic_vars:
                    text = self.money_text(value) if value else ""
                    if self.basic_vars[key].get() != text:
                        self.basic_vars[key].set(text)

            today = date.today()
            elapsed = (today - start).days + 1 if start and today >= start else 0
            workday_no = self.count_project_workdays_until(start, today) if start else 0

            self.summary_vars["start"].set(fmt_date(start))
            self.summary_vars["finish1"].set(fmt_date(finish_holiday))
            self.summary_vars["finish2"].set(fmt_date(finish_transport))
            self.summary_vars["elapsed"].set(str(elapsed))
            self.summary_vars["workday_no"].set(f"{float(workday_no):.1f}")
            self.summary_vars["contract_total"].set(self.money_text(self.tree_amount_total("payment_contract_tree")))
            self.summary_vars["other_total"].set(self.money_text(self.tree_amount_total("payment_other_tree")))
            self.summary_vars["admin_total"].set(self.money_text(self.tree_amount_total("payment_admin_tree")))
            self.summary_vars["execution_status"].set(self.execution_status_var.get() if hasattr(self, "execution_status_var") else "")

            if not self.loading:
                self.render_calendar()
        finally:
            self.recalculating = False

    def save_current(self):
        if not self.current_project_id:
            return
        if not self.can_edit() and self.dirty:
            self.status_var.set("編輯已鎖定，未儲存變更")
            return

        data = {k: v.get().strip() for k, v in self.basic_vars.items()}
        if hasattr(self, "project_description_text"):
            data["project_description"] = self.project_description_text.get("1.0", "end-1c").strip()
        data["day_type"] = self.day_type_var.get()
        try:
            data["contract_days"] = int(float(data.get("contract_days") or 0))
        except ValueError:
            data["contract_days"] = 0

        self.db.save_project(self.current_project_id, data)

        bids = []
        for r in self.bid_tree.get_rows():
            bids.append({"round_no": r[0] or 1, "online_date": r[1], "open_date": r[2]})
        self.db.replace_rows("bids", self.current_project_id, bids)

        holidays = [
            {"excluded": 1 if len(r) > 0 and r[0] == "✓" else 0, "day": r[1] if len(r) > 1 else "", "name": r[2] if len(r) > 2 else ""}
            for r in self.holiday_tree.get_rows()
        ]
        self.db.replace_rows("holidays", self.current_project_id, holidays)
        workdays = [
            {"excluded": 1 if len(r) > 0 and r[0] == "✓" else 0, "day": r[1] if len(r) > 1 else "", "name": r[2] if len(r) > 2 else ""}
            for r in self.workday_tree.get_rows()
        ]
        self.db.replace_rows("workdays", self.current_project_id, workdays)

        weather = []
        for r in self.weather_tree.get_rows():
            weather.append({
                "day": r[0], "morning": r[1] or 0, "afternoon": r[2] or 0,
                "typhoon": r[3] or 0, "site": r[4] or 0, "note": r[5] if len(r) > 5 else ""
            })
        self.db.replace_rows("weather", self.current_project_id, weather)

        railway = [
            {"excluded": 1 if len(r) > 0 and r[0] == "✓" else 0, "day": r[1] if len(r) > 1 else "", "note": r[2] if len(r) > 2 else ""}
            for r in self.railway_tree.get_rows()
        ]
        self.db.replace_rows("railway", self.current_project_id, railway)

        for table_name, tree_attr in [
            ("payment_contract", "payment_contract_tree"),
            ("payment_other", "payment_other_tree"),
            ("payment_admin", "payment_admin_tree"),
        ]:
            rows = []
            for r in getattr(self, tree_attr).get_rows():
                rows.append({
                    "day": r[0] if len(r) > 0 else "",
                    "item": r[1] if len(r) > 1 else "",
                    "voucher_no": r[2] if len(r) > 2 else "",
                    "amount": self.safe_amount(r[3] if len(r) > 3 else 0),
                    "note": r[4] if len(r) > 4 else "",
                })
            self.db.replace_rows(table_name, self.current_project_id, rows)

        execution_rows = []
        for r in self.execution_tree.get_rows():
            execution_rows.append({
                "day": r[0] if len(r) > 0 else "",
                "record_type": r[1] if len(r) > 1 else "",
                "subject": "",
                "content": r[2] if len(r) > 2 else "",
                "note": r[3] if len(r) > 3 else "",
            })
        self.db.replace_rows("execution_records", self.current_project_id, execution_rows)
        self.refresh_milestone_rows()
        milestone_rows = []
        for r in self.milestone_tree.get_rows():
            row = self.calc_milestone_row(r)
            milestone_rows.append({
                "item_no": row[0], "contract_item": row[1], "start_date": row[2],
                "deadline_days": row[3] or 0, "deadline_date": row[4], "received_date": row[5],
                "overdue": row[6], "received_no": row[7], "note": row[8], "day_adjust": row[9] or 0
            })
        self.db.replace_rows("project_milestones", self.current_project_id, milestone_rows)
        self.db.save_status(self.current_project_id, self.execution_status_var.get() if hasattr(self, "execution_status_var") else "")

        self.db.set_setting("last_project_id", self.current_project_id)
        self.dirty = False
        self.last_state = self.capture_state()
        self.status_var.set("已自動儲存：" + datetime.now().strftime("%H:%M:%S"))

    def backup_database(self):
        self.save_current()
        folder = filedialog.askdirectory(title="選擇備份儲存資料夾")
        if not folder:
            return
        base = simpledialog.askstring("備份檔名", "請輸入備份檔名前綴：", initialvalue="TR_FxWork_備份", parent=self)
        if not base:
            return
        safe_base = "".join(ch if ch not in r'\/:*?"<>|' else "_" for ch in base.strip()) or "TR_FxWork_備份"
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        out_path = os.path.join(folder, f"{safe_base}_{timestamp}.zip")
        tmp_db = os.path.join(tempfile.gettempdir(), f"TR_FxWork_backup_{timestamp}.db")
        try:
            with sqlite3.connect(DB_FILE) as src, sqlite3.connect(tmp_db) as dst:
                src.backup(dst)
            with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                zf.write(tmp_db, arcname=DB_FILE_NAME)
                zf.writestr("README_備份說明.txt", f"臺鐵監造紀錄小本資料庫備份\n備份時間：{timestamp}\n")
            messagebox.showinfo("備份完成", f"已完成備份：\n{out_path}")
        except Exception as exc:
            messagebox.showerror("備份失敗", str(exc))
        finally:
            try:
                if os.path.exists(tmp_db):
                    os.remove(tmp_db)
            except OSError:
                pass

    def backup_database_offsite(self):
        self.save_current()
        folder = filedialog.askdirectory(title="選擇異地備份儲存資料夾")
        if not folder:
            return
        timestamp = datetime.now().strftime("%Y%m%d%H%M")
        out_path = os.path.join(folder, f"TR_FxWork_backup_{timestamp}.zip")
        tmp_db = os.path.join(tempfile.gettempdir(), f"TR_FxWork_backup_{timestamp}.db")
        try:
            with sqlite3.connect(DB_FILE) as src, sqlite3.connect(tmp_db) as dst:
                src.backup(dst)
            with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                zf.write(tmp_db, arcname=DB_FILE_NAME)
                zf.writestr("README_異地備份說明.txt", f"TR_FxWork 異地備份\n備份時間：{timestamp}\n")
            messagebox.showinfo("異地備份完成", f"已完成異地備份：\n{out_path}")
        except Exception as exc:
            messagebox.showerror("異地備份失敗", str(exc))
        finally:
            try:
                if os.path.exists(tmp_db):
                    os.remove(tmp_db)
            except OSError:
                pass

    def exportable_pages(self):
        return {
            "工程基本資料": None,
            "招標資訊": self.bid_tree,
            "假期表": self.holiday_tree,
            "補班日表": self.workday_tree,
            "晴雨表": self.weather_tree,
            "鐵路疏運表": self.railway_tree,
            "發包工程費計價": self.payment_contract_tree,
            "發包以外計價": self.payment_other_tree,
            "管理費計價": self.payment_admin_tree,
            "工程執行紀錄表": self.execution_tree,
            "工程大事記": self.milestone_tree,
        }

    def basic_page_export_rows(self):
        rows = [["欄位", "內容"]]
        labels = {
            "name": "工程名稱", "exec_no": "工程執行號", "budget_no": "動支請示單號",
            "purchase_contract_no": "採購契約號碼", "award_date": "決標日期", "contract_date": "簽約日期",
            "planned_start": "預訂開工日", "actual_start": "實際開工日", "contract_days": "契約工期",
            "planned_finish_holiday": "預訂竣工日（例假表）", "planned_finish_transport": "預訂竣工日（疏運表）",
            "actual_finish": "實際竣工日", "contractor": "承攬商", "company_address": "公司地址",
            "responsible_person": "負責人", "contact_person": "聯絡人", "phone": "電話",
            "fax": "傳真電話", "tax_id": "統一編號", "project_description": "工程說明",
            "contract_budget_net": "發包工程費-預算(未稅)", "contract_award_net": "發包工程費-決標(未稅)",
            "contract_budget_tax": "發包工程費-稅金(預算)", "contract_award_tax": "發包工程費-稅金(決標)",
            "contract_budget_total": "發包工程費-預算(含稅)", "contract_award_total": "發包工程費-決標(契約金額含稅)",
            "labor_budget": "包工費-預算", "labor_award": "包工費-決標",
            "deposit_difference": "差額保證金", "deposit_performance": "履約保證金", "deposit_total": "保證金總額",
            "final_contract_amount": "竣工發包工程費", "warranty_rate": "保固金比例", "warranty_deposit": "保固保證金",
        }
        for key, label in labels.items():
            if key == "project_description" and hasattr(self, "project_description_text"):
                value = self.project_description_text.get("1.0", "end-1c")
            else:
                value = self.basic_vars[key].get() if key in self.basic_vars else ""
            rows.append([label, value])
        rows.append(["工期類型", self.day_type_var.get()])
        rows.append(["工程執行狀態", self.execution_status_var.get() if hasattr(self, "execution_status_var") else ""])
        rows.append([])
        rows.append(["招標資訊"])
        rows.append([self.bid_tree.tree.heading(col)["text"] for col in self.bid_tree.columns])
        rows.extend(self.bid_tree.get_rows())
        return rows

    def table_page_export_rows(self, page_name):
        tree = self.exportable_pages()[page_name]
        return [[tree.tree.heading(col)["text"] for col in tree.columns]] + tree.get_rows()

    def export_page_excel(self):
        self.save_current()
        pages = self.exportable_pages()
        win = tk.Toplevel(self)
        win.title("匯出分頁檔案")
        win.transient(self)
        win.grab_set()
        vars_by_page = {}
        box = ttk.Frame(win, padding=12)
        box.pack(fill="both", expand=True)
        ttk.Label(box, text="請勾選要匯出的分頁").pack(anchor="w", pady=(0, 6))
        for page_name in pages:
            var = tk.BooleanVar(value=False)
            vars_by_page[page_name] = var
            ttk.Checkbutton(box, text=page_name, variable=var).pack(anchor="w", pady=2)

        def do_export():
            selected = [name for name, var in vars_by_page.items() if var.get()]
            if not selected:
                messagebox.showwarning("尚未選擇", "請至少勾選一個分頁。", parent=win)
                return
            folder = filedialog.askdirectory(title="選擇匯出 Excel 儲存資料夾", parent=win)
            if not folder:
                return
            timestamp = datetime.now().strftime("%Y%m%d%H%M")
            exported = []
            try:
                for page_name in selected:
                    data = self.basic_page_export_rows() if page_name == "工程基本資料" else self.table_page_export_rows(page_name)
                    safe_name = "".join(ch if ch not in r'\/:*?"<>|' else "_" for ch in page_name)
                    path = os.path.join(folder, f"{safe_name}_{timestamp}.xlsx")
                    write_simple_xlsx(path, page_name, data)
                    exported.append(path)
                messagebox.showinfo("匯出完成", "已匯出：\n" + "\n".join(exported), parent=win)
                win.destroy()
            except Exception as exc:
                messagebox.showerror("匯出失敗", str(exc), parent=win)

        btns = ttk.Frame(box)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="匯出", command=do_export).pack(side="right", padx=4)
        ttk.Button(btns, text="取消", command=win.destroy).pack(side="right", padx=4)

    def import_basic_page_rows(self, rows):
        label_to_key = {
            "工程名稱": "name", "工程執行號": "exec_no", "動支請示單號": "budget_no",
            "採購契約號碼": "purchase_contract_no", "決標日期": "award_date", "簽約日期": "contract_date",
            "預訂開工日": "planned_start", "實際開工日": "actual_start", "契約工期": "contract_days",
            "預訂竣工日（例假表）": "planned_finish_holiday", "預訂竣工日（疏運表）": "planned_finish_transport",
            "實際竣工日": "actual_finish", "承攬商": "contractor", "公司地址": "company_address",
            "負責人": "responsible_person", "聯絡人": "contact_person", "電話": "phone",
            "傳真電話": "fax", "統一編號": "tax_id", "工程說明": "project_description",
            "發包工程費-預算(未稅)": "contract_budget_net", "發包工程費-決標(未稅)": "contract_award_net",
            "發包工程費-稅金(預算)": "contract_budget_tax", "發包工程費-稅金(決標)": "contract_award_tax",
            "發包工程費-預算(含稅)": "contract_budget_total", "發包工程費-決標(契約金額含稅)": "contract_award_total",
            "包工費-預算": "labor_budget", "包工費-決標": "labor_award",
            "差額保證金": "deposit_difference", "履約保證金": "deposit_performance", "保證金總額": "deposit_total",
            "竣工發包工程費": "final_contract_amount", "保固金比例": "warranty_rate", "保固保證金": "warranty_deposit",
        }
        for row in rows:
            if len(row) < 2:
                continue
            key = label_to_key.get(row[0])
            if not key:
                if row[0] == "工期類型":
                    self.day_type_var.set(row[1])
                elif row[0] == "工程執行狀態" and hasattr(self, "execution_status_var"):
                    self.execution_status_var.set(row[1])
                continue
            if key == "project_description" and hasattr(self, "project_description_text"):
                self.project_description_text.delete("1.0", "end")
                self.project_description_text.insert("1.0", row[1])
            elif key in self.basic_vars:
                self.basic_vars[key].set(row[1])

    def import_page_excel(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "請先解除鎖定後再匯入。")
            return
        path = filedialog.askopenfilename(
            title="選擇要匯入的 Excel 檔",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not path:
            return
        try:
            rows = read_simple_xlsx(path)
            if not rows or not rows[0]:
                raise RuntimeError("Excel 第一行找不到分頁名稱")
            page_name = rows[0][0].strip()
            pages = self.exportable_pages()
            if page_name not in pages:
                raise RuntimeError(f"無法辨識分頁名稱：{page_name}")
            data_rows = rows[2:] if len(rows) >= 2 else []
            if page_name == "工程基本資料":
                self.import_basic_page_rows(data_rows)
                self.mark_dirty()
                self.save_current()
                messagebox.showinfo("匯入完成", f"已匯入到分頁：{page_name}")
                return
            expected_cols = len(pages[page_name].columns)
            normalized = [(r + [""] * expected_cols)[:expected_cols] for r in data_rows if any(str(v).strip() for v in r)]
            pages[page_name].set_rows(normalized)
            if page_name == "工程大事記":
                self.refresh_milestone_rows()
            self.mark_dirty()
            self.save_current()
            messagebox.showinfo("匯入完成", f"已匯入到分頁：{page_name}")
        except Exception as exc:
            messagebox.showerror("匯入失敗", str(exc))

    def _table_exists(self, conn, table):
        row = conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone()
        return row is not None

    def _copy_project_from_conn(self, src, src_pid, overwrite_pid=None):
        p = src.execute("SELECT * FROM projects WHERE id=?", (src_pid,)).fetchone()
        if not p:
            return None
        if overwrite_pid:
            self.db.delete_project(overwrite_pid)
        fields = [
            "name", "exec_no", "budget_no", "award_date", "planned_start", "actual_start",
            "contract_days", "day_type", "planned_finish_holiday", "planned_finish_transport",
            "actual_finish", "updated_at", "password_hash"
        ] + PROJECT_EXTRA_FIELDS
        vals = [p[f] if f in p.keys() else "" for f in fields]
        cur = self.db.conn.execute(
            "INSERT INTO projects(" + ",".join(fields) + ") VALUES(" + ",".join("?" for _ in fields) + ")",
            vals
        )
        new_pid = cur.lastrowid

        copy_specs = {
            "bids": ["round_no", "online_date", "open_date"],
            "holidays": ["day", "name"],
            "workdays": ["day", "name"],
            "weather": ["day", "morning", "afternoon", "typhoon", "site", "note"],
            "railway": ["day", "note"],
            "payment_contract": ["day", "item", "voucher_no", "amount", "note"],
            "payment_other": ["day", "item", "voucher_no", "amount", "note"],
            "payment_admin": ["day", "item", "voucher_no", "amount", "note"],
            "execution_records": ["day", "record_type", "subject", "content", "note"],
            "project_milestones": ["item_no", "contract_item", "start_date", "deadline_days", "deadline_date", "received_date", "overdue", "received_no", "note", "day_adjust"],
        }
        for table, cols in copy_specs.items():
            if not self._table_exists(src, table):
                continue
            for r in src.execute(f"SELECT * FROM {table} WHERE project_id=?", (src_pid,)).fetchall():
                values = [new_pid] + [r[c] if c in r.keys() else "" for c in cols]
                self.db.conn.execute(
                    f"INSERT INTO {table}(project_id,{','.join(cols)}) VALUES({','.join('?' for _ in values)})",
                    values
                )
        if self._table_exists(src, "execution_status"):
            st = src.execute("SELECT status FROM execution_status WHERE project_id=?", (src_pid,)).fetchone()
            if st:
                self.db.conn.execute(
                    "INSERT OR REPLACE INTO execution_status(project_id,status,updated_at) VALUES(?,?,?)",
                    (new_pid, st["status"], datetime.now().isoformat(timespec="seconds"))
                )
        self.db.conn.commit()
        return new_pid

    def import_database(self):
        self.save_current()
        path = filedialog.askopenfilename(
            title="選擇要匯入的備份檔",
            filetypes=[("TR_FxWork 備份", "*.zip *.db"), ("ZIP", "*.zip"), ("SQLite DB", "*.db"), ("所有檔案", "*.*")]
        )
        if not path:
            return
        temp_dir = None
        db_path = path
        try:
            if path.lower().endswith(".zip"):
                temp_dir = tempfile.mkdtemp(prefix="TR_FxWork_import_")
                with zipfile.ZipFile(path, "r") as zf:
                    names = zf.namelist()
                    compatible_names = []
                    for target_name in (DB_FILE_NAME, LEGACY_DB_FILE_NAME):
                        compatible_names.extend(
                            n for n in names
                            if os.path.basename(n) == target_name
                        )
                    db_names = compatible_names or [n for n in names if n.lower().endswith(".db")]
                    if not db_names:
                        raise RuntimeError("ZIP 內找不到 TRFxWork_db 或舊版 .db 資料庫檔")
                    zf.extract(db_names[0], temp_dir)
                    db_path = os.path.join(temp_dir, db_names[0])

            src = sqlite3.connect(db_path)
            src.row_factory = sqlite3.Row
            if not self._table_exists(src, "projects"):
                raise RuntimeError("匯入檔不是 TR_FxWork 資料庫")

            imported = skipped = overwritten = 0
            for p in src.execute("SELECT * FROM projects ORDER BY id").fetchall():
                same = self.db.get_project_by_name(p["name"])
                overwrite_pid = None
                if same:
                    ans = messagebox.askyesnocancel(
                        "工程名稱重複",
                        f"匯入檔內工程「{p['name']}」已存在。\n\n是：覆蓋既有工程\n否：不匯入此工程，繼續下一筆\n取消：停止匯入"
                    )
                    if ans is None:
                        break
                    if ans is False:
                        skipped += 1
                        continue
                    overwrite_pid = same["id"]
                    overwritten += 1
                self._copy_project_from_conn(src, p["id"], overwrite_pid=overwrite_pid)
                imported += 1
            src.close()
            self.load_projects()
            messagebox.showinfo("匯入完成", f"匯入：{imported} 筆\n覆蓋：{overwritten} 筆\n略過：{skipped} 筆")
        except Exception as exc:
            messagebox.showerror("匯入失敗", str(exc))
        finally:
            if temp_dir:
                shutil.rmtree(temp_dir, ignore_errors=True)

    def delete_current_project(self):
        if not self.current_project_id:
            return
        p = self.db.get_project(self.current_project_id)
        if not p:
            return
        ok = messagebox.askyesno(
            "警示確認",
            f"確定要刪除工程「{p['name']}」嗎？\n\n此動作會刪除該工程所有分頁資料，且無法復原。"
        )
        if not ok:
            return
        self.db.delete_project(self.current_project_id)
        self.current_project_id = None
        self.load_projects()
        self.status_var.set("已刪除工程")

    def auto_save_loop(self):
        if self.dirty:
            self.save_current()
        self.after(3000, self.auto_save_loop)

    def on_close(self):
        self.save_current()
        self.destroy()

    def weather_text_map(self):
        out = {}
        for r in self.weather_tree.get_rows():
            d = parse_date(r[0])
            if not d:
                continue
            tags = []
            try:
                if float(r[1] or 0) > 0:
                    tags.append("上午雨")
            except ValueError:
                pass
            try:
                if float(r[2] or 0) > 0:
                    tags.append("下午雨")
            except ValueError:
                pass
            try:
                if float(r[3] or 0) == 1:
                    tags.append("颱風")
            except ValueError:
                pass
            try:
                if float(r[4] or 0) == 1:
                    tags.append("場地")
            except ValueError:
                pass
            if tags:
                out[d] = "、".join(tags)
        return out

    def render_calendar(self):
        if not hasattr(self, "cal_canvas"):
            return
        c = self.cal_canvas
        c.delete("all")

        try:
            y, m = map(int, self.cal_month_var.get().split("-"))
        except Exception:
            y, m = date.today().year, date.today().month

        holidays = {}
        for r in self.holiday_tree.get_rows():
            if r and r[0] == "✓":
                continue
            d = parse_date(r[1] if len(r) > 1 else "")
            if d:
                holidays[d] = (r[2] if len(r) > 2 else "") or "假日"

        railway = self.collect_railway_dates()
        workdays = self.collect_workday_dates()

        weather = self.weather_text_map()
        weather_rows = self.collect_weather_deductions()

        canvas_w = c.winfo_width()
        canvas_h = c.winfo_height()
        width = canvas_w if canvas_w > 1 else 900
        height = canvas_h if canvas_h > 1 else 430
        left_w = max(58, min(88, int(width * 0.09)))
        right_pad = 8
        top_h = 28
        weekday_h = 22
        note_h = 22
        cell_w = max(50, (width - left_w - right_pad) / 7)

        colors = {
            "normal_date": "#fff2cc",      # 粉黃色
            "weekend": "#f4cccc",          # 紅粉色
            "holiday": "#d9ead3",          # 粉綠色
            "transport": "#ead1dc",        # 粉棕色
            "weather": "#ead1dc",
            "white": "#ffffff",
            "grid": "#888888",
            "header": "#ddebf7",
        }

        cal = calendar.Calendar(firstweekday=0)
        weeks = cal.monthdatescalendar(y, m)
        available_h = max(150, height - top_h - weekday_h - note_h - 6)
        week_h = available_h / max(1, len(weeks))
        row_h = week_h / 5
        main_font_size = max(7, min(9, int(row_h * 0.44)))
        label_font_size = max(7, min(9, int(row_h * 0.42)))
        title_font_size = max(11, min(14, int(top_h * 0.48)))
        note_font_size = max(7, min(9, int(note_h * 0.42)))
        content_height = top_h + weekday_h + len(weeks) * week_h
        note_y = content_height + 24
        c.configure(scrollregion=(0, 0, width, height))

        c.create_text(width/2, 15, text=f"{y} 年 {m:02d} 月 施工日曆", font=("Microsoft JhengHei UI", title_font_size, "bold"))

        for i, wd in enumerate(WEEKDAY_NAMES):
            x0 = left_w + i * cell_w
            c.create_rectangle(x0, top_h, x0 + cell_w, top_h + weekday_h, fill=colors["header"], outline=colors["grid"])
            c.create_text(x0 + cell_w/2, top_h + weekday_h/2, text=f"星期{wd}", font=("Microsoft JhengHei UI", main_font_size, "bold"))

        row_labels = ["假日", "疏運日", "雨天", "工作日數"]
        work_count = 0
        holiday_ex = self.collect_exclude_dates(True)
        holiday_dates = self.collect_exclude_dates(False)
        start_count_date = parse_date(self.basic_vars["actual_start"].get()) or parse_date(self.basic_vars["planned_start"].get())
        finish_count_date = (
            parse_date(self.basic_vars["planned_finish_transport"].get())
            or parse_date(self.basic_vars["planned_finish_holiday"].get())
            or parse_date(self.basic_vars["actual_finish"].get())
        )

        def day_increment(d):
            return self.daily_work_increment(d, holiday_dates, railway, workdays, weather_rows)

        if start_count_date:
            cur = start_count_date
            first_of_month = date(y, m, 1)
            while cur < first_of_month and (not finish_count_date or cur <= finish_count_date):
                work_count += day_increment(cur)
                cur += timedelta(days=1)

        for wi, week in enumerate(weeks):
            y0 = top_h + weekday_h + wi * week_h
            for ri, label in enumerate(row_labels):
                c.create_rectangle(5, y0 + ri*row_h, left_w, y0 + (ri+1)*row_h, fill="#f2f2f2", outline=colors["grid"])
                c.create_text(left_w - 6, y0 + ri*row_h + row_h/2, text=label, anchor="e", font=("Microsoft JhengHei UI", label_font_size))

            for di, d in enumerate(week):
                x0 = left_w + di * cell_w
                in_month = d.month == m
                alpha_fill = colors["weekend"] if d.weekday() >= 5 else colors["normal_date"]
                if not in_month:
                    alpha_fill = "#eeeeee"

                # 第1行：日期
                c.create_rectangle(x0, y0, x0+cell_w, y0+row_h, fill=alpha_fill, outline=colors["grid"])
                date_fill = "red" if in_month and d == date.today() else "black"
                c.create_text(
                    x0+cell_w-5, y0+row_h/2,
                    text=d.strftime("%m/%d") if in_month else "",
                    anchor="e",
                    font=("Microsoft JhengHei UI", main_font_size, "bold"),
                    fill=date_fill
                )

                # 第2行：假日
                htxt = holidays.get(d, "") if in_month else ""
                fill = colors["holiday"] if htxt else colors["white"]
                c.create_rectangle(x0, y0+row_h, x0+cell_w, y0+2*row_h, fill=fill, outline=colors["grid"])
                c.create_text(x0+cell_w/2, y0+row_h*1.5, text=htxt, font=("Microsoft JhengHei UI", main_font_size))

                # 第3行：疏運
                rtxt = "疏運" if in_month and d in railway else ""
                fill = colors["transport"] if rtxt else colors["white"]
                c.create_rectangle(x0, y0+2*row_h, x0+cell_w, y0+3*row_h, fill=fill, outline=colors["grid"])
                c.create_text(x0+cell_w/2, y0+row_h*2.5, text=rtxt, font=("Microsoft JhengHei UI", main_font_size))

                # 第4行：雨天
                wtxt = weather.get(d, "") if in_month else ""
                fill = colors["weather"] if wtxt else colors["white"]
                c.create_rectangle(x0, y0+3*row_h, x0+cell_w, y0+4*row_h, fill=fill, outline=colors["grid"])
                c.create_text(x0+cell_w/2, y0+row_h*3.5, text=wtxt, font=("Microsoft JhengHei UI", main_font_size))

                # 第5個視覺區塊：工作日數累計
                in_contract_period = (
                    in_month
                    and start_count_date
                    and finish_count_date
                    and start_count_date <= d <= finish_count_date
                )
                if in_contract_period:
                    work_count += day_increment(d)
                    txt = f"{work_count:g}" if work_count else ""
                else:
                    txt = ""
                c.create_rectangle(x0, y0+4*row_h, x0+cell_w, y0+5*row_h, fill="#ffffff", outline=colors["grid"])
                c.create_text(x0+cell_w/2, y0+row_h*4.5, text=txt, font=("Microsoft JhengHei UI", main_font_size, "bold"))

        c.create_text(
            10, note_y,
            text="說明：週六週日紅粉色；假日粉綠色；疏運與雨天粉棕色；資料關閉前與編輯中會自動儲存。",
            anchor="sw",
            font=("Microsoft JhengHei UI", note_font_size)
        )


if __name__ == "__main__":
    app = App()
    app.mainloop()
