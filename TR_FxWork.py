# -*- coding: utf-8 -*-
"""
臺鐵監造紀錄小本 V0.1.d
- Python 標準函式庫版本：tkinter + sqlite3
- 關閉前自動儲存
- 可建立多個工程
- 開啟時自動載入上次編輯工程
- 基本資料、假期表、晴雨表、鐵路疏運表、週曆總表、計價資料、工程執行紀錄表
"""

import os
import hashlib
import shutil
import sqlite3
import tempfile
import zipfile
import calendar
from datetime import date, datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog


APP_TITLE = "臺鐵監造紀錄小本 V0.1.d"
DB_FILE = "TR_FxWork.db"

WEEKDAY_NAMES = ["一", "二", "三", "四", "五", "六", "日"]
PASSWORD_SALT = "1981"


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
            "bids", "holidays", "weather", "railway",
            "payment_contract", "payment_other", "payment_admin",
            "execution_records", "execution_status"
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
        ]
        sets = ", ".join(f"{f}=?" for f in fields)
        vals = [data.get(f, "") for f in fields] + [pid]
        self.conn.execute(f"UPDATE projects SET {sets} WHERE id=?", vals)
        self.conn.commit()

    def rows(self, table, pid):
        return self.conn.execute(f"SELECT * FROM {table} WHERE project_id=? ORDER BY day, id", (pid,)).fetchall()

    def bids(self, pid):
        return self.conn.execute("SELECT * FROM bids WHERE project_id=? ORDER BY round_no, id", (pid,)).fetchall()

    def replace_rows(self, table, pid, rows):
        self.conn.execute(f"DELETE FROM {table} WHERE project_id=?", (pid,))
        if table == "bids":
            for r in rows:
                self.conn.execute(
                    "INSERT INTO bids(project_id, round_no, online_date, open_date) VALUES(?,?,?,?)",
                    (pid, r.get("round_no", 1), r.get("online_date", ""), r.get("open_date", ""))
                )
        elif table == "holidays":
            for r in rows:
                self.conn.execute(
                    "INSERT INTO holidays(project_id, day, name) VALUES(?,?,?)",
                    (pid, r.get("day", ""), r.get("name", ""))
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
            for r in rows:
                self.conn.execute(
                    "INSERT INTO railway(project_id, day, note) VALUES(?,?,?)",
                    (pid, r.get("day", ""), r.get("note", ""))
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
        self.conn.commit()


class EditableTree(ttk.Frame):
    def __init__(self, master, columns, headings, widths, on_changed=None, can_edit=None):
        super().__init__(master)
        self.columns = columns
        self.on_changed = on_changed
        self.can_edit = can_edit or (lambda: True)
        self.tree = ttk.Treeview(self, columns=columns, show="headings", height=12)
        vs = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)

        for col, head, width in zip(columns, headings, widths):
            self.tree.heading(col, text=head)
            self.tree.column(col, width=width, anchor="center", stretch=True)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        btns = ttk.Frame(self)
        btns.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
        ttk.Button(btns, text="新增一列", command=self.add_row).pack(side="left", padx=3)
        ttk.Button(btns, text="編輯選取列", command=self.edit_row).pack(side="left", padx=3)
        ttk.Button(btns, text="刪除選取列", command=self.delete_row).pack(side="left", padx=3)
        self.tree.bind("<Double-1>", lambda e: self.edit_row())

    def add_row(self, values=None):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        values = values or [""] * len(self.columns)
        self.tree.insert("", "end", values=values)
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
            ttk.Label(win, text=col).grid(row=i, column=0, sticky="e", padx=8, pady=4)
            e = ttk.Entry(win, width=32)
            e.grid(row=i, column=1, sticky="ew", padx=8, pady=4)
            e.insert(0, old[i] if i < len(old) else "")
            entries.append(e)

        def ok():
            vals = [e.get().strip() for e in entries]
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

    def get_rows(self):
        return [list(self.tree.item(i, "values")) for i in self.tree.get_children()]

    def changed(self):
        if self.on_changed:
            self.on_changed()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1250x820")
        self.minsize(1100, 720)

        self.db = DB(DB_FILE)
        self.current_project_id = None
        self.loading = False
        self.dirty = False
        self.project_password_hash = ""
        self.edit_unlocked = True
        self.edit_widgets = []

        self.style = ttk.Style()
        self.style.configure("Top.TLabelframe.Label", font=("Microsoft JhengHei UI", 11, "bold"))
        self.style.configure("TLabel", font=("Microsoft JhengHei UI", 10))
        self.style.configure("TButton", font=("Microsoft JhengHei UI", 10))
        self.style.configure("Treeview", rowheight=26, font=("Microsoft JhengHei UI", 10))
        self.style.configure("Treeview.Heading", font=("Microsoft JhengHei UI", 10, "bold"))

        self.build_ui()
        self.load_projects()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(3000, self.auto_save_loop)

    def build_ui(self):
        top_select = ttk.Frame(self, padding=8)
        top_select.pack(fill="x")

        ttk.Label(top_select, text="工程：").pack(side="left")
        self.project_combo = ttk.Combobox(top_select, state="readonly", width=55)
        self.project_combo.pack(side="left", padx=5)
        self.project_combo.bind("<<ComboboxSelected>>", self.on_project_selected)
        ttk.Button(top_select, text="新增工程", command=self.new_project).pack(side="left", padx=5)
        ttk.Button(top_select, text="立即儲存", command=self.save_current).pack(side="left", padx=5)
        ttk.Button(top_select, text="資料庫打包備份", command=self.backup_database).pack(side="left", padx=5)
        ttk.Button(top_select, text="匯入備份", command=self.import_database).pack(side="left", padx=5)
        ttk.Button(top_select, text="刪除工程", command=self.delete_current_project).pack(side="left", padx=5)

        ttk.Label(top_select, text="編輯密碼：").pack(side="left", padx=(12, 2))
        self.edit_password_var = tk.StringVar()
        self.edit_password_entry = ttk.Entry(top_select, textvariable=self.edit_password_var, show="*", width=12)
        self.edit_password_entry.pack(side="left", padx=2)
        ttk.Button(top_select, text="解鎖", command=self.unlock_project).pack(side="left", padx=3)
        self.lock_state_var = tk.StringVar(value="未鎖定")
        ttk.Label(top_select, textvariable=self.lock_state_var, foreground="#a64d00").pack(side="left", padx=3)

        self.status_var = tk.StringVar(value="")
        ttk.Label(top_select, textvariable=self.status_var).pack(side="right")

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
        self.tab_status = ttk.Frame(self.nb, padding=8)
        self.tab_settings = ttk.Frame(self.nb, padding=8)

        self.nb.add(self.tab_basic, text="第一分頁：工程基本資料")
        self.nb.add(self.tab_holiday, text="第二分頁：假期表")
        self.nb.add(self.tab_weather, text="第三分頁：晴雨表")
        self.nb.add(self.tab_railway, text="第四分頁：鐵路疏運表")
        self.nb.add(self.tab_calendar, text="第五分頁：週曆總表")
        self.nb.add(self.tab_payment_contract, text="第六分頁：發包工程費計價")
        self.nb.add(self.tab_payment_other, text="第七分頁：發包以外計價")
        self.nb.add(self.tab_payment_admin, text="第八分頁：管理費計價")
        self.nb.add(self.tab_execution, text="第九分頁：工程執行紀錄表")
        self.nb.add(self.tab_status, text="第十分頁：執行狀態")
        self.nb.add(self.tab_settings, text="第十一分頁：設定")

        self.build_basic_tab()
        self.build_holiday_tab()
        self.build_weather_tab()
        self.build_railway_tab()
        self.build_calendar_tab()
        self.build_payment_tabs()
        self.build_execution_tab()
        self.build_status_tab()
        self.build_settings_tab()
        self.assign_tree_edit_guards()

    def can_edit(self):
        return self.edit_unlocked or not self.project_password_hash

    def mark_dirty(self, *_):
        if not self.loading:
            if not self.can_edit():
                self.status_var.set("編輯已鎖定，請先輸入正確編輯密碼")
                return
            self.dirty = True
            self.recalculate()

    def entry(self, parent, row, col, label, key, width=28):
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="e", padx=6, pady=5)
        var = tk.StringVar()
        var.trace_add("write", self.mark_dirty)
        ent = ttk.Entry(parent, textvariable=var, width=width)
        ent.grid(row=row, column=col+1, sticky="ew", padx=6, pady=5)
        self.edit_widgets.append(ent)
        self.basic_vars[key] = var
        return ent

    def build_basic_tab(self):
        self.basic_vars = {}
        form = ttk.LabelFrame(self.tab_basic, text="手動輸入工程基本資料", padding=8)
        form.pack(fill="x")

        self.entry(form, 0, 0, "工程名稱", "name", 45)
        self.entry(form, 0, 2, "工程執行號", "exec_no")
        self.entry(form, 1, 0, "工程預算單號", "budget_no")
        self.entry(form, 1, 2, "決標日期", "award_date")
        self.entry(form, 2, 0, "預訂開工日", "planned_start")
        self.entry(form, 2, 2, "實際開工日", "actual_start")
        self.entry(form, 3, 0, "契約工期", "contract_days")
        ttk.Label(form, text="工期類型").grid(row=3, column=2, sticky="e", padx=6, pady=5)
        self.day_type_var = tk.StringVar(value="工作日")
        self.day_type_var.trace_add("write", self.mark_dirty)
        self.day_type = ttk.Combobox(form, textvariable=self.day_type_var, state="readonly", values=["工作日", "日曆天"], width=25)
        self.day_type.grid(row=3, column=3, sticky="ew", padx=6, pady=5)
        self.edit_widgets.append(self.day_type)

        self.entry(form, 4, 0, "預訂竣工日（例假表）", "planned_finish_holiday")
        self.entry(form, 4, 2, "預訂竣工日（疏運表）", "planned_finish_transport")
        self.entry(form, 5, 0, "實際竣工日", "actual_finish")

        for i in range(4):
            form.grid_columnconfigure(i, weight=1)

        bid_box = ttk.LabelFrame(self.tab_basic, text="招標上網日 / 開標日期（可建立多次開標）", padding=8)
        bid_box.pack(fill="both", expand=True, pady=8)
        self.bid_tree = EditableTree(
            bid_box,
            ["round_no", "online_date", "open_date"],
            ["第幾次", "招標上網日", "開標日期"],
            [100, 180, 180],
            self.mark_dirty
        )
        self.bid_tree.pack(fill="both", expand=True)

        ttk.Button(self.tab_basic, text="重新計算預訂竣工日", command=self.recalculate).pack(anchor="e", pady=4)

    def build_holiday_tab(self):
        ttk.Label(self.tab_holiday, text="輸入格式：日期請用 YYYY-MM-DD，名稱例如：中秋節、國定假日、停工日").pack(anchor="w")
        self.holiday_tree = EditableTree(
            self.tab_holiday,
            ["day", "name"],
            ["日期", "假日名稱"],
            [160, 360],
            self.mark_dirty
        )
        self.holiday_tree.pack(fill="both", expand=True, pady=6)

    def build_weather_tab(self):
        ttk.Label(
            self.tab_weather,
            text="晴雨表：上午欄位寫 1 = 上午雨；下午欄位寫 0.5 = 下午雨；天氣欄位寫 1 = 颱風；場地欄位寫 1 = 場地"
        ).pack(anchor="w")
        self.weather_tree = EditableTree(
            self.tab_weather,
            ["day", "morning", "afternoon", "typhoon", "site", "note"],
            ["日期", "上午", "下午", "天氣", "場地", "備註"],
            [130, 80, 80, 80, 80, 300],
            self.mark_dirty
        )
        self.weather_tree.pack(fill="both", expand=True, pady=6)

    def build_railway_tab(self):
        ttk.Label(self.tab_railway, text="鐵路疏運停工日期：日期請用 YYYY-MM-DD；列入預訂竣工日（疏運表）排除。").pack(anchor="w")
        self.railway_tree = EditableTree(
            self.tab_railway,
            ["day", "note"],
            ["日期", "備註"],
            [160, 420],
            self.mark_dirty
        )
        self.railway_tree.pack(fill="both", expand=True, pady=6)

    def build_calendar_tab(self):
        top = ttk.Frame(self.tab_calendar)
        top.pack(fill="x")
        ttk.Label(top, text="顯示月份：").pack(side="left")
        self.cal_month_var = tk.StringVar(value=date.today().strftime("%Y-%m"))
        ttk.Entry(top, textvariable=self.cal_month_var, width=12).pack(side="left", padx=4)
        ttk.Button(top, text="產生週曆", command=self.render_calendar).pack(side="left", padx=4)
        ttk.Button(top, text="上個月", command=lambda: self.shift_month(-1)).pack(side="left", padx=4)
        ttk.Button(top, text="下個月", command=lambda: self.shift_month(1)).pack(side="left", padx=4)

        self.cal_canvas = tk.Canvas(self.tab_calendar, background="white")
        self.cal_canvas.pack(fill="both", expand=True, pady=8)
        self.cal_canvas.bind("<Configure>", lambda e: self.render_calendar())

    def build_payment_tabs(self):
        self.payment_contract_tree = self.make_payment_tree(
            self.tab_payment_contract,
            "發包工程費計價：先建立資料欄位，後續再依需求增加計價公式與報表。"
        )
        self.payment_other_tree = self.make_payment_tree(
            self.tab_payment_other,
            "發包以外計價：先建立資料欄位，後續再依需求增加分類與核銷流程。"
        )
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
        ttk.Label(self.tab_execution, text="工程執行紀錄表：先建立紀錄欄位，後續可再擴充為查驗、督導、變更、會勘等分類。行內可雙擊編輯。").pack(anchor="w")
        self.execution_tree = EditableTree(
            self.tab_execution,
            ["day", "record_type", "subject", "content", "note"],
            ["日期", "類別", "主旨", "執行內容", "備註"],
            [120, 140, 220, 420, 260],
            self.mark_dirty
        )
        self.execution_tree.pack(fill="both", expand=True, pady=6)

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
            "bid_tree", "holiday_tree", "weather_tree", "railway_tree",
            "payment_contract_tree", "payment_other_tree", "payment_admin_tree",
            "execution_tree"
        ]:
            if hasattr(self, name):
                getattr(self, name).can_edit = self.can_edit

    def apply_edit_lock_state(self):
        unlocked = self.can_edit()
        for w in getattr(self, "edit_widgets", []):
            try:
                w.configure(state="normal" if unlocked else "disabled")
            except tk.TclError:
                pass
        self.assign_tree_edit_guards()
        if not self.project_password_hash:
            self.lock_state_var.set("未設定密碼")
        elif unlocked:
            self.lock_state_var.set("已解鎖")
        else:
            self.lock_state_var.set("已鎖定")

    def unlock_project(self):
        if not self.current_project_id:
            return
        if not self.project_password_hash:
            self.edit_unlocked = True
            self.apply_edit_lock_state()
            self.status_var.set("此工程未設定密碼，可直接編輯")
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
            return float(str(value).replace(",", "").strip() or 0)
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
        return f"{value:,.0f}"

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
        self.day_type_var.set(p["day_type"] or "工作日")

        self.bid_tree.set_rows([[r["round_no"], r["online_date"], r["open_date"]] for r in self.db.bids(pid)])
        self.holiday_tree.set_rows([[r["day"], r["name"]] for r in self.db.rows("holidays", pid)])
        self.weather_tree.set_rows([[r["day"], r["morning"], r["afternoon"], r["typhoon"], r["site"], r["note"]] for r in self.db.rows("weather", pid)])
        self.railway_tree.set_rows([[r["day"], r["note"]] for r in self.db.rows("railway", pid)])
        self.payment_contract_tree.set_rows([[r["day"], r["item"], r["voucher_no"], r["amount"], r["note"]] for r in self.db.rows("payment_contract", pid)])
        self.payment_other_tree.set_rows([[r["day"], r["item"], r["voucher_no"], r["amount"], r["note"]] for r in self.db.rows("payment_other", pid)])
        self.payment_admin_tree.set_rows([[r["day"], r["item"], r["voucher_no"], r["amount"], r["note"]] for r in self.db.rows("payment_admin", pid)])
        self.execution_tree.set_rows([[r["day"], r["record_type"], r["subject"], r["content"], r["note"]] for r in self.db.rows("execution_records", pid)])
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
        self.status_var.set(f"已載入：{p['name']}")

    def collect_exclude_dates(self, include_railway=False):
        days = set()
        for row in self.holiday_tree.get_rows():
            d = parse_date(row[0])
            if d:
                days.add(d)
        if include_railway:
            for row in self.railway_tree.get_rows():
                d = parse_date(row[0])
                if d:
                    days.add(d)
        return days

    def recalculate(self):
        if not self.current_project_id:
            return
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

        today = date.today()
        elapsed = (today - start).days + 1 if start and today >= start else 0
        workday_no = count_work_days_until(start, today, holiday_ex | set(d for d in transport_ex if d)) if start else 0

        self.summary_vars["start"].set(fmt_date(start))
        self.summary_vars["finish1"].set(fmt_date(finish_holiday))
        self.summary_vars["finish2"].set(fmt_date(finish_transport))
        self.summary_vars["elapsed"].set(str(elapsed))
        self.summary_vars["workday_no"].set(str(workday_no))
        self.summary_vars["contract_total"].set(self.money_text(self.tree_amount_total("payment_contract_tree")))
        self.summary_vars["other_total"].set(self.money_text(self.tree_amount_total("payment_other_tree")))
        self.summary_vars["admin_total"].set(self.money_text(self.tree_amount_total("payment_admin_tree")))
        self.summary_vars["execution_status"].set(self.execution_status_var.get() if hasattr(self, "execution_status_var") else "")

        if not self.loading:
            self.render_calendar()

    def save_current(self):
        if not self.current_project_id:
            return
        if not self.can_edit() and self.dirty:
            self.status_var.set("編輯已鎖定，未儲存變更")
            return

        data = {k: v.get().strip() for k, v in self.basic_vars.items()}
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

        holidays = [{"day": r[0], "name": r[1]} for r in self.holiday_tree.get_rows()]
        self.db.replace_rows("holidays", self.current_project_id, holidays)

        weather = []
        for r in self.weather_tree.get_rows():
            weather.append({
                "day": r[0], "morning": r[1] or 0, "afternoon": r[2] or 0,
                "typhoon": r[3] or 0, "site": r[4] or 0, "note": r[5] if len(r) > 5 else ""
            })
        self.db.replace_rows("weather", self.current_project_id, weather)

        railway = [{"day": r[0], "note": r[1]} for r in self.railway_tree.get_rows()]
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
                "subject": r[2] if len(r) > 2 else "",
                "content": r[3] if len(r) > 3 else "",
                "note": r[4] if len(r) > 4 else "",
            })
        self.db.replace_rows("execution_records", self.current_project_id, execution_rows)
        self.db.save_status(self.current_project_id, self.execution_status_var.get() if hasattr(self, "execution_status_var") else "")

        self.db.set_setting("last_project_id", self.current_project_id)
        self.dirty = False
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
                zf.write(tmp_db, arcname="TR_FxWork.db")
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
        ]
        vals = [p[f] if f in p.keys() else "" for f in fields]
        cur = self.db.conn.execute(
            "INSERT INTO projects(" + ",".join(fields) + ") VALUES(" + ",".join("?" for _ in fields) + ")",
            vals
        )
        new_pid = cur.lastrowid

        copy_specs = {
            "bids": ["round_no", "online_date", "open_date"],
            "holidays": ["day", "name"],
            "weather": ["day", "morning", "afternoon", "typhoon", "site", "note"],
            "railway": ["day", "note"],
            "payment_contract": ["day", "item", "voucher_no", "amount", "note"],
            "payment_other": ["day", "item", "voucher_no", "amount", "note"],
            "payment_admin": ["day", "item", "voucher_no", "amount", "note"],
            "execution_records": ["day", "record_type", "subject", "content", "note"],
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
                    db_names = [n for n in zf.namelist() if n.lower().endswith(".db")]
                    if not db_names:
                        raise RuntimeError("ZIP 內找不到 .db 資料庫檔")
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
                if float(r[1] or 0) == 1:
                    tags.append("上午雨")
            except ValueError:
                pass
            try:
                if float(r[2] or 0) == 0.5:
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
            d = parse_date(r[0])
            if d:
                holidays[d] = r[1] or "假日"

        railway = set()
        for r in self.railway_tree.get_rows():
            d = parse_date(r[0])
            if d:
                railway.add(d)

        weather = self.weather_text_map()

        width = max(c.winfo_width(), 900)
        height = max(c.winfo_height(), 520)
        left_w = 88
        top_h = 30
        cell_w = (width - left_w - 20) / 7
        week_h = 115
        row_h = 23

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

        c.create_text(width/2, 15, text=f"{y} 年 {m:02d} 月 週曆總表", font=("Microsoft JhengHei UI", 14, "bold"))

        for i, wd in enumerate(WEEKDAY_NAMES):
            x0 = left_w + i * cell_w
            c.create_rectangle(x0, top_h, x0 + cell_w, top_h + 25, fill=colors["header"], outline=colors["grid"])
            c.create_text(x0 + cell_w/2, top_h + 12, text=f"星期{wd}", font=("Microsoft JhengHei UI", 10, "bold"))

        cal = calendar.Calendar(firstweekday=0)
        weeks = cal.monthdatescalendar(y, m)

        row_labels = ["假日", "疏運日", "雨天", "工作日數"]
        work_count = 0
        holiday_ex = self.collect_exclude_dates(True)

        for wi, week in enumerate(weeks):
            y0 = top_h + 25 + wi * week_h
            for ri, label in enumerate(row_labels):
                c.create_rectangle(5, y0 + ri*row_h, left_w, y0 + (ri+1)*row_h, fill="#f2f2f2", outline=colors["grid"])
                c.create_text(left_w - 8, y0 + ri*row_h + row_h/2, text=label, anchor="e", font=("Microsoft JhengHei UI", 9))

            for di, d in enumerate(week):
                x0 = left_w + di * cell_w
                in_month = d.month == m
                alpha_fill = colors["weekend"] if d.weekday() >= 5 else colors["normal_date"]
                if not in_month:
                    alpha_fill = "#eeeeee"

                # 第1行：日期
                c.create_rectangle(x0, y0, x0+cell_w, y0+row_h, fill=alpha_fill, outline=colors["grid"])
                c.create_text(x0+cell_w-5, y0+row_h/2, text=d.strftime("%m/%d") if in_month else "", anchor="e", font=("Microsoft JhengHei UI", 9, "bold"))

                # 第2行：假日
                htxt = holidays.get(d, "") if in_month else ""
                fill = colors["holiday"] if htxt else colors["white"]
                c.create_rectangle(x0, y0+row_h, x0+cell_w, y0+2*row_h, fill=fill, outline=colors["grid"])
                c.create_text(x0+cell_w/2, y0+row_h*1.5, text=htxt, font=("Microsoft JhengHei UI", 9))

                # 第3行：疏運
                rtxt = "疏運" if in_month and d in railway else ""
                fill = colors["transport"] if rtxt else colors["white"]
                c.create_rectangle(x0, y0+2*row_h, x0+cell_w, y0+3*row_h, fill=fill, outline=colors["grid"])
                c.create_text(x0+cell_w/2, y0+row_h*2.5, text=rtxt, font=("Microsoft JhengHei UI", 9))

                # 第4行：雨天
                wtxt = weather.get(d, "") if in_month else ""
                fill = colors["weather"] if wtxt else colors["white"]
                c.create_rectangle(x0, y0+3*row_h, x0+cell_w, y0+4*row_h, fill=fill, outline=colors["grid"])
                c.create_text(x0+cell_w/2, y0+row_h*3.5, text=wtxt, font=("Microsoft JhengHei UI", 9))

                # 第5個視覺區塊：工作日數累計
                is_work = in_month and d.weekday() < 5 and d not in holiday_ex
                if is_work:
                    work_count += 1
                    txt = str(work_count)
                else:
                    txt = ""
                c.create_rectangle(x0, y0+4*row_h, x0+cell_w, y0+5*row_h, fill="#ffffff", outline=colors["grid"])
                c.create_text(x0+cell_w/2, y0+row_h*4.5, text=txt, font=("Microsoft JhengHei UI", 9, "bold"))

        c.create_text(
            10, height - 10,
            text="說明：週六週日紅粉色；假日粉綠色；疏運與雨天粉棕色；資料關閉前與編輯中會自動儲存。",
            anchor="sw",
            font=("Microsoft JhengHei UI", 9)
        )


if __name__ == "__main__":
    app = App()
    app.mainloop()
