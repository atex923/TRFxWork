# -*- coding: utf-8 -*-
"""
KAGAMI 臺鐵工程本本 V0.3.4.3
- Python 標準函式庫版本：tkinter + sqlite3
- 關閉前自動儲存
- 可建立多個工程
- 開啟時自動載入上次編輯工程
- 基本資料、假期表、晴雨表、鐵路疏運表、週曆總表、計價資料、會議記錄表
- V0.2.6.3：第一分頁保證金新增履約保證金手動修改與保證金型式欄位。
- V0.2.6.7：加寬第一分頁預算/契約金額區塊指定金額欄位，可輸入十億元。
- V0.3.0：改黑白高對比配色，新增選擇/新增資料庫與預設資料庫來源記憶。
- V0.3.1：視窗底色改淺灰，新增資料庫備份按鈕與檔名規則，第一分頁欄位自適應，第三分頁日期底色調整。
- V0.3.2：第四分頁發包工程費計價欄位重整，新增日期選擇、累計計算與金額小數格式。
- V0.3.2.1：第四分頁固定第一欄、金額右對齊、累計依期數計算、日期欄移到瀏覽區最後。
- V0.3.2.2：修正第四分頁點擊空白區死當，點選資料列時整行淺灰底色識別。
- V0.3.3：摘要金額來源調整、第五/第六分頁計價欄位重整、第一分頁新增變更後契約金額。
- V0.3.3.1：修正第五/第六分頁既有資料編輯，並新增第六分頁累計稅金。
- V0.3.3.2：第六分頁可支用額度超額提示改為可支用金額紅字加粗。
- V0.3.3.3：晴雨表備註 -1 強制施工日曆不計工期，施工日曆註記加入天氣與場地內容。
- V0.3.3.4：施工日曆晴雨註記僅顯示有內容的天氣/場地，並移除天氣與場地前綴詞。
- V0.3.3.5：施工日曆晴雨註記過濾晴雨表中的 0.0 / 0 / 空值。
- V0.3.4：新增工程進度估算分頁，依發包契約金額與進度百分比自動計算施作金額。
- V0.3.4.1：修正疏運表變更後施工日曆即時同步，並顯示疏運名稱。
- V0.3.4.2：回復動作歷史擴充為最多 10 次。
- V0.3.4.3：列印設定強化，會議記錄表同步工程大事記，並調整工程名稱與大事記排序。
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
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import json
import html
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog


APP_VERSION = "V0.3.4.3"
APP_RELEASE_SUMMARY = "列印設定強化，會議記錄表同步工程大事記，並調整工程名稱與大事記排序。"
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
APP_CONFIG_FILE = os.path.join(APP_DIR, "TRFxWork_config.json")


def normalize_db_path(path):
    if not path:
        return ""
    return os.path.abspath(os.path.expanduser(path))


def load_app_config():
    try:
        with open(APP_CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def save_app_config(data):
    os.makedirs(APP_DIR, exist_ok=True)
    tmp_path = APP_CONFIG_FILE + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, APP_CONFIG_FILE)


def set_default_database_path(path):
    cfg = load_app_config()
    cfg["database_path"] = normalize_db_path(path)
    cfg["updated_at"] = datetime.now().isoformat(timespec="seconds")
    save_app_config(cfg)


def timestamp_suffix():
    return datetime.now().strftime("%Y%m%d%H%M%S")


def append_timestamp_suffix(path, prefix="_"):
    """確保檔名主體最後帶有時間戳後綴。"""
    folder, filename = os.path.split(normalize_db_path(path))
    stem, ext = os.path.splitext(filename)
    if re.search(r"_\d{14}$", stem):
        return os.path.join(folder, filename)
    return os.path.join(folder, f"{stem}{prefix}{timestamp_suffix()}{ext}")

WEEKDAY_NAMES = ["一", "二", "三", "四", "五", "六", "日"]
PASSWORD_SALT = "1981"

# 黑白高對比 + 淺灰底色配色
OSX_WINDOW_BG = "#eeeeee"
OSX_PANEL_BG = "#f7f7f7"
OSX_PANEL_ALT_BG = "#ffffff"
OSX_BORDER = "#000000"
OSX_TEXT = "#000000"
OSX_MUTED_TEXT = "#333333"
OSX_ACCENT = "#000000"
OSX_ACCENT_HOVER = "#222222"
OSX_SELECTION = "#d9d9d9"
OSX_ENTRY_BG = "#ffffff"
OSX_READONLY_BG = "#eeeeee"
OSX_INPUT_BORDER = "#000000"
OSX_BUDGET_BG = "#f7f7f7"
OSX_AWARD_BG = "#ffffff"
OSX_BUDGET_BOOK_BG = "#f7f7f7"
OSX_CONTRACT_BOOK_BG = "#ffffff"
OSX_WARNING = "#000000"

PROJECT_EXTRA_FIELDS = [
    "contractor", "company_address", "responsible_person", "contact_person",
    "phone", "fax", "tax_id", "purchase_contract_no", "contract_date", "project_description",
    "contract_budget_net", "contract_award_net", "contract_budget_tax", "contract_award_tax",
    "contract_budget_total", "contract_award_total",
    "labor_budget", "labor_award", "deposit_difference", "deposit_performance", "deposit_performance_manual",
    "performance_bond_type", "warranty_bond_type", "deposit_total",
    "final_contract_amount", "warranty_rate", "warranty_deposit",
    "planned_precheck_date", "actual_precheck_date", "planned_acceptance_date", "actual_acceptance_date",
    "settlement_date", "warranty_years", "warranty_end_date", "warranty_note", "performance_bond_rate",
    "budget_total_amount", "budget_unfinished_amount", "budget_input_tax",
    "budget_contract_amount", "budget_contract_tax", "budget_contract_total",
    "budget_labor",
    "budget_mgmt_fee", "budget_self_labor", "budget_self_material", "budget_spare_material",
    "budget_railway_material", "budget_supervision_fee", "budget_freight", "budget_air_pollution_fee",
    "budget_other",
    "award_total_amount", "award_unfinished_amount", "award_input_tax",
    "award_contract_amount", "award_contract_tax", "award_contract_total", "award_base_price",
    "award_contract_budget_ratio", "award_contract_base_ratio", "award_base_budget_ratio",
    "award_labor",
    "award_mgmt_fee", "award_self_labor", "award_self_material", "award_spare_material",
    "award_railway_material", "award_supervision_fee", "award_freight", "award_air_pollution_fee",
    "award_other",
    "admin_change_mgmt_fee",
    "admin_alloc1_amount", "admin_alloc1_0c12", "admin_alloc1_0c11", "admin_alloc1_0c14",
    "admin_alloc2_amount", "admin_alloc2_0c12", "admin_alloc2_0c11", "admin_alloc2_0c14",
    "admin_alloc3_amount", "admin_alloc3_0c12", "admin_alloc3_0c11", "admin_alloc3_0c14",
    "admin_alloc4_amount", "admin_alloc4_0c12", "admin_alloc4_0c11", "admin_alloc4_0c14",
]

MONEY_FIELDS = {
    "contract_budget_net", "contract_award_net", "contract_budget_tax", "contract_award_tax",
    "contract_budget_total", "contract_award_total",
    "labor_budget", "labor_award", "deposit_difference", "deposit_performance", "deposit_total",
    "final_contract_amount", "warranty_deposit",
    "budget_total_amount", "budget_unfinished_amount", "budget_input_tax",
    "budget_contract_amount", "budget_contract_tax", "budget_contract_total", "budget_labor",
    "budget_mgmt_fee", "budget_self_labor", "budget_self_material", "budget_spare_material",
    "budget_railway_material", "budget_supervision_fee", "budget_freight", "budget_air_pollution_fee",
    "budget_other",
    "award_total_amount", "award_unfinished_amount", "award_input_tax",
    "award_contract_amount", "award_contract_tax", "award_contract_total", "award_base_price",
    "award_labor", "award_mgmt_fee", "award_self_labor", "award_self_material", "award_spare_material",
    "award_railway_material", "award_supervision_fee", "award_freight", "award_air_pollution_fee",
    "award_other", "admin_change_mgmt_fee",
    "admin_alloc1_amount", "admin_alloc1_0c12", "admin_alloc1_0c11", "admin_alloc1_0c14",
    "admin_alloc2_amount", "admin_alloc2_0c12", "admin_alloc2_0c11", "admin_alloc2_0c14",
    "admin_alloc3_amount", "admin_alloc3_0c12", "admin_alloc3_0c11", "admin_alloc3_0c14",
    "admin_alloc4_amount", "admin_alloc4_0c12", "admin_alloc4_0c11", "admin_alloc4_0c14",
}


PAYMENT_CONTRACT_FIELDS = [
    "period_no", "estimated_amount_taxed", "billing_amount_untaxed",
    "billing_business_tax", "paid_amount_untaxed", "paid_business_tax", "retention_amount",
    "cumulative_billing_amount", "cumulative_billing_tax", "cumulative_retention_amount",
    "cumulative_paid_amount", "cumulative_paid_tax", "progress_percent", "payment_period", "worked_days",
    "contractor_submit_date", "supervision_submit_date", "owner_payment_date", "reimbursement_submit_date",
]
PAYMENT_CONTRACT_HEADINGS = [
    "期數", "估驗金額(含稅)", "計價金額(未稅)", "計價營業稅", "實發金額(未稅)", "實發營業稅",
    "保留款", "累計計價金額", "累計計價營業稅", "累計保留款", "累計實發金額",
    "累計實發營業稅", "計價工程進度", "計價期間", "已工作日數",
    "廠商提送日期", "監造提送日期", "主辦計價日期", "送件核銷日期",
]
PAYMENT_CONTRACT_WIDTHS = [
    70, 145, 145, 130, 145, 130,
    120, 150, 150, 140, 150,
    150, 120, 180, 100,
    120, 120, 120, 120,
]
PAYMENT_CONTRACT_DATE_FIELDS = {
    "contractor_submit_date", "supervision_submit_date", "owner_payment_date", "reimbursement_submit_date"
}
PAYMENT_CONTRACT_MONEY_FIELDS = {
    "estimated_amount_taxed", "billing_amount_untaxed", "billing_business_tax",
    "paid_amount_untaxed", "paid_business_tax", "retention_amount",
    "cumulative_billing_amount", "cumulative_billing_tax", "cumulative_retention_amount",
    "cumulative_paid_amount", "cumulative_paid_tax",
}
PAYMENT_CONTRACT_CUMULATIVE_FIELDS = {
    "cumulative_billing_amount", "cumulative_billing_tax", "cumulative_retention_amount",
    "cumulative_paid_amount", "cumulative_paid_tax",
}


PAYMENT_OTHER_FIELDS = [
    "period_no", "payment_date", "payment_item", "payment_amount", "cumulative_amount",
    "management_fee_allocation", "cumulative_management_fee_allocation", "allocated_fee",
    "travel_fee_0c12", "overtime_fee_0c11", "other_0c14",
]
PAYMENT_OTHER_HEADINGS = [
    "期數", "計價日期", "計價項目", "計價金額", "累計金額",
    "提撥管理費", "累計提撥管理費", "分攤後費用",
    "0C12差費", "0C11加班費", "0C14其他",
]
PAYMENT_OTHER_WIDTHS = [80, 120, 140, 140, 140, 140, 160, 140, 130, 130, 130]
PAYMENT_OTHER_DATE_FIELDS = {"payment_date"}
PAYMENT_OTHER_LEFT_FIELDS = {"payment_item"}
PAYMENT_OTHER_MONEY_FIELDS = {
    "payment_amount", "cumulative_amount", "management_fee_allocation",
    "cumulative_management_fee_allocation", "allocated_fee", "travel_fee_0c12",
    "overtime_fee_0c11", "other_0c14",
}
PAYMENT_OTHER_CUMULATIVE_FIELDS = {"cumulative_amount", "cumulative_management_fee_allocation"}
PAYMENT_OTHER_ITEM_OPTIONS = ["管理費", "差費", "延時工資"]

PAYMENT_ADMIN_FIELDS = [
    "period_no", "payment_date", "travel_fee_billing", "overtime_fee_billing", "other_fee_billing",
    "tax_amount", "cumulative_tax_amount", "current_amount", "cumulative_amount",
]
PAYMENT_ADMIN_HEADINGS = [
    "期數", "計價日期", "0C12差費計價", "0C11加班費計價", "0C14其他計價",
    "稅金", "累計稅金", "本次計價金額", "累計金額",
]
PAYMENT_ADMIN_WIDTHS = [80, 120, 150, 150, 150, 120, 140, 150, 150]
PAYMENT_ADMIN_DATE_FIELDS = {"payment_date"}
PAYMENT_ADMIN_MONEY_FIELDS = {
    "travel_fee_billing", "overtime_fee_billing", "other_fee_billing",
    "tax_amount", "cumulative_tax_amount", "current_amount", "cumulative_amount",
}
PAYMENT_ADMIN_CUMULATIVE_FIELDS = {"cumulative_tax_amount", "current_amount", "cumulative_amount"}


PROGRESS_ESTIMATE_FIELDS = [
    "item_no", "month", "estimated_progress", "estimated_amount", "actual_progress", "actual_amount",
]
PROGRESS_ESTIMATE_HEADINGS = [
    "項次", "月份", "預估進度", "預估施作金額", "實際進度", "實際施作金額",
]
PROGRESS_ESTIMATE_WIDTHS = [90, 120, 120, 170, 120, 170]
PROGRESS_ESTIMATE_PERCENT_FIELDS = {"estimated_progress", "actual_progress"}
PROGRESS_ESTIMATE_MONEY_FIELDS = {"estimated_amount", "actual_amount"}
PROGRESS_ESTIMATE_READONLY_FIELDS = {"estimated_amount", "actual_amount"}

CHANGE_AWARD_FIELDS = [
    "change_award_total_amount", "change_award_unfinished_amount", "change_award_input_tax",
    "change_award_contract_total", "change_award_contract_amount", "change_award_contract_tax",
    "change_award_base_price", "change_award_contract_budget_ratio", "change_award_contract_base_ratio", "change_award_base_budget_ratio",
    "change_award_labor", "change_award_mgmt_fee", "change_award_self_labor", "change_award_self_material",
    "change_award_spare_material", "change_award_railway_material", "change_award_supervision_fee",
    "change_award_freight", "change_award_other", "change_award_air_pollution_fee",
]
CHANGE_AWARD_MONEY_FIELDS = {
    "change_award_total_amount", "change_award_unfinished_amount", "change_award_input_tax",
    "change_award_contract_total", "change_award_contract_amount", "change_award_contract_tax",
    "change_award_base_price", "change_award_labor", "change_award_mgmt_fee", "change_award_self_labor",
    "change_award_self_material", "change_award_spare_material", "change_award_railway_material",
    "change_award_supervision_fee", "change_award_freight", "change_award_other", "change_award_air_pollution_fee",
}


def period_sort_key(period_text, original_index=0):
    text = str(period_text or "").strip()
    if not text:
        return (2, original_index)
    cleaned = text.replace("第", "").replace("期", "").strip()
    try:
        return (0, Decimal(cleaned), original_index)
    except InvalidOperation:
        parts = re.split(r"(\d+(?:\.\d+)?)", cleaned)
        key = []
        for part in parts:
            if not part:
                continue
            try:
                key.append((0, Decimal(part)))
            except InvalidOperation:
                key.append((1, part.lower()))
        return (1, key, original_index)


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


def add_years(d, years):
    if not d:
        return None
    years = int(float(years or 0))
    try:
        return d.replace(year=d.year + years)
    except ValueError:
        return d.replace(month=2, day=28, year=d.year + years)


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
    cfg = load_app_config()
    configured = normalize_db_path(cfg.get("database_path", ""))
    if configured:
        folder = os.path.dirname(configured)
        if os.path.exists(configured) or (folder and os.path.isdir(folder)):
            return configured
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
            open_date TEXT,
            award_date TEXT DEFAULT ''
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
            note TEXT,
            period_no TEXT DEFAULT '',
            contractor_submit_date TEXT DEFAULT '',
            supervision_submit_date TEXT DEFAULT '',
            owner_payment_date TEXT DEFAULT '',
            reimbursement_submit_date TEXT DEFAULT '',
            estimated_amount_taxed TEXT DEFAULT '',
            billing_amount_untaxed TEXT DEFAULT '',
            billing_business_tax TEXT DEFAULT '',
            paid_amount_untaxed TEXT DEFAULT '',
            paid_business_tax TEXT DEFAULT '',
            retention_amount TEXT DEFAULT '',
            cumulative_billing_amount TEXT DEFAULT '',
            cumulative_billing_tax TEXT DEFAULT '',
            cumulative_retention_amount TEXT DEFAULT '',
            cumulative_paid_amount TEXT DEFAULT '',
            cumulative_paid_tax TEXT DEFAULT '',
            progress_percent TEXT DEFAULT '',
            payment_period TEXT DEFAULT '',
            worked_days TEXT DEFAULT ''
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
        CREATE TABLE IF NOT EXISTS progress_estimates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            item_no TEXT DEFAULT '',
            month TEXT DEFAULT '',
            estimated_progress TEXT DEFAULT '',
            estimated_amount TEXT DEFAULT '',
            actual_progress TEXT DEFAULT '',
            actual_amount TEXT DEFAULT ''
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
        c.execute("""
        CREATE TABLE IF NOT EXISTS budget_books (
            project_id INTEGER,
            area TEXT,
            rows_json TEXT,
            source_file TEXT,
            updated_at TEXT,
            PRIMARY KEY(project_id, area)
        )
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS change_records (
            project_id INTEGER,
            change_no TEXT,
            fields_json TEXT,
            demand_json TEXT,
            confirm_json TEXT,
            budget_json TEXT,
            source_file TEXT,
            updated_at TEXT,
            PRIMARY KEY(project_id, change_no)
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
        for field in PAYMENT_CONTRACT_FIELDS:
            try:
                c.execute(f"ALTER TABLE payment_contract ADD COLUMN {field} TEXT DEFAULT ''")
            except sqlite3.OperationalError:
                pass
        for field in PAYMENT_OTHER_FIELDS:
            try:
                c.execute(f"ALTER TABLE payment_other ADD COLUMN {field} TEXT DEFAULT ''")
            except sqlite3.OperationalError:
                pass
        for field in PAYMENT_ADMIN_FIELDS:
            try:
                c.execute(f"ALTER TABLE payment_admin ADD COLUMN {field} TEXT DEFAULT ''")
            except sqlite3.OperationalError:
                pass
        for field in PROGRESS_ESTIMATE_FIELDS:
            try:
                c.execute(f"ALTER TABLE progress_estimates ADD COLUMN {field} TEXT DEFAULT ''")
            except sqlite3.OperationalError:
                pass
        try:
            c.execute("ALTER TABLE bids ADD COLUMN award_date TEXT DEFAULT ''")
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
            "payment_contract", "payment_other", "payment_admin", "progress_estimates",
            "execution_records", "execution_status", "project_milestones",
            "holiday_project_excludes", "workday_project_excludes", "railway_project_excludes",
            "budget_books", "change_records"
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

    def save_budget_book(self, pid, area, rows, source_file=""):
        self.conn.execute(
            "INSERT OR REPLACE INTO budget_books(project_id, area, rows_json, source_file, updated_at) VALUES(?,?,?,?,?)",
            (pid, area, json.dumps(rows, ensure_ascii=False), source_file, datetime.now().isoformat(timespec="seconds"))
        )
        self.conn.commit()

    def budget_book(self, pid, area):
        row = self.conn.execute("SELECT rows_json, source_file FROM budget_books WHERE project_id=? AND area=?", (pid, area)).fetchone()
        if not row:
            return [], ""
        try:
            return json.loads(row["rows_json"] or "[]"), row["source_file"] or ""
        except json.JSONDecodeError:
            return [], row["source_file"] or ""

    def change_numbers(self, pid):
        rows = self.conn.execute(
            "SELECT change_no FROM change_records WHERE project_id=? ORDER BY CAST(change_no AS INTEGER), change_no",
            (pid,)
        ).fetchall()
        return [r["change_no"] for r in rows]

    def change_record(self, pid, change_no):
        row = self.conn.execute("SELECT * FROM change_records WHERE project_id=? AND change_no=?", (pid, str(change_no))).fetchone()
        if not row:
            return None
        def loads(key, default):
            try:
                return json.loads(row[key] or default)
            except json.JSONDecodeError:
                return json.loads(default)
        return {
            "change_no": row["change_no"],
            "fields": loads("fields_json", "{}"),
            "demand": loads("demand_json", "[]"),
            "confirm": loads("confirm_json", "[]"),
            "budget": loads("budget_json", "[]"),
            "source_file": row["source_file"] or "",
        }

    def save_change_record(self, pid, change_no, fields, demand_rows, confirm_rows, budget_rows, source_file=""):
        self.conn.execute(
            """
            INSERT OR REPLACE INTO change_records
            (project_id, change_no, fields_json, demand_json, confirm_json, budget_json, source_file, updated_at)
            VALUES(?,?,?,?,?,?,?,?)
            """,
            (
                pid, str(change_no), json.dumps(fields, ensure_ascii=False),
                json.dumps(demand_rows, ensure_ascii=False), json.dumps(confirm_rows, ensure_ascii=False),
                json.dumps(budget_rows, ensure_ascii=False), source_file, datetime.now().isoformat(timespec="seconds")
            )
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
        if table == "progress_estimates":
            return self.conn.execute("SELECT * FROM progress_estimates WHERE project_id=? ORDER BY id", (pid,)).fetchall()
        return self.conn.execute(f"SELECT * FROM {table} WHERE project_id=? ORDER BY day, id", (pid,)).fetchall()

    def bids(self, pid):
        return self.conn.execute("SELECT * FROM bids WHERE project_id=? ORDER BY CAST(round_no AS INTEGER) DESC, id DESC", (pid,)).fetchall()

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
                    "INSERT INTO bids(project_id, round_no, online_date, open_date, award_date) VALUES(?,?,?,?,?)",
                    (pid, r.get("round_no", 1), r.get("online_date", ""), r.get("open_date", ""), r.get("award_date", ""))
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
        elif table == "payment_contract":
            cols = ["project_id", "day", "item", "voucher_no", "amount", "note"] + PAYMENT_CONTRACT_FIELDS
            placeholders = ",".join("?" for _ in cols)
            sql = f"INSERT INTO payment_contract({','.join(cols)}) VALUES({placeholders})"
            for r in rows:
                legacy_day = r.get("owner_payment_date") or r.get("contractor_submit_date", "")
                legacy_item = r.get("period_no", "")
                legacy_amount = r.get("estimated_amount_taxed") or r.get("paid_amount_untaxed") or 0
                legacy_note = r.get("payment_period", "")
                values = [pid, legacy_day, legacy_item, "", legacy_amount, legacy_note]
                values.extend(r.get(field, "") for field in PAYMENT_CONTRACT_FIELDS)
                self.conn.execute(sql, values)
        elif table == "payment_other":
            cols = ["project_id", "day", "item", "voucher_no", "amount", "note"] + PAYMENT_OTHER_FIELDS
            placeholders = ",".join("?" for _ in cols)
            sql = f"INSERT INTO payment_other({','.join(cols)}) VALUES({placeholders})"
            for r in rows:
                values = [
                    pid,
                    r.get("payment_date") or r.get("day", ""),
                    r.get("payment_item") or r.get("item", ""),
                    r.get("period_no", ""),
                    r.get("payment_amount") or r.get("amount", 0),
                    r.get("note", ""),
                ]
                values.extend(r.get(field, "") for field in PAYMENT_OTHER_FIELDS)
                self.conn.execute(sql, values)
        elif table == "payment_admin":
            cols = ["project_id", "day", "item", "voucher_no", "amount", "note"] + PAYMENT_ADMIN_FIELDS
            placeholders = ",".join("?" for _ in cols)
            sql = f"INSERT INTO payment_admin({','.join(cols)}) VALUES({placeholders})"
            for r in rows:
                values = [
                    pid,
                    r.get("payment_date") or r.get("day", ""),
                    r.get("period_no") or r.get("item", ""),
                    r.get("period_no", ""),
                    r.get("current_amount") or r.get("amount", 0),
                    r.get("note", ""),
                ]
                values.extend(r.get(field, "") for field in PAYMENT_ADMIN_FIELDS)
                self.conn.execute(sql, values)
        elif table == "progress_estimates":
            cols = ["project_id"] + PROGRESS_ESTIMATE_FIELDS
            placeholders = ",".join("?" for _ in cols)
            sql = f"INSERT INTO progress_estimates({','.join(cols)}) VALUES({placeholders})"
            for r in rows:
                values = [pid]
                values.extend(r.get(field, "") for field in PROGRESS_ESTIMATE_FIELDS)
                self.conn.execute(sql, values)
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
    def __init__(self, master, columns, headings, widths, on_changed=None, can_edit=None, add_command=None, edit_command=None, height=12):
        super().__init__(master)
        self.columns = columns
        self.headings = headings
        self.on_changed = on_changed
        self.can_edit = can_edit or (lambda: True)
        self.add_command = add_command or self.add_row
        self.edit_command = edit_command
        self.sort_descending = {}
        self.tree = ttk.Treeview(self, columns=columns, show="headings", height=height, style="Grid.Treeview")
        vs = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.tag_configure("pink", background="#e6e6e6")
        self.tree.tag_configure("red", foreground="#000000")
        self.tree.tag_configure("year_sep", background="#000000", foreground="#ffffff")

        for col, head, width in zip(columns, headings, widths):
            self.tree.heading(col, text=head, command=lambda c=col: self.sort_by_column(c))
            fixed_cols = {
                "exclude", "day", "name", "note", "morning", "afternoon", "typhoon", "site",
                "awarded", "round_no", "online_date", "open_date", "award_date"
            }
            self.tree.column(col, width=width, minwidth=width, anchor="center", stretch=False if col in fixed_cols else True)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        btns = ttk.Frame(self)
        btns.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
        ttk.Button(btns, text="新增一列", command=self.add_command).pack(side="left", padx=3)
        ttk.Button(btns, text="編輯選取列", command=self.run_edit_command).pack(side="left", padx=3)
        ttk.Button(btns, text="刪除選取列", command=self.delete_row).pack(side="left", padx=3)
        ttk.Button(btns, text="上移", command=lambda: self.move_selected(-1)).pack(side="left", padx=3)
        ttk.Button(btns, text="下移", command=lambda: self.move_selected(1)).pack(side="left", padx=3)
        self.tree.bind("<Double-1>", self.run_edit_command)
        self.tree.bind("<Button-1>", self.on_tree_click)

    def run_edit_command(self, event=None):
        if event is not None:
            row = self.tree.identify_row(event.y)
            if row:
                self.tree.selection_set(row)
                self.tree.focus(row)
            else:
                return "break"
        if not self.tree.focus():
            selected = self.tree.selection()
            if selected:
                self.tree.focus(selected[0])
        if self.edit_command:
            return self.edit_command()
        return self.edit_row()

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
        item = self.tree.insert("", "end", values=values)
        self.tree.selection_set(item)
        self.tree.focus(item)
        self.tree.see(item)
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

    def move_selected(self, delta):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        selected = list(self.tree.selection())
        if not selected:
            return
        children = list(self.tree.get_children(""))
        items = selected if delta < 0 else list(reversed(selected))
        moved = False
        for item in items:
            index = children.index(item)
            new_index = index + delta
            if new_index < 0 or new_index >= len(children):
                continue
            self.tree.move(item, "", new_index)
            children = list(self.tree.get_children(""))
            moved = True
        if moved:
            self.tree.selection_set(selected)
            self.tree.focus(selected[0])
            self.tree.see(selected[0])
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
        reverse = self.sort_descending.get(col, False)
        dated_cols = {"day", "online_date", "open_date", "award_date", "start_date", "deadline_date", "received_date", "contractor_submit_date", "supervision_submit_date", "owner_payment_date", "reimbursement_submit_date"}
        valid_rows = []
        blank_rows = []
        for item in self.tree.get_children(""):
            raw_value = self.tree.set(item, col)
            if col in dated_cols:
                parsed = parse_date(raw_value)
                if parsed:
                    valid_rows.append((parsed, item))
                else:
                    blank_rows.append(((raw_value or "").strip(), item))
            else:
                text = (raw_value or "").strip()
                try:
                    key = (0, float(text.replace(",", "").replace("元", "")))
                except ValueError:
                    key = (1, text.lower())
                valid_rows.append((key, item))
        valid_rows.sort(key=lambda x: x[0], reverse=reverse)
        for index, (_, item) in enumerate(valid_rows + blank_rows):
            self.tree.move(item, "", index)
        self.sort_descending[col] = not reverse
        self.changed()


class FixedFirstColumnTree(ttk.Frame):
    """第一欄固定、其餘欄位可橫向捲動的 Treeview 包裝元件。"""
    def __init__(self, master, columns, headings, widths, on_changed=None, can_edit=None,
                 add_command=None, edit_command=None, height=12, money_columns=None, date_columns=None):
        super().__init__(master)
        self.columns = list(columns)
        self.headings = list(headings)
        self.widths = list(widths)
        self.on_changed = on_changed
        self.can_edit = can_edit or (lambda: True)
        self.add_command = add_command or self.add_row
        self.edit_command = edit_command or self.edit_row
        self.money_columns = set(money_columns or [])
        self.date_columns = set(date_columns or [])
        self.sort_descending = {}
        self._syncing_selection = False
        self._iid_counter = 0

        self.fixed_column = self.columns[0]
        self.scroll_columns = self.columns[1:]
        fixed_heading = self.headings[0]
        scroll_headings = self.headings[1:]
        fixed_width = self.widths[0]
        scroll_widths = self.widths[1:]

        table = ttk.Frame(self)
        table.grid(row=0, column=0, columnspan=2, sticky="nsew")
        table.grid_rowconfigure(0, weight=1)
        table.grid_columnconfigure(1, weight=1)

        self.fixed_tree = ttk.Treeview(table, columns=[self.fixed_column], show="headings", height=height, style="Grid.Treeview")
        self.tree = ttk.Treeview(table, columns=self.scroll_columns, show="headings", height=height, style="Grid.Treeview")
        vs = ttk.Scrollbar(table, orient="vertical", command=self._yview)
        hs = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.fixed_tree.configure(yscrollcommand=lambda *args: None)

        self.fixed_tree.heading(self.fixed_column, text=fixed_heading, command=lambda c=self.fixed_column: self.sort_by_column(c))
        self.fixed_tree.column(self.fixed_column, width=fixed_width, minwidth=fixed_width, anchor="center", stretch=False)
        for col, head, width in zip(self.scroll_columns, scroll_headings, scroll_widths):
            self.tree.heading(col, text=head, command=lambda c=col: self.sort_by_column(c))
            anchor = "e" if col in self.money_columns else "center"
            self.tree.column(col, width=width, minwidth=width, anchor=anchor, stretch=False)

        self.fixed_tree.grid(row=0, column=0, sticky="ns")
        self.tree.grid(row=0, column=1, sticky="nsew")
        vs.grid(row=0, column=2, sticky="ns")
        hs.grid(row=1, column=1, sticky="ew")

        self.fixed_tree.tag_configure("gridline", background="#ffffff")
        self.tree.tag_configure("gridline", background="#ffffff")
        self.fixed_tree.tag_configure("selected_row", background="#e6e6e6", foreground="#000000")
        self.tree.tag_configure("selected_row", background="#e6e6e6", foreground="#000000")
        # 不使用 <<TreeviewSelect>> 雙向同步，避免兩個 Treeview 在空白區或重複選取時互相觸發事件。
        # 改由 <Button-1> 明確判斷是否點到資料列，只有點到資料列才同步選取。
        self.fixed_tree.bind("<Button-1>", lambda e: self._on_tree_click(e, self.fixed_tree, self.tree))
        self.tree.bind("<Button-1>", lambda e: self._on_tree_click(e, self.tree, self.fixed_tree))
        self.fixed_tree.bind("<Double-1>", lambda e: self.run_edit_command())
        self.tree.bind("<Double-1>", lambda e: self.run_edit_command())
        self.fixed_tree.bind("<MouseWheel>", self._on_mousewheel)
        self.tree.bind("<MouseWheel>", self._on_mousewheel)

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        btns = ttk.Frame(self)
        btns.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
        ttk.Button(btns, text="新增一列", command=self.add_command).pack(side="left", padx=3)
        ttk.Button(btns, text="編輯選取列", command=self.run_edit_command).pack(side="left", padx=3)
        ttk.Button(btns, text="刪除選取列", command=self.delete_row).pack(side="left", padx=3)

    def _yview(self, *args):
        self.fixed_tree.yview(*args)
        self.tree.yview(*args)

    def _on_mousewheel(self, event):
        delta = -1 if event.delta > 0 else 1
        self.fixed_tree.yview_scroll(delta, "units")
        self.tree.yview_scroll(delta, "units")
        return "break"

    def _clear_visual_selection(self):
        for item in self.fixed_tree.get_children(""):
            self.fixed_tree.item(item, tags=("gridline",))
        for item in self.tree.get_children(""):
            self.tree.item(item, tags=("gridline",))

    def _apply_visual_selection(self, item):
        self._clear_visual_selection()
        if item and self.fixed_tree.exists(item):
            self.fixed_tree.item(item, tags=("selected_row",))
        if item and self.tree.exists(item):
            self.tree.item(item, tags=("selected_row",))

    def _safe_select_item(self, item):
        if not item or not self.fixed_tree.exists(item) or not self.tree.exists(item):
            self._syncing_selection = True
            try:
                self.fixed_tree.selection_remove(self.fixed_tree.selection())
                self.tree.selection_remove(self.tree.selection())
                self.fixed_tree.focus("")
                self.tree.focus("")
                self._clear_visual_selection()
            finally:
                self._syncing_selection = False
            return
        self._syncing_selection = True
        try:
            self.fixed_tree.selection_set(item)
            self.tree.selection_set(item)
            self.fixed_tree.focus(item)
            self.tree.focus(item)
            self._apply_visual_selection(item)
        finally:
            self._syncing_selection = False

    def _on_tree_click(self, event, source, target):
        region = source.identify("region", event.x, event.y)
        if region == "heading":
            return None
        item = source.identify_row(event.y)
        if not item:
            self._safe_select_item("")
            return "break"
        self._safe_select_item(item)
        return "break"

    def _sync_selection(self, source, target):
        if self._syncing_selection:
            return
        self._syncing_selection = True
        try:
            selection = source.selection()
            valid_selection = tuple(item for item in selection if source.exists(item) and target.exists(item))
            if valid_selection:
                target.selection_set(valid_selection)
                target.focus(valid_selection[0])
                self._apply_visual_selection(valid_selection[0])
            else:
                target.selection_remove(target.selection())
                self._clear_visual_selection()
        finally:
            self._syncing_selection = False

    def run_edit_command(self):
        if self.edit_command:
            return self.edit_command()

    def focus(self):
        item = self.tree.focus() or self.fixed_tree.focus()
        if not item:
            selected = self.tree.selection() or self.fixed_tree.selection()
            item = selected[0] if selected else ""
        return item

    def selection(self):
        return self.tree.selection() or self.fixed_tree.selection()

    def add_row(self, values=None):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        self.insert_row(values or [""] * len(self.columns), select=True)
        self.changed()

    def edit_row(self):
        return None

    def insert_row(self, values, index="end", select=False):
        values = (list(values) + [""] * len(self.columns))[:len(self.columns)]
        self._iid_counter += 1
        iid = f"row_{self._iid_counter}"
        self.fixed_tree.insert("", index, iid=iid, values=[values[0]], tags=("gridline",))
        self.tree.insert("", index, iid=iid, values=values[1:], tags=("gridline",))
        if select:
            self.fixed_tree.selection_set(iid)
            self.tree.selection_set(iid)
            self.fixed_tree.focus(iid)
            self.tree.focus(iid)
            self.fixed_tree.see(iid)
            self.tree.see(iid)
        return iid

    def set_item_values(self, item, values):
        if not item:
            return
        values = (list(values) + [""] * len(self.columns))[:len(self.columns)]
        is_selected = item in self.selection()
        row_tags = ("selected_row",) if is_selected else ("gridline",)
        if self.fixed_tree.exists(item):
            self.fixed_tree.item(item, values=[values[0]], tags=row_tags)
        if self.tree.exists(item):
            self.tree.item(item, values=values[1:], tags=row_tags)

    def get_item_values(self, item):
        if not item or not self.tree.exists(item):
            return [""] * len(self.columns)
        first = list(self.fixed_tree.item(item, "values")) if self.fixed_tree.exists(item) else [""]
        rest = list(self.tree.item(item, "values"))
        return (first + rest + [""] * len(self.columns))[:len(self.columns)]

    def get_focused_values(self):
        return self.get_item_values(self.focus())

    def delete_row(self):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        for item in list(self.selection()):
            if self.fixed_tree.exists(item):
                self.fixed_tree.delete(item)
            if self.tree.exists(item):
                self.tree.delete(item)
        self.changed()

    def set_rows(self, rows):
        self.fixed_tree.delete(*self.fixed_tree.get_children())
        self.tree.delete(*self.tree.get_children())
        self._iid_counter = 0
        for row in rows:
            self.insert_row(row, select=False)

    def get_rows(self):
        return [self.get_item_values(item) for item in self.tree.get_children()]

    def apply_row_tags(self, tag_func):
        for item in self.tree.get_children():
            values = self.get_item_values(item)
            tag = tag_func(values)
            tags = (tag,) if tag else ()
            self.fixed_tree.item(item, tags=tags)
            self.tree.item(item, tags=tags)

    def changed(self):
        if self.on_changed:
            self.on_changed()

    def sort_by_column(self, col):
        reverse = self.sort_descending.get(col, False)
        indexes = {field: idx for idx, field in enumerate(self.columns)}
        idx = indexes.get(col, 0)
        rows = self.get_rows()

        def natural_key(text):
            text = (text or "").strip()
            parts = re.split(r"(\d+(?:\.\d+)?)", text)
            key = []
            for part in parts:
                if not part:
                    continue
                try:
                    key.append((0, Decimal(part)))
                except InvalidOperation:
                    key.append((1, part.lower()))
            return key or [(2, "")]

        def key(row):
            raw = row[idx] if idx < len(row) else ""
            if col in self.date_columns:
                parsed = parse_date(raw)
                return (0, parsed.toordinal()) if parsed else (2, str(raw or ""))
            if col in self.money_columns:
                cleaned = str(raw or "").replace(",", "").replace("元", "").replace("%", "").strip()
                try:
                    return (0, Decimal(cleaned or "0"))
                except InvalidOperation:
                    return (2, cleaned)
            return (1, natural_key(str(raw or "")))

        rows.sort(key=key, reverse=reverse)
        self.set_rows(rows)
        self.sort_descending[col] = not reverse
        self.changed()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1250x820")
        self.minsize(1100, 720)
        self.resizable(True, True)

        self.database_path = resolve_database_path()
        self.db = DB(self.database_path)
        set_default_database_path(self.database_path)
        self.current_project_id = None
        self.loading = False
        self.dirty = False
        self.project_password_hash = ""
        self.edit_unlocked = True
        self.edit_widgets = []
        self.undo_history = []
        self.last_state = None
        self.restoring = False
        self.recalculating = False
        self.save_after_id = None
        self.readonly_basic_keys = set()
        self.data_edit_enabled_var = tk.BooleanVar(value=True)

        self.configure(bg=OSX_WINDOW_BG)
        self.option_add("*Font", "{Microsoft JhengHei UI} 10")
        self.style = ttk.Style()
        self.setup_osx_style()

        self.build_ui()
        self.load_projects()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(3000, self.auto_save_loop)

    def setup_osx_style(self):
        """套用黑白高對比 + 淺灰底色配色。"""
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        base_font = ("Microsoft JhengHei UI", 10)
        bold_font = ("Microsoft JhengHei UI", 10, "bold")
        title_font = ("Microsoft JhengHei UI", 11, "bold")

        self.style.configure(".", background=OSX_WINDOW_BG, foreground=OSX_TEXT, font=base_font)
        self.style.configure("TFrame", background=OSX_WINDOW_BG)
        self.style.configure("TLabel", background=OSX_WINDOW_BG, foreground=OSX_TEXT, font=base_font)
        self.style.configure("TCheckbutton", background=OSX_WINDOW_BG, foreground=OSX_TEXT, font=base_font)
        self.style.configure("TRadiobutton", background=OSX_WINDOW_BG, foreground=OSX_TEXT, font=base_font)

        self.style.configure(
            "TButton",
            background=OSX_PANEL_ALT_BG,
            foreground=OSX_TEXT,
            bordercolor=OSX_BORDER,
            lightcolor=OSX_PANEL_BG,
            darkcolor=OSX_BORDER,
            focuscolor=OSX_ACCENT,
            padding=(8, 4),
            relief="flat",
            font=base_font,
        )
        self.style.map(
            "TButton",
            background=[("pressed", "#000000"), ("active", "#e6e6e6"), ("disabled", OSX_PANEL_ALT_BG)],
            foreground=[("pressed", "#ffffff"), ("disabled", OSX_MUTED_TEXT)],
        )

        self.style.configure(
            "TEntry",
            fieldbackground=OSX_ENTRY_BG,
            background=OSX_ENTRY_BG,
            foreground=OSX_TEXT,
            bordercolor=OSX_BORDER,
            lightcolor=OSX_ENTRY_BG,
            darkcolor=OSX_BORDER,
            insertcolor=OSX_TEXT,
            padding=3,
            relief="flat",
        )
        self.style.map(
            "TEntry",
            fieldbackground=[("readonly", OSX_READONLY_BG), ("disabled", OSX_PANEL_ALT_BG), ("!disabled", OSX_ENTRY_BG)],
            foreground=[("disabled", OSX_TEXT), ("readonly", OSX_TEXT)],
            bordercolor=[("focus", OSX_ACCENT), ("!focus", OSX_BORDER)],
        )
        self.style.configure(
            "AdminAvailable.TEntry",
            fieldbackground=OSX_READONLY_BG,
            background=OSX_READONLY_BG,
            foreground=OSX_TEXT,
            font=base_font,
            bordercolor=OSX_BORDER,
            padding=3,
            relief="flat",
        )
        self.style.map(
            "AdminAvailable.TEntry",
            fieldbackground=[("readonly", OSX_READONLY_BG), ("disabled", OSX_READONLY_BG), ("!disabled", OSX_READONLY_BG)],
            foreground=[("readonly", OSX_TEXT), ("disabled", OSX_TEXT), ("!disabled", OSX_TEXT)],
            bordercolor=[("focus", OSX_ACCENT), ("!focus", OSX_BORDER)],
        )
        self.style.configure(
            "AdminAvailableWarning.TEntry",
            fieldbackground=OSX_READONLY_BG,
            background=OSX_READONLY_BG,
            foreground="#c00000",
            font=bold_font,
            bordercolor=OSX_BORDER,
            padding=3,
            relief="flat",
        )
        self.style.map(
            "AdminAvailableWarning.TEntry",
            fieldbackground=[("readonly", OSX_READONLY_BG), ("disabled", OSX_READONLY_BG), ("!disabled", OSX_READONLY_BG)],
            foreground=[("readonly", "#c00000"), ("disabled", "#c00000"), ("!disabled", "#c00000")],
            bordercolor=[("focus", OSX_ACCENT), ("!focus", OSX_BORDER)],
        )

        self.style.configure(
            "TCombobox",
            fieldbackground=OSX_ENTRY_BG,
            background=OSX_ENTRY_BG,
            foreground=OSX_TEXT,
            bordercolor=OSX_BORDER,
            arrowcolor=OSX_MUTED_TEXT,
            padding=3,
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", OSX_READONLY_BG), ("disabled", OSX_PANEL_ALT_BG), ("!disabled", OSX_ENTRY_BG)],
            foreground=[("disabled", OSX_TEXT), ("readonly", OSX_TEXT)],
            bordercolor=[("focus", OSX_ACCENT), ("!focus", OSX_BORDER)],
        )

        self.style.configure("TLabelframe", background=OSX_PANEL_BG, bordercolor=OSX_BORDER, relief="solid")
        self.style.configure("TLabelframe.Label", background=OSX_PANEL_BG, foreground=OSX_TEXT, font=title_font)
        self.style.configure("Top.TLabelframe", background=OSX_PANEL_BG, bordercolor=OSX_BORDER, relief="solid")
        self.style.configure("Top.TLabelframe.Label", background=OSX_PANEL_BG, foreground=OSX_TEXT, font=title_font)
        self.style.configure("DayTable.TLabelframe", background=OSX_PANEL_BG, bordercolor=OSX_BORDER, relief="solid")
        self.style.configure("DayTable.TLabelframe.Label", background=OSX_PANEL_BG, foreground=OSX_TEXT, font=("Microsoft JhengHei UI", 14, "bold"))

        self.style.configure("TNotebook", background=OSX_WINDOW_BG, borderwidth=0, tabmargins=(4, 4, 4, 0))
        self.style.configure(
            "TNotebook.Tab",
            background=OSX_PANEL_ALT_BG,
            foreground=OSX_MUTED_TEXT,
            padding=(14, 7),
            bordercolor=OSX_BORDER,
            font=base_font,
        )
        self.style.map(
            "TNotebook.Tab",
            background=[("selected", "#000000"), ("active", "#ffffff")],
            foreground=[("selected", "#ffffff"), ("active", OSX_TEXT)],
        )

        self.style.configure(
            "Treeview",
            background=OSX_PANEL_BG,
            fieldbackground=OSX_PANEL_BG,
            foreground=OSX_TEXT,
            rowheight=26,
            bordercolor=OSX_BORDER,
            relief="flat",
            font=base_font,
        )
        self.style.configure(
            "Treeview.Heading",
            background="#ffffff",
            foreground=OSX_TEXT,
            bordercolor=OSX_BORDER,
            relief="flat",
            font=bold_font,
        )
        self.style.map("Treeview", background=[("selected", OSX_SELECTION)], foreground=[("selected", OSX_TEXT)])
        self.style.configure("Grid.Treeview", background=OSX_PANEL_BG, fieldbackground=OSX_PANEL_BG, rowheight=26, font=base_font, borderwidth=1, relief="solid")
        self.style.map("Grid.Treeview", background=[("selected", "#e6e6e6")], foreground=[("selected", "#000000")])
        self.style.configure("Grid.Treeview.Heading", background="#ffffff", foreground=OSX_TEXT, font=bold_font, borderwidth=1, relief="solid")

    def build_ui(self):
        top_select = ttk.Frame(self, padding=8)
        top_select.pack(fill="x")

        project_select = ttk.Frame(self, padding=(8, 8, 8, 0))
        project_select.pack(fill="x")
        ttk.Label(project_select, text="工程名稱：").pack(side="left")
        self.project_combo = ttk.Combobox(project_select, state="readonly", width=55)
        self.project_combo.pack(side="left", fill="x", expand=True, padx=5)
        self.project_combo.bind("<<ComboboxSelected>>", self.on_project_selected)
        ttk.Button(project_select, text="新增工程", command=self.new_project).pack(side="left", padx=5)

        top_select = ttk.Frame(self, padding=8)
        top_select.pack(fill="x")
        ttk.Button(top_select, text="立即儲存", command=self.save_current).pack(side="left", padx=5)
        ttk.Button(top_select, text="回復上一個動作", command=self.undo_last_action).pack(side="left", padx=5)

        self.edit_password_var = tk.StringVar()
        self.lock_state_var = tk.StringVar(value="未鎖定")
        ttk.Label(top_select, textvariable=self.lock_state_var, foreground=OSX_WARNING).pack(side="left", padx=3)

        ttk.Button(top_select, text="▶", width=3, command=self.toggle_function_panel).pack(side="right", padx=(6, 0))
        self.summary_visible = True
        self.summary_toggle_btn = ttk.Button(top_select, text="▲", width=3, command=self.toggle_summary_area)
        self.summary_toggle_btn.pack(side="right", padx=(6, 0))
        self.status_var = tk.StringVar(value="")
        ttk.Label(top_select, textvariable=self.status_var).pack(side="right")
        self.function_panel = None

        db_bar = ttk.Frame(self, padding=(8, 0, 8, 6))
        db_bar.pack(fill="x")
        ttk.Button(db_bar, text="選擇資料庫", command=self.select_database).pack(side="left", padx=(0, 5))
        ttk.Button(db_bar, text="新增資料庫", command=self.new_database).pack(side="left", padx=(0, 5))
        ttk.Button(db_bar, text="備份資料庫", command=self.backup_database).pack(side="left", padx=(0, 8))
        ttk.Label(db_bar, text="現在使用的資料庫：").pack(side="left")
        self.database_path_var = tk.StringVar(value=self.database_path)
        self.database_path_label = ttk.Label(db_bar, textvariable=self.database_path_var, foreground=OSX_TEXT)
        self.database_path_label.pack(side="left", fill="x", expand=True)

        self.summary = ttk.LabelFrame(self, text="工程辦理情形摘要", padding=8, style="Top.TLabelframe")
        self.summary.pack(fill="x", padx=8, pady=(0, 8))
        self.summary_vars = {}
        labels = [
            ("開工時間", "start"),
            ("預定完工時間", "finish1"),
            ("修正後預訂完工時間", "finish2"),
            ("已經過多少施工日數", "elapsed"),
            ("到今天日期是第幾工作日", "workday_no"),
            ("發包計價(未稅)", "contract_total"),
            ("發包以外計價", "other_total"),
            ("管理費計價(未稅)", "admin_total"),
            ("工程執行狀態", "execution_status"),
        ]
        for i, (text, key) in enumerate(labels):
            row = 0 if i < 5 else 1
            col = i if i < 5 else i - 5
            ttk.Label(self.summary, text=text + "：").grid(row=row, column=col*2, sticky="e", padx=(5, 2), pady=3)
            v = tk.StringVar()
            self.summary_vars[key] = v
            ttk.Label(self.summary, textvariable=v, foreground=OSX_ACCENT, font=("Microsoft JhengHei UI", 11, "bold")).grid(
                row=row, column=col*2+1, sticky="w", padx=(0, 12), pady=3
            )

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)

        self.tab_basic = ttk.Frame(self.nb, padding=8)
        self.tab_day_tables = ttk.Frame(self.nb, padding=8)
        self.tab_calendar = ttk.Frame(self.nb, padding=8)
        self.tab_payment_contract = ttk.Frame(self.nb, padding=8)
        self.tab_payment_other = ttk.Frame(self.nb, padding=8)
        self.tab_payment_admin = ttk.Frame(self.nb, padding=8)
        self.tab_progress_estimate = ttk.Frame(self.nb, padding=8)
        self.tab_execution = ttk.Frame(self.nb, padding=8)
        self.tab_milestone = ttk.Frame(self.nb, padding=8)
        self.tab_budget_data = ttk.Frame(self.nb, padding=8)
        self.tab_change_data = ttk.Frame(self.nb, padding=8)

        self.tab_day_tables.grid_rowconfigure(0, weight=1)
        self.tab_day_tables.grid_columnconfigure(0, weight=1)
        self.tab_day_tables.grid_rowconfigure(1, weight=0)
        day_tables_row = ttk.Frame(self.tab_day_tables)
        day_tables_row.grid(row=0, column=0, sticky="nsew")
        day_tables_row.grid_rowconfigure(0, weight=1)
        for i in range(4):
            day_tables_row.grid_columnconfigure(i, weight=1, uniform="day_tables")

        self.tab_holiday = ttk.LabelFrame(day_tables_row, text="假期表", padding=8, style="DayTable.TLabelframe")
        self.tab_workday = ttk.LabelFrame(day_tables_row, text="補班日表", padding=8, style="DayTable.TLabelframe")
        self.tab_railway = ttk.LabelFrame(day_tables_row, text="鐵路疏運表", padding=8, style="DayTable.TLabelframe")
        self.tab_weather = ttk.LabelFrame(day_tables_row, text="晴雨表", padding=8, style="DayTable.TLabelframe")
        self.tab_holiday.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self.tab_workday.grid(row=0, column=1, sticky="nsew", padx=4)
        self.tab_railway.grid(row=0, column=2, sticky="nsew", padx=4)
        self.tab_weather.grid(row=0, column=3, sticky="nsew", padx=(4, 0))
        day_lock = ttk.Frame(self.tab_day_tables)
        day_lock.grid(row=1, column=0, sticky="e", pady=(6, 0))
        ttk.Checkbutton(
            day_lock,
            text="資料編輯鎖定解除（勾選才可編輯）",
            variable=self.data_edit_enabled_var,
            command=self.apply_edit_lock_state
        ).pack(side="right")

        self.nb.add(self.tab_basic, text="工程基本資料")
        self.nb.add(self.tab_day_tables, text="假期/晴雨/疏運表")
        self.nb.add(self.tab_calendar, text="施工日曆")
        self.nb.add(self.tab_progress_estimate, text="工程進度估算")
        self.nb.add(self.tab_payment_contract, text="發包工程費計價")
        self.nb.add(self.tab_payment_other, text="發包以外計價")
        self.nb.add(self.tab_payment_admin, text="管理費計價")
        self.nb.add(self.tab_execution, text="會議記錄表")
        self.nb.add(self.tab_milestone, text="工程大事記")
        self.nb.add(self.tab_budget_data, text="預算資料")
        self.nb.add(self.tab_change_data, text="變更資料")

        self.build_basic_tab()
        self.build_holiday_tab()
        self.build_weather_tab()
        self.build_railway_tab()
        self.build_calendar_tab()
        self.build_progress_estimate_tab()
        self.build_payment_tabs()
        self.build_execution_tab()
        self.build_milestone_tab()
        self.build_budget_data_tab()
        self.build_change_data_tab()
        self.assign_tree_edit_guards()

    def toggle_summary_area(self):
        if self.summary_visible:
            self.summary.pack_forget()
            self.summary_visible = False
            self.summary_toggle_btn.configure(text="▼")
        else:
            self.summary.pack(fill="x", padx=8, pady=(0, 8), before=self.nb)
            self.summary_visible = True
            self.summary_toggle_btn.configure(text="▲")

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

    def push_undo_snapshot(self, state):
        """保留最多 10 次可回復動作。"""
        if not state:
            return
        if self.undo_history and self.undo_history[-1] == state:
            return
        self.undo_history.append(state)
        if len(self.undo_history) > 10:
            self.undo_history = self.undo_history[-10:]

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
                    self.push_undo_snapshot(self.last_state)
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
        if hasattr(self, "warranty_note_text"):
            state["basic"]["warranty_note"] = self.warranty_note_text.get("1.0", "end-1c")
        for name in [
            "bid_tree", "holiday_tree", "workday_tree", "weather_tree", "railway_tree",
            "payment_contract_tree", "payment_other_tree", "payment_admin_tree", "progress_estimate_tree",
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
            if hasattr(self, "warranty_note_text"):
                self.warranty_note_text.delete("1.0", "end")
                self.warranty_note_text.insert("1.0", state.get("basic", {}).get("warranty_note", ""))
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
        if not self.undo_history:
            self.status_var.set("目前沒有可回復的動作")
            return
        snapshot = self.undo_history.pop()
        self.restore_state(snapshot)
        self.save_current()
        remain = len(self.undo_history)
        self.status_var.set(f"已回復上一個動作並自動儲存，尚可回復 {remain} 次")

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

    def money_decimal(self, value):
        text = str(value or "").replace(",", "").replace("元", "").replace("%", "").strip()
        if not text:
            return Decimal("0")
        try:
            return Decimal(text)
        except (InvalidOperation, ValueError):
            try:
                return Decimal(str(float(text)))
            except (ValueError, InvalidOperation):
                return Decimal("0")

    def format_money_value(self, value):
        raw = str(value or "").strip().replace("元", "").replace(",", "")
        if not raw:
            return ""
        amount = self.money_decimal(value).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if amount == amount.to_integral_value():
            return f"{int(amount):,}元"
        return f"{amount:,.2f}元"

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
        ttk.Button(box, text="列印設定/產生列印檔", command=self.open_print_dialog).pack(fill="x", pady=4)
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
        if hasattr(self, "basic_widgets"):
            self.basic_widgets[key] = ent
        if key in MONEY_FIELDS:
            ent.bind("<FocusOut>", lambda e, k=key: self.format_money_field(k))
            ent.bind("<Return>", lambda e, k=key: self.format_money_field(k))
        return ent

    def readonly_entry(self, parent, row, col, label, key, width=14):
        ent = self.entry(parent, row, col, label, key, width)
        ent.configure(state="readonly")
        self.readonly_basic_keys.add(key)
        return ent

    def section_title(self, parent, text, row, columnspan=8):
        ttk.Label(parent, text=text, anchor="center", font=("Microsoft JhengHei UI", 12, "bold")).grid(
            row=row, column=0, columnspan=columnspan, sticky="ew", padx=6, pady=(10, 6)
        )

    def multiline_entry(self, parent, row, col, label, key, height=3):
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="ne", padx=3, pady=2)
        txt = tk.Text(parent, height=height, width=40, wrap="word", font=("Microsoft JhengHei UI", 10))
        txt.configure(borderwidth=1, relief="solid", highlightthickness=1, highlightbackground=OSX_BORDER, background=OSX_ENTRY_BG, foreground=OSX_TEXT, insertbackground=OSX_TEXT)
        txt.grid(row=row, column=col+1, columnspan=7, sticky="ew", padx=3, pady=2)
        txt.bind("<KeyRelease>", lambda e: self.mark_dirty())
        txt.bind("<FocusOut>", lambda e: self.mark_dirty())
        self.edit_widgets.append(txt)
        setattr(self, key + "_text", txt)
        return txt

    def money_section_title(self, parent, text, row, columnspan=6):
        bg = parent.cget("bg") if hasattr(parent, "cget") else None
        tk.Label(parent, text=text, anchor="center", bg=bg, font=("Microsoft JhengHei UI", 11, "bold")).grid(
            row=row, column=0, columnspan=columnspan, sticky="ew", padx=4, pady=(8, 4)
        )

    def money_entry(self, parent, row, col, label, key, width=12, readonly=False, suffix="", copy_from=None):
        bg = parent.cget("bg") if hasattr(parent, "cget") else None
        tk.Label(parent, text=label, anchor="e", bg=bg).grid(row=row, column=col, sticky="e", padx=3, pady=2)
        var = tk.StringVar()
        var.trace_add("write", self.mark_dirty)

        holder = tk.Frame(parent, bg=bg)
        holder.grid(row=row, column=col + 1, sticky="ew", padx=3, pady=2)
        holder.grid_columnconfigure(0, weight=1)

        # 第一分頁下半部：可輸入的金額欄位只保留黑色底線，不再使用四邊黑框。
        if readonly:
            ent = ttk.Entry(holder, textvariable=var, width=width)
            ent.grid(row=0, column=0, sticky="ew")
        else:
            underline_holder = tk.Frame(holder, bg=bg)
            underline_holder.grid(row=0, column=0, sticky="ew")
            underline_holder.grid_columnconfigure(0, weight=1)
            ent = ttk.Entry(underline_holder, textvariable=var, width=width)
            ent.grid(row=0, column=0, sticky="ew")
            bottom_line = tk.Frame(underline_holder, bg=OSX_INPUT_BORDER, height=2)
            bottom_line.grid(row=1, column=0, sticky="ew", pady=(1, 0))
            bottom_line.grid_propagate(False)

        if suffix:
            tk.Label(holder, text=suffix, anchor="w", bg=bg).grid(row=0, column=1, sticky="w", padx=(4, 0))
        if copy_from:
            btn = ttk.Button(holder, text="複製", width=5, command=lambda s=copy_from, t=key: self.copy_budget_value(s, t))
            btn.grid(row=0, column=2, sticky="w", padx=(4, 0))
            self.edit_widgets.append(btn)
        if readonly:
            ent.configure(state="readonly")
        self.edit_widgets.append(ent)
        self.basic_vars[key] = var
        if hasattr(self, "basic_widgets"):
            self.basic_widgets[key] = ent
        if readonly:
            self.readonly_basic_keys.add(key)
        if key in MONEY_FIELDS:
            ent.bind("<FocusOut>", lambda e, k=key: self.format_money_field(k))
            ent.bind("<Return>", lambda e, k=key: self.format_money_field(k))
        return ent

    def copy_budget_value(self, source_key, target_key):
        if source_key in self.basic_vars and target_key in self.basic_vars:
            self.basic_vars[target_key].set(self.basic_vars[source_key].get())
            self.format_money_field(target_key)
            self.recalculate()

    def award_money_entry(self, parent, row, col, label, key, source_key=None, width=12, readonly=False, suffix=""):
        self.money_entry(parent, row, col, label, key, width, readonly, suffix=suffix, copy_from=source_key)

    def build_money_sections(self, parent):
        wrap = ttk.Frame(parent)
        wrap.grid(row=19, column=0, columnspan=8, sticky="ew", padx=3, pady=(10, 4))
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_columnconfigure(1, weight=1)

        budget = tk.LabelFrame(wrap, text="預算金額", bg=OSX_BUDGET_BG, fg=OSX_TEXT, bd=1, relief="solid", padx=8, pady=8, font=("Microsoft JhengHei UI", 11, "bold"))
        award = tk.LabelFrame(wrap, text="契約金額", bg=OSX_AWARD_BG, fg=OSX_TEXT, bd=1, relief="solid", padx=8, pady=8, font=("Microsoft JhengHei UI", 11, "bold"))
        budget.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        award.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        for frame in (budget, award):
            for idx in range(6):
                # 一行三欄：每欄由「標籤 + 輸入框」組成。
                frame.grid_columnconfigure(idx, weight=1 if idx % 2 else 0)

        # 總工程預算、發包工程費、包工費的金額欄位加寬，
        # 可完整輸入/顯示十億元等級金額，例如 1,000,000,000.00元。
        wide_money_width = 18

        self.money_section_title(budget, "總工程預算", 0, columnspan=6)
        self.money_entry(budget, 1, 0, "總預算", "budget_total_amount", width=wide_money_width, readonly=True, suffix="(含稅)")
        self.money_entry(budget, 1, 2, "未完工程", "budget_unfinished_amount", width=wide_money_width, readonly=True)
        self.money_entry(budget, 1, 4, "進項稅額", "budget_input_tax", width=wide_money_width, readonly=True)
        self.money_section_title(budget, "發包工程費", 2, columnspan=6)
        self.money_entry(budget, 3, 0, "預算總計", "budget_contract_total", width=wide_money_width, readonly=True, suffix="(含稅)")
        self.money_entry(budget, 3, 2, "預算金額", "budget_contract_amount", width=wide_money_width)
        self.money_entry(budget, 3, 4, "稅金", "budget_contract_tax", width=wide_money_width, readonly=True)
        self.money_section_title(budget, "包工費", 4, columnspan=6)
        self.money_entry(budget, 5, 0, "預算", "budget_labor", width=wide_money_width)
        self.money_section_title(budget, "發包以外", 6, columnspan=6)
        self.money_entry(budget, 7, 0, "工程管理費", "budget_mgmt_fee")
        self.money_entry(budget, 7, 2, "自辦工費", "budget_self_labor")
        self.money_entry(budget, 7, 4, "自購材料費", "budget_self_material")
        self.money_entry(budget, 8, 0, "路備材料費", "budget_spare_material")
        self.money_entry(budget, 8, 2, "路購材料費", "budget_railway_material")
        self.money_entry(budget, 8, 4, "監理費", "budget_supervision_fee")
        self.money_entry(budget, 9, 0, "運雜費", "budget_freight")
        self.money_entry(budget, 9, 2, "其他", "budget_other")
        self.money_entry(budget, 9, 4, "空汙費", "budget_air_pollution_fee")

        self.money_section_title(award, "總工程費用", 0, columnspan=6)
        self.money_entry(award, 1, 0, "總預算", "award_total_amount", width=wide_money_width, readonly=True, suffix="(含稅)")
        self.money_entry(award, 1, 2, "未完工程", "award_unfinished_amount", width=wide_money_width, readonly=True)
        self.money_entry(award, 1, 4, "進項稅額", "award_input_tax", width=wide_money_width, readonly=True)
        self.money_section_title(award, "發包工程費", 2, columnspan=6)
        self.money_entry(award, 3, 0, "契約金額總計", "award_contract_total", width=wide_money_width, readonly=True, suffix="(含稅)")
        self.money_entry(award, 3, 2, "發包契約金額", "award_contract_amount", width=wide_money_width)
        self.money_entry(award, 3, 4, "營業稅", "award_contract_tax", width=wide_money_width, readonly=True)
        self.money_entry(award, 4, 0, "底價", "award_base_price", width=wide_money_width)
        self.money_entry(award, 4, 2, "決標/預算=", "award_contract_budget_ratio", readonly=True)
        self.money_entry(award, 4, 4, "決標/底價=", "award_contract_base_ratio", readonly=True)
        self.money_entry(award, 5, 0, "底價/預算=", "award_base_budget_ratio", readonly=True)
        self.money_section_title(award, "包工費", 6, columnspan=6)
        self.money_entry(award, 7, 0, "發包", "award_labor", width=wide_money_width)
        self.money_section_title(award, "發包以外", 8, columnspan=6)
        self.award_money_entry(award, 9, 0, "工程管理費", "award_mgmt_fee", "budget_mgmt_fee")
        self.award_money_entry(award, 9, 2, "自辦工費", "award_self_labor", "budget_self_labor")
        self.award_money_entry(award, 9, 4, "自購材料費", "award_self_material", "budget_self_material")
        self.award_money_entry(award, 10, 0, "路備材料費", "award_spare_material", "budget_spare_material")
        self.award_money_entry(award, 10, 2, "路購材料費", "award_railway_material", "budget_railway_material")
        self.award_money_entry(award, 10, 4, "監理費", "award_supervision_fee", "budget_supervision_fee")
        self.award_money_entry(award, 11, 0, "運雜費", "award_freight", "budget_freight")
        self.award_money_entry(award, 11, 2, "其他", "award_other", "budget_other")
        self.award_money_entry(award, 11, 4, "空汙費", "award_air_pollution_fee", "budget_air_pollution_fee")

        self.build_change_award_money_section(wrap, wide_money_width)

    def build_change_award_money_section(self, wrap, wide_money_width):
        change_box = tk.LabelFrame(wrap, text="變更後契約金額", bg=OSX_PANEL_BG, fg=OSX_TEXT, bd=1, relief="solid", padx=8, pady=8, font=("Microsoft JhengHei UI", 11, "bold"))
        change_box.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        for idx in range(6):
            change_box.grid_columnconfigure(idx, weight=1 if idx % 2 else 0)
        self.change_award_vars = {}
        self.change_award_widgets = {}
        self.change_award_select_var = tk.StringVar()
        header = tk.Frame(change_box, bg=OSX_PANEL_BG)
        header.grid(row=0, column=0, columnspan=6, sticky="ew", pady=(0, 4))
        tk.Label(header, text="第幾次變更：", bg=OSX_PANEL_BG, fg=OSX_TEXT).pack(side="left")
        self.change_award_combo = ttk.Combobox(header, textvariable=self.change_award_select_var, state="readonly", width=12)
        self.change_award_combo.pack(side="left", padx=(4, 8))
        self.change_award_combo.bind("<<ComboboxSelected>>", lambda e: self.load_selected_change_award_fields())
        ttk.Button(header, text="儲存此變更金額", command=self.save_current_change_award_fields).pack(side="left", padx=(4, 0))
        self.change_award_section_title(change_box, "總工程費用", 1)
        self.change_award_entry(change_box, 2, 0, "總預算", "change_award_total_amount", wide_money_width, True, "(含稅)")
        self.change_award_entry(change_box, 2, 2, "未完工程", "change_award_unfinished_amount", wide_money_width, True)
        self.change_award_entry(change_box, 2, 4, "進項稅額", "change_award_input_tax", wide_money_width, True)
        self.change_award_section_title(change_box, "發包工程費", 3)
        self.change_award_entry(change_box, 4, 0, "契約金額總計", "change_award_contract_total", wide_money_width, True, "(含稅)")
        self.change_award_entry(change_box, 4, 2, "發包契約金額", "change_award_contract_amount", wide_money_width)
        self.change_award_entry(change_box, 4, 4, "營業稅", "change_award_contract_tax", wide_money_width, True)
        self.change_award_entry(change_box, 5, 0, "底價", "change_award_base_price", wide_money_width)
        self.change_award_entry(change_box, 5, 2, "決標/預算=", "change_award_contract_budget_ratio", 12, True)
        self.change_award_entry(change_box, 5, 4, "決標/底價=", "change_award_contract_base_ratio", 12, True)
        self.change_award_entry(change_box, 6, 0, "底價/預算=", "change_award_base_budget_ratio", 12, True)
        self.change_award_section_title(change_box, "包工費", 7)
        self.change_award_entry(change_box, 8, 0, "發包", "change_award_labor", wide_money_width)
        self.change_award_section_title(change_box, "發包以外", 9)
        self.change_award_entry(change_box, 10, 0, "工程管理費", "change_award_mgmt_fee")
        self.change_award_entry(change_box, 10, 2, "自辦工費", "change_award_self_labor")
        self.change_award_entry(change_box, 10, 4, "自購材料費", "change_award_self_material")
        self.change_award_entry(change_box, 11, 0, "路備材料費", "change_award_spare_material")
        self.change_award_entry(change_box, 11, 2, "路購材料費", "change_award_railway_material")
        self.change_award_entry(change_box, 11, 4, "監理費", "change_award_supervision_fee")
        self.change_award_entry(change_box, 12, 0, "運雜費", "change_award_freight")
        self.change_award_entry(change_box, 12, 2, "其他", "change_award_other")
        self.change_award_entry(change_box, 12, 4, "空汙費", "change_award_air_pollution_fee")

    def change_award_section_title(self, parent, text, row):
        bg = parent.cget("bg")
        tk.Label(parent, text=text, anchor="center", bg=bg, fg=OSX_TEXT, font=("Microsoft JhengHei UI", 11, "bold")).grid(row=row, column=0, columnspan=6, sticky="ew", padx=4, pady=(8, 4))

    def change_award_entry(self, parent, row, col, label, key, width=12, readonly=False, suffix=""):
        bg = parent.cget("bg")
        tk.Label(parent, text=label, anchor="e", bg=bg, fg=OSX_TEXT).grid(row=row, column=col, sticky="e", padx=3, pady=2)
        var = tk.StringVar()
        var.trace_add("write", self.on_change_award_field_changed)
        holder = tk.Frame(parent, bg=bg)
        holder.grid(row=row, column=col + 1, sticky="ew", padx=3, pady=2)
        holder.grid_columnconfigure(0, weight=1)
        ent = ttk.Entry(holder, textvariable=var, width=width)
        ent.grid(row=0, column=0, sticky="ew")
        if suffix:
            tk.Label(holder, text=suffix, anchor="w", bg=bg, fg=OSX_TEXT).grid(row=0, column=1, sticky="w", padx=(4, 0))
        if readonly:
            ent.configure(state="readonly")
        elif key in CHANGE_AWARD_MONEY_FIELDS:
            ent.bind("<FocusOut>", lambda e, k=key: self.format_change_award_money(k))
            ent.bind("<Return>", lambda e, k=key: self.format_change_award_money(k))
        self.change_award_vars[key] = var
        self.change_award_widgets[key] = ent
        self.edit_widgets.append(ent)
        return ent

    def performance_deposit_entry(self, parent, row, col):
        """履約保證金欄位：可由比例自動計算，也可勾選後手動輸入。"""
        ttk.Label(parent, text="履約保證金").grid(row=row, column=col, sticky="e", padx=3, pady=2)
        holder = ttk.Frame(parent)
        holder.grid(row=row, column=col + 1, sticky="ew", padx=3, pady=2)
        holder.grid_columnconfigure(0, weight=1)

        var = tk.StringVar()
        var.trace_add("write", self.mark_dirty)
        ent = ttk.Entry(holder, textvariable=var, width=12)
        ent.grid(row=0, column=0, sticky="ew")

        manual_var = tk.StringVar(value="0")
        manual_var.trace_add("write", self.on_deposit_performance_manual_changed)
        chk = ttk.Checkbutton(holder, text="手動修改", variable=manual_var, onvalue="1", offvalue="0")
        chk.grid(row=0, column=1, sticky="w", padx=(6, 0))

        self.edit_widgets.append(ent)
        self.edit_widgets.append(chk)
        self.basic_vars["deposit_performance"] = var
        self.basic_vars["deposit_performance_manual"] = manual_var
        if hasattr(self, "basic_widgets"):
            self.basic_widgets["deposit_performance"] = ent
            self.basic_widgets["deposit_performance_manual"] = chk
        ent.bind("<FocusOut>", lambda e: self.format_money_field("deposit_performance"))
        ent.bind("<Return>", lambda e: self.format_money_field("deposit_performance"))
        self.deposit_performance_entry_widget = ent
        self.deposit_performance_manual_check = chk
        self.update_deposit_performance_manual_state()
        return ent

    def on_deposit_performance_manual_changed(self, *_):
        if self.loading or self.restoring:
            self.update_deposit_performance_manual_state()
            return
        self.update_deposit_performance_manual_state()
        self.mark_dirty()

    def update_deposit_performance_manual_state(self):
        ent = getattr(self, "deposit_performance_entry_widget", None)
        if not ent:
            return
        manual = self.basic_vars.get("deposit_performance_manual", tk.StringVar(value="0")).get() == "1"
        unlocked = self.can_edit()
        try:
            if not unlocked:
                ent.configure(state="disabled")
            elif manual:
                ent.configure(state="normal")
            else:
                ent.configure(state="readonly")
        except tk.TclError:
            pass
        chk = getattr(self, "deposit_performance_manual_check", None)
        if chk:
            try:
                chk.configure(state="normal" if unlocked else "disabled")
            except tk.TclError:
                pass

    def build_basic_tab(self):
        self.basic_vars = {}
        self.basic_widgets = {}
        canvas = tk.Canvas(self.tab_basic, highlightthickness=0, background=OSX_WINDOW_BG)
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
        # 第一分頁輸入區自動貼合目前視窗寬度，避免每行資料被擠到非視區外。
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=max(1, e.width)))
        def on_mousewheel(event):
            delta = -1 if event.delta > 0 else 1
            canvas.yview_scroll(delta * 3, "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        form = ttk.Frame(content, padding=8)
        form.pack(fill="x", expand=True)
        for idx in range(8):
            form.grid_columnconfigure(idx, weight=1 if idx % 2 else 0, minsize=0)
        self.section_title(form, "工程基本資料", 0)

        ttk.Checkbutton(
            form,
            text="資料編輯鎖定解除（勾選才可編輯）",
            variable=self.data_edit_enabled_var,
            command=self.apply_edit_lock_state
        ).grid(row=1, column=0, columnspan=8, sticky="w", padx=3, pady=2)

        name_entry = self.entry(form, 2, 0, "工程名稱", "name", 48)
        name_entry.grid_configure(columnspan=7)
        self.entry(form, 3, 0, "工程執行號", "exec_no", 14)
        self.entry(form, 3, 2, "動支請示單號", "budget_no", 14)
        self.entry(form, 3, 4, "採購契約號碼", "purchase_contract_no", 14)
        self.readonly_entry(form, 4, 0, "決標日期", "award_date", 14)
        self.entry(form, 4, 2, "簽約日期", "contract_date", 14, date_picker=True)
        self.entry(form, 4, 4, "預訂開工日", "planned_start", 14, date_picker=True)
        self.entry(form, 4, 6, "實際開工日", "actual_start", 14, date_picker=True)
        self.entry(form, 5, 0, "契約工期", "contract_days", 10)
        ttk.Label(form, text="工期類型").grid(row=5, column=2, sticky="e", padx=3, pady=2)
        self.day_type_var = tk.StringVar(value="工作日")
        self.day_type_var.trace_add("write", self.mark_dirty)
        self.day_type = ttk.Combobox(form, textvariable=self.day_type_var, state="readonly", values=["工作日", "日曆天"], width=10)
        self.day_type.grid(row=5, column=3, sticky="ew", padx=3, pady=2)
        self.edit_widgets.append(self.day_type)

        self.entry(form, 6, 0, "預訂竣工日（例假表）", "planned_finish_holiday", 14, date_picker=True)
        self.entry(form, 6, 2, "預訂竣工日（疏運表）", "planned_finish_transport", 14, date_picker=True)
        self.entry(form, 6, 4, "實際竣工日", "actual_finish", 14, date_picker=True)

        self.entry(form, 7, 0, "預定初驗日", "planned_precheck_date", 14, date_picker=True)
        self.entry(form, 7, 2, "實際初驗日", "actual_precheck_date", 14, date_picker=True)
        self.entry(form, 7, 4, "預定驗收日", "planned_acceptance_date", 14, date_picker=True)
        self.entry(form, 7, 6, "實際驗收日", "actual_acceptance_date", 14, date_picker=True)

        self.entry(form, 8, 0, "決算日", "settlement_date", 14, date_picker=True)
        self.entry(form, 8, 2, "保固年限", "warranty_years", 8)
        self.readonly_entry(form, 8, 4, "保固結束日", "warranty_end_date", 14)
        ttk.Label(form, text="保固備註").grid(row=8, column=6, sticky="ne", padx=3, pady=2)
        warranty_note = tk.Text(form, height=4, width=14, wrap="word", font=("Microsoft JhengHei UI", 10))
        warranty_note.configure(borderwidth=1, relief="solid", highlightthickness=1, highlightbackground=OSX_BORDER, background=OSX_ENTRY_BG, foreground=OSX_TEXT, insertbackground=OSX_TEXT)
        warranty_note.grid(row=8, column=7, sticky="ew", padx=3, pady=2)
        warranty_note.bind("<KeyRelease>", lambda e: self.mark_dirty())
        warranty_note.bind("<FocusOut>", lambda e: self.mark_dirty())
        self.edit_widgets.append(warranty_note)
        self.warranty_note_text = warranty_note

        self.entry(form, 9, 0, "承攬商", "contractor", 14)
        self.entry(form, 9, 2, "公司地址", "company_address", 14)
        self.entry(form, 9, 4, "負責人", "responsible_person", 14)
        self.entry(form, 9, 6, "聯絡人", "contact_person", 14)
        self.entry(form, 10, 0, "電話", "phone", 14)
        self.entry(form, 10, 2, "傳真電話", "fax", 14)
        self.entry(form, 10, 4, "統一編號", "tax_id", 14)

        ttk.Label(form, text="工程執行狀態").grid(row=11, column=0, sticky="e", padx=3, pady=2)
        self.execution_status_var = tk.StringVar(value="規劃中")
        self.execution_status_var.trace_add("write", self.mark_dirty)
        self.execution_status_combo = ttk.Combobox(
            form,
            textvariable=self.execution_status_var,
            values=["規劃中", "招標中", "決標完成", "開工中", "施工中", "停工中", "復工中", "竣工中", "已竣工", "驗收中", "驗收完成", "結案"],
            width=14
        )
        self.execution_status_combo.grid(row=11, column=1, sticky="ew", padx=3, pady=2)
        self.edit_widgets.append(self.execution_status_combo)

        self.multiline_entry(form, 12, 0, "工程說明", "project_description", height=3)

        self.section_title(form, "保證金", 13)
        self.entry(form, 14, 0, "差額保證金", "deposit_difference", 12)
        self.performance_deposit_entry(form, 14, 2)
        self.entry(form, 14, 4, "履保金比例", "performance_bond_rate", 8)
        self.readonly_entry(form, 14, 6, "保證金總額", "deposit_total", 12)
        self.readonly_entry(form, 15, 0, "保固保證金", "warranty_deposit", 12)
        self.entry(form, 15, 2, "保固金比例", "warranty_rate", 8)
        self.entry(form, 15, 4, "履保金型式", "performance_bond_type", 12)
        self.entry(form, 15, 6, "保固金型式", "warranty_bond_type", 12)

        self.build_money_sections(form)

        # 欄位寬度已在 form 建立時設定為自適應。

        bid_box = ttk.LabelFrame(content, text="招標情形", padding=8)
        bid_box.pack(fill="x", expand=False, pady=8)
        self.bid_tree = EditableTree(
            bid_box,
            ["awarded", "round_no", "online_date", "open_date", "award_date"],
            ["決標", "次數", "公告上網日", "開標日", "決標日"],
            [50, 30, 120, 120, 120],
            self.mark_dirty,
            add_command=self.open_bid_calendar_dialog
        )
        self.bid_tree.tree.configure(height=6)
        self.bid_tree.pack(fill="x", expand=False)

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
        date_type_var = tk.StringVar(value="公告上網日")

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
        ttk.Label(win, textvariable=selected_day_var, foreground=OSX_ACCENT).pack(anchor="w", padx=12)

        form = ttk.Frame(win, padding=10)
        form.pack(fill="x")
        ttk.Label(form, text="第幾次").grid(row=0, column=0, sticky="e", padx=(0, 6), pady=4)
        ttk.Entry(form, textvariable=round_var, width=10).grid(row=0, column=1, sticky="w", pady=4)
        ttk.Label(form, text="輸入狀況").grid(row=1, column=0, sticky="e", padx=(0, 6), pady=4)
        ttk.Combobox(form, textvariable=date_type_var, values=["公告上網日", "開標日", "決標日"], state="readonly", width=16).grid(row=1, column=1, sticky="w", pady=4)

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
            values = ["", round_text, "", "", ""]
            if date_type_var.get() == "公告上網日":
                values[2] = day_text
            elif date_type_var.get() == "開標日":
                values[3] = day_text
            else:
                values[0] = "V"
                values[4] = day_text
            self.bid_tree.add_row_after_selection(values)
            self.update_bid_award_state()
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
        toolbar = ttk.Frame(self.tab_holiday)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Label(toolbar, text="假期表：按「新增一列」用日曆新增。").pack(side="left")
        ttk.Button(toolbar, text="複製前一年度假期", command=self.copy_previous_year_holidays).pack(side="right", padx=4)
        ttk.Button(toolbar, text="確認假期", command=self.confirm_holidays).pack(side="right", padx=4)

        self.holiday_tree = EditableTree(
            self.tab_holiday,
            ["exclude", "day", "name"],
            ["排除", "日期", "假日名稱"],
            [36, 138, 135],
            self.mark_dirty,
            add_command=self.open_holiday_calendar_dialog
        )
        self.holiday_tree.pack(fill="both", expand=True, pady=6)

        workday_toolbar = ttk.Frame(self.tab_workday)
        workday_toolbar.pack(fill="x", pady=(0, 6))
        ttk.Label(workday_toolbar, text="補班日表：按「新增一列」用日曆新增。").pack(side="left")
        self.workday_tree = EditableTree(
            self.tab_workday,
            ["exclude", "day", "name"],
            ["排除", "日期", "補班名稱"],
            [36, 138, 135],
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
        ttk.Label(win, textvariable=selected_day_var, foreground=OSX_ACCENT).pack(anchor="w", padx=12)

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
        ttk.Label(win, textvariable=selected_day_var, foreground=OSX_ACCENT).pack(anchor="w", padx=12)
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
        ttk.Label(
            self.tab_weather,
            text="晴雨表：按「新增一列」會開啟年、月日曆表；下方可分別輸入上午、下午、天氣、場地與備註。"
        ).pack(anchor="w")
        self.weather_tree = EditableTree(
            self.tab_weather,
            ["day", "morning", "afternoon", "typhoon", "site", "note"],
            ["日期", "上午", "下午", "天氣", "場地", "備註"],
            [88, 56, 56, 56, 56, 72],
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
        ttk.Label(win, textvariable=selected_day_var, foreground=OSX_ACCENT).pack(anchor="w", padx=12)

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
        top = ttk.Frame(self.tab_railway)
        top.pack(fill="x")
        ttk.Label(top, text="鐵路疏運停工日期：可讀入第二分頁假期表；新增時會插在目前選取列下方。").pack(side="left")
        ttk.Button(top, text="讀入第二分頁假期表", command=self.import_holidays_to_railway).pack(side="right")
        self.railway_tree = EditableTree(
            self.tab_railway,
            ["exclude", "day", "note"],
            ["排除", "日期", "疏運名稱"],
            [36, 138, 135],
            self.on_railway_changed,
            add_command=self.open_railway_calendar_dialog
        )
        self.railway_tree.pack(fill="both", expand=True, pady=6)

    def on_railway_changed(self):
        self.mark_dirty()
        if not self.loading and not self.restoring:
            try:
                self.after_idle(self.render_calendar)
                self.status_var.set("疏運表已更新，施工日曆已同步重算")
            except tk.TclError:
                pass

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
        ttk.Label(win, textvariable=selected_day_var, foreground=OSX_ACCENT).pack(anchor="w", padx=12)

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
        self.cal_canvas = tk.Canvas(cal_area, background=OSX_PANEL_BG)
        self.cal_canvas.grid(row=0, column=0, sticky="nsew")
        cal_area.grid_rowconfigure(0, weight=1)
        cal_area.grid_columnconfigure(0, weight=1)
        self.cal_canvas.bind("<Configure>", lambda e: self.render_calendar())
        self.cal_canvas.bind("<MouseWheel>", lambda e: self.shift_month(-1 if e.delta > 0 else 1))

    def build_progress_estimate_tab(self):
        self.add_page_edit_toggle(self.tab_progress_estimate)
        ttk.Label(
            self.tab_progress_estimate,
            text="工程進度估算：月份請用 YYY/MM 格式；預估/實際施作金額會依第一分頁的發包契約金額與進度百分比自動計算。"
        ).pack(anchor="w", pady=(0, 6))
        self.progress_estimate_tree = EditableTree(
            self.tab_progress_estimate,
            PROGRESS_ESTIMATE_FIELDS,
            PROGRESS_ESTIMATE_HEADINGS,
            PROGRESS_ESTIMATE_WIDTHS,
            self.on_progress_estimate_changed,
            add_command=self.open_progress_estimate_add_dialog,
            edit_command=self.open_progress_estimate_edit_dialog,
            height=14,
        )
        self.configure_tree_alignment(
            self.progress_estimate_tree.tree,
            PROGRESS_ESTIMATE_FIELDS,
            PROGRESS_ESTIMATE_MONEY_FIELDS,
        )
        self.progress_estimate_tree.pack(fill="both", expand=True, pady=6)

    def default_roc_month_text(self):
        today = date.today()
        return f"{today.year - 1911:03d}/{today.month:02d}"

    def normalize_progress_month(self, text):
        text = str(text or "").strip().replace("-", "/")
        if not text:
            return ""
        m = re.match(r"^(\d{2,4})\s*/\s*(\d{1,2})$", text)
        if not m:
            return text
        year = int(m.group(1))
        month = max(1, min(12, int(m.group(2))))
        if year >= 1912:
            year -= 1911
        return f"{year:03d}/{month:02d}"

    def normalize_percent_value(self, value):
        raw = str(value or "").strip().replace("%", "").replace(",", "")
        if not raw:
            return ""
        try:
            dec = Decimal(raw).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        except (InvalidOperation, ValueError):
            return str(value or "").strip()
        if dec == dec.to_integral_value():
            return f"{int(dec)}%"
        return f"{dec:.2f}%"

    def percent_decimal_fraction(self, value):
        raw = str(value or "").strip().replace("%", "").replace(",", "")
        if not raw:
            return Decimal("0")
        try:
            return Decimal(raw) / Decimal("100")
        except (InvalidOperation, ValueError):
            return Decimal("0")

    def progress_contract_base_amount(self):
        if hasattr(self, "basic_vars") and "award_contract_amount" in self.basic_vars:
            return self.money_decimal(self.basic_vars["award_contract_amount"].get())
        return Decimal("0")

    def on_progress_estimate_changed(self):
        self.refresh_progress_estimate_amounts(mark_dirty=False)
        self.mark_dirty()

    def open_progress_estimate_add_dialog(self):
        self.open_progress_estimate_dialog(item=None)

    def open_progress_estimate_edit_dialog(self):
        if not hasattr(self, "progress_estimate_tree"):
            return
        item = self.progress_estimate_tree.tree.focus()
        if not item:
            return
        self.open_progress_estimate_dialog(item=item)

    def open_progress_estimate_dialog(self, item=None):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        tree = self.progress_estimate_tree
        is_edit = item is not None
        old_values = list(tree.tree.item(item, "values")) if is_edit else [""] * len(PROGRESS_ESTIMATE_FIELDS)
        old_values = (old_values + [""] * len(PROGRESS_ESTIMATE_FIELDS))[:len(PROGRESS_ESTIMATE_FIELDS)]
        idx_map = {field: i for i, field in enumerate(PROGRESS_ESTIMATE_FIELDS)}
        if not is_edit:
            if not old_values[idx_map["item_no"]]:
                old_values[idx_map["item_no"]] = str(len(tree.get_rows()) + 1)
            if not old_values[idx_map["month"]]:
                old_values[idx_map["month"]] = self.default_roc_month_text()

        win = tk.Toplevel(self)
        win.title("編輯工程進度估算" if is_edit else "新增工程進度估算")
        win.transient(self)
        win.grab_set()
        win.geometry("620x320")
        win.minsize(560, 280)
        form = ttk.Frame(win, padding=12)
        form.pack(fill="both", expand=True)
        vars_map = {}
        entries = {}
        for i, (field, heading) in enumerate(zip(PROGRESS_ESTIMATE_FIELDS, PROGRESS_ESTIMATE_HEADINGS)):
            row = i
            ttk.Label(form, text=heading + "：").grid(row=row, column=0, sticky="e", padx=6, pady=5)
            var = tk.StringVar(value=old_values[i] if i < len(old_values) else "")
            vars_map[field] = var
            ent = ttk.Entry(form, textvariable=var, width=28)
            ent.grid(row=row, column=1, sticky="ew", padx=6, pady=5)
            entries[field] = ent
            if field in PROGRESS_ESTIMATE_READONLY_FIELDS:
                ent.configure(state="readonly")
        form.grid_columnconfigure(1, weight=1)

        def update_amounts(*_):
            base = self.progress_contract_base_amount()
            for progress_field, amount_field in (("estimated_progress", "estimated_amount"), ("actual_progress", "actual_amount")):
                amount = base * self.percent_decimal_fraction(vars_map[progress_field].get())
                text = self.format_money_value(amount) if amount else ""
                vars_map[amount_field].set(text)

        for field in ("estimated_progress", "actual_progress"):
            entries[field].bind("<KeyRelease>", update_amounts)
            entries[field].bind("<FocusOut>", lambda e, f=field: (vars_map[f].set(self.normalize_percent_value(vars_map[f].get())), update_amounts()))
            entries[field].bind("<Return>", lambda e, f=field: (vars_map[f].set(self.normalize_percent_value(vars_map[f].get())), update_amounts()))
        entries["month"].bind("<FocusOut>", lambda e: vars_map["month"].set(self.normalize_progress_month(vars_map["month"].get())))
        entries["month"].bind("<Return>", lambda e: vars_map["month"].set(self.normalize_progress_month(vars_map["month"].get())))
        update_amounts()

        def ok():
            vars_map["month"].set(self.normalize_progress_month(vars_map["month"].get()))
            vars_map["estimated_progress"].set(self.normalize_percent_value(vars_map["estimated_progress"].get()))
            vars_map["actual_progress"].set(self.normalize_percent_value(vars_map["actual_progress"].get()))
            update_amounts()
            values = [vars_map[field].get().strip() for field in PROGRESS_ESTIMATE_FIELDS]
            win.destroy()
            if is_edit:
                tree.tree.item(item, values=values)
                tree.tree.selection_set(item)
                tree.tree.focus(item)
            else:
                tree.add_row(values)
            self.refresh_progress_estimate_amounts(mark_dirty=True)

        btns = ttk.Frame(win, padding=(12, 0, 12, 12))
        btns.pack(fill="x")
        ttk.Button(btns, text="取消", command=win.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="確定", command=ok).pack(side="right", padx=4)
        entries["item_no"].focus_set()

    def refresh_progress_estimate_amounts(self, mark_dirty=False):
        if not hasattr(self, "progress_estimate_tree"):
            return
        restore_focus = self._remember_and_restore_tree_focus(self.progress_estimate_tree, PROGRESS_ESTIMATE_FIELDS, "item_no")
        rows = self.progress_estimate_tree.get_rows()
        if not rows:
            return
        indexes = {field: idx for idx, field in enumerate(PROGRESS_ESTIMATE_FIELDS)}
        base = self.progress_contract_base_amount()
        calculated_rows = []
        for raw_row in rows:
            row = (list(raw_row) + [""] * len(PROGRESS_ESTIMATE_FIELDS))[:len(PROGRESS_ESTIMATE_FIELDS)]
            row[indexes["month"]] = self.normalize_progress_month(row[indexes["month"]])
            for progress_field, amount_field in (("estimated_progress", "estimated_amount"), ("actual_progress", "actual_amount")):
                p_text = self.normalize_percent_value(row[indexes[progress_field]])
                row[indexes[progress_field]] = p_text
                amount = base * self.percent_decimal_fraction(p_text)
                row[indexes[amount_field]] = self.format_money_value(amount) if amount else ""
            calculated_rows.append(row)
        self.progress_estimate_tree.set_rows(calculated_rows)
        self.configure_tree_alignment(self.progress_estimate_tree.tree, PROGRESS_ESTIMATE_FIELDS, PROGRESS_ESTIMATE_MONEY_FIELDS)
        restore_focus()
        if mark_dirty:
            self.mark_dirty()

    def progress_estimate_db_rows_to_tree(self, db_rows):
        rows = []
        for r in db_rows:
            rows.append([r[field] if field in r.keys() and r[field] is not None else "" for field in PROGRESS_ESTIMATE_FIELDS])
        return rows

    def progress_estimate_tree_rows_to_db(self):
        self.refresh_progress_estimate_amounts(mark_dirty=False)
        rows = []
        for r in self.progress_estimate_tree.get_rows():
            vals = (list(r) + [""] * len(PROGRESS_ESTIMATE_FIELDS))[:len(PROGRESS_ESTIMATE_FIELDS)]
            rows.append({field: vals[idx] for idx, field in enumerate(PROGRESS_ESTIMATE_FIELDS)})
        return rows

    def build_payment_tabs(self):
        self.add_page_edit_toggle(self.tab_payment_contract)
        self.payment_contract_tree = self.make_contract_payment_tree(self.tab_payment_contract)

        self.add_page_edit_toggle(self.tab_payment_other)
        self.payment_other_tree = self.make_payment_other_tree(self.tab_payment_other)

        self.add_page_edit_toggle(self.tab_payment_admin)
        self.build_payment_admin_tab(self.tab_payment_admin)

    def make_contract_payment_tree(self, parent):
        ttk.Label(
            parent,
            text="發包工程費計價：第一欄固定，第二欄以後可橫向捲動；累計欄位固定依期數由舊到新計算。"
        ).pack(anchor="w")
        tree = FixedFirstColumnTree(
            parent,
            PAYMENT_CONTRACT_FIELDS,
            PAYMENT_CONTRACT_HEADINGS,
            PAYMENT_CONTRACT_WIDTHS,
            self.on_payment_contract_changed,
            add_command=self.open_contract_payment_add_dialog,
            edit_command=self.open_contract_payment_edit_dialog,
            height=14,
            money_columns=PAYMENT_CONTRACT_MONEY_FIELDS,
            date_columns=PAYMENT_CONTRACT_DATE_FIELDS,
        )
        tree.pack(fill="both", expand=True, pady=6)
        return tree

    def configure_tree_alignment(self, tree_widget, fields, money_fields=None, left_fields=None):
        money_fields = set(money_fields or [])
        left_fields = set(left_fields or [])
        for field in fields:
            try:
                if field in money_fields:
                    tree_widget.column(field, anchor="e")
                elif field in left_fields:
                    tree_widget.column(field, anchor="w")
                else:
                    tree_widget.column(field, anchor="center")
            except tk.TclError:
                pass

    def make_payment_other_tree(self, parent):
        ttk.Label(parent, text="發包以外計價：累計欄位依期數由舊到新計算，不受目前瀏覽排序影響。滑鼠雙擊或按編輯可開啟日期/下拉輸入視窗。" ).pack(anchor="w")
        tree = EditableTree(
            parent,
            PAYMENT_OTHER_FIELDS,
            PAYMENT_OTHER_HEADINGS,
            PAYMENT_OTHER_WIDTHS,
            self.on_payment_other_changed,
            add_command=self.open_payment_other_add_dialog,
            edit_command=self.open_payment_other_edit_dialog,
            height=14,
        )
        self.configure_tree_alignment(tree.tree, PAYMENT_OTHER_FIELDS, PAYMENT_OTHER_MONEY_FIELDS, PAYMENT_OTHER_LEFT_FIELDS)
        tree.pack(fill="both", expand=True, pady=6)
        return tree

    def build_payment_admin_tab(self, parent):
        self.admin_top_frame = ttk.LabelFrame(parent, text="管理費提撥與可支用額度", padding=8)
        self.admin_top_frame.pack(fill="x", pady=(0, 8))
        self.admin_calc_vars = {}

        def add_admin_var(key, value=""):
            var = tk.StringVar(value=value)
            var.trace_add("write", self.on_admin_setup_changed)
            self.admin_calc_vars[key] = var
            return var

        ttk.Label(self.admin_top_frame, text="管理費：").grid(row=0, column=0, sticky="e", padx=4, pady=3)
        self.admin_calc_vars["mgmt_fee"] = tk.StringVar()
        mgmt_ent = ttk.Entry(self.admin_top_frame, textvariable=self.admin_calc_vars["mgmt_fee"], width=18, state="readonly")
        mgmt_ent.grid(row=0, column=1, sticky="ew", padx=4, pady=3)
        ttk.Label(self.admin_top_frame, text="變更後管理費：").grid(row=0, column=2, sticky="e", padx=4, pady=3)
        self.admin_calc_vars["change_mgmt_fee"] = add_admin_var("change_mgmt_fee")
        change_ent = ttk.Entry(self.admin_top_frame, textvariable=self.admin_calc_vars["change_mgmt_fee"], width=18)
        change_ent.grid(row=0, column=3, sticky="ew", padx=4, pady=3)
        ttk.Label(self.admin_top_frame, text="（暫不自動讀取變更資料）").grid(row=0, column=4, sticky="w", padx=4, pady=3)

        heads = ["提撥階段", "提撥金額", "0C12差費", "0C11加班費", "0C14其他"]
        for c, head in enumerate(heads):
            ttk.Label(self.admin_top_frame, text=head, font=("Microsoft JhengHei UI", 10, "bold")).grid(row=1, column=c, sticky="ew", padx=4, pady=3)
        labels = [
            "第一次提撥(開工_60%)", "第二次提撥(60%_30%)",
            "第三次提撥(80%_10%)", "第四次提撥(竣工_負計)",
        ]
        for idx, label in enumerate(labels, start=1):
            r = idx + 1
            ttk.Label(self.admin_top_frame, text=label + "：").grid(row=r, column=0, sticky="e", padx=4, pady=3)
            for c, suffix in enumerate(["amount", "0c12", "0c11", "0c14"], start=1):
                key = f"alloc{idx}_{suffix}"
                ent = ttk.Entry(self.admin_top_frame, textvariable=add_admin_var(key), width=16)
                ent.grid(row=r, column=c, sticky="ew", padx=4, pady=3)
                ent.bind("<FocusOut>", lambda e, k=key: self.format_admin_setup_money(k))
                ent.bind("<Return>", lambda e, k=key: self.format_admin_setup_money(k))

        total_row = 6
        ttk.Label(self.admin_top_frame, text="可支用合計：", font=("Microsoft JhengHei UI", 10, "bold")).grid(row=total_row, column=0, sticky="e", padx=4, pady=3)
        self.admin_calc_vars["available_0c12"] = tk.StringVar()
        self.admin_calc_vars["available_0c11"] = tk.StringVar()
        self.admin_calc_vars["available_0c14"] = tk.StringVar()
        ttk.Label(self.admin_top_frame, text="0C12差費合計").grid(row=total_row, column=1, sticky="e", padx=4, pady=3)
        self.admin_available_0c12_entry = ttk.Entry(self.admin_top_frame, textvariable=self.admin_calc_vars["available_0c12"], state="readonly", width=16, style="AdminAvailable.TEntry")
        self.admin_available_0c12_entry.grid(row=total_row, column=2, sticky="ew", padx=4, pady=3)
        ttk.Label(self.admin_top_frame, text="0C11加班費合計").grid(row=total_row, column=3, sticky="e", padx=4, pady=3)
        self.admin_available_0c11_entry = ttk.Entry(self.admin_top_frame, textvariable=self.admin_calc_vars["available_0c11"], state="readonly", width=16, style="AdminAvailable.TEntry")
        self.admin_available_0c11_entry.grid(row=total_row, column=4, sticky="ew", padx=4, pady=3)
        ttk.Label(self.admin_top_frame, text="0C14其他合計").grid(row=total_row, column=5, sticky="e", padx=4, pady=3)
        self.admin_available_0c14_entry = ttk.Entry(self.admin_top_frame, textvariable=self.admin_calc_vars["available_0c14"], state="readonly", width=16, style="AdminAvailable.TEntry")
        self.admin_available_0c14_entry.grid(row=total_row, column=6, sticky="ew", padx=4, pady=3)

        used_row = 7
        ttk.Label(self.admin_top_frame, text="已計價累計：", font=("Microsoft JhengHei UI", 10, "bold")).grid(row=used_row, column=0, sticky="e", padx=4, pady=3)
        self.admin_calc_vars["used_0c12"] = tk.StringVar()
        self.admin_calc_vars["used_0c11"] = tk.StringVar()
        self.admin_calc_vars["used_0c14"] = tk.StringVar()
        self.admin_used_0c12_entry = ttk.Entry(self.admin_top_frame, textvariable=self.admin_calc_vars["used_0c12"], state="readonly", width=16)
        self.admin_used_0c12_entry.grid(row=used_row, column=2, sticky="ew", padx=4, pady=3)
        self.admin_used_0c11_entry = ttk.Entry(self.admin_top_frame, textvariable=self.admin_calc_vars["used_0c11"], state="readonly", width=16)
        self.admin_used_0c11_entry.grid(row=used_row, column=4, sticky="ew", padx=4, pady=3)
        self.admin_used_0c14_entry = ttk.Entry(self.admin_top_frame, textvariable=self.admin_calc_vars["used_0c14"], state="readonly", width=16)
        self.admin_used_0c14_entry.grid(row=used_row, column=6, sticky="ew", padx=4, pady=3)
        for c in range(7):
            self.admin_top_frame.grid_columnconfigure(c, weight=1 if c else 0)

        lower = ttk.LabelFrame(parent, text="管理費計價明細", padding=8)
        lower.pack(fill="both", expand=True)
        self.payment_admin_tree = EditableTree(
            lower,
            PAYMENT_ADMIN_FIELDS,
            PAYMENT_ADMIN_HEADINGS,
            PAYMENT_ADMIN_WIDTHS,
            self.on_payment_admin_changed,
            add_command=self.open_payment_admin_add_dialog,
            edit_command=self.open_payment_admin_edit_dialog,
            height=12,
        )
        self.configure_tree_alignment(self.payment_admin_tree.tree, PAYMENT_ADMIN_FIELDS, PAYMENT_ADMIN_MONEY_FIELDS)
        self.payment_admin_tree.pack(fill="both", expand=True, pady=6)

    def on_payment_contract_changed(self):
        self.refresh_contract_payment_cumulatives(mark_dirty=False)
        self.mark_dirty()

    def open_contract_payment_add_dialog(self):
        self.open_contract_payment_dialog(item=None)

    def open_contract_payment_edit_dialog(self):
        if not hasattr(self, "payment_contract_tree"):
            return
        item = self.payment_contract_tree.focus()
        if not item:
            return
        self.open_contract_payment_dialog(item=item)

    def open_contract_payment_dialog(self, item=None):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        tree = self.payment_contract_tree
        is_edit = item is not None
        old_values = tree.get_item_values(item) if is_edit else [""] * len(PAYMENT_CONTRACT_FIELDS)
        old_values = (old_values + [""] * len(PAYMENT_CONTRACT_FIELDS))[:len(PAYMENT_CONTRACT_FIELDS)]
        if not is_edit and not old_values[0]:
            old_values[0] = str(len(tree.get_rows()) + 1)

        win = tk.Toplevel(self)
        win.title("編輯發包工程費計價資料" if is_edit else "新增發包工程費計價資料")
        win.transient(self)
        win.grab_set()
        win.geometry("780x640")
        win.minsize(640, 480)

        canvas = tk.Canvas(win, highlightthickness=0, background=OSX_WINDOW_BG)
        vs = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        form = ttk.Frame(canvas, padding=12)
        canvas_window = canvas.create_window((0, 0), window=form, anchor="nw")
        canvas.configure(yscrollcommand=vs.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        win.grid_rowconfigure(0, weight=1)
        win.grid_columnconfigure(0, weight=1)
        form.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=max(e.width, form.winfo_reqwidth())))

        vars_map = {}
        entries = {}
        for idx, (field, heading) in enumerate(zip(PAYMENT_CONTRACT_FIELDS, PAYMENT_CONTRACT_HEADINGS)):
            row = idx // 2
            col = (idx % 2) * 3
            ttk.Label(form, text=heading + "：").grid(row=row, column=col, sticky="e", padx=(4, 2), pady=4)
            var = tk.StringVar(value=old_values[idx])
            vars_map[field] = var
            holder = ttk.Frame(form)
            holder.grid(row=row, column=col + 1, sticky="ew", padx=(2, 8), pady=4)
            form.grid_columnconfigure(col + 1, weight=1)
            if field in PAYMENT_CONTRACT_DATE_FIELDS:
                ent = ttk.Entry(holder, textvariable=var, width=16)
                ent.pack(side="left", fill="x", expand=True)
                ttk.Button(holder, text="▼", width=3, command=lambda v=var, w=win, h=heading: self.open_payment_date_picker(v, w, h)).pack(side="left", padx=(3, 0))
            else:
                ent = ttk.Entry(holder, textvariable=var, width=22)
                ent.pack(fill="x", expand=True)
            entries[field] = ent
            if field in PAYMENT_CONTRACT_CUMULATIVE_FIELDS:
                ent.configure(state="readonly")
            if field in PAYMENT_CONTRACT_MONEY_FIELDS and field not in PAYMENT_CONTRACT_CUMULATIVE_FIELDS:
                ent.bind("<FocusOut>", lambda e, v=var: v.set(self.format_money_value(v.get()) if v.get().strip() else ""))
                ent.bind("<Return>", lambda e, v=var: v.set(self.format_money_value(v.get()) if v.get().strip() else ""))

        def ok():
            values = []
            for field in PAYMENT_CONTRACT_FIELDS:
                raw = vars_map[field].get().strip()
                if field in PAYMENT_CONTRACT_MONEY_FIELDS and raw:
                    raw = self.format_money_value(raw)
                values.append(raw)
            if is_edit:
                tree.set_item_values(item, values)
            else:
                tree.insert_row(values, select=True)
            win.destroy()
            self.refresh_contract_payment_cumulatives(mark_dirty=True)

        btns = ttk.Frame(win, padding=10)
        btns.grid(row=1, column=0, columnspan=2, sticky="ew")
        ttk.Button(btns, text="取消", command=win.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="確定", command=ok).pack(side="right", padx=4)
        entries.get("period_no", next(iter(entries.values()))).focus_set()

    def open_payment_date_picker(self, target_var, parent_win=None, title="選擇日期"):
        base_date = parse_date(target_var.get()) or date.today()
        win = tk.Toplevel(parent_win or self)
        win.title(title)
        win.transient(parent_win or self)
        win.grab_set()
        win.resizable(False, False)

        year_var = tk.IntVar(value=base_date.year)
        month_var = tk.IntVar(value=base_date.month)
        top = ttk.Frame(win, padding=10)
        top.pack(fill="x")
        cal_frame = ttk.Frame(win, padding=(10, 0, 10, 10))
        cal_frame.pack()

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

        def choose(day):
            target_var.set(day.strftime("%Y-%m-%d"))
            win.destroy()

        ttk.Button(top, text="上一月", command=lambda: shift_month(-1)).pack(side="left", padx=(0, 8))
        ttk.Label(top, text="年").pack(side="left")
        year_spin = ttk.Spinbox(top, from_=base_date.year - 30, to=base_date.year + 30, textvariable=year_var, width=8, command=render_calendar)
        year_spin.pack(side="left", padx=(2, 8))
        ttk.Label(top, text="月").pack(side="left")
        month_spin = ttk.Spinbox(top, from_=1, to=12, textvariable=month_var, width=5, command=render_calendar)
        month_spin.pack(side="left", padx=(2, 8))
        ttk.Button(top, text="下一月", command=lambda: shift_month(1)).pack(side="left")
        ttk.Button(top, text="今天", command=lambda: (year_var.set(date.today().year), month_var.set(date.today().month), render_calendar())).pack(side="left", padx=8)
        year_spin.bind("<Return>", lambda e: render_calendar())
        month_spin.bind("<Return>", lambda e: render_calendar())
        render_calendar()

    def refresh_contract_payment_cumulatives(self, mark_dirty=False):
        if not hasattr(self, "payment_contract_tree"):
            return
        rows = self.payment_contract_tree.get_rows()
        if not rows:
            return
        indexes = {field: idx for idx, field in enumerate(PAYMENT_CONTRACT_FIELDS)}

        def period_key(indexed_row):
            original_index, row = indexed_row
            text = str(row[indexes["period_no"]] if indexes["period_no"] < len(row) else "").strip()
            if not text:
                return (2, original_index)
            cleaned = text.replace("第", "").replace("期", "").strip()
            try:
                return (0, Decimal(cleaned), original_index)
            except InvalidOperation:
                parts = re.split(r"(\d+(?:\.\d+)?)", cleaned)
                key = []
                for part in parts:
                    if not part:
                        continue
                    try:
                        key.append((0, Decimal(part)))
                    except InvalidOperation:
                        key.append((1, part.lower()))
                return (1, key, original_index)

        prepared_rows = []
        for raw_row in rows:
            row = (list(raw_row) + [""] * len(PAYMENT_CONTRACT_FIELDS))[:len(PAYMENT_CONTRACT_FIELDS)]
            for field in PAYMENT_CONTRACT_MONEY_FIELDS - PAYMENT_CONTRACT_CUMULATIVE_FIELDS:
                idx = indexes[field]
                if str(row[idx]).strip():
                    row[idx] = self.format_money_value(row[idx])
            prepared_rows.append(row)

        calculated_rows = [list(row) for row in prepared_rows]
        cum_billing = Decimal("0")
        cum_billing_tax = Decimal("0")
        cum_retention = Decimal("0")
        cum_paid = Decimal("0")
        cum_paid_tax = Decimal("0")

        for original_index, row in sorted(enumerate(prepared_rows), key=period_key):
            cum_billing += self.money_decimal(row[indexes["billing_amount_untaxed"]])
            cum_billing_tax += self.money_decimal(row[indexes["billing_business_tax"]])
            cum_retention += self.money_decimal(row[indexes["retention_amount"]])
            cum_paid += self.money_decimal(row[indexes["paid_amount_untaxed"]])
            cum_paid_tax += self.money_decimal(row[indexes["paid_business_tax"]])
            calculated_rows[original_index][indexes["cumulative_billing_amount"]] = self.format_money_value(cum_billing) if cum_billing else ""
            calculated_rows[original_index][indexes["cumulative_billing_tax"]] = self.format_money_value(cum_billing_tax) if cum_billing_tax else ""
            calculated_rows[original_index][indexes["cumulative_retention_amount"]] = self.format_money_value(cum_retention) if cum_retention else ""
            calculated_rows[original_index][indexes["cumulative_paid_amount"]] = self.format_money_value(cum_paid) if cum_paid else ""
            calculated_rows[original_index][indexes["cumulative_paid_tax"]] = self.format_money_value(cum_paid_tax) if cum_paid_tax else ""

        self.payment_contract_tree.set_rows(calculated_rows)
        if mark_dirty:
            self.mark_dirty()

    def payment_contract_db_rows_to_tree(self, db_rows):
        rows = []
        for r in db_rows:
            has_new_data = any((r[field] if field in r.keys() and r[field] is not None else "") for field in PAYMENT_CONTRACT_FIELDS)
            if has_new_data:
                rows.append([r[field] if field in r.keys() and r[field] is not None else "" for field in PAYMENT_CONTRACT_FIELDS])
            else:
                legacy_amount = r["amount"] if "amount" in r.keys() else ""
                row = [""] * len(PAYMENT_CONTRACT_FIELDS)
                idx = {field: i for i, field in enumerate(PAYMENT_CONTRACT_FIELDS)}
                row[idx["period_no"]] = r["item"] if "item" in r.keys() else ""
                row[idx["owner_payment_date"]] = r["day"] if "day" in r.keys() else ""
                row[idx["estimated_amount_taxed"]] = self.format_money_value(legacy_amount) if str(legacy_amount).strip() else ""
                row[idx["payment_period"]] = r["note"] if "note" in r.keys() else ""
                rows.append(row)
        return rows

    def payment_contract_tree_rows_to_db(self):
        self.refresh_contract_payment_cumulatives(mark_dirty=False)
        rows = []
        for r in self.payment_contract_tree.get_rows():
            vals = (list(r) + [""] * len(PAYMENT_CONTRACT_FIELDS))[:len(PAYMENT_CONTRACT_FIELDS)]
            rows.append({field: vals[idx] for idx, field in enumerate(PAYMENT_CONTRACT_FIELDS)})
        return rows


    def on_payment_other_changed(self):
        self.refresh_payment_other_cumulatives(mark_dirty=False)
        self.mark_dirty()

    def on_payment_admin_changed(self):
        self.refresh_payment_admin_cumulatives(mark_dirty=False)
        self.refresh_admin_setup_totals()
        self.mark_dirty()

    def open_payment_other_add_dialog(self):
        self.open_payment_other_dialog(item=None)

    def open_payment_other_edit_dialog(self):
        if not hasattr(self, "payment_other_tree"):
            return
        tree = self.payment_other_tree.tree
        item = tree.focus()
        if not item or not tree.exists(item):
            selected = tree.selection()
            item = selected[0] if selected else ""
        if not item or not tree.exists(item):
            messagebox.showinfo("編輯資料", "請先選取要修改的資料列。")
            return
        tree.selection_set(item)
        tree.focus(item)
        self.open_payment_other_dialog(item=item)

    def open_payment_admin_add_dialog(self):
        self.open_payment_admin_dialog(item=None)

    def open_payment_admin_edit_dialog(self):
        if not hasattr(self, "payment_admin_tree"):
            return
        tree = self.payment_admin_tree.tree
        item = tree.focus()
        if not item or not tree.exists(item):
            selected = tree.selection()
            item = selected[0] if selected else ""
        if not item or not tree.exists(item):
            messagebox.showinfo("編輯資料", "請先選取要修改的資料列。")
            return
        tree.selection_set(item)
        tree.focus(item)
        self.open_payment_admin_dialog(item=item)

    def open_payment_other_dialog(self, item=None):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        tree = self.payment_other_tree
        is_edit = item is not None
        old_values = list(tree.tree.item(item, "values")) if is_edit else [""] * len(PAYMENT_OTHER_FIELDS)
        old_values = (old_values + [""] * len(PAYMENT_OTHER_FIELDS))[:len(PAYMENT_OTHER_FIELDS)]
        if not is_edit and not old_values[0]:
            old_values[0] = str(len(tree.get_rows()) + 1)
        self.open_structured_payment_dialog(
            title="編輯發包以外計價資料" if is_edit else "新增發包以外計價資料",
            fields=PAYMENT_OTHER_FIELDS,
            headings=PAYMENT_OTHER_HEADINGS,
            old_values=old_values,
            date_fields=PAYMENT_OTHER_DATE_FIELDS,
            money_fields=PAYMENT_OTHER_MONEY_FIELDS,
            readonly_fields=PAYMENT_OTHER_CUMULATIVE_FIELDS,
            combo_fields={"payment_item": PAYMENT_OTHER_ITEM_OPTIONS},
            on_ok=lambda values: self._finish_payment_other_dialog(tree, item, values, is_edit),
        )

    def _finish_payment_other_dialog(self, tree, item, values, is_edit):
        if is_edit:
            tree.tree.item(item, values=values)
        else:
            tree.add_row(values)
        self.refresh_payment_other_cumulatives(mark_dirty=True)

    def open_payment_admin_dialog(self, item=None):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "此工程已設定密碼，請先在上半部輸入正確編輯密碼解鎖。")
            return
        tree = self.payment_admin_tree
        is_edit = item is not None
        old_values = list(tree.tree.item(item, "values")) if is_edit else [""] * len(PAYMENT_ADMIN_FIELDS)
        old_values = (old_values + [""] * len(PAYMENT_ADMIN_FIELDS))[:len(PAYMENT_ADMIN_FIELDS)]
        if not is_edit and not old_values[0]:
            old_values[0] = str(len(tree.get_rows()) + 1)
        self.open_structured_payment_dialog(
            title="編輯管理費計價資料" if is_edit else "新增管理費計價資料",
            fields=PAYMENT_ADMIN_FIELDS,
            headings=PAYMENT_ADMIN_HEADINGS,
            old_values=old_values,
            date_fields=PAYMENT_ADMIN_DATE_FIELDS,
            money_fields=PAYMENT_ADMIN_MONEY_FIELDS,
            readonly_fields=PAYMENT_ADMIN_CUMULATIVE_FIELDS,
            combo_fields={},
            on_ok=lambda values: self._finish_payment_admin_dialog(tree, item, values, is_edit),
        )

    def _finish_payment_admin_dialog(self, tree, item, values, is_edit):
        if is_edit:
            tree.tree.item(item, values=values)
        else:
            tree.add_row(values)
        self.refresh_payment_admin_cumulatives(mark_dirty=True)

    def open_structured_payment_dialog(self, title, fields, headings, old_values, date_fields, money_fields, readonly_fields, combo_fields, on_ok):
        win = tk.Toplevel(self)
        win.title(title)
        win.transient(self)
        win.grab_set()
        win.geometry("760x520")
        win.minsize(620, 420)
        canvas = tk.Canvas(win, highlightthickness=0, background=OSX_WINDOW_BG)
        vs = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        form = ttk.Frame(canvas, padding=12)
        canvas_window = canvas.create_window((0, 0), window=form, anchor="nw")
        canvas.configure(yscrollcommand=vs.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        win.grid_rowconfigure(0, weight=1)
        win.grid_columnconfigure(0, weight=1)
        form.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=max(e.width, form.winfo_reqwidth())))

        vars_map = {}
        first_entry = None
        for idx, (field, heading) in enumerate(zip(fields, headings)):
            row = idx // 2
            col = (idx % 2) * 3
            ttk.Label(form, text=heading + "：").grid(row=row, column=col, sticky="e", padx=(4, 2), pady=4)
            var = tk.StringVar(value=old_values[idx] if idx < len(old_values) else "")
            vars_map[field] = var
            holder = ttk.Frame(form)
            holder.grid(row=row, column=col + 1, sticky="ew", padx=(2, 8), pady=4)
            form.grid_columnconfigure(col + 1, weight=1)
            if field in date_fields:
                ent = ttk.Entry(holder, textvariable=var, width=16)
                ent.pack(side="left", fill="x", expand=True)
                ttk.Button(holder, text="▼", width=3, command=lambda v=var, w=win, h=heading: self.open_payment_date_picker(v, w, h)).pack(side="left", padx=(3, 0))
            elif field in combo_fields:
                ent = ttk.Combobox(holder, textvariable=var, values=combo_fields[field], state="readonly", width=18)
                ent.pack(fill="x", expand=True)
            else:
                ent = ttk.Entry(holder, textvariable=var, width=22)
                ent.pack(fill="x", expand=True)
            if first_entry is None:
                first_entry = ent
            if field in readonly_fields:
                ent.configure(state="readonly")
            if field in money_fields and field not in readonly_fields:
                ent.bind("<FocusOut>", lambda e, v=var: v.set(self.format_money_value(v.get()) if v.get().strip() else ""))
                ent.bind("<Return>", lambda e, v=var: v.set(self.format_money_value(v.get()) if v.get().strip() else ""))

        def ok():
            values = []
            for field in fields:
                raw = vars_map[field].get().strip()
                if field in money_fields and raw:
                    raw = self.format_money_value(raw)
                values.append(raw)
            win.destroy()
            on_ok(values)

        btns = ttk.Frame(win, padding=10)
        btns.grid(row=1, column=0, columnspan=2, sticky="ew")
        ttk.Button(btns, text="取消", command=win.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="確定", command=ok).pack(side="right", padx=4)
        if first_entry:
            first_entry.focus_set()

    def _remember_and_restore_tree_focus(self, editable_tree, fields, key_field, before_values=None):
        """Return a restorer that re-selects the same logical row after set_rows()."""
        target_key = ""
        if before_values:
            idx = {field: i for i, field in enumerate(fields)}
            target_key = str(before_values[idx.get(key_field, 0)] or "").strip()
        if not target_key:
            tree = editable_tree.tree
            item = tree.focus()
            if item and tree.exists(item):
                vals = list(tree.item(item, "values"))
                idx = {field: i for i, field in enumerate(fields)}
                target_key = str(vals[idx.get(key_field, 0)] or "").strip()
        def restore():
            if not target_key:
                return
            tree = editable_tree.tree
            idx = {field: i for i, field in enumerate(fields)}
            key_idx = idx.get(key_field, 0)
            for child in tree.get_children():
                vals = list(tree.item(child, "values"))
                if key_idx < len(vals) and str(vals[key_idx] or "").strip() == target_key:
                    tree.selection_set(child)
                    tree.focus(child)
                    tree.see(child)
                    break
        return restore

    def refresh_payment_other_cumulatives(self, mark_dirty=False):
        if not hasattr(self, "payment_other_tree"):
            return
        restore_focus = self._remember_and_restore_tree_focus(self.payment_other_tree, PAYMENT_OTHER_FIELDS, "period_no")
        rows = self.payment_other_tree.get_rows()
        if not rows:
            return
        indexes = {field: idx for idx, field in enumerate(PAYMENT_OTHER_FIELDS)}
        prepared_rows = []
        for raw_row in rows:
            row = (list(raw_row) + [""] * len(PAYMENT_OTHER_FIELDS))[:len(PAYMENT_OTHER_FIELDS)]
            for field in PAYMENT_OTHER_MONEY_FIELDS - PAYMENT_OTHER_CUMULATIVE_FIELDS:
                idx = indexes[field]
                if str(row[idx]).strip():
                    row[idx] = self.format_money_value(row[idx])
            prepared_rows.append(row)
        calculated_rows = [list(row) for row in prepared_rows]
        cum_amount = Decimal("0")
        cum_alloc = Decimal("0")
        for original_index, row in sorted(enumerate(prepared_rows), key=lambda x: period_sort_key(x[1][indexes["period_no"]], x[0])):
            cum_amount += self.money_decimal(row[indexes["payment_amount"]])
            cum_alloc += self.money_decimal(row[indexes["management_fee_allocation"]])
            calculated_rows[original_index][indexes["cumulative_amount"]] = self.format_money_value(cum_amount) if cum_amount else ""
            calculated_rows[original_index][indexes["cumulative_management_fee_allocation"]] = self.format_money_value(cum_alloc) if cum_alloc else ""
        self.payment_other_tree.set_rows(calculated_rows)
        self.configure_tree_alignment(self.payment_other_tree.tree, PAYMENT_OTHER_FIELDS, PAYMENT_OTHER_MONEY_FIELDS, PAYMENT_OTHER_LEFT_FIELDS)
        restore_focus()
        if mark_dirty:
            self.mark_dirty()

    def refresh_payment_admin_cumulatives(self, mark_dirty=False):
        if not hasattr(self, "payment_admin_tree"):
            return
        restore_focus = self._remember_and_restore_tree_focus(self.payment_admin_tree, PAYMENT_ADMIN_FIELDS, "period_no")
        rows = self.payment_admin_tree.get_rows()
        if not rows:
            self.refresh_admin_setup_totals()
            return
        indexes = {field: idx for idx, field in enumerate(PAYMENT_ADMIN_FIELDS)}
        prepared_rows = []
        for raw_row in rows:
            row = (list(raw_row) + [""] * len(PAYMENT_ADMIN_FIELDS))[:len(PAYMENT_ADMIN_FIELDS)]
            current_amount = (
                self.money_decimal(row[indexes["travel_fee_billing"]]) +
                self.money_decimal(row[indexes["overtime_fee_billing"]]) +
                self.money_decimal(row[indexes["other_fee_billing"]])
            )
            row[indexes["current_amount"]] = self.format_money_value(current_amount) if current_amount else ""
            for field in PAYMENT_ADMIN_MONEY_FIELDS - PAYMENT_ADMIN_CUMULATIVE_FIELDS:
                idx = indexes[field]
                if str(row[idx]).strip():
                    row[idx] = self.format_money_value(row[idx])
            prepared_rows.append(row)
        calculated_rows = [list(row) for row in prepared_rows]
        cumulative = Decimal("0")
        cumulative_tax = Decimal("0")
        for original_index, row in sorted(enumerate(prepared_rows), key=lambda x: period_sort_key(x[1][indexes["period_no"]], x[0])):
            cumulative_tax += self.money_decimal(row[indexes["tax_amount"]])
            cumulative += self.money_decimal(row[indexes["current_amount"]])
            calculated_rows[original_index][indexes["cumulative_tax_amount"]] = self.format_money_value(cumulative_tax) if cumulative_tax else ""
            calculated_rows[original_index][indexes["cumulative_amount"]] = self.format_money_value(cumulative) if cumulative else ""
        self.payment_admin_tree.set_rows(calculated_rows)
        self.configure_tree_alignment(self.payment_admin_tree.tree, PAYMENT_ADMIN_FIELDS, PAYMENT_ADMIN_MONEY_FIELDS)
        restore_focus()
        self.refresh_admin_setup_totals()
        if mark_dirty:
            self.mark_dirty()

    def payment_other_db_rows_to_tree(self, db_rows):
        rows = []
        for r in db_rows:
            has_new_data = any((r[field] if field in r.keys() and r[field] is not None else "") for field in PAYMENT_OTHER_FIELDS)
            if has_new_data:
                rows.append([r[field] if field in r.keys() and r[field] is not None else "" for field in PAYMENT_OTHER_FIELDS])
            else:
                row = [""] * len(PAYMENT_OTHER_FIELDS)
                idx = {field: i for i, field in enumerate(PAYMENT_OTHER_FIELDS)}
                row[idx["period_no"]] = r["voucher_no"] if "voucher_no" in r.keys() else ""
                row[idx["payment_date"]] = r["day"] if "day" in r.keys() else ""
                row[idx["payment_item"]] = r["item"] if "item" in r.keys() else ""
                row[idx["payment_amount"]] = self.format_money_value(r["amount"]) if "amount" in r.keys() and str(r["amount"]).strip() else ""
                rows.append(row)
        return rows

    def payment_other_tree_rows_to_db(self):
        self.refresh_payment_other_cumulatives(mark_dirty=False)
        rows = []
        for r in self.payment_other_tree.get_rows():
            vals = (list(r) + [""] * len(PAYMENT_OTHER_FIELDS))[:len(PAYMENT_OTHER_FIELDS)]
            rows.append({field: vals[idx] for idx, field in enumerate(PAYMENT_OTHER_FIELDS)})
        return rows

    def payment_admin_db_rows_to_tree(self, db_rows):
        rows = []
        for r in db_rows:
            has_new_data = any((r[field] if field in r.keys() and r[field] is not None else "") for field in PAYMENT_ADMIN_FIELDS)
            if has_new_data:
                rows.append([r[field] if field in r.keys() and r[field] is not None else "" for field in PAYMENT_ADMIN_FIELDS])
            else:
                row = [""] * len(PAYMENT_ADMIN_FIELDS)
                idx = {field: i for i, field in enumerate(PAYMENT_ADMIN_FIELDS)}
                row[idx["period_no"]] = r["voucher_no"] if "voucher_no" in r.keys() else ""
                row[idx["payment_date"]] = r["day"] if "day" in r.keys() else ""
                row[idx["current_amount"]] = self.format_money_value(r["amount"]) if "amount" in r.keys() and str(r["amount"]).strip() else ""
                rows.append(row)
        return rows

    def payment_admin_tree_rows_to_db(self):
        self.refresh_payment_admin_cumulatives(mark_dirty=False)
        rows = []
        for r in self.payment_admin_tree.get_rows():
            vals = (list(r) + [""] * len(PAYMENT_ADMIN_FIELDS))[:len(PAYMENT_ADMIN_FIELDS)]
            rows.append({field: vals[idx] for idx, field in enumerate(PAYMENT_ADMIN_FIELDS)})
        return rows

    def format_admin_setup_money(self, key):
        if not hasattr(self, "admin_calc_vars") or key not in self.admin_calc_vars:
            return
        raw = self.admin_calc_vars[key].get().strip()
        if raw:
            self.admin_calc_vars[key].set(self.format_money_value(raw))

    def on_admin_setup_changed(self, *_):
        if self.loading or self.restoring:
            return
        self.refresh_admin_setup_totals()
        self.mark_dirty()

    def refresh_admin_management_fee_source(self):
        if not hasattr(self, "admin_calc_vars"):
            return
        self.admin_calc_vars["mgmt_fee"].set(self.basic_vars.get("award_mgmt_fee", tk.StringVar()).get())

    def refresh_admin_setup_totals(self):
        if not hasattr(self, "admin_calc_vars"):
            return
        def d(key):
            return self.money_decimal(self.admin_calc_vars.get(key, tk.StringVar()).get())
        available_0c12 = sum(d(f"alloc{i}_0c12") for i in range(1, 5))
        available_0c11 = sum(d(f"alloc{i}_0c11") for i in range(1, 5))
        available_0c14 = sum(d(f"alloc{i}_0c14") for i in range(1, 5))
        self.admin_calc_vars["available_0c12"].set(self.format_money_value(available_0c12) if available_0c12 else "")
        self.admin_calc_vars["available_0c11"].set(self.format_money_value(available_0c11) if available_0c11 else "")
        self.admin_calc_vars["available_0c14"].set(self.format_money_value(available_0c14) if available_0c14 else "")

        used_0c12 = used_0c11 = used_0c14 = Decimal("0")
        if hasattr(self, "payment_admin_tree"):
            idx = {field: i for i, field in enumerate(PAYMENT_ADMIN_FIELDS)}
            for row in self.payment_admin_tree.get_rows():
                vals = (list(row) + [""] * len(PAYMENT_ADMIN_FIELDS))[:len(PAYMENT_ADMIN_FIELDS)]
                used_0c12 += self.money_decimal(vals[idx["travel_fee_billing"]])
                used_0c11 += self.money_decimal(vals[idx["overtime_fee_billing"]])
                used_0c14 += self.money_decimal(vals[idx["other_fee_billing"]])
        self.admin_calc_vars["used_0c12"].set(self.format_money_value(used_0c12) if used_0c12 else "")
        self.admin_calc_vars["used_0c11"].set(self.format_money_value(used_0c11) if used_0c11 else "")
        self.admin_calc_vars["used_0c14"].set(self.format_money_value(used_0c14) if used_0c14 else "")

        for key, used, available in [
            ("0c12", used_0c12, available_0c12),
            ("0c11", used_0c11, available_0c11),
            ("0c14", used_0c14, available_0c14),
        ]:
            available_widget = getattr(self, f"admin_available_{key}_entry", None)
            if available_widget:
                try:
                    style_name = "AdminAvailableWarning.TEntry" if available and used > available else "AdminAvailable.TEntry"
                    available_widget.configure(style=style_name)
                except tk.TclError:
                    pass



    def on_change_award_field_changed(self, *_):
        if self.loading or self.restoring:
            return
        self.refresh_change_award_calculated_fields()
        self.dirty = True
        self.schedule_auto_save()

    def format_change_award_money(self, key):
        if hasattr(self, "change_award_vars") and key in self.change_award_vars:
            raw = self.change_award_vars[key].get().strip()
            if raw:
                self.change_award_vars[key].set(self.format_money_value(raw))

    def refresh_change_award_select(self):
        if not hasattr(self, "change_award_combo") or not self.current_project_id:
            return
        nums = self.db.change_numbers(self.current_project_id)
        self.change_award_combo["values"] = nums
        if nums:
            if self.change_award_select_var.get() not in nums:
                self.change_award_select_var.set(nums[-1])
            self.load_selected_change_award_fields()
        else:
            self.change_award_select_var.set("")
            for var in getattr(self, "change_award_vars", {}).values():
                var.set("")

    def load_selected_change_award_fields(self):
        if not hasattr(self, "change_award_vars") or not self.current_project_id:
            return
        change_no = self.change_award_select_var.get().strip()
        if not change_no:
            return
        record = self.db.change_record(self.current_project_id, change_no)
        fields = record.get("fields", {}) if record else {}
        self.loading = True
        try:
            for key, var in self.change_award_vars.items():
                var.set(fields.get(key, ""))
        finally:
            self.loading = False
        self.refresh_change_award_calculated_fields()

    def save_current_change_award_fields(self):
        if not hasattr(self, "change_award_vars") or not self.current_project_id:
            return
        change_no = self.change_award_select_var.get().strip()
        if not change_no:
            nums = self.db.change_numbers(self.current_project_id)
            change_no = nums[-1] if nums else "1"
            self.change_award_select_var.set(change_no)
        record = self.db.change_record(self.current_project_id, change_no)
        fields = dict(record.get("fields", {})) if record else {"change_no": change_no}
        for key, var in self.change_award_vars.items():
            fields[key] = var.get().strip()
        demand = record.get("demand", []) if record else []
        confirm = record.get("confirm", []) if record else []
        budget = record.get("budget", []) if record else []
        source = record.get("source_file", "") if record else ""
        self.db.save_change_record(self.current_project_id, change_no, fields, demand, confirm, budget, source)
        self.refresh_change_select()
        self.refresh_change_award_select()

    def refresh_change_award_calculated_fields(self):
        if not hasattr(self, "change_award_vars"):
            return
        net = self.safe_amount(self.change_award_vars.get("change_award_contract_amount", tk.StringVar()).get())
        tax = net * 0.05
        total = net + tax
        air = self.safe_amount(self.change_award_vars.get("change_award_air_pollution_fee", tk.StringVar()).get())
        outside_without_air = sum(
            self.safe_amount(self.change_award_vars.get(key, tk.StringVar()).get())
            for key in (
                "change_award_mgmt_fee", "change_award_self_labor", "change_award_self_material", "change_award_spare_material",
                "change_award_railway_material", "change_award_supervision_fee", "change_award_freight", "change_award_other"
            )
        )
        unfinished = net + outside_without_air
        base = self.safe_amount(self.change_award_vars.get("change_award_base_price", tk.StringVar()).get())
        budget_total = self.safe_amount(self.basic_vars.get("budget_contract_total", tk.StringVar()).get())
        readonly_values = {
            "change_award_contract_tax": tax,
            "change_award_contract_total": total,
            "change_award_input_tax": tax,
            "change_award_unfinished_amount": unfinished,
            "change_award_total_amount": unfinished + air + tax,
        }
        for key, value in readonly_values.items():
            if key in self.change_award_vars:
                text = self.money_text(value) if value else ""
                if self.change_award_vars[key].get() != text:
                    self.change_award_vars[key].set(text)
        ratios = {
            "change_award_contract_budget_ratio": self.percent_text(total, budget_total),
            "change_award_contract_base_ratio": self.percent_text(total, base),
            "change_award_base_budget_ratio": self.percent_text(base, budget_total),
        }
        for key, value in ratios.items():
            if key in self.change_award_vars and self.change_award_vars[key].get() != value:
                self.change_award_vars[key].set(value)

    def load_admin_setup_from_project(self, project_row):
        if not hasattr(self, "admin_calc_vars"):
            return
        self.refresh_admin_management_fee_source()
        mapping = {
            "change_mgmt_fee": "admin_change_mgmt_fee",
            "alloc1_amount": "admin_alloc1_amount", "alloc1_0c12": "admin_alloc1_0c12", "alloc1_0c11": "admin_alloc1_0c11", "alloc1_0c14": "admin_alloc1_0c14",
            "alloc2_amount": "admin_alloc2_amount", "alloc2_0c12": "admin_alloc2_0c12", "alloc2_0c11": "admin_alloc2_0c11", "alloc2_0c14": "admin_alloc2_0c14",
            "alloc3_amount": "admin_alloc3_amount", "alloc3_0c12": "admin_alloc3_0c12", "alloc3_0c11": "admin_alloc3_0c11", "alloc3_0c14": "admin_alloc3_0c14",
            "alloc4_amount": "admin_alloc4_amount", "alloc4_0c12": "admin_alloc4_0c12", "alloc4_0c11": "admin_alloc4_0c11", "alloc4_0c14": "admin_alloc4_0c14",
        }
        for var_key, db_key in mapping.items():
            if var_key in self.admin_calc_vars:
                self.admin_calc_vars[var_key].set(project_row[db_key] if db_key in project_row.keys() and project_row[db_key] is not None else "")
        self.refresh_admin_setup_totals()

    def collect_admin_setup_to_project_data(self, data):
        if not hasattr(self, "admin_calc_vars"):
            return
        mapping = {
            "admin_change_mgmt_fee": "change_mgmt_fee",
            "admin_alloc1_amount": "alloc1_amount", "admin_alloc1_0c12": "alloc1_0c12", "admin_alloc1_0c11": "alloc1_0c11", "admin_alloc1_0c14": "alloc1_0c14",
            "admin_alloc2_amount": "alloc2_amount", "admin_alloc2_0c12": "alloc2_0c12", "admin_alloc2_0c11": "alloc2_0c11", "admin_alloc2_0c14": "alloc2_0c14",
            "admin_alloc3_amount": "alloc3_amount", "admin_alloc3_0c12": "alloc3_0c12", "admin_alloc3_0c11": "alloc3_0c11", "admin_alloc3_0c14": "alloc3_0c14",
            "admin_alloc4_amount": "alloc4_amount", "admin_alloc4_0c12": "alloc4_0c12", "admin_alloc4_0c11": "alloc4_0c11", "admin_alloc4_0c14": "alloc4_0c14",
        }
        for db_key, var_key in mapping.items():
            data[db_key] = self.admin_calc_vars.get(var_key, tk.StringVar()).get().strip()

    def build_execution_tab(self):
        self.add_page_edit_toggle(self.tab_execution)
        ttk.Label(self.tab_execution, text="會議記錄表：按「新增一列」用行事曆新增；時間與標題會同步加入工程大事記。").pack(anchor="w")
        self.execution_tree = EditableTree(
            self.tab_execution,
            ["day", "record_type", "subject", "content", "note"],
            ["時間", "類型", "標題", "內容", "備註"],
            [120, 140, 220, 420, 220],
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
        win.title("新增會議記錄")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)

        year_var = tk.IntVar(value=base_date.year)
        month_var = tk.IntVar(value=base_date.month)
        selected_day_var = tk.StringVar(value="")
        type_var = tk.StringVar(value="工作會議")
        subject_var = tk.StringVar(value="")
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
        ttk.Label(win, textvariable=selected_day_var, foreground=OSX_ACCENT).pack(anchor="w", padx=12)

        form = ttk.Frame(win, padding=10)
        form.pack(fill="x")
        ttk.Label(form, text="類型").grid(row=0, column=0, sticky="e", padx=(0, 6), pady=4)
        ttk.Combobox(
            form,
            textvariable=type_var,
            values=["工作會議", "會勘", "變更需求會議", "變更確認會議", "其他"],
            state="readonly",
            width=20
        ).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(form, text="標題").grid(row=1, column=0, sticky="e", padx=(0, 6), pady=4)
        ttk.Entry(form, textvariable=subject_var, width=48).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Label(form, text="內容").grid(row=2, column=0, sticky="ne", padx=(0, 6), pady=4)
        content_text = tk.Text(form, height=5, width=48, wrap="word", font=("Microsoft JhengHei UI", 10), background=OSX_ENTRY_BG, foreground=OSX_TEXT, insertbackground=OSX_TEXT, relief="solid", bd=1)
        content_text.grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Label(form, text="備註").grid(row=3, column=0, sticky="e", padx=(0, 6), pady=4)
        note_entry = ttk.Entry(form, textvariable=note_var, width=48)
        note_entry.grid(row=3, column=1, sticky="ew", pady=4)
        form.grid_columnconfigure(1, weight=1)

        def select_day(day):
            selected_day_var.set(day.strftime("%Y-%m-%d"))
            subject_var.set(subject_var.get().strip())
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
            subject = subject_var.get().strip()
            if not subject:
                messagebox.showwarning("尚未輸入標題", "請輸入標題。", parent=win)
                return
            self.execution_tree.add_row_after_selection([day_text, type_var.get(), subject, content, note_var.get().strip()])
            self.add_meeting_to_milestones(day_text, subject, type_var.get(), content)
            selected_day_var.set("")
            subject_var.set("")
            content_text.delete("1.0", "end")
            note_var.set("")
            self.refresh_milestone_rows()
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

    def normalize_milestone_rows(self, rows):
        normalized = [self.calc_milestone_row(list(r)) for r in rows]

        def sort_key(row):
            start = parse_date(row[2] if len(row) > 2 else "")
            deadline = parse_date(row[4] if len(row) > 4 else "")
            try:
                original_no = int(float(row[0] or 0))
            except ValueError:
                original_no = 0
            return (
                start or date.max,
                deadline or date.max,
                original_no,
                row[1] if len(row) > 1 else "",
            )

        normalized.sort(key=sort_key)
        for index, row in enumerate(normalized, start=1):
            row[0] = str(index)
        return normalized

    def refresh_milestone_rows(self):
        if not hasattr(self, "milestone_tree"):
            return
        rows = self.normalize_milestone_rows(self.milestone_tree.get_rows())
        self.milestone_tree.set_rows(rows)
        self.milestone_tree.apply_row_tags(lambda r: "pink" if len(r) > 5 and not r[5] else "")

    def add_meeting_to_milestones(self, day_text, subject, record_type="", content=""):
        if not hasattr(self, "milestone_tree"):
            return
        note_parts = [record_type.strip(), content.strip()]
        note = " / ".join(part for part in note_parts if part)
        self.milestone_tree.add_row([
            "", subject.strip(), day_text.strip(), "", "", "", "", "", note, ""
        ])

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
        vars_["item_no"].set(str(len(self.milestone_tree.get_rows()) + 1))
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

    def build_budget_data_tab(self):
        self.add_page_edit_toggle(self.tab_budget_data)
        container = ttk.Frame(self.tab_budget_data)
        container.pack(fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1, uniform="budget_data")
        container.grid_columnconfigure(1, weight=1, uniform="budget_data")

        budget_box = tk.LabelFrame(
            container,
            text="預算",
            bg=OSX_BUDGET_BOOK_BG,
            fg=OSX_TEXT,
            bd=2,
            relief="solid",
            labelanchor="n",
            font=("Microsoft JhengHei UI", 12, "bold"),
            padx=10,
            pady=10,
        )
        contract_box = tk.LabelFrame(
            container,
            text="契約",
            bg=OSX_CONTRACT_BOOK_BG,
            fg=OSX_TEXT,
            bd=2,
            relief="solid",
            labelanchor="n",
            font=("Microsoft JhengHei UI", 12, "bold"),
            padx=10,
            pady=10,
        )
        budget_box.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=6)
        contract_box.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=6)
        for box in (budget_box, contract_box):
            box.grid_rowconfigure(0, weight=1)
            box.grid_columnconfigure(0, weight=1)
        self.budget_book_trees = {}
        self.budget_book_sources = {}
        self.create_budget_book_panel(budget_box, "budget", "讀入預算書", OSX_BUDGET_BOOK_BG)
        self.create_budget_book_panel(contract_box, "contract", "讀入預算書", OSX_CONTRACT_BOOK_BG)

    def create_budget_book_panel(self, parent, area, button_text, bg):
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        top = tk.Frame(parent, bg=bg)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(top, text=button_text, command=lambda a=area: self.import_budget_book(a)).pack(side="left")
        source_var = tk.StringVar(value="")
        self.budget_book_sources[area] = source_var
        ttk.Label(top, textvariable=source_var).pack(side="left", padx=8)
        tree = EditableTree(
            parent,
            ["c1", "c2", "c3", "c4", "c5", "c6"],
            ["欄1", "欄2", "欄3", "欄4", "欄5", "欄6"],
            [120, 120, 120, 120, 120, 120],
            self.mark_dirty,
        )
        tree.grid(row=1, column=0, sticky="nsew")
        self.budget_book_trees[area] = tree

    def set_budget_book_rows(self, area, rows, source_file=""):
        if area not in getattr(self, "budget_book_trees", {}):
            return
        normalized = [(list(row) + [""] * 6)[:6] for row in rows]
        self.budget_book_trees[area].set_rows(normalized)
        if area in self.budget_book_sources:
            self.budget_book_sources[area].set(os.path.basename(source_file) if source_file else "")

    def import_budget_book(self, area):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "請先解除鎖定。")
            return
        path = filedialog.askopenfilename(
            title="選擇預算書 Excel 檔",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not path:
            return
        try:
            rows = read_simple_xlsx(path)
            if rows and len(rows[0]) == 1:
                rows = rows[1:]
            self.set_budget_book_rows(area, rows, path)
            if self.current_project_id:
                self.db.save_budget_book(self.current_project_id, area, rows, path)
            self.mark_dirty()
            self.status_var.set(f"已讀入預算書：{os.path.basename(path)}")
        except Exception as exc:
            messagebox.showerror("讀入預算書失敗", str(exc))

    def build_change_data_tab(self):
        self.add_page_edit_toggle(self.tab_change_data)
        self.change_upper_visible = True
        self.change_vars = {k: tk.StringVar() for k in [
            "change_no", "change_total", "add_original", "deduct_original", "add_new",
            "send_public_works", "send_accounting", "accounting_approved",
            "public_works_approved", "company_approved", "aa_review_done",
        ]}
        self.change_vars["change_no"].set("1")

        top_bar = ttk.Frame(self.tab_change_data)
        top_bar.pack(fill="x", pady=(0, 6))
        ttk.Button(top_bar, text="▲", width=3, command=self.toggle_change_upper).pack(side="right")
        self.change_toggle_btn = top_bar.winfo_children()[-1]
        ttk.Button(top_bar, text="儲存本次變更", command=self.save_change_data).pack(side="right", padx=4)

        self.change_upper = ttk.Frame(self.tab_change_data)
        self.change_upper.pack(fill="x")
        header = ttk.Frame(self.change_upper)
        header.pack(fill="x", pady=(0, 6))
        ttk.Label(header, text="第幾次變更預算").pack(side="left")
        ttk.Spinbox(header, from_=1, to=999, textvariable=self.change_vars["change_no"], width=8).pack(side="left", padx=6)
        ttk.Button(header, text="讀入預算書", command=self.import_change_budget_book).pack(side="left", padx=8)

        history = ttk.LabelFrame(self.change_upper, text="變更歷程", padding=8)
        history.pack(fill="x", pady=4)
        hist_left = ttk.Frame(history)
        hist_left.pack(side="left", fill="both", expand=True, padx=(0, 4))
        hist_right = ttk.Frame(history)
        hist_right.pack(side="left", fill="both", expand=True, padx=(4, 0))
        ttk.Button(hist_left, text="新增需求會議", command=lambda: self.open_change_meeting_dialog("demand")).pack(anchor="w")
        self.change_demand_tree = EditableTree(hist_left, ["no", "day", "doc_no"], ["第幾次需求會議", "會議時間", "文號"], [130, 120, 180], self.mark_dirty)
        self.change_demand_tree.pack(fill="both", expand=True, pady=4)
        ttk.Button(hist_right, text="新增確認會議", command=lambda: self.open_change_meeting_dialog("confirm")).pack(anchor="w")
        self.change_confirm_tree = EditableTree(hist_right, ["no", "day", "doc_no"], ["第幾次確認會議", "會議時間", "文號"], [130, 120, 180], self.mark_dirty)
        self.change_confirm_tree.pack(fill="both", expand=True, pady=4)

        money_box = ttk.LabelFrame(self.change_upper, text="變更金額", padding=8)
        money_box.pack(fill="x", pady=4)
        money_fields = [
            ("變更金額總計", "change_total"), ("原工項追加金額", "add_original"),
            ("原工項追減金額", "deduct_original"), ("新增工項金額", "add_new"),
        ]
        for i, (label, key) in enumerate(money_fields):
            ttk.Label(money_box, text=label).grid(row=0, column=i*2, sticky="e", padx=4, pady=3)
            ent = ttk.Entry(money_box, textvariable=self.change_vars[key], width=16)
            ent.grid(row=0, column=i*2+1, sticky="ew", padx=4, pady=3)
            self.edit_widgets.append(ent)
        for i in range(8):
            money_box.grid_columnconfigure(i, weight=1)

        reason_box = ttk.LabelFrame(self.change_upper, text="變更事由", padding=8)
        reason_box.pack(fill="x", pady=4)
        self.change_reason_text = tk.Text(reason_box, height=5, width=80, wrap="word", relief="solid", bd=1, background=OSX_ENTRY_BG, foreground=OSX_TEXT, insertbackground=OSX_TEXT)
        self.change_reason_text.pack(fill="x", expand=True)
        self.edit_widgets.append(self.change_reason_text)

        send_box = ttk.LabelFrame(self.change_upper, text="送預算書時間", padding=8)
        send_box.pack(fill="x", pady=4)
        date_fields = [
            ("工務處", "send_public_works"), ("主計單位", "send_accounting"),
            ("主計單位簽準", "accounting_approved"), ("工務處簽準", "public_works_approved"),
            ("公司簽準", "company_approved"), ("AA預算書審核完成", "aa_review_done"),
        ]
        for i, (label, key) in enumerate(date_fields):
            row, col = divmod(i, 3)
            ttk.Label(send_box, text=label).grid(row=row, column=col*3, sticky="e", padx=4, pady=3)
            ent = ttk.Entry(send_box, textvariable=self.change_vars[key], width=14)
            ent.grid(row=row, column=col*3+1, sticky="ew", padx=2, pady=3)
            ttk.Button(send_box, text="▼", width=3, command=lambda v=self.change_vars[key]: self.open_date_picker(v, "選擇日期")).grid(row=row, column=col*3+2, padx=2, pady=3)
            self.edit_widgets.append(ent)
        for i in range(9):
            send_box.grid_columnconfigure(i, weight=1)

        lower = ttk.LabelFrame(self.tab_change_data, text="變更情形與預算書資料區", padding=8)
        lower.pack(fill="both", expand=True, pady=(8, 0))
        select_row = ttk.Frame(lower)
        select_row.pack(fill="x", pady=(0, 6))
        ttk.Label(select_row, text="第幾次預算變更").pack(side="left")
        self.change_select_var = tk.StringVar()
        self.change_select_combo = ttk.Combobox(select_row, textvariable=self.change_select_var, state="readonly", width=12)
        self.change_select_combo.pack(side="left", padx=6)
        self.change_select_combo.bind("<<ComboboxSelected>>", lambda e: self.show_selected_change())
        ttk.Button(select_row, text="顯示", command=self.show_selected_change).pack(side="left")
        self.change_summary = tk.Text(lower, height=8, wrap="word", background=OSX_READONLY_BG, foreground=OSX_TEXT, insertbackground=OSX_TEXT, relief="solid", bd=1)
        self.change_summary.pack(fill="x", pady=(0, 6))
        self.change_budget_tree = EditableTree(lower, ["c1", "c2", "c3", "c4", "c5", "c6"], ["欄1", "欄2", "欄3", "欄4", "欄5", "欄6"], [120, 120, 120, 120, 120, 120], self.mark_dirty)
        self.change_budget_tree.pack(fill="both", expand=True)

    def toggle_change_upper(self):
        if self.change_upper_visible:
            self.change_upper.pack_forget()
            self.change_upper_visible = False
            self.change_toggle_btn.configure(text="▼")
        else:
            self.change_upper.pack(fill="x", before=self.tab_change_data.winfo_children()[-1])
            self.change_upper_visible = True
            self.change_toggle_btn.configure(text="▲")

    def open_change_meeting_dialog(self, kind):
        if not self.can_edit():
            messagebox.showwarning("編輯鎖定", "請先解除鎖定。")
            return
        win = tk.Toplevel(self)
        win.title("新增需求會議" if kind == "demand" else "新增確認會議")
        win.transient(self)
        win.grab_set()
        no_var = tk.StringVar()
        day_var = tk.StringVar()
        doc_var = tk.StringVar()
        form = ttk.Frame(win, padding=10)
        form.pack(fill="x")
        labels = [("第幾次需求會議" if kind == "demand" else "第幾次確認會議", no_var), ("會議時間", day_var), ("文號", doc_var)]
        for i, (label, var) in enumerate(labels):
            ttk.Label(form, text=label).grid(row=i, column=0, sticky="e", padx=4, pady=3)
            ttk.Entry(form, textvariable=var, width=28).grid(row=i, column=1, sticky="ew", padx=4, pady=3)
            if i == 1:
                ttk.Button(form, text="▼", width=3, command=lambda: self.open_date_picker(day_var, "選擇會議時間")).grid(row=i, column=2, padx=2)
        form.grid_columnconfigure(1, weight=1)
        def ok():
            tree = self.change_demand_tree if kind == "demand" else self.change_confirm_tree
            tree.add_row_after_selection([no_var.get().strip(), day_var.get().strip(), doc_var.get().strip()])
            win.destroy()
        ttk.Button(win, text="新增", command=ok).pack(side="right", padx=10, pady=10)
        ttk.Button(win, text="取消", command=win.destroy).pack(side="right", pady=10)

    def import_change_budget_book(self):
        path = filedialog.askopenfilename(title="選擇變更預算書 Excel 檔", filetypes=[("Excel Workbook", "*.xlsx")])
        if not path:
            return
        try:
            rows = read_simple_xlsx(path)
            if rows and len(rows[0]) == 1:
                rows = rows[1:]
            self.change_budget_tree.set_rows([(list(row) + [""] * 6)[:6] for row in rows])
            self.change_budget_source = path
            self.mark_dirty()
            self.status_var.set(f"已讀入變更預算書：{os.path.basename(path)}")
        except Exception as exc:
            messagebox.showerror("讀入預算書失敗", str(exc))

    def save_change_data(self):
        if not self.current_project_id:
            return
        change_no = self.change_vars["change_no"].get().strip() or "1"
        fields = {k: v.get().strip() for k, v in self.change_vars.items()}
        fields["reason"] = self.change_reason_text.get("1.0", "end-1c")
        self.db.save_change_record(
            self.current_project_id,
            change_no,
            fields,
            self.change_demand_tree.get_rows(),
            self.change_confirm_tree.get_rows(),
            self.change_budget_tree.get_rows(),
            getattr(self, "change_budget_source", ""),
        )
        self.refresh_change_select()
        self.refresh_change_award_select()
        self.status_var.set(f"已儲存第 {change_no} 次預算變更")

    def refresh_change_select(self):
        if not self.current_project_id or not hasattr(self, "change_select_combo"):
            return
        nums = self.db.change_numbers(self.current_project_id)
        self.change_select_combo["values"] = nums
        if nums and self.change_select_var.get() not in nums:
            self.change_select_var.set(nums[-1])

    def clear_change_editor(self):
        if not hasattr(self, "change_vars"):
            return
        for key, var in self.change_vars.items():
            var.set("1" if key == "change_no" else "")
        self.change_reason_text.delete("1.0", "end")
        self.change_demand_tree.set_rows([])
        self.change_confirm_tree.set_rows([])
        self.change_budget_tree.set_rows([])
        self.change_summary.delete("1.0", "end")
        self.change_budget_source = ""

    def load_change_to_editor(self, change_no):
        record = self.db.change_record(self.current_project_id, change_no)
        if not record:
            return
        fields = record["fields"]
        for key, var in self.change_vars.items():
            var.set(fields.get(key, record["change_no"] if key == "change_no" else ""))
        self.change_reason_text.delete("1.0", "end")
        self.change_reason_text.insert("1.0", fields.get("reason", ""))
        self.change_demand_tree.set_rows(record["demand"])
        self.change_confirm_tree.set_rows(record["confirm"])
        self.change_budget_tree.set_rows([(list(row) + [""] * 6)[:6] for row in record["budget"]])
        self.change_budget_source = record["source_file"]

    def show_selected_change(self):
        if not self.current_project_id:
            return
        change_no = self.change_select_var.get().strip()
        if not change_no:
            return
        record = self.db.change_record(self.current_project_id, change_no)
        if not record:
            return
        self.load_change_to_editor(change_no)
        fields = record["fields"]
        lines = [
            f"第 {record['change_no']} 次預算變更",
            f"變更金額總計：{fields.get('change_total', '')}",
            f"原工項追加金額：{fields.get('add_original', '')}",
            f"原工項追減金額：{fields.get('deduct_original', '')}",
            f"新增工項金額：{fields.get('add_new', '')}",
            f"變更事由：{fields.get('reason', '')}",
            f"送預算書時間-工務處：{fields.get('send_public_works', '')}",
            f"送預算書時間-主計單位：{fields.get('send_accounting', '')}",
            f"主計單位簽準：{fields.get('accounting_approved', '')}",
            f"工務處簽準：{fields.get('public_works_approved', '')}",
            f"公司簽準：{fields.get('company_approved', '')}",
            f"AA預算書審核完成：{fields.get('aa_review_done', '')}",
        ]
        self.change_summary.delete("1.0", "end")
        self.change_summary.insert("1.0", "\n".join(lines))
        self.change_budget_tree.set_rows([(list(row) + [""] * 6)[:6] for row in record["budget"]])

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
            text="此欄位會即時顯示在上半部工程辦理情形摘要。後續可再擴充狀態日期、狀態說明與歷程紀錄。"
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
            "payment_contract_tree", "payment_other_tree", "payment_admin_tree", "progress_estimate_tree",
            "execution_tree", "milestone_tree", "change_demand_tree", "change_confirm_tree",
            "change_budget_tree"
        ]:
            if hasattr(self, name):
                getattr(self, name).can_edit = self.can_edit
        for tree in getattr(self, "budget_book_trees", {}).values():
            tree.can_edit = self.can_edit

    def apply_edit_lock_state(self):
        unlocked = self.can_edit()
        for w in getattr(self, "edit_widgets", []):
            try:
                w.configure(state="normal" if unlocked else "disabled")
                if isinstance(w, tk.Text):
                    w.configure(foreground=OSX_TEXT)
            except tk.TclError:
                pass
        if unlocked:
            for key in getattr(self, "readonly_basic_keys", set()):
                widget = getattr(self, "basic_widgets", {}).get(key)
                if widget:
                    try:
                        widget.configure(state="readonly")
                    except tk.TclError:
                        pass
        self.update_deposit_performance_manual_state()
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
        if tree == "payment_contract_tree":
            return float(self.latest_contract_cumulative_billing())
        if tree == "payment_other_tree":
            idx = {field: i for i, field in enumerate(PAYMENT_OTHER_FIELDS)}
            latest = self.latest_cumulative_by_period(getattr(self, tree).get_rows(), PAYMENT_OTHER_FIELDS, "cumulative_amount")
            return float(latest)
        if tree == "payment_admin_tree":
            idx = {field: i for i, field in enumerate(PAYMENT_ADMIN_FIELDS)}
            latest = self.latest_cumulative_by_period(getattr(self, tree).get_rows(), PAYMENT_ADMIN_FIELDS, "cumulative_amount")
            return float(latest)
        for row in getattr(self, tree).get_rows():
            if len(row) >= 4:
                total += self.safe_amount(row[3])
        return total

    def latest_cumulative_by_period(self, rows, fields, cumulative_field):
        if not rows:
            return Decimal("0")
        indexes = {field: idx for idx, field in enumerate(fields)}
        prepared = []
        for original_index, raw in enumerate(rows):
            row = (list(raw) + [""] * len(fields))[:len(fields)]
            prepared.append((original_index, row))
        latest_row = sorted(prepared, key=lambda x: period_sort_key(x[1][indexes.get("period_no", 0)], x[0]))[-1][1]
        return self.money_decimal(latest_row[indexes[cumulative_field]])

    def latest_contract_cumulative_billing(self):
        if not hasattr(self, "payment_contract_tree"):
            return Decimal("0")
        return self.latest_cumulative_by_period(self.payment_contract_tree.get_rows(), PAYMENT_CONTRACT_FIELDS, "cumulative_billing_amount")

    def money_text(self, value):
        return self.format_money_value(value)

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


    def update_database_path_display(self):
        if hasattr(self, "database_path_var"):
            self.database_path_var.set(self.database_path)

    def _open_database_for_switch(self, path):
        path = normalize_db_path(path)
        if not path:
            raise RuntimeError("資料庫路徑不可空白")
        folder = os.path.dirname(path)
        if folder:
            os.makedirs(folder, exist_ok=True)
        new_db = DB(path)
        return path, new_db

    def switch_database(self, path, status_text="已切換資料庫"):
        self.save_current()
        old_db = self.db
        old_path = self.database_path
        try:
            new_path, new_db = self._open_database_for_switch(path)
            self.db = new_db
            self.database_path = new_path
            set_default_database_path(new_path)
            self.current_project_id = None
            self.project_password_hash = ""
            self.edit_unlocked = True
            self.dirty = False
            self.undo_history = []
            self.last_state = None
            self.update_database_path_display()
            self.load_projects()
            try:
                old_db.conn.close()
            except Exception:
                pass
            self.status_var.set(f"{status_text}：{new_path}")
        except Exception as exc:
            self.db = old_db
            self.database_path = old_path
            self.update_database_path_display()
            messagebox.showerror("資料庫切換失敗", str(exc))

    def select_database(self):
        path = filedialog.askopenfilename(
            title="選擇資料庫",
            filetypes=[("TR_FxWork 資料庫", ("TRFxWork_db", "TR_FxWork.db", "*.db")), ("SQLite DB", "*.db"), ("所有檔案", "*.*")]
        )
        if not path:
            return
        try:
            with sqlite3.connect(path) as conn:
                conn.execute("SELECT name FROM sqlite_master LIMIT 1").fetchall()
        except Exception as exc:
            messagebox.showerror("選擇資料庫失敗", f"這個檔案不是可讀取的 SQLite 資料庫：\n{exc}")
            return
        self.switch_database(path, "已選擇資料庫，並設為預設來源")

    def new_database(self):
        timestamp = timestamp_suffix()
        path = filedialog.asksaveasfilename(
            title="新增資料庫",
            initialfile=f"{DB_FILE_NAME}_{timestamp}.db",
            defaultextension=".db",
            filetypes=[("SQLite DB", "*.db"), ("所有檔案", "*.*")]
        )
        if not path:
            return
        path = append_timestamp_suffix(path, prefix="_")
        if os.path.exists(path):
            ok = messagebox.askyesno("覆蓋資料庫", "檔案已存在，是否覆蓋並建立新的空白資料庫？")
            if not ok:
                return
            try:
                os.remove(path)
            except OSError as exc:
                messagebox.showerror("新增資料庫失敗", str(exc))
                return
        self.switch_database(path, "已新增資料庫，並設為預設來源")

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
        if hasattr(self, "warranty_note_text"):
            self.warranty_note_text.delete("1.0", "end")
            self.warranty_note_text.insert("1.0", p["warranty_note"] if "warranty_note" in p.keys() and p["warranty_note"] is not None else "")
        self.day_type_var.set(p["day_type"] or "工作日")

        bid_rows = []
        for r in self.db.bids(pid):
            award_day = r["award_date"] if "award_date" in r.keys() else ""
            bid_rows.append(["V" if award_day else "", r["round_no"], r["online_date"], r["open_date"], award_day])
        self.bid_tree.set_rows(bid_rows)
        self.holiday_tree.set_rows([["✓" if r["excluded"] else "", r["day"], r["name"]] for r in self.db.rows("holidays", pid)])
        self.workday_tree.set_rows([["✓" if r["excluded"] else "", r["day"], r["name"]] for r in self.db.rows("workdays", pid)])
        self.apply_year_separators(self.holiday_tree)
        self.apply_year_separators(self.workday_tree)
        self.weather_tree.set_rows([[r["day"], r["morning"], r["afternoon"], r["typhoon"], r["site"], r["note"]] for r in self.db.rows("weather", pid)])
        self.railway_tree.set_rows([["✓" if r["excluded"] else "", r["day"], r["note"]] for r in self.db.rows("railway", pid)])
        self.payment_contract_tree.set_rows(self.payment_contract_db_rows_to_tree(self.db.rows("payment_contract", pid)))
        self.refresh_contract_payment_cumulatives(mark_dirty=False)
        self.payment_other_tree.set_rows(self.payment_other_db_rows_to_tree(self.db.rows("payment_other", pid)))
        self.refresh_payment_other_cumulatives(mark_dirty=False)
        self.payment_admin_tree.set_rows(self.payment_admin_db_rows_to_tree(self.db.rows("payment_admin", pid)))
        self.load_admin_setup_from_project(p)
        self.refresh_payment_admin_cumulatives(mark_dirty=False)
        if hasattr(self, "progress_estimate_tree"):
            self.progress_estimate_tree.set_rows(self.progress_estimate_db_rows_to_tree(self.db.rows("progress_estimates", pid)))
            self.refresh_progress_estimate_amounts(mark_dirty=False)
        self.execution_tree.set_rows([
            [r["day"], r["record_type"], r["subject"] if "subject" in r.keys() else "", r["content"], r["note"]]
            for r in self.db.rows("execution_records", pid)
        ])
        self.milestone_tree.set_rows([
            self.calc_milestone_row([r["item_no"], r["contract_item"], r["start_date"], r["deadline_days"], r["deadline_date"], r["received_date"], r["overdue"], r["received_no"], r["note"], r["day_adjust"]])
            for r in self.db.rows("project_milestones", pid)
        ])
        self.refresh_milestone_rows()
        for area in ("budget", "contract"):
            rows, source = self.db.budget_book(pid, area)
            self.set_budget_book_rows(area, rows, source)
        self.refresh_change_select()
        self.refresh_change_award_select()
        nums = self.db.change_numbers(pid)
        if nums:
            self.change_select_var.set(nums[-1])
            self.load_change_to_editor(nums[-1])
            self.show_selected_change()
        else:
            self.clear_change_editor()
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
        self.undo_history = []
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

    def collect_railway_text_map(self):
        railway_map = {}
        for row in self.railway_tree.get_rows():
            if row and row[0] == "✓":
                continue
            d = parse_date(row[1] if len(row) > 1 else "")
            if not d:
                continue
            note = str(row[2] if len(row) > 2 else "").strip() or "疏運"
            railway_map[d] = note
        return railway_map

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
                note_text = str(row[5] if len(row) > 5 else "").strip()
                rows[d] = {
                    "morning": self.safe_amount(row[1] if len(row) > 1 else 0),
                    "afternoon": self.safe_amount(row[2] if len(row) > 2 else 0),
                    "force_no_work": note_text == "-1",
                }
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
        weather_info = weather_rows.get(d) if isinstance(weather_rows, dict) else None
        if isinstance(weather_info, dict) and weather_info.get("force_no_work"):
            return 0
        base = 1 if self.day_type_var.get() == "日曆天" else (1 if d.weekday() < 5 else 0)
        if d in workday_dates:
            base = 1
        if d in holiday_dates:
            return 0
        if weather_info:
            if isinstance(weather_info, dict):
                morning = self.safe_amount(weather_info.get("morning", 0))
                afternoon = self.safe_amount(weather_info.get("afternoon", 0))
            else:
                morning, afternoon = weather_info
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

    def percent_text(self, numerator, denominator):
        if not denominator:
            return ""
        return f"{round((numerator / denominator) * 100):.0f}%"

    def set_basic_value(self, key, value):
        if key in self.basic_vars and self.basic_vars[key].get() != value:
            self.basic_vars[key].set(value)

    def update_bid_award_state(self):
        if not hasattr(self, "bid_tree"):
            return
        rows = []
        for row in self.bid_tree.get_rows():
            raw = list(row)
            if len(raw) == 3:
                vals = ["", raw[0], raw[1], raw[2], ""]
            elif len(raw) == 4:
                vals = ["", raw[0], raw[1], raw[2], raw[3]]
            else:
                vals = (raw + [""] * 5)[:5]
            vals[0] = "V" if vals[4] else ""
            rows.append(vals)

        def sort_key(row):
            try:
                return int(float(row[1] or 0))
            except ValueError:
                return 0

        rows.sort(key=sort_key, reverse=True)
        self.bid_tree.set_rows(rows)
        award_day = next((row[4] for row in rows if row[4]), "")
        self.set_basic_value("award_date", award_day)

    def recalculate(self):
        if not self.current_project_id:
            return
        self.recalculating = True
        try:
            self.update_bid_award_state()
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

            if not self.basic_vars.get("performance_bond_rate", tk.StringVar()).get().strip():
                self.basic_vars["performance_bond_rate"].set("10%")
            if not self.basic_vars.get("warranty_rate", tk.StringVar()).get().strip():
                self.basic_vars["warranty_rate"].set("3%")

            budget_contract = self.safe_amount(self.basic_vars.get("budget_contract_amount", tk.StringVar()).get())
            budget_air_pollution = self.safe_amount(self.basic_vars.get("budget_air_pollution_fee", tk.StringVar()).get())
            budget_outside_without_air = sum(
                self.safe_amount(self.basic_vars.get(key, tk.StringVar()).get())
                for key in (
                    "budget_mgmt_fee", "budget_self_labor", "budget_self_material", "budget_spare_material",
                    "budget_railway_material", "budget_supervision_fee", "budget_freight", "budget_other"
                )
            )
            # 未完工程：發包工程費預算金額 + 發包以外金額（空汙費除外）。
            budget_unfinished = budget_contract + budget_outside_without_air
            budget_tax = budget_contract * 0.05
            budget_total = budget_contract + budget_tax

            award_net = self.safe_amount(self.basic_vars.get("award_contract_amount", tk.StringVar()).get())
            award_tax = award_net * 0.05
            award_total = award_net + award_tax
            award_air_pollution = self.safe_amount(self.basic_vars.get("award_air_pollution_fee", tk.StringVar()).get())
            award_outside_without_air = sum(
                self.safe_amount(self.basic_vars.get(key, tk.StringVar()).get())
                for key in (
                    "award_mgmt_fee", "award_self_labor", "award_self_material", "award_spare_material",
                    "award_railway_material", "award_supervision_fee", "award_freight", "award_other"
                )
            )
            # 未完工程：發包工程費發包契約金額 + 發包以外金額（空汙費除外）。
            award_unfinished = award_net + award_outside_without_air
            award_base = self.safe_amount(self.basic_vars.get("award_base_price", tk.StringVar()).get())
            perf_rate = self.safe_amount(self.basic_vars.get("performance_bond_rate", tk.StringVar()).get()) / 100
            warranty_rate = self.safe_amount(self.basic_vars.get("warranty_rate", tk.StringVar()).get()) / 100
            manual_performance_deposit = self.basic_vars.get("deposit_performance_manual", tk.StringVar(value="0")).get() == "1"
            performance_deposit_value = (
                self.safe_amount(self.basic_vars.get("deposit_performance", tk.StringVar()).get())
                if manual_performance_deposit
                else award_total * perf_rate
            )
            acceptance_date = parse_date(self.basic_vars.get("actual_acceptance_date", tk.StringVar()).get())
            warranty_years = self.safe_amount(self.basic_vars.get("warranty_years", tk.StringVar()).get())
            warranty_end = add_years(acceptance_date, warranty_years) - timedelta(days=1) if acceptance_date and warranty_years else None

            auto_totals = {
                "budget_contract_tax": budget_tax,
                "budget_contract_total": budget_total,
                "budget_input_tax": budget_tax,
                "budget_unfinished_amount": budget_unfinished,
                "budget_total_amount": budget_unfinished + budget_air_pollution + budget_tax,
                "award_contract_tax": award_tax,
                "award_contract_total": award_total,
                "award_input_tax": award_tax,
                "award_unfinished_amount": award_unfinished,
                "award_total_amount": award_unfinished + award_air_pollution + award_tax,
                "contract_budget_net": budget_contract,
                "contract_budget_tax": budget_tax,
                "contract_budget_total": budget_total,
                "contract_award_net": award_net,
                "contract_award_tax": award_tax,
                "contract_award_total": award_total,
                "labor_budget": self.safe_amount(self.basic_vars.get("budget_labor", tk.StringVar()).get()),
                "labor_award": self.safe_amount(self.basic_vars.get("award_labor", tk.StringVar()).get()),
                "deposit_total": self.safe_amount(self.basic_vars.get("deposit_difference", tk.StringVar()).get()) + performance_deposit_value,
                "warranty_deposit": award_total * warranty_rate,
                "final_contract_amount": award_total,
            }
            if not manual_performance_deposit:
                auto_totals["deposit_performance"] = performance_deposit_value

            for key, value in auto_totals.items():
                if key in self.basic_vars:
                    text = self.money_text(value) if value else ""
                    if self.basic_vars[key].get() != text:
                        self.basic_vars[key].set(text)
            self.update_deposit_performance_manual_state()
            self.refresh_admin_management_fee_source()
            self.refresh_admin_setup_totals()
            self.refresh_progress_estimate_amounts(mark_dirty=False)

            ratio_values = {
                "award_contract_budget_ratio": self.percent_text(award_total, budget_total),
                "award_contract_base_ratio": self.percent_text(award_total, award_base),
                "award_base_budget_ratio": self.percent_text(award_base, budget_total),
                "warranty_end_date": fmt_date(warranty_end),
            }
            for key, text in ratio_values.items():
                self.set_basic_value(key, text)

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
        self.collect_admin_setup_to_project_data(data)
        self.save_current_change_award_fields()
        if hasattr(self, "project_description_text"):
            data["project_description"] = self.project_description_text.get("1.0", "end-1c").strip()
        if hasattr(self, "warranty_note_text"):
            data["warranty_note"] = self.warranty_note_text.get("1.0", "end-1c").strip()
        data["day_type"] = self.day_type_var.get()
        try:
            data["contract_days"] = int(float(data.get("contract_days") or 0))
        except ValueError:
            data["contract_days"] = 0

        self.db.save_project(self.current_project_id, data)

        bids = []
        for r in self.bid_tree.get_rows():
            vals = (r + [""] * 5)[:5]
            bids.append({"round_no": vals[1] or 1, "online_date": vals[2], "open_date": vals[3], "award_date": vals[4]})
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

        self.db.replace_rows("payment_contract", self.current_project_id, self.payment_contract_tree_rows_to_db())
        self.db.replace_rows("payment_other", self.current_project_id, self.payment_other_tree_rows_to_db())
        self.db.replace_rows("payment_admin", self.current_project_id, self.payment_admin_tree_rows_to_db())
        if hasattr(self, "progress_estimate_tree"):
            self.db.replace_rows("progress_estimates", self.current_project_id, self.progress_estimate_tree_rows_to_db())

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
        if hasattr(self, "budget_book_trees"):
            for area, tree in self.budget_book_trees.items():
                source = self.budget_book_sources.get(area).get() if area in self.budget_book_sources else ""
                self.db.save_budget_book(self.current_project_id, area, tree.get_rows(), source)
        if hasattr(self, "change_vars"):
            change_no = self.change_vars["change_no"].get().strip()
            has_change_data = any(v.get().strip() for v in self.change_vars.values())
            has_change_data = has_change_data or bool(self.change_reason_text.get("1.0", "end-1c").strip())
            has_change_data = has_change_data or bool(self.change_demand_tree.get_rows() or self.change_confirm_tree.get_rows() or self.change_budget_tree.get_rows())
            if change_no and has_change_data:
                existing_record = self.db.change_record(self.current_project_id, change_no)
                fields = dict(existing_record.get("fields", {})) if existing_record else {}
                fields.update({k: v.get().strip() for k, v in self.change_vars.items()})
                fields["reason"] = self.change_reason_text.get("1.0", "end-1c")
                if hasattr(self, "change_award_vars") and self.change_award_select_var.get().strip() == change_no:
                    fields.update({k: v.get().strip() for k, v in self.change_award_vars.items()})
                self.db.save_change_record(
                    self.current_project_id,
                    change_no,
                    fields,
                    self.change_demand_tree.get_rows(),
                    self.change_confirm_tree.get_rows(),
                    self.change_budget_tree.get_rows(),
                    getattr(self, "change_budget_source", ""),
                )
                self.refresh_change_select()
                self.refresh_change_award_select()

        self.db.set_setting("last_project_id", self.current_project_id)
        self.dirty = False
        self.last_state = self.capture_state()
        self.status_var.set("已自動儲存：" + datetime.now().strftime("%H:%M:%S"))

    def backup_database(self):
        self.save_current()
        folder = filedialog.askdirectory(title="選擇資料庫備份儲存資料夾")
        if not folder:
            return
        timestamp = timestamp_suffix()
        source_stem = os.path.splitext(os.path.basename(self.database_path))[0] or "TRFxWork_db"
        safe_base = "".join(ch if ch not in r'\/:*?"<>|' else "_" for ch in source_stem)
        out_path = os.path.join(folder, f"{safe_base}BK{timestamp}.zip")
        tmp_db = os.path.join(tempfile.gettempdir(), f"TR_FxWork_backup_{timestamp}.db")
        try:
            with sqlite3.connect(self.database_path) as src, sqlite3.connect(tmp_db) as dst:
                src.backup(dst)
            with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                zf.write(tmp_db, arcname=f"{safe_base}BK{timestamp}.db")
                zf.writestr(
                    "README_備份說明.txt",
                    f"臺鐵監造紀錄小本資料庫備份\n備份時間：{timestamp}\n資料庫來源：{self.database_path}\n備份檔名規則：原資料庫檔名 + BKYYYYMMDDHHMMSS\n"
                )
            messagebox.showinfo("備份完成", f"已完成資料庫備份：\n{out_path}")
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
        timestamp = timestamp_suffix()
        source_stem = os.path.splitext(os.path.basename(self.database_path))[0] or "TRFxWork_db"
        out_path = os.path.join(folder, f"{source_stem}BK{timestamp}.zip")
        tmp_db = os.path.join(tempfile.gettempdir(), f"TR_FxWork_backup_{timestamp}.db")
        try:
            with sqlite3.connect(self.database_path) as src, sqlite3.connect(tmp_db) as dst:
                src.backup(dst)
            with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                zf.write(tmp_db, arcname=DB_FILE_NAME)
                zf.writestr("README_異地備份說明.txt", f"TR_FxWork 異地備份\n備份時間：{timestamp}\n資料庫來源：{self.database_path}\n")
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
            "工程進度估算": self.progress_estimate_tree,
            "發包工程費計價": self.payment_contract_tree,
            "發包以外計價": self.payment_other_tree,
            "管理費計價": self.payment_admin_tree,
            "會議記錄表": self.execution_tree,
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
            "planned_precheck_date": "預定初驗日", "actual_precheck_date": "實際初驗日",
            "planned_acceptance_date": "預定驗收日", "actual_acceptance_date": "實際驗收日",
            "settlement_date": "決算日", "warranty_years": "保固年限", "warranty_end_date": "保固結束日",
            "warranty_note": "保固備註", "performance_bond_rate": "履保金比例",
            "deposit_performance_manual": "履約保證金手動修改", "performance_bond_type": "履保金型式", "warranty_bond_type": "保固金型式",
            "contract_budget_net": "發包工程費-預算(未稅)", "contract_award_net": "發包工程費-決標(未稅)",
            "contract_budget_tax": "發包工程費-稅金(預算)", "contract_award_tax": "發包工程費-稅金(決標)",
            "contract_budget_total": "發包工程費-預算(含稅)", "contract_award_total": "發包工程費-決標(契約金額含稅)",
            "labor_budget": "包工費-預算", "labor_award": "包工費-決標",
            "deposit_difference": "差額保證金", "deposit_performance": "履約保證金", "deposit_total": "保證金總額",
            "final_contract_amount": "竣工發包工程費", "warranty_rate": "保固金比例", "warranty_deposit": "保固保證金",
            "budget_total_amount": "總工程預算-總預算(含稅)", "budget_unfinished_amount": "總工程預算-未完工程",
            "budget_input_tax": "總工程預算-進項稅額", "budget_contract_amount": "預算發包工程費-預算金額",
            "budget_contract_tax": "預算發包工程費-稅金", "budget_contract_total": "預算發包工程費-預算總計(含稅)",
            "budget_labor": "預算包工費-預算", "budget_mgmt_fee": "預算發包以外-工程管理費",
            "budget_self_labor": "預算發包以外-自辦工費", "budget_self_material": "預算發包以外-自購材料費",
            "budget_spare_material": "預算發包以外-路備材料費", "budget_railway_material": "預算發包以外-路購材料費",
            "budget_supervision_fee": "預算發包以外-監理費", "budget_freight": "預算發包以外-運雜費",
            "budget_air_pollution_fee": "預算發包以外-空汙費", "budget_other": "預算發包以外-其他",
            "award_total_amount": "總工程費用-總預算(含稅)", "award_unfinished_amount": "總工程費用-未完工程",
            "award_input_tax": "總工程費用-進項稅額", "award_contract_amount": "決標發包工程費-發包契約金額",
            "award_contract_tax": "決標發包工程費-營業稅", "award_contract_total": "決標發包工程費-契約金額總計(含稅)",
            "award_base_price": "決標發包工程費-底價", "award_contract_budget_ratio": "決標/預算",
            "award_contract_base_ratio": "決標/底價", "award_base_budget_ratio": "底價/預算",
            "award_labor": "決標包工費-發包", "award_mgmt_fee": "決標發包以外-工程管理費",
            "award_self_labor": "決標發包以外-自辦工費", "award_self_material": "決標發包以外-自購材料費",
            "award_spare_material": "決標發包以外-路備材料費", "award_railway_material": "決標發包以外-路購材料費",
            "award_supervision_fee": "決標發包以外-監理費", "award_freight": "決標發包以外-運雜費",
            "award_air_pollution_fee": "決標發包以外-空汙費", "award_other": "決標發包以外-其他",
        }
        for key, label in labels.items():
            if key == "project_description" and hasattr(self, "project_description_text"):
                value = self.project_description_text.get("1.0", "end-1c")
            elif key == "warranty_note" and hasattr(self, "warranty_note_text"):
                value = self.warranty_note_text.get("1.0", "end-1c")
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
        return [list(getattr(tree, "headings", []))] + tree.get_rows()

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

    def open_print_dialog(self):
        self.save_current()
        pages = self.exportable_pages()
        win = tk.Toplevel(self)
        win.title("列印設定")
        win.transient(self)
        win.grab_set()
        win.resizable(False, True)

        box = ttk.Frame(win, padding=12)
        box.pack(fill="both", expand=True)
        ttk.Label(box, text="請勾選要列印的分頁").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 6))
        vars_by_page = {}
        for i, page_name in enumerate(pages):
            var = tk.BooleanVar(value=page_name in ("工程基本資料", "會議記錄表", "工程大事記"))
            vars_by_page[page_name] = var
            ttk.Checkbutton(box, text=page_name, variable=var).grid(row=1 + i // 2, column=(i % 2) * 2, columnspan=2, sticky="w", pady=2)

        option_row = 1 + (len(pages) + 1) // 2
        ttk.Separator(box).grid(row=option_row, column=0, columnspan=4, sticky="ew", pady=8)
        paper_var = tk.StringVar(value="A4")
        orientation_var = tk.StringVar(value="直式")
        duplex_var = tk.StringVar(value="單面")
        margin_var = tk.StringVar(value="12")
        font_size_var = tk.StringVar(value="10")
        include_project_var = tk.BooleanVar(value=True)
        include_time_var = tk.BooleanVar(value=True)

        ttk.Label(box, text="紙張").grid(row=option_row + 1, column=0, sticky="e", padx=4, pady=3)
        ttk.Combobox(box, textvariable=paper_var, values=["A4", "A3", "Letter"], state="readonly", width=10).grid(row=option_row + 1, column=1, sticky="w", pady=3)
        ttk.Label(box, text="方向").grid(row=option_row + 1, column=2, sticky="e", padx=4, pady=3)
        ttk.Combobox(box, textvariable=orientation_var, values=["直式", "橫式"], state="readonly", width=10).grid(row=option_row + 1, column=3, sticky="w", pady=3)
        ttk.Label(box, text="正反面").grid(row=option_row + 2, column=0, sticky="e", padx=4, pady=3)
        ttk.Combobox(box, textvariable=duplex_var, values=["單面", "雙面長邊", "雙面短邊"], state="readonly", width=10).grid(row=option_row + 2, column=1, sticky="w", pady=3)
        ttk.Label(box, text="邊界(mm)").grid(row=option_row + 2, column=2, sticky="e", padx=4, pady=3)
        ttk.Spinbox(box, from_=5, to=30, textvariable=margin_var, width=8).grid(row=option_row + 2, column=3, sticky="w", pady=3)
        ttk.Label(box, text="字級(pt)").grid(row=option_row + 3, column=0, sticky="e", padx=4, pady=3)
        ttk.Spinbox(box, from_=8, to=16, textvariable=font_size_var, width=8).grid(row=option_row + 3, column=1, sticky="w", pady=3)
        ttk.Checkbutton(box, text="列印工程名稱", variable=include_project_var).grid(row=option_row + 3, column=2, sticky="w", pady=3)
        ttk.Checkbutton(box, text="列印產生時間", variable=include_time_var).grid(row=option_row + 3, column=3, sticky="w", pady=3)

        def do_print_export():
            selected = [name for name, var in vars_by_page.items() if var.get()]
            if not selected:
                messagebox.showwarning("尚未選擇", "請至少勾選一個分頁。", parent=win)
                return
            folder = filedialog.askdirectory(title="選擇列印檔儲存資料夾", parent=win)
            if not folder:
                return
            try:
                path = self.write_print_html(
                    folder,
                    selected,
                    paper_var.get(),
                    orientation_var.get(),
                    duplex_var.get(),
                    margin_var.get(),
                    font_size_var.get(),
                    include_project_var.get(),
                    include_time_var.get(),
                )
                messagebox.showinfo("列印檔已產生", f"已產生列印檔：\n{path}\n\n請用瀏覽器開啟後列印。", parent=win)
                win.destroy()
            except Exception as exc:
                messagebox.showerror("列印檔產生失敗", str(exc), parent=win)

        btns = ttk.Frame(box)
        btns.grid(row=option_row + 4, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        ttk.Button(btns, text="產生列印檔", command=do_print_export).pack(side="right", padx=4)
        ttk.Button(btns, text="取消", command=win.destroy).pack(side="right", padx=4)

    def write_print_html(self, folder, page_names, paper, orientation, duplex, margin_mm, font_size, include_project, include_time):
        timestamp = datetime.now().strftime("%Y%m%d%H%M")
        project_name = self.basic_vars.get("name", tk.StringVar()).get().strip() if hasattr(self, "basic_vars") else ""
        safe_project = "".join(ch if ch not in r'\/:*?"<>|' else "_" for ch in (project_name or "TRFxWork"))
        path = os.path.join(folder, f"{safe_project}_print_{timestamp}.html")
        paper_css = html.escape(paper or "A4")
        orientation_css = "landscape" if orientation == "橫式" else "portrait"
        try:
            margin = max(5, min(30, int(float(str(margin_mm).strip() or 12))))
        except ValueError:
            margin = 12
        try:
            size = max(8, min(16, int(float(str(font_size).strip() or 10))))
        except ValueError:
            size = 10
        sections = []
        for page_name in page_names:
            rows = self.basic_page_export_rows() if page_name == "工程基本資料" else self.table_page_export_rows(page_name)
            body_rows = []
            for row in rows:
                if not row:
                    body_rows.append("<tr class=\"blank\"><td colspan=\"20\"></td></tr>")
                    continue
                cells = "".join(f"<td>{html.escape(str(cell))}</td>" for cell in row)
                body_rows.append(f"<tr>{cells}</tr>")
            sections.append(
                "<section class=\"page-section\">"
                f"<h2>{html.escape(page_name)}</h2>"
                "<table>"
                + "\n".join(body_rows)
                + "</table></section>"
            )
        meta = []
        if include_project:
            meta.append(f"工程名稱：{html.escape(project_name)}")
        if include_time:
            meta.append("產生時間：" + datetime.now().strftime("%Y-%m-%d %H:%M"))
        meta.append("列印方式：" + html.escape(duplex or "單面"))
        meta.append("紙張：" + html.escape(paper or "A4") + " " + html.escape(orientation or "直式"))
        content = f"""<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<title>{html.escape(project_name or "TRFxWork")} 列印檔</title>
<style>
@page {{ size: {paper_css} {orientation_css}; margin: {margin}mm; }}
body {{ font-family: "Microsoft JhengHei UI", "PingFang TC", sans-serif; font-size: {size}pt; color: #000; }}
h1 {{ text-align: center; font-size: {size + 6}pt; margin: 0 0 8mm; }}
h2 {{ font-size: {size + 3}pt; margin: 0 0 4mm; }}
.meta {{ border: 1px solid #000; padding: 6px 8px; margin-bottom: 8mm; line-height: 1.6; }}
.page-section {{ page-break-after: always; }}
.page-section:last-child {{ page-break-after: auto; }}
table {{ border-collapse: collapse; width: 100%; table-layout: auto; }}
td, th {{ border: 1px solid #000; padding: 4px 5px; vertical-align: top; white-space: pre-wrap; word-break: break-word; }}
tr.blank td {{ height: 8mm; border: 0; }}
@media print {{ .no-print {{ display: none; }} }}
</style>
</head>
<body>
<div class="no-print meta">瀏覽器列印時請依這裡的設定選擇印表機選項；正反面選項為「{html.escape(duplex or "單面")}」。</div>
<h1>{html.escape(project_name or "臺鐵工程本本")}</h1>
<div class="meta">{'<br>'.join(meta)}</div>
{''.join(sections)}
</body>
</html>
"""
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        return path

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
            "差額保證金": "deposit_difference", "履約保證金": "deposit_performance", "履約保證金手動修改": "deposit_performance_manual",
            "履保金型式": "performance_bond_type", "保固金型式": "warranty_bond_type", "保證金總額": "deposit_total",
            "竣工發包工程費": "final_contract_amount", "保固金比例": "warranty_rate", "保固保證金": "warranty_deposit",
        }
        label_to_key.update({
            "預定初驗日": "planned_precheck_date", "實際初驗日": "actual_precheck_date",
            "預定驗收日": "planned_acceptance_date", "實際驗收日": "actual_acceptance_date",
            "決算日": "settlement_date", "保固年限": "warranty_years", "保固結束日": "warranty_end_date",
            "保固備註": "warranty_note", "履保金比例": "performance_bond_rate",
            "總工程預算-總預算(含稅)": "budget_total_amount", "總工程費用-總預算(含稅)": "award_total_amount",
            "決標發包工程費-契約金額總計(含稅)": "award_contract_total",
            "總工程預算-總預算": "budget_total_amount", "總工程預算-未完工程": "budget_unfinished_amount",
            "總工程預算-進項稅額": "budget_input_tax", "預算發包工程費-預算金額": "budget_contract_amount",
            "預算發包工程費-稅金": "budget_contract_tax", "預算發包工程費-預算總計(含稅)": "budget_contract_total",
            "預算包工費-預算": "budget_labor", "預算發包以外-工程管理費": "budget_mgmt_fee",
            "預算發包以外-自辦工費": "budget_self_labor", "預算發包以外-自購材料費": "budget_self_material",
            "預算發包以外-路備材料費": "budget_spare_material", "預算發包以外-路購材料費": "budget_railway_material",
            "預算發包以外-監理費": "budget_supervision_fee", "預算發包以外-運雜費": "budget_freight",
            "預算發包以外-空汙費": "budget_air_pollution_fee", "預算發包以外-其他": "budget_other",
            "總工程費用-總預算": "award_total_amount", "總工程費用-未完工程": "award_unfinished_amount",
            "總工程費用-進項稅額": "award_input_tax", "決標發包工程費-發包契約金額": "award_contract_amount",
            "決標發包工程費-營業稅": "award_contract_tax", "決標發包工程費-決標金額": "award_contract_total",
            "決標發包工程費-底價": "award_base_price", "決標/預算": "award_contract_budget_ratio",
            "決標/底價": "award_contract_base_ratio", "底價/預算": "award_base_budget_ratio",
            "決標包工費-發包": "award_labor", "決標發包以外-工程管理費": "award_mgmt_fee",
            "決標發包以外-自辦工費": "award_self_labor", "決標發包以外-自購材料費": "award_self_material",
            "決標發包以外-路備材料費": "award_spare_material", "決標發包以外-路購材料費": "award_railway_material",
            "決標發包以外-監理費": "award_supervision_fee", "決標發包以外-運雜費": "award_freight",
            "決標發包以外-空汙費": "award_air_pollution_fee", "決標發包以外-其他": "award_other",
        })
        label_to_key.update({
            "總工程預算-總預算": "budget_total_amount", "總工程費用-總預算": "award_total_amount",
            "決標發包工程費-決標金額": "award_contract_total",
            "預算內容-總預算": "budget_total_amount", "預算內容-未完工程": "budget_unfinished_amount",
            "預算內容-進項稅額": "budget_input_tax",
            "決標內容-總預算": "award_total_amount", "決標內容-未完工程": "award_unfinished_amount",
            "決標內容-進項稅額": "award_input_tax",
        })
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
            elif key == "warranty_note" and hasattr(self, "warranty_note_text"):
                self.warranty_note_text.delete("1.0", "end")
                self.warranty_note_text.insert("1.0", row[1])
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
            if page_name == "工程執行紀錄表":
                page_name = "會議記錄表"
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
            normalized = []
            for r in data_rows:
                if not any(str(v).strip() for v in r):
                    continue
                if page_name == "會議記錄表" and len(r) == 4:
                    normalized.append([r[0], r[1], "", r[2], r[3]])
                else:
                    normalized.append((r + [""] * expected_cols)[:expected_cols])
            pages[page_name].set_rows(normalized)
            if page_name == "工程大事記":
                self.refresh_milestone_rows()
            if page_name == "工程進度估算":
                self.refresh_progress_estimate_amounts(mark_dirty=False)
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
            "payment_contract": ["day", "item", "voucher_no", "amount", "note"] + PAYMENT_CONTRACT_FIELDS,
            "payment_other": ["day", "item", "voucher_no", "amount", "note"],
            "payment_admin": ["day", "item", "voucher_no", "amount", "note"] + PAYMENT_ADMIN_FIELDS,
            "progress_estimates": PROGRESS_ESTIMATE_FIELDS,
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
        try:
            self.db.conn.close()
        except Exception:
            pass
        self.destroy()

    def weather_text_map(self):
        out = {}

        def normalize_weather_field(value):
            # 施工日曆晴雨註記：空白、0、0.0、0.00 不顯示；有內容就直接顯示內容，不加前綴詞。
            text = str(value or "").strip()
            if not text:
                return ""
            try:
                if Decimal(text.replace(",", "")) == 0:
                    return ""
            except (InvalidOperation, ValueError):
                pass
            return text

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

            weather_text = normalize_weather_field(r[3] if len(r) > 3 else "")
            site_text = normalize_weather_field(r[4] if len(r) > 4 else "")
            note_text = normalize_weather_field(r[5] if len(r) > 5 else "")
            raw_note_text = str(r[5] if len(r) > 5 else "").strip()
            if weather_text:
                tags.append(weather_text)
            if site_text:
                tags.append(site_text)
            if raw_note_text == "-1":
                tags.append("不計工期")
            elif note_text:
                tags.append(f"備註:{note_text}")
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

        railway_text_map = self.collect_railway_text_map()
        railway = set(railway_text_map.keys())
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
            "normal_date": "#d9edf7",      # 週一到週五日期底色：粉藍色
            "weekend": "#f4cccc",          # 週六週日日期底色：粉紅色
            "holiday_date": "#fce4d6",     # 遇有假期日期底色：粉橘色
            "holiday": "#fce4d6",          # 假日列提示
            "transport": "#e6e6e6",        # 疏運/雨天
            "weather": "#e6e6e6",
            "white": "#ffffff",
            "grid": "#000000",
            "header": "#ffffff",
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

        row_labels = ["假日", "疏運日", "晴雨註記", "工作日數"]
        work_count = 0
        holiday_ex = self.collect_exclude_dates(True)
        holiday_dates = self.collect_exclude_dates(False)
        start_count_date = parse_date(self.basic_vars["actual_start"].get()) or parse_date(self.basic_vars["planned_start"].get())
        finish_count_date = (
            parse_date(self.basic_vars["planned_finish_transport"].get())
            or parse_date(self.basic_vars["planned_finish_holiday"].get())
            or parse_date(self.basic_vars["actual_finish"].get())
        )
        try:
            target_work_days = float(self.basic_vars["contract_days"].get() or 0)
        except ValueError:
            target_work_days = 0.0

        def day_increment(d):
            return self.daily_work_increment(d, holiday_dates, railway, workdays, weather_rows)

        if start_count_date and target_work_days > 0:
            cur = start_count_date
            total = 0.0
            target_finish_date = None
            for _ in range(20000):
                total += day_increment(cur)
                if total >= target_work_days:
                    target_finish_date = cur
                    break
                cur += timedelta(days=1)
            if target_finish_date and (not finish_count_date or target_finish_date > finish_count_date):
                finish_count_date = target_finish_date

        if start_count_date:
            cur = start_count_date
            first_of_month = date(y, m, 1)
            while cur < first_of_month and (not finish_count_date or cur <= finish_count_date):
                work_count += day_increment(cur)
                cur += timedelta(days=1)

        for wi, week in enumerate(weeks):
            y0 = top_h + weekday_h + wi * week_h
            for ri, label in enumerate(row_labels):
                label_y = y0 + (ri + 1) * row_h
                c.create_rectangle(5, label_y, left_w, label_y + row_h, fill="#f2f2f2", outline=colors["grid"])
                c.create_text(left_w - 6, label_y + row_h/2, text=label, anchor="e", font=("Microsoft JhengHei UI", label_font_size))

            for di, d in enumerate(week):
                x0 = left_w + di * cell_w
                in_month = d.month == m
                alpha_fill = colors["weekend"] if d.weekday() >= 5 else colors["normal_date"]
                if in_month and d in holidays:
                    alpha_fill = colors["holiday_date"]
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
                rtxt = railway_text_map.get(d, "") if in_month else ""
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
                    display_count = min(work_count, target_work_days) if target_work_days > 0 else work_count
                    txt = f"{display_count:g}" if display_count else ""
                else:
                    txt = ""
                c.create_rectangle(x0, y0+4*row_h, x0+cell_w, y0+5*row_h, fill="#ffffff", outline=colors["grid"])
                c.create_text(x0+cell_w/2, y0+row_h*4.5, text=txt, font=("Microsoft JhengHei UI", main_font_size, "bold"))

        c.create_text(
            10, note_y,
            text="說明：週一到週五日期粉藍色；週六週日日期粉紅色；遇假期日期粉橘色；資料關閉前與編輯中會自動儲存。",
            anchor="sw",
            font=("Microsoft JhengHei UI", note_font_size)
        )


if __name__ == "__main__":
    app = App()
    app.mainloop()
