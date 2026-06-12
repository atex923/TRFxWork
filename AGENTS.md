# AGENTS.md

## 專案名稱

- 中文：臺鐵監造紀錄小本
- 英文：TR_FxWork
- 目前版本：V0.1.q

## 版號規則

第三碼使用累積編：

- V0.1.a
- V0.1.b
- V0.1.c
- V0.1.d
- V0.1.e
- V0.1.f
- V0.1.g
- V0.1.h
- V0.1.i
- V0.1.j
- V0.1.k
- V0.1.l
- V0.1.m
- V0.1.n
- V0.1.o
- V0.1.p
- V0.1.q

## 改版同步規則

- 每次產出新版 `.py` 與 `.pyw` 後，需同步複製到 Google 雲端硬碟備份資料夾。
- 目標帳號：`atex.lin@gmail.com`
- 目標資料夾：`/Users/atex1/Library/CloudStorage/GoogleDrive-atex.lin@gmail.com/我的雲端硬碟/12.Codex/`
- 需複製當次新版檔名，例如 `TR_FxWork_V0_1_q.py` 與 `TR_FxWork_V0_1_q.pyw`。
- 若沙盒限制無法寫入，需要求授權後再複製。

## 技術限制

- 優先使用 Python 標準函式庫。
- 目前 GUI 使用 tkinter。
- 資料庫使用 SQLite。
- 優先維持 `.py` 與 `.pyw` 可直接執行。

## 開發方向

- 第六分頁：發包工程費計價。
- 第七分頁：發包以外計價。
- 第八分頁：管理費計價。
- 第九分頁：工程執行紀錄表。
- 第十分頁：執行狀態。
- 第十一分頁：設定，管理每個工程的編輯密碼。
- 密碼規則：SHA256(PASSWORD_SALT + password)，PASSWORD_SALT = `1981`；每個工程獨立儲存 `password_hash`。
- 後續需補計價欄位、統計規則、匯出 Excel/PDF、工程執行狀態歷程。
