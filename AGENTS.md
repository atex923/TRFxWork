# AGENTS.md

## 專案名稱

- 中文：臺鐵監造紀錄小本
- 英文：TR_FxWork
- 目前版本：V0.1.6

## 版號規則

第三碼使用數字累積編；每次修改程式後，第三碼加 1。

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
- V0.1.1
- V0.1.2
- V0.1.3
- V0.1.4
- V0.1.6

## 技術限制

- 優先使用 Python 標準函式庫。
- 目前 GUI 使用 tkinter。
- 資料庫使用 SQLite。
- 優先維持 `.py` 與 `.pyw` 可直接執行。
- 不要每次修改都重新轉 exe；只有使用者明確要求轉 exe 時才執行 Nuitka。
- Nuitka 轉 exe 時使用 `assets/icon_result.ico` 作為圖示，詳細指令記在 `BUILD_NOTES.md`。

## 開發方向

- 第六分頁：發包工程費計價。
- 第七分頁：發包以外計價。
- 第八分頁：管理費計價。
- 第九分頁：工程執行紀錄表。
- 第十分頁：執行狀態。
- 第十一分頁：設定，管理每個工程的編輯密碼。
- 密碼規則：SHA256(PASSWORD_SALT + password)，PASSWORD_SALT = `1981`；每個工程獨立儲存 `password_hash`。
- 後續需補計價欄位、統計規則、匯出 Excel/PDF、工程執行狀態歷程。
