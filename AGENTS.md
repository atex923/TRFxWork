# AGENTS.md

## 專案名稱

- 中文：臺鐵監造紀錄小本
- 英文：TR_FxWork
- 目前版本：V0.3.4.3
- 本版摘要：列印設定強化，會議記錄表同步工程大事記，並調整工程名稱與大事記排序。

## 版號規則

第三碼使用數字累積編；每次修改程式後，第三碼加 1。
本檔只保留目前最新版號，完整改版摘要集中在 `HISTORY.md`。

- 目前最新版：V0.3.4.3

## 歷史區規則

- 工作區只保留最新版號程式檔，例如 `TR_FxWork_V0_3_4_3.py` 與 `TR_FxWork_V0_3_4_3.pyw`。
- 舊版號程式檔集中存放在 `history/old_versions/`，並由 Git commit 與 tag 保存。
- 改版前先建立 Git tag 保存舊版，例如 `v0.1.7`。
- README 主頁、啟動檔、主程式常數永遠指向最新版。
- 每次改版都要在主程式 `APP_RELEASE_SUMMARY` 與 `HISTORY.md` 新增摘要。
- 第一碼或第二碼進版時，將上一個大版/小版的第三碼摘要濃縮成大版摘要。

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
- 會議記錄表：時間、類型、標題、內容、備註。
- 工程大事記：日期排序、自動項次與履約期限追蹤。
- 密碼規則：SHA256(PASSWORD_SALT + password)，PASSWORD_SALT = `1981`；每個工程獨立儲存 `password_hash`。
- 後續需補計價欄位、統計規則、匯出 Excel/PDF、工程執行狀態歷程。
