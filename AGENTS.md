# 專案代理人工作規則 (AGENTS.md) - Outlook AOI Mail Sync

## 📌 專案基本設定
- **專案名稱**：Outlook 圖檢異常郵件自動備份與清單同步
- **實體路徑**：`D:\Antigravity Project\Outlook圖檢異常mail自動彙整`
- **Obsidian 專案駕駛艙**：`C:\Users\eric.ho\Documents\Secondbrain\知識庫\Outlook圖檢異常郵件自動備份專案.md`
- **GitHub 儲存庫**：https://github.com/ericho-0213/outlook-mail-backup (Public)
- **對接信箱**：預設日常信箱（`archive.pst` / 預設寄件備份）與歷史備份（`2026.pst`）
- **語言規範**：繁體中文（Taiwan）

## ⚙️ 核心運作規則
1. **清單維護**：
   - 清單總表為 `圖檢異常_寄件備份清單.xlsx` 與 `圖檢異常_寄件備份清單.csv`。
   - 欄位限定 5 個：`寄件時間`、`主旨`、`[Issue]`、`[會議決議]`、`開啟郵件`。
   - `[Issue]` 與 `[會議決議]` 格式必須為自動換行 (WrapText)、文字靠上對齊。
   - `開啟郵件` 為本機實體 `.msg` 原生超連結，路徑分隔符號使用 `/`。
2. **自動排程**：
   - 透過 Windows 工作排程器 `Outlook_AOI_Mail_AutoSync` 於每日 16:00:00 執行 `Auto_Sync_Mail.ps1`。
3. **編碼安全**：
   - PowerShell 腳本內避免無 BOM 中文字串常數，一律使用 `Get-UStr` 或系統屬性帶入。

## 🔄 開工與收工規範
- **開工**：讀取本檔、檢查 Obsidian 駕駛艙與 `sync_log.txt`。
- **收工**：更新 Obsidian 駕駛艙與每日筆記，確保排程狀態正常。
