# ANTIGRAVITY.md - Outlook AOI Mail Sync

## 🚀 專案概述
本專案為 Outlook 寄件備份中「圖檢異常」信件的自動化歸檔與清單整理系統，支援自動剖析內文 Issue 與會議決議、儲存實體 `.msg`、維護 Excel 超連結清單，並透過 Windows 每日排程達成增量同步。

## 📁 檔案結構
- `Auto_Sync_Mail.ps1`：核心自動化增量同步腳本
- `圖檢異常_寄件備份/`：存出的 `.msg` 實體郵件資料夾
- `圖檢異常_寄件備份清單.xlsx`：主要 Excel 5 欄位清單總表
- `圖檢異常_寄件備份清單.csv`：文字版 CSV 清單
- `sync_log.txt`：排程執行日誌
