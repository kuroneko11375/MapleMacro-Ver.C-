# 楓之谷自動化腳本 (C# WinForms)

**作者：** SchwarzeKatze_R ｜ **版本：** 1.1.1 ｜ **語言：** C# (.NET 8, WinForms)

## 📖 專案簡介

基於 C# 的楓之谷自動化輔助工具，支援錄製與播放鍵盤動作。透過 Windows API 實現底層按鍵注入，支援前景與背景兩種模式。

<img width="678" height="628" alt="image" src="https://github.com/user-attachments/assets/994d7641-b3c9-466a-a4cc-1b068e25d4b0" />

## ⚡ 功能特色

| 功能 | 說明 |
|------|------|
| **錄製 / 播放** | 全域鍵盤鉤子，精確擷取按鍵與時間間隔 |
| **前景模式** | `SendInput` 模擬物理按鍵 |
| **背景模式** | `PostMessage` + `AttachThreadInput` 發送至指定視窗 |
| **自定義按鍵 (15 格)** | 按間隔自動觸發，支援延遲 / 暫停腳本 |
| **全局熱鍵** | F9 播放 / F10 停止（可自訂） |
| **定時執行** | 設定時間自動播放，即時倒數 |
| **設定管理** | 自動儲存載入，支援匯入匯出 |

### 方向鍵發送模式

| 模式 | 說明 |
|------|------|
| **Rust FF (推薦)** | Rust DLL Flash Focus，混合 SendInput + PostMessage |
| **S2C** | ThreadAttach + PostMessage，純背景 |
| **TAB** | ThreadAttach + Blocker |
| **SWB** | SendInput + Blocker |

## 🛠️ 環境需求

* Windows 10 / 11
* [.NET 8 Runtime](https://dotnet.microsoft.com/download/dotnet/8.0)
* **必須以系統管理員身分執行**

## ▶️ 快速開始

1. 右鍵以 **管理員身分** 啟動程式
2. 點擊「鎖定視窗」選取遊戲視窗
3. 「開始錄製」→ 操作鍵盤 →「停止錄製」
4. 設定循環次數，按 **F9** 開始播放、**F10** 停止

## 📅 更新紀錄

| 日期 | 版本 | 重點 |
|------|------|------|
| 2025/02/08 | v1.1.1 | 新增 Rust DLL 引擎，背景按鍵混合策略 |
| 2025/02/06 | v1.1.0 | 自定義按鍵系統 (15 格)，改用 SendInput |
| 2025/02/04 | v1.0.4 | 改善按鍵釋放邏輯 |
| 2025/01/30 | v1.0.3 | 全局熱鍵、UI 更新 |
| 2025/01/20 | v1.0.0 | 正式轉換至 C# WinForms |

## ⚠️ 重要聲明

本工具僅供 **技術研究與個人學習** 之用。使用自動化腳本可能違反遊戲服務條款，使用者需自負一切後果。

## 📄 授權

[MIT License](https://github.com/kuroneko11375/MapleMacro-Ver.C-/blob/master/LICENSE.txt)



