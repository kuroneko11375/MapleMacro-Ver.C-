# MapleHook DLL 編譯說明

這個 DLL 用於 Hook Windows API，讓 V113 楓之谷在背景時仍然可以接收按鍵輸入。

## 原理

V113 版本的楓之谷會在主迴圈中檢查 `GetForegroundWindow()` 是否等於自己的視窗句柄。
如果不是，就會忽略移動鍵（方向鍵）的輸入。

這個 DLL 透過 IAT Hook 技術，攔截遊戲對以下 API 的呼叫：
- `GetForegroundWindow` - 永遠回傳遊戲視窗句柄
- `GetAsyncKeyState` - 回傳模擬的按鍵狀態
- `GetKeyboardState` - 回傳模擬的鍵盤狀態

## 編譯方法

### 方法一：使用 Visual Studio

1. 開啟 Visual Studio
2. 建立新專案 → **Dynamic-Link Library (DLL)**
3. 將 `dllmain.cpp` 的內容複製到專案中
4. 專案設定：
   - 平台：**x86** (32位元，配合 V113)
   - 設定類型：**動態程式庫 (.dll)**
   - 字元集：**使用 Unicode 字元集**
   - 執行階段程式庫：**多執行緒 DLL (/MD)** 或 **多執行緒 (/MT)**
5. 建置 → 建置解決方案
6. 將產生的 `MapleHook.dll` 複製到 Macro 程式的目錄中

### 方法二：使用 MinGW

```bash
g++ -shared -o MapleHook.dll dllmain.cpp -m32 -static -luser32 -lkernel32
```

### 方法三：使用 cl.exe (Visual Studio 命令提示字元)

```bash
cl /LD /O2 dllmain.cpp /link user32.lib kernel32.lib /OUT:MapleHook.dll
```

## 使用方式

1. 將編譯好的 `MapleHook.dll` 放到 Macro 程式執行檔的同一目錄
2. 在 Macro 程式中選擇「Hook模式」
3. 開始播放時，程式會自動將 DLL 注入到遊戲進程

## 注意事項

- ?? **僅供私服使用**：這個 DLL 會修改遊戲的記憶體，請勿在官方伺服器使用
- ?? **需要管理員權限**：DLL 注入需要管理員權限才能正常運作
- 如果注入失敗，程式會自動回退到純 PostMessage 模式

## 故障排除

### DLL 注入失敗
- 確保以管理員身分執行 Macro 程式
- 確保 `MapleHook.dll` 存在於程式目錄中
- 確保遊戲視窗已正確選取

### 方向鍵仍然無法操作
- 確認已選擇「Hook模式」
- 確認日誌顯示「Hook DLL 注入成功」
- 嘗試重新啟動遊戲後再注入

## 進階：使用 MinHook 庫

如果 IAT Hook 不夠穩定，可以考慮使用 [MinHook](https://github.com/TsudaKageworthy/minhook) 庫進行更穩定的 Inline Hook。

```cpp
#include "MinHook.h"

// 在 DllMain 中：
MH_Initialize();
MH_CreateHook(&GetForegroundWindow, &HookedGetForegroundWindow, (LPVOID*)&pOriginalGetForegroundWindow);
MH_EnableHook(&GetForegroundWindow);
```
