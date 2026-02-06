// MapleHook.h - 頭文件
#pragma once

#ifdef MAPLEHOOK_EXPORTS
#define MAPLEHOOK_API __declspec(dllexport)
#else
#define MAPLEHOOK_API __declspec(dllimport)
#endif

extern "C" {
    // 設定遊戲視窗句柄
    MAPLEHOOK_API void SetGameHWND(HWND hwnd);

    // 設定按鍵狀態 (vKey: 虛擬鍵碼, pressed: true=按下, false=放開)
    MAPLEHOOK_API void SetKeyState(int vKey, bool pressed);

    // 清除所有按鍵狀態
    MAPLEHOOK_API void ClearAllKeyStates();

    // 檢查 Hook 是否已啟動
    MAPLEHOOK_API bool IsHookActive();
}
