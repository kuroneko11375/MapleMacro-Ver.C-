//! WH_KEYBOARD_LL Hook - split filtering by key type.
//! Zero allocation, zero locks, atomic-only state reads.
//!
//! Strategy (when game is in background):
//!   Extended keys (arrows, Ins, Del, etc.):
//!     PASS through (CallNextHookEx) - game gets WM_KEYDOWN via ATT+SetFocus.
//!   Regular keys (X, Z, A-Z, 0-9, etc.):
//!     BLOCK (LRESULT(1)) - prevents foreground typing.
//!     SendInput already updated GetAsyncKeyState BEFORE hook fires.
//!     Game reads key state via GetAsyncKeyState polling.

use std::ffi::c_void;
use std::sync::atomic::Ordering;

use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    CallNextHookEx, GetForegroundWindow, SetWindowsHookExW, UnhookWindowsHookEx,
    HHOOK, KBDLLHOOKSTRUCT,
    WH_KEYBOARD_LL, WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP,
};

use crate::{
    BLOCKED_COUNT, BLOCKING_ENABLED, HOOK_HANDLE, MACRO_KEY_MARKER,
    PASSED_COUNT, TARGET_HWND,
};

pub fn install_hook() -> bool {
    let current = HOOK_HANDLE.load(Ordering::SeqCst);
    if current != 0 {
        return true;
    }

    let hook = unsafe {
        SetWindowsHookExW(
            WH_KEYBOARD_LL,
            Some(low_level_keyboard_proc),
            GetModuleHandleW(None).unwrap_or_default(),
            0,
        )
    };

    match hook {
        Ok(h) => {
            HOOK_HANDLE.store(h.0 as isize, Ordering::SeqCst);
            true
        }
        Err(_) => false,
    }
}

pub fn uninstall_hook() {
    let handle = HOOK_HANDLE.swap(0, Ordering::SeqCst);
    if handle != 0 {
        unsafe {
            let _ = UnhookWindowsHookEx(HHOOK(handle as *mut c_void));
        }
    }
}

/// The core hook callback. Must complete in <100us.
/// Direct pointer deref instead of marshalling. No heap allocation.
unsafe extern "system" fn low_level_keyboard_proc(
    code: i32,
    w_param: WPARAM,
    l_param: LPARAM,
) -> LRESULT {
    if code < 0 || !BLOCKING_ENABLED.load(Ordering::Relaxed) {
        return call_next(code, w_param, l_param);
    }

    let msg = w_param.0 as u32;
    if msg != WM_KEYDOWN && msg != WM_KEYUP && msg != WM_SYSKEYDOWN && msg != WM_SYSKEYUP {
        return call_next(code, w_param, l_param);
    }

    // Direct pointer deref - 100x faster than Marshal.PtrToStructure
    let kbd = &*(l_param.0 as *const KBDLLHOOKSTRUCT);

    // Detection: dwExtraInfo marker (our macro key)
    if kbd.dwExtraInfo == MACRO_KEY_MARKER {
        let target = HWND(TARGET_HWND.load(Ordering::Relaxed) as *mut c_void);

        // If game is foreground, pass everything through
        if !target.0.is_null() {
            let foreground = GetForegroundWindow();
            if foreground == target {
                PASSED_COUNT.fetch_add(1, Ordering::Relaxed);
                return call_next(code, w_param, l_param);
            }
        }

        // Game is background - split by key type:
        if is_extended_key(kbd.vkCode) {
            // Extended keys (arrows, navigation):
            // PASS - WM_KEYDOWN goes to game via ATT+SetFocus.
            PASSED_COUNT.fetch_add(1, Ordering::Relaxed);
            return call_next(code, w_param, l_param);
        } else {
            // Regular keys (X, Z, A-Z, 0-9, etc.):
            // BLOCK - prevents typing in Chrome.
            // Game reads via GetAsyncKeyState (already updated by SendInput).
            BLOCKED_COUNT.fetch_add(1, Ordering::Relaxed);
            return LRESULT(1);
        }
    }

    // No MACRO_KEY_MARKER => not our key, always pass through.
    // Avoids false positives from other injected software (IME, drivers, etc.)
    call_next(code, w_param, l_param)
}

/// Extended keys pass through, regular keys get blocked.
#[inline(always)]
fn is_extended_key(vk: u32) -> bool {
    // 0x21..=0x2E: PgUp, PgDn, End, Home, Left, Up, Right, Down, Ins, Del, etc.
    (vk >= 0x21 && vk <= 0x2E)
        || vk == 0xA1  // VK_RSHIFT
        || vk == 0xA3  // VK_RCONTROL
        || vk == 0xA5  // VK_RMENU
        || vk == 0x90  // VK_NUMLOCK
}

#[inline(always)]
unsafe fn call_next(code: i32, w_param: WPARAM, l_param: LPARAM) -> LRESULT {
    let handle = HOOK_HANDLE.load(Ordering::Relaxed);
    CallNextHookEx(HHOOK(handle as *mut c_void), code, w_param, l_param)
}
