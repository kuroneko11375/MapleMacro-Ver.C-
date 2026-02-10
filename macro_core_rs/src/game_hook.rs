//! WH_KEYBOARD thread-level hook injected into the game's thread.
//!
//! Strategy:
//!   SetWindowsHookEx(WH_KEYBOARD, callback, hmodule, gameThreadId)
//!   injects this DLL into the game process. The callback runs in the
//!   game's thread context, so Get/SetKeyboardState affects ONLY the
//!   game's thread-local keyboard state table.
//!
//! When PostMessage delivers WM_KEYDOWN/WM_KEYUP for arrow keys,
//! the game calls GetMessage/PeekMessage which triggers this hook
//! BEFORE WndProc. We update the keyboard state so GetKeyState
//! returns the correct pressed/released state for arrow keys.
//!
//! This is completely invisible to foreground apps - no ATT, no
//! SendInput, no SetForegroundWindow.

use std::ffi::c_void;
use std::sync::atomic::{AtomicIsize, Ordering};

use windows::Win32::Foundation::{LPARAM, LRESULT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::{GetKeyboardState, SetKeyboardState};
use windows::Win32::UI::WindowsAndMessaging::{
    CallNextHookEx, SetWindowsHookExW, UnhookWindowsHookEx,
    HHOOK, WH_KEYBOARD,
};
use windows::core::PCWSTR;

use crate::types;

/// Handle for the game thread hook (WH_KEYBOARD, NOT WH_KEYBOARD_LL).
static GAME_HOOK_HANDLE: AtomicIsize = AtomicIsize::new(0);

/// Install WH_KEYBOARD hook on a specific game thread.
///
/// Unlike WH_KEYBOARD_LL (system-wide), WH_KEYBOARD is per-thread and
/// requires a DLL module handle. Windows loads the DLL into the target
/// process and calls the hook proc in that thread's context.
pub fn install_game_hook(thread_id: u32) -> bool {
    let current = GAME_HOOK_HANDLE.load(Ordering::SeqCst);
    if current != 0 {
        // Already installed - uninstall first
        uninstall_game_hook();
    }

    // Get our DLL's module handle
    let dll_name: Vec<u16> = "macro_core.dll\0".encode_utf16().collect();
    let hmodule = unsafe { GetModuleHandleW(PCWSTR(dll_name.as_ptr())) };

    let hmodule = match hmodule {
        Ok(h) => h,
        Err(_) => return false,
    };

    // Install WH_KEYBOARD hook on the game's thread
    // This is NOT WH_KEYBOARD_LL (system-wide).
    // WH_KEYBOARD fires per-thread when GetMessage/PeekMessage retrieves a keyboard message.
    let hook = unsafe {
        SetWindowsHookExW(
            WH_KEYBOARD,
            Some(game_keyboard_proc),
            hmodule,
            thread_id,
        )
    };

    match hook {
        Ok(h) => {
            GAME_HOOK_HANDLE.store(h.0 as isize, Ordering::SeqCst);
            true
        }
        Err(_) => false,
    }
}

/// Uninstall the game thread hook.
pub fn uninstall_game_hook() {
    let handle = GAME_HOOK_HANDLE.swap(0, Ordering::SeqCst);
    if handle != 0 {
        unsafe {
            let _ = UnhookWindowsHookEx(HHOOK(handle as *mut c_void));
        }
    }
}

/// WH_KEYBOARD callback - runs in the GAME'S thread context.
///
/// This fires when the game calls GetMessage/PeekMessage and retrieves a
/// keyboard message (WM_KEYDOWN/WM_KEYUP from our PostMessage).
///
/// For arrow keys, we update the keyboard state table so GetKeyState
/// returns the correct state when the game polls it.
///
/// Parameters (WH_KEYBOARD):
///   wParam: virtual key code
///   lParam: bit 31 = transition state (0=pressed, 1=released)
///           other bits = repeat count, scan code, etc.
unsafe extern "system" fn game_keyboard_proc(
    code: i32,
    w_param: WPARAM,
    l_param: LPARAM,
) -> LRESULT {
    // HC_ACTION = 0: the message is a real keyboard message to process
    if code >= 0 {
        let vk = w_param.0 as u16;

        // Only intercept arrow keys - they need GetKeyState updating
        // Other keys (ZXC etc.) work fine with PostMessage alone
        if types::is_arrow_key(vk) {
            // lParam bit 31: transition state
            //   0 = key is being pressed (WM_KEYDOWN)
            //   1 = key is being released (WM_KEYUP)
            let is_key_down = (l_param.0 as u32 & (1 << 31)) == 0;

            // Get current keyboard state for this thread (game's thread)
            let mut key_state = [0u8; 256];
            let _ = GetKeyboardState(&mut key_state);

            // Update the arrow key's state
            // High bit (0x80) = key is currently pressed
            if is_key_down {
                key_state[vk as usize] |= 0x80;
            } else {
                key_state[vk as usize] &= !0x80;
            }

            // Apply the updated state - GetKeyState will now return correct value
            let _ = SetKeyboardState(&key_state);
        }
    }

    // Always call next hook in chain
    CallNextHookEx(HHOOK::default(), code, w_param, l_param)
}
