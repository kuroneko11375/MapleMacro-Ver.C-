//! Background Key Sender
//!
//! Strategy:
//!   - Foreground: Direct SendInput (hook passes through since game is foreground)
//!   - Background (extended keys / arrows): ATT + SetFocus + SendInput with MACRO_KEY_MARKER
//!     Hook passes these through. Game receives WM_KEYDOWN via focused window.
//!   - Background (regular keys / alphanumeric): PostMessage WM_KEYDOWN/WM_KEYUP
//!     Sends directly to game window message queue. Does NOT trigger LL hooks.
//!     Does NOT affect foreground window at all.

use windows::Win32::Foundation::{HWND, WPARAM, LPARAM};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    MapVirtualKeyW, SendInput, SetFocus, INPUT, INPUT_KEYBOARD, KEYBDINPUT,
    KEYBD_EVENT_FLAGS, KEYEVENTF_EXTENDEDKEY, KEYEVENTF_KEYUP,
    MAP_VIRTUAL_KEY_TYPE, VIRTUAL_KEY,
};
use windows::Win32::UI::WindowsAndMessaging::{
    GetForegroundWindow, GetWindowThreadProcessId,
    PostMessageW, WM_KEYDOWN, WM_KEYUP,
};
use windows::Win32::System::Threading::{AttachThreadInput, GetCurrentThreadId};

use crate::{types, MACRO_KEY_MARKER};

/// Send key to game (background-safe).
///
/// When game is foreground: direct SendInput (hook passes through).
/// When game is background:
///   - Extended keys (arrows, nav): ATT + SetFocus + SendInput
///     Hook passes these through, game receives WM_KEYDOWN.
///   - Regular keys (A-Z, 0-9, etc.): PostMessage WM_KEYDOWN/WM_KEYUP
///     Directly to game message queue, no LL hook, no foreground leak.
///
/// Returns: 0=success, -3=SendInput failed, -4=PostMessage failed
pub fn flash_focus_send_key(game_hwnd: HWND, vk_code: u16, is_key_down: bool) -> i32 {
    unsafe {
        let foreground = GetForegroundWindow();
        if foreground == game_hwnd {
            // Game is foreground - direct SendInput, hook will pass through
            return send_input_raw(vk_code, is_key_down, true);
        }

        // Game is background - split by key type
        if types::is_extended_key(vk_code) {
            // Extended keys (arrows, nav): ATT + SetFocus + SendInput
            // Hook passes these through, game receives via WM_KEYDOWN
            let our_thread = GetCurrentThreadId();
            let game_thread = GetWindowThreadProcessId(game_hwnd, None);
            let attached = AttachThreadInput(our_thread, game_thread, true).as_bool();

            if attached {
                let _ = SetFocus(game_hwnd);
            }

            let result = send_input_raw(vk_code, is_key_down, true);

            if attached {
                let _ = AttachThreadInput(our_thread, game_thread, false);
            }

            result
        } else {
            // Regular keys (A-Z, 0-9, etc.): PostMessage directly to game
            // Does NOT trigger LL hooks. Does NOT affect foreground window.
            post_message_key(game_hwnd, vk_code, is_key_down)
        }
    }
}

/// Send regular key via PostMessage (background-safe, no foreground leak).
///
/// PostMessage sends WM_KEYDOWN/WM_KEYUP directly to target window's message queue.
/// Does NOT trigger LL hooks. Does NOT affect foreground window at all.
/// Suitable for alphanumeric keys (A-Z, 0-9, etc.) that the game reads via WM_KEYDOWN.
///
/// Returns: 0=success, -4=PostMessage failed
fn post_message_key(game_hwnd: HWND, vk_code: u16, is_key_down: bool) -> i32 {
    let scan_code = unsafe {
        MapVirtualKeyW(vk_code as u32, MAP_VIRTUAL_KEY_TYPE(0)) as u32
    };

    let msg = if is_key_down { WM_KEYDOWN } else { WM_KEYUP };

    // Construct lParam:
    // Bits 0-15:  repeat count (1)
    // Bits 16-23: scan code
    // Bit 24:     extended key flag (0 for regular keys)
    // Bit 30:     previous key state (1 for key up)
    // Bit 31:     transition state (0=pressed, 1=released)
    let lparam_val: isize = if is_key_down {
        (1i64 | ((scan_code as i64) << 16)) as isize
    } else {
        (1i64 | ((scan_code as i64) << 16) | (1i64 << 30) | (1i64 << 31)) as isize
    };

    let result = unsafe {
        PostMessageW(
            game_hwnd,
            msg,
            WPARAM(vk_code as usize),
            LPARAM(lparam_val),
        )
    };

    match result {
        Ok(_) => 0,
        Err(_) => -4,
    }
}

/// Foreground SendInput with MACRO_TAG
pub fn send_input_foreground(vk_code: u16, is_key_down: bool) -> i32 {
    send_input_raw(vk_code, is_key_down, true)
}

/// Low-level SendInput wrapper
fn send_input_raw(vk_code: u16, is_key_down: bool, with_marker: bool) -> i32 {
    let scan_code = unsafe {
        MapVirtualKeyW(vk_code as u32, MAP_VIRTUAL_KEY_TYPE(0)) as u16
    };

    let mut flags = KEYBD_EVENT_FLAGS::default();
    if !is_key_down {
        flags |= KEYEVENTF_KEYUP;
    }
    if types::is_extended_key(vk_code) {
        flags |= KEYEVENTF_EXTENDEDKEY;
    }

    let extra_info = if with_marker { MACRO_KEY_MARKER } else { 0 };

    let input = INPUT {
        r#type: INPUT_KEYBOARD,
        Anonymous: windows::Win32::UI::Input::KeyboardAndMouse::INPUT_0 {
            ki: KEYBDINPUT {
                wVk: VIRTUAL_KEY(vk_code),
                wScan: scan_code,
                dwFlags: flags,
                time: 0,
                dwExtraInfo: extra_info,
            },
        },
    };

    let inputs = [input];
    let result = unsafe { SendInput(&inputs, std::mem::size_of::<INPUT>() as i32) };

    if result == 1 { 0 } else { -3 }
}
