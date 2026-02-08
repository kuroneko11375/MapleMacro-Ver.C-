//! macro_core - Rust native DLL for MapleStoryMacro
//!
//! Three core features:
//! 1. Hardware-level WH_KEYBOARD_LL hook (microsecond callback, never times out)
//! 2. Flash Focus key sending (AttachThreadInput + SetFocus + SendInput in <1ms)
//! 3. dwExtraInfo MAGIC_TAG filtering (zero key leakage)

mod hook;
mod sender;
mod types;

use std::ffi::c_void;
use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU32, Ordering};
use windows::Win32::Foundation::HWND;

pub use hook::*;
pub use sender::*;
pub use types::*;

// Global state (lock-free, read directly from hook callback)
static TARGET_HWND: AtomicIsize = AtomicIsize::new(0);
static BLOCKING_ENABLED: AtomicBool = AtomicBool::new(false);
static HOOK_HANDLE: AtomicIsize = AtomicIsize::new(0);
static BLOCKED_COUNT: AtomicU32 = AtomicU32::new(0);
static PASSED_COUNT: AtomicU32 = AtomicU32::new(0);

/// Must match C# KeyboardBlocker.MACRO_KEY_MARKER
const MACRO_KEY_MARKER: usize = 0x12345678;

// ---- FFI exports (called from C# via P/Invoke) ----

#[no_mangle]
pub extern "C" fn mc_set_target_hwnd(hwnd: isize) {
    TARGET_HWND.store(hwnd, Ordering::SeqCst);
}

#[no_mangle]
pub extern "C" fn mc_get_target_hwnd() -> isize {
    TARGET_HWND.load(Ordering::SeqCst)
}

#[no_mangle]
pub extern "C" fn mc_install_hook() -> bool {
    hook::install_hook()
}

#[no_mangle]
pub extern "C" fn mc_uninstall_hook() {
    hook::uninstall_hook();
}

#[no_mangle]
pub extern "C" fn mc_set_blocking(enabled: bool) {
    BLOCKING_ENABLED.store(enabled, Ordering::SeqCst);
}

#[no_mangle]
pub extern "C" fn mc_is_blocking() -> bool {
    BLOCKING_ENABLED.load(Ordering::SeqCst)
}

#[no_mangle]
pub extern "C" fn mc_flash_send_key(vk_code: u16, is_key_down: bool) -> i32 {
    let hwnd = HWND(TARGET_HWND.load(Ordering::SeqCst) as *mut c_void);
    if hwnd.0.is_null() {
        return -1;
    }
    sender::flash_focus_send_key(hwnd, vk_code, is_key_down)
}

#[no_mangle]
pub extern "C" fn mc_send_input_foreground(vk_code: u16, is_key_down: bool) -> i32 {
    sender::send_input_foreground(vk_code, is_key_down)
}

#[no_mangle]
pub extern "C" fn mc_get_blocked_count() -> u32 {
    BLOCKED_COUNT.load(Ordering::Relaxed)
}

#[no_mangle]
pub extern "C" fn mc_get_passed_count() -> u32 {
    PASSED_COUNT.load(Ordering::Relaxed)
}

#[no_mangle]
pub extern "C" fn mc_reset_stats() {
    BLOCKED_COUNT.store(0, Ordering::Relaxed);
    PASSED_COUNT.store(0, Ordering::Relaxed);
}

/// Health check - returns 0xDEADBEEF if DLL loaded OK
#[no_mangle]
pub extern "C" fn mc_ping() -> u32 {
    0xDEAD_BEEF
}
