//! 共用型別定義

/// 按鍵是否為延伸鍵（方向鍵、Insert、Delete 等）
pub fn is_extended_key(vk: u16) -> bool {
    matches!(
        vk,
        0x25 | 0x26 | 0x27 | 0x28 | // Left, Up, Right, Down
        0x2D | 0x2E |                 // Insert, Delete
        0x24 | 0x23 |                 // Home, End
        0x21 | 0x22 |                 // PageUp, PageDown
        0x90 | 0x2C |                 // NumLock
        0xA5 | 0xA3 | 0xA1            // RMenu (右 Alt), RCtrl, RShift
    )
}

/// 按鍵是否為方向鍵
pub fn is_arrow_key(vk: u16) -> bool {
    matches!(vk, 0x25 | 0x26 | 0x27 | 0x28)
}

/// 取得掃描碼（硬編碼常用鍵 + MapVirtualKey fallback）
pub fn get_scan_code(vk: u16) -> u16 {
    match vk {
        0x25 => 0x4B, // Left
        0x26 => 0x48, // Up
        0x27 => 0x4D, // Right
        0x28 => 0x50, // Down
        0x2D => 0x52, // Insert
        0x2E => 0x53, // Delete
        0x24 => 0x47, // Home
        0x23 => 0x4F, // End
        0x21 => 0x49, // PageUp
        0x22 => 0x51, // PageDown
        0x12 | 0xA4 | 0xA5 => 0x38, // Alt / LAlt, RAlt (same scan code, extended flag differs)
        _ => {
            // MapVirtualKey fallback
            use windows::Win32::UI::Input::KeyboardAndMouse::{MapVirtualKeyW, MAP_VIRTUAL_KEY_TYPE};
            unsafe { MapVirtualKeyW(vk as u32, MAP_VIRTUAL_KEY_TYPE(0)) as u16 }
        }
    }
}

/// 檢查 VK 是否應被攔截（與 C# KeyboardBlocker.IsArrowOrTargetKey 對應）
pub fn is_target_key(vk: u32) -> bool {
    matches!(vk, 0x25..=0x28)
        || matches!(vk, 0x2D | 0x2E | 0x24 | 0x23 | 0x21 | 0x22)
        || (0x41..=0x5A).contains(&vk)
        || (0x30..=0x39).contains(&vk)
        || (0x70..=0x7B).contains(&vk)
        || matches!(vk, 0x20 | 0x0D | 0x1B)
        || matches!(vk, 0x10 | 0x11 | 0x12 | 0xA0..=0xA5)
}
