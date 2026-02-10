# Copilot Instructions

## General Guidelines
- First general instruction
- Second general instruction

## Game-Specific Instructions
- For MapleStory private server background macros, avoid using pure PostMessage for arrow keys. Instead, use AttachThreadInput combined with PostMessage, as the game polls arrow keys via GetKeyState, not WM_KEYDOWN messages. Non-arrow keys (skills, potions) can be sent using pure PostMessage without issues. Alphanumeric keys can also be sent using PostMessage without issues. Using pure PostMessage for all keys can disrupt arrow key movement. Note that a fundamentally different approach is required for arrow keys, as they do not work with PostMessage variants due to the game's polling method.
- ZXC skill keys work perfectly with pure PostMessage in background mode.
- Since the user is operating on a MapleStory private server (¨pªA), methods such as WH_GETMESSAGE hooks, DLL injection, and in-process API hooking are viable options for macros without concerns about anti-cheat detection from GameGuard or XIGNCODE. Modifying memory is permissible, ensuring flexibility in implementation. This means that there are no compliance issues aside from memory modification, allowing for a range of intervention methods without worrying about anti-cheat detection. Additionally, there are no compliance restrictions during development, as the focus is solely on the private server environment.
- The Rust DLL (macro_core.dll) loads correctly, but be aware that AttachThreadInput can leak arrow key state to foreground applications (e.g., Chrome, VS).