# Research Findings: Windows 11 Theme Switcher Tray App

## 1. Theme Switching Mechanism
Windows 11 stores theme preferences in the Registry. To change the theme programmatically, we need to modify the following keys:
- **System Theme:** `HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\SystemUsesLightTheme`
- **App Theme:** `HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme`
- **Values:** `0` for Dark Mode, `1` for Light Mode.
- **Notification:** After changing the registry, we should broadcast a `WM_SETTINGCHANGE` message to notify running applications to update their UI.

## 2. Technology Stack Selection
To achieve "near 0% resource usage" and use "modern tools":
- **Language:** **Rust** is the best candidate. It provides memory safety, zero-cost abstractions, and produces a single, small binary with minimal runtime overhead.
- **Windows Integration:** The `windows-rs` crate (official Microsoft Rust bindings) allows direct access to Win32 APIs.
- **Tray Icon:** The `tray-icon` crate is a modern, cross-platform (but Windows-first) library for system tray management.
- **Sunrise/Sunset Calculation:** The `sunrise-sunset-calculator` or `sun-times` crate can perform calculations locally without network requests after the initial location is determined.
- **Location:** Windows Location API via `windows-rs` or a simple IP-based lookup (though Windows API is more "native").

## 3. Resource Optimization Strategy
- **Event-Driven:** Instead of polling every second, calculate the next sunrise/sunset time and use a timer (or `WaitableTimer` in Win32) to wake up the app only when a switch is needed.
- **No GUI:** The app will only have a system tray icon and a context menu, avoiding the overhead of a full UI framework like Electron or even WPF/WinUI 3.
- **Static Linking:** Compile as a standalone `.exe` to avoid dependency issues.

## 4. Windows 11 25H2 Specifics
- Ensure compatibility with the latest Win32 APIs.
- Use modern iconography (Fluent Design) for the tray icon.
- Support "Mica" or "Acrylic" if any settings window is eventually added (though a simple context menu is lighter).
