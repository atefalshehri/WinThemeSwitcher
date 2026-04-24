# Project Plan: Windows 11 25H2 Theme Switcher Tray App

## 1. Project Overview
The goal is to develop a lightweight, modern Windows 11 system tray application that automatically switches between Light and Dark themes based on the user's local sunrise and sunset times. The application will be designed for maximum efficiency, aiming for near-zero resource consumption by utilizing native Windows APIs and an event-driven architecture.

## 2. Technical Architecture
The application will be built using the **Rust** programming language, leveraging the official **Microsoft `windows-rs`** crate for direct Win32 API access. This approach ensures a small binary size and minimal memory footprint.

### 2.1 Core Components
| Component | Technology | Description |
| :--- | :--- | :--- |
| **Language** | Rust | Provides memory safety and zero-cost abstractions. |
| **Windows API** | `windows-rs` | Direct interaction with Windows Registry and System Tray. |
| **Tray UI** | `tray-icon` | Modern, lightweight system tray management. |
| **Astronomy** | `sun-times` | Local calculation of sunrise and sunset times. |
| **Location** | Windows Location API | Native geolocation to determine coordinates. |
| **Persistence** | Windows Registry | Stores user preferences and auto-start settings. |

### 2.2 Resource Optimization Strategy
To achieve the "near 0% resource" goal, the application will follow these principles:
- **Event-Driven Execution:** The app will calculate the next transition time (sunrise or sunset) and set a high-precision system timer to wake up only at that exact moment.
- **No UI Framework:** By avoiding heavy frameworks like Electron, WPF, or WinUI 3, the app will maintain a memory footprint of only a few megabytes.
- **Native Integration:** Direct Registry manipulation and `WM_SETTINGCHANGE` broadcasting will be used for theme switching, which is the most efficient method available.

## 3. Implementation Phases

### Phase 1: Core Logic Development
- Implement geolocation retrieval using the native Windows Location API, with a fallback to manual coordinate entry if Location Services are disabled.
- Develop a lightweight city-to-coordinate lookup feature for manual location selection.
- Integrate the `sun-times` library to calculate daily sunrise and sunset based on coordinates.
- Develop the theme-switching module that modifies the Windows Registry and notifies the system of changes.

### Phase 2: System Tray & UI
- Create a modern system tray icon with a context menu (Fluent Design style).
- Add options to manually toggle themes, refresh location, and access settings.
- Implement a "Settings" flyout using native Win32 dialogs for minimal overhead.

### Phase 3: Automation & Packaging
- Implement an "Auto-start on Login" feature via the Windows Registry.
- Create a robust error-handling system for location services and registry access.
- Package the application as a single, statically-linked `.exe` file for easy distribution.

## 4. Deliverables
- **Source Code:** Fully commented Rust project.
- **Executable:** A standalone, optimized `.exe` file.
- **Documentation:** Instructions for installation, usage, and manual configuration.
