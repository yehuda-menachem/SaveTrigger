# SaveTrigger - Windows File Activity Smart Explorer System

## Project Goal

Build a Windows desktop background application that monitors file creation events in predefined folders and automatically opens the corresponding folder in Windows Explorer when a new file is created locally on the current machine.

The system intelligently:
- Ignores files created by other machines on shared/network drives
- Avoids opening duplicate windows
- Manages Explorer windows and tabs efficiently
- Prevents UI clutter
- Provides robust logging for debugging and improvement

## Platform & Stack

- **Language:** C#
- **Framework:** .NET 8
- **UI:** WPF (preferred) or WinUI 3
- **App Type:** Tray Application (Must run in user context)
- **Deployment:** Installable via installer (MSI or equivalent) with automatic Windows startup support.

## Core Features

### 1. Folder Monitoring
- Monitors a predefined static list of root directories recursively.
- Uses `FileSystemWatcher` to track `Created` and `Renamed` events.
- Handles large directory trees seamlessly.

### 2. Event Processing Pipeline
- **Raw Event Intake:** Collects filesystem events into a queue.
- **Debounce:** Groups events within a ~1000-2000 ms window to merge duplicate events.
- **File Stabilization:** Ensures the file is no longer being written (checks file size and locks).
- **Local Origin Correlation (Critical):** Determines if the file was created locally on this machine using heuristics (recent user activity, timing correlation) to avoid triggering on synced or network-created files.

### 3. Explorer Window Management
- Opens the containing folder in Windows Explorer when a valid file event is confirmed.
- **Window Ownership Tracking:** Tracks windows opened by the app (handle, folder path, open timestamp).
- **Smart Window Logic:** Brings existing managed windows to the front, or opens new tabs if possible (fallback to new window).
- **Window Limits:** Manages a maximum of 5 Explorer windows, automatically closing the oldest *managed* window if the limit is exceeded. Never closes unmanaged windows.
- Automatically selects the newly created file when the folder is opened.
- Moves the Explorer window to a specific monitor and brings it to the foreground.

### 4. System Tray UI
Minimal tray interface providing:
- Start / Resume
- Pause
- Exit
- Open Logs

### 5. Logging System
Two-level logging system (File-based, daily rotation, structured format):
- **User Log:** Readable events (File detected, Folder opened, Window reused/closed).
- **Debug Log:** Detailed internal state (Raw events, Debounce grouping, Stabilization checks, Correlation decisions, Explorer actions, and Errors).

## Architecture
The application is structured into logical components:
1. **Core:** Event models, Processing pipeline, Decision engine.
2. **Infrastructure:** FileSystemWatcher wrapper, Explorer control, Window management, OS interaction.
3. **UI:** Tray app, Basic controls.
4. **Logging:** Central logging service.

## Performance & Safety Constraints
- **Lightweight & Efficient:** Handles high-frequency events without consuming excessive CPU/RAM or blocking the UI thread.
- **Safety Rules:** 
  - NEVER interfere with user-opened Explorer windows.
  - NEVER close windows not owned by the app.
  - Avoid aggressive OS hooks.

## Edge Cases Handled
- Rapid creation of multiple files (treated as single event).
- File created then renamed.
- Temporary files created during the save process.
- Explorer crashes or closes.
- Shared/network drive noise and OneDrive sync duplication.

---

*This application is designed with a strict focus on correctness, stability, non-intrusiveness, and clear logging.*
