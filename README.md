---
## 📋 Table of Contents

- [Overview](#overview)
- [Screenshots & UI Layout](#screenshots--ui-layout)
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Running the Application](#running-the-application)
- [Building a Standalone EXE (Windows)](#building-a-standalone-exe-windows)
- [Application Sections — Full Reference](#application-sections--full-reference)
  - [Connection Panel](#1-connection-panel)
  - [Sampling Panel](#2-sampling-panel)
  - [Output Panel](#3-output-panel)
  - [Continuous Save Panel](#4-continuous-save-panel)
  - [Action Buttons](#5-action-buttons)
  - [Live Data View](#6-live-data-view)
- [Continuous Save Modes — Deep Dive](#continuous-save-modes--deep-dive)
  - [Mode ① Continuous (no time window)](#mode--continuous-no-time-window)
  - [Mode ② Simple Window](#mode--simple-window)
  - [Mode ③ Pre / Post Window](#mode--pre--post-window)
- [Save Every N Inputs — Decimation](#save-every-n-inputs--decimation)
- [Excel Output Format](#excel-output-format)
- [Usage Examples](#usage-examples)
- [Code Architecture](#code-architecture)
  - [Helper Functions](#helper-functions)
  - [SerialLogger Class](#seriallogger-class)
  - [App Class (GUI)](#app-class-gui)
- [Threading Model](#threading-model)
- [Troubleshooting](#troubleshooting)
- [License](#license)

---

## Overview

Serial Data Logger is a standalone GUI application that connects to any serial (COM/tty) port, displays incoming data in real time, and saves it to colour-formatted Excel files. It is designed for engineers, researchers, and makers who need to log data from microcontrollers (Arduino, ESP32, STM32, etc.), sensors, or any other serial device — without writing a single line of code.

Key design goals:

- **Zero configuration** — auto-detects available ports; sensible defaults for all settings
- **No data loss** — data is written to a CSV file immediately as it arrives; Excel conversion happens afterwards
- **Flexible capture** — five different ways to decide *which* readings get saved to Excel
- **Portable** — can be compiled to a single `.exe` for distribution to users who don't have Python

---

## Screenshots & UI Layout

```
┌──────────────────────────────────────────────────────────────────────┐
│  ⚡ Serial Data Logger          made by Emad    ● Disconnected       │
├─────────────────────────┬────────────────────────────────────────────┤
│  Connection             │  #      Time              Value            │
│  Sampling               │  ────────────────────────────────────────  │
│  Output                 │    1    10:42:01.234      23.5             │
│  Continuous Save        │    2    10:42:01.334      23.6             │
│  ─────────────────────  │    3    10:42:01.434      23.4  ← orange   │
│  💾 Save Snapshot       │    4    10:42:01.534      23.7             │
│  🗑 Clear Live View     │    ...                                     │
│  ⏹ Stop / Disconnect   │                                            │
│  Records: 0 | Saved: 0  │                                            │
├─────────────────────────┴────────────────────────────────────────────┤
│  Serial Data Logger  v1.0                           Made by Emad     │
└──────────────────────────────────────────────────────────────────────┘
```

---

## Features

| Feature | Description |
|---|---|
| 🔌 Auto port detection | Lists all available COM/tty ports; refresh with one click |
| ⚡ Configurable baud rate | Supports 300 – 921600 baud |
| 📈 Configurable sample rate | Set how many times per second to poll the serial port (Hz) |
| 🖥️ Live data view | Colour-coded scrolling terminal showing index, timestamp, and value |
| 💾 Snapshot save | Export the last N seconds of data to Excel at any moment |
| 🔁 Auto-save | Periodically save snapshots automatically at a configurable interval |
| 📂 Continuous save (3 modes) | Stream data to Excel continuously with time-based or event-based windowing |
| 🔢 N-input decimation | Save only every Nth reading — works with all continuous save modes |
| 🟢🟠 Colour-coded Excel output | Pre-buffer rows in green, post-window rows in orange, alternating blue otherwise |
| 📊 Summary sheet | Each Excel file includes a Summary tab with metadata |
| 🏗️ EXE build script | Included `build_exe.bat` to compile to a single Windows executable |

---

## Requirements

### Python packages

```bash
pip install pyserial openpyxl
```

| Package | Version | Purpose |
|---|---|---|
| `pyserial` | ≥ 3.5 | Serial port communication |
| `openpyxl` | ≥ 3.0 | Reading and writing `.xlsx` files |
| `tkinter` | built-in | GUI framework (included with standard Python) |

### Python version

Python **3.8 or higher** is recommended.

---

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/yourusername/serial-data-logger.git
cd serial-data-logger

# 2. Install dependencies
pip install pyserial openpyxl

# 3. Run
python serial_logger.py
```

> **Windows users:** If `tkinter` is missing, reinstall Python from [python.org](https://www.python.org/downloads/) and tick **"tcl/tk and IDLE"** during installation.

> **Linux users:** You may need to install tkinter separately:
> ```bash
> sudo apt install python3-tk
> ```
> You may also need to add your user to the `dialout` group to access serial ports:
> ```bash
> sudo usermod -a -G dialout $USER
> ```

---

## Running the Application

```bash
python serial_logger.py
```

The application window opens immediately. No serial device needs to be connected to launch the app.

---

## Building a Standalone EXE (Windows)

A `build_exe.bat` script is included. It automatically installs PyInstaller and packages the application into a single `.exe` file that works on any Windows PC — no Python installation required.

**Steps:**

1. Place `serial_logger.py` and `build_exe.bat` in the same folder
2. Double-click `build_exe.bat`
3. Wait ~30–60 seconds for the build to complete
4. Find your executable at `dist\SerialDataLogger.exe`
5. Copy and share `SerialDataLogger.exe` — it is completely self-contained

> **Note:** Windows Defender may show a SmartScreen warning the first time the `.exe` is run on a new machine. This is a common false positive for PyInstaller-built applications. Click **"More info → Run anyway"** to proceed.

---

## Application Sections — Full Reference

### 1. Connection Panel

| Control | Description |
|---|---|
| **Port** dropdown | Lists all available serial ports (COM1, COM3, /dev/ttyUSB0, etc.). Click **↻** to refresh the list at any time. |
| **Baud Rate** dropdown | Select the communication speed. Must match the baud rate configured on your serial device. Common values: `9600`, `115200`. |
| **Connect / Disconnect** button | Opens the serial port and starts the reading loop. The button text and the top-right status indicator toggle between connected and disconnected states. |

### 2. Sampling Panel

| Control | Default | Description |
|---|---|---|
| **Sample Rate (Hz)** | `10` | How many times per second the application reads from the serial buffer. `10` = poll 10 times/sec. Does not affect how fast your device sends data — only how often the app checks for new data. |
| **Snapshot Interval (s)** | `60` | When you click **Save Snapshot**, this many seconds of buffered data are exported. E.g. `60` exports the last 1 minute of readings. |
| **Auto-Save Every (s)** | `0` | If > 0, automatically saves a snapshot at this interval while connected. `0` disables auto-save. |
| **Save every N inputs** | `1` | Decimation factor for Continuous Save. `1` = save every reading. `10` = save every 10th reading. See [Save Every N Inputs](#save-every-n-inputs--decimation). |

### 3. Output Panel

| Control | Description |
|---|---|
| **Save Folder** | Directory where all Excel files are saved. Click **…** to browse. Defaults to your home directory. |
| **Filename Prefix** | Prefix for all output filenames. Files are named `<prefix>_<timestamp>.xlsx`. Default: `serial_log`. |

### 4. Continuous Save Panel

Controls how data is streamed to an Excel file while the connection is active. See [Continuous Save Modes](#continuous-save-modes--deep-dive) for full details.

| Control | Description |
|---|---|
| **① Continuous (no time window)** | Records all readings (subject to N decimation) from the moment you press Start until you press Stop. |
| **② Simple window** | Records for a burst of `window` seconds, pauses, then records again every `period` seconds. |
| **③ Pre / Post window** | At each trigger (every `period` seconds), saves the last `pre` seconds from the buffer AND records live for the next `post` seconds. |
| **▶ Start Continuous Save** | Begins streaming. A CSV file is opened immediately so no data is lost. |
| **■ Stop & Save Excel** | Stops streaming and converts the CSV to a formatted Excel file in a background thread. |
| **Gate indicator** | Shows the current phase: `◯ waiting for trigger`, `⬤ flushing pre-buffer`, `⬤ POST window`, `⬤ RECORDING`, or `◯ waiting…` |

### 5. Action Buttons

| Button | Description |
|---|---|
| **💾 Save Snapshot to Excel** | Immediately exports the last N seconds (set by Snapshot Interval) to a new Excel file. Does not interrupt continuous save. |
| **🗑 Clear Live View** | Clears the scrolling terminal. Does not delete any saved data. Resets the live row counter. |
| **⏹ Stop / Disconnect** | Stops continuous save (if running), closes the serial port, and disconnects cleanly. |

### 6. Live Data View

The right panel shows every incoming reading in real time.

| Column | Description |
|---|---|
| **#** | Sequential row number since the last clear. Grey text. |
| **Time** | Timestamp of when the reading was received, in `HH:MM:SS.mmm` format. Blue text. |
| **Value** | The raw string received from the serial port. |

**Colour coding of the Value column:**

| Colour | Meaning |
|---|---|
| 🟢 Green (normal) | Reading received, not being saved to continuous save file |
| 🟠 Orange bold | Reading is being saved to the continuous save file (Mode ① or ② window active) |
| 🟢 Green bold | Pre-buffer flush in Mode ③ (these rows were already in the buffer before the trigger) |
| 🟠 Orange bold | Post-window in Mode ③ (live recording after trigger) |

The status bar below the live view shows: `Records: <total seen>  |  Saved: <total written to CSV>`

---

## Continuous Save Modes — Deep Dive

### Mode ① Continuous (no time window)

**Use case:** You want a complete log of everything the device sends for the entire session.

**How it works:** Every reading that arrives is immediately written to the CSV file. Combined with the N-decimation setting, you can reduce the file size by saving only every Nth reading.

**Timeline:**
```
t=0s   ▶ Start pressed
       ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ (all readings saved)
t=300s ■ Stop pressed → Excel file generated
```

---

### Mode ② Simple Window

**Use case:** Your device generates high-frequency data, but you only need periodic samples — e.g. record 10 seconds of data every 5 minutes.

**Settings:**
- **Record for** `window` seconds
- **Every** `period` seconds

**How it works:** A background scheduler thread alternates between an OPEN phase (gate = True, readings saved) and a CLOSED phase (gate = False, readings skipped). The gate starts open as soon as you press Start.

**Timeline example (`window=10`, `period=60`):**
```
t=0s   Gate OPENS  → 10s of data saved (shown orange in live view)
t=10s  Gate CLOSES → 50s of data skipped
t=60s  Gate OPENS  → 10s of data saved
t=70s  Gate CLOSES → 50s of data skipped
...
```

**Validation:** `window` must be greater than 0 and strictly less than `period`.

---

### Mode ③ Pre / Post Window

**Use case:** You want to capture what happened *just before* and *just after* a periodic event — like a sensor reading, a system state change, or a scheduled trigger. This is the most powerful mode.

**Settings:**
- **Trigger every** `period` seconds — how often the capture fires
- **Pre-buffer** `pre` seconds — how much historical data to save before the trigger
- **Post-window** `post` seconds — how long to record live after the trigger

**How it works:**

The application maintains a rolling in-memory buffer of all readings at all times. When the trigger fires:

1. **Pre-buffer flush:** The scheduler looks back `pre` seconds into the buffer and writes all those readings to the CSV with flag `"pre"` (green in Excel)
2. **Post-window open:** The gate opens for `post` seconds; any reading that arrives during this time is written with flag `"post"` (orange in Excel)
3. **Sleep:** The gate closes and the scheduler waits for the remainder of the period, then the cycle repeats

**Timeline example (`period=120`, `pre=20`, `post=30`):**
```
t=0s    Session starts — data flows into memory buffer (nothing saved yet)
        ◯ waiting for trigger...

t=120s  TRIGGER fires:
        ├─ Pre-flush: writes readings from t=100s → t=120s  (green rows)
        ├─ Post-window OPENS
        │   readings from t=120s → t=150s saved             (orange rows)
        └─ Post-window CLOSES at t=150s

t=150s  ◯ waiting for next trigger (70s gap)

t=240s  TRIGGER fires again → same cycle repeats
```

**Result in Excel:** Each trigger produces a 50-second window of data (20s before + 30s after), colour-coded so you can immediately see the boundary.

**Validation:** `pre + post` must not exceed `period`. `period` must be > 0. At least one of `pre` or `post` must be > 0.

---

## Save Every N Inputs — Decimation

The **Save every N inputs** field in the Sampling panel is a decimation filter that applies **on top of all continuous save modes**.

| N value | Behaviour |
|---|---|
| `1` (default) | Every reading that passes the gate check is saved — no decimation |
| `5` | 1 out of every 5 gate-passing readings is saved |
| `10` | 1 out of every 10 gate-passing readings is saved |
| `100` | 1 out of every 100 gate-passing readings is saved |

**Important:** The counter only increments on readings that *pass the gate*. This means:
- In Mode ②, the counter resets conceptually to track readings within each open window
- In Mode ③, the counter tracks readings within the post-window
- The live view always shows all readings regardless of decimation

**Combination examples:**

| Mode | N | Result |
|---|---|---|
| ① Continuous | 1 | Every single reading saved |
| ① Continuous | 10 | Every 10th reading saved; 90% of data discarded |
| ② Simple window `10s / 60s` | 5 | During each 10s window, save every 5th reading |
| ③ Pre/Post `120s / 20s pre / 30s` | 1 | All 50s of data around each trigger saved |
| ③ Pre/Post `120s / 20s pre / 30s` | 10 | Every 10th reading in the post-window saved (pre-buffer always saves all) |

> **Note:** The N-decimation setting does not affect the pre-buffer flush in Mode ③. Pre-buffer rows are always written in full to preserve the historical context.

---

## Excel Output Format

Every exported Excel file follows the same structure:

### Sheet 1: Serial Data

| Column | Header | Type | Description |
|---|---|---|---|
| A | `#` | Integer | Sequential row number |
| B | `Timestamp` | String | Full timestamp: `YYYY-MM-DD HH:MM:SS.mmm` |
| C | `Date` | String | Date portion: `YYYY-MM-DD` |
| D | `Time` | String | Time portion: `HH:MM:SS.mmm` |
| E | `Raw Value` | String | The raw string received from the serial port |

**Row colour coding:**

| Colour | Hex | Meaning |
|---|---|---|
| Dark blue header | `#1F4E79` | Column headers |
| Alternating blue | `#D6E4F0` | Normal even rows (Mode ① or unflagged) |
| White | — | Normal odd rows |
| Soft green | `#C8E6C9` | Pre-buffer rows (Mode ③, data before trigger) |
| Soft orange | `#FFE0B2` | Post-window rows (Mode ③, data after trigger) or windowed rows (Mode ②) |

### Sheet 2: Summary

Contains metadata about the export:

| Field | Description |
|---|---|
| Export Time | When the file was generated |
| Total Records | Number of data rows in the Serial Data sheet |
| Source CSV | Filename of the intermediate CSV (continuous save only) |
| Green rows | Explanation of colour coding |
| Orange rows | Explanation of colour coding |

### File naming convention

| Save type | Filename pattern |
|---|---|
| Manual snapshot | `<prefix>_YYYYMMDD_HHMMSS.xlsx` |
| Auto-save | `<prefix>_auto_YYYYMMDD_HHMMSS.xlsx` |
| Continuous save | `<prefix>_continuous_YYYYMMDD_HHMMSS.xlsx` |

---

## Usage Examples

### Example 1: Basic logging from an Arduino

**Scenario:** An Arduino sends temperature readings every 100ms over USB serial at 9600 baud. You want to log for 5 minutes.

**Settings:**
```
Port:               COM3
Baud Rate:          9600
Sample Rate (Hz):   10
Snapshot Interval:  300
Save Folder:        C:\Data\Arduino
Filename Prefix:    temperature_log
```

**Steps:**
1. Select `COM3` and `9600` baud
2. Click **Connect**
3. Watch readings appear in the live view
4. After 5 minutes, click **💾 Save Snapshot to Excel**
5. Open `temperature_log_20241201_103000.xlsx`

---

### Example 2: Periodic burst logging (low storage)

**Scenario:** A device sends 100 readings/second but you only need a 5-second sample every 2 minutes to monitor trends over 8 hours.

**Settings:**
```
Continuous Save Mode:   ② Simple window
Record for:             5 s
Every:                  120 s
Save every N inputs:    1
```

**What happens:**
- 5 seconds of data saved every 2 minutes = 150 saves per 8-hour session
- Each save contains ~500 readings (5s × 100 readings/s)
- Total: ~75,000 rows in the Excel file instead of 2,880,000

---

### Example 3: Event-context capture (Pre/Post mode)

**Scenario:** A motor controller sends telemetry at 50Hz. You want to capture what the system was doing 10 seconds before and 20 seconds after a periodic diagnostic trigger every 5 minutes.

**Settings:**
```
Continuous Save Mode:   ③ Pre / Post window
Trigger every:          300 s
Pre-buffer:             10 s
Post-window:            20 s
Save every N inputs:    1
```

**What happens:**
- At t=300s: saves readings from t=290s→300s (green, 500 rows) + t=300s→320s (orange, 1000 rows)
- At t=600s: saves readings from t=590s→600s (green) + t=600s→620s (orange)
- Each trigger produces 30 seconds (1500 rows) of context data
- Over 1 hour: 12 triggers × 1500 rows = 18,000 rows total

---

### Example 4: Decimated high-frequency logging

**Scenario:** A sensor sends data at 1000Hz but you only need 1 reading per second for a long-term trend.

**Settings:**
```
Sample Rate (Hz):       100   (poll the port 100×/s)
Continuous Save Mode:   ① Continuous
Save every N inputs:    100
```

**What happens:**
- The app checks the port 100 times/second
- Every 100th reading is saved → effectively 1 reading/second
- Over 1 hour: 3,600 rows instead of 360,000

---

### Example 5: Auto-save snapshots overnight

**Scenario:** You want the app to automatically save a fresh snapshot every 10 minutes while your device runs overnight.

**Settings:**
```
Snapshot Interval (s):   600    (last 10 minutes of data)
Auto-Save Every (s):     600    (trigger every 10 minutes)
Filename Prefix:         overnight_log
```

**What happens:**
- Every 10 minutes a new `.xlsx` file is created with the previous 600 seconds of data
- Files are named: `overnight_log_auto_20241201_220000.xlsx`, `overnight_log_auto_20241201_221000.xlsx`, ...

---

## Code Architecture

The application is structured as a single Python file with two main classes and two helper functions.

```
serial_logger.py
├── sanitize()          — strips illegal Excel characters from serial data
├── csv_to_excel()      — converts the continuous-save CSV to formatted Excel
├── SerialLogger        — handles all serial I/O and data saving logic
│   ├── connect() / disconnect()
│   ├── start_reading() / _read_loop()
│   ├── save_snapshot_to_excel()
│   ├── start_continuous_save() / stop_continuous_save()
│   ├── _cs_write() / _cs_write_row_direct()
│   ├── _scheduler_window()     — Mode B gate thread
│   └── _scheduler_prepost()    — Mode C trigger thread
└── App (tk.Tk)         — the entire GUI
    ├── _build_ui()
    ├── _connect() / _disconnect()
    ├── _on_data() / _append_live_row()
    ├── _save_snapshot()
    ├── _start_continuous_save() / _stop_continuous_save()
    └── _schedule_autosave()
```

---

### Helper Functions

#### `sanitize(v: str) → str`

```python
_ILLEGAL = re.compile(
    r"[\x00-\x08\x0b-\x0c\x0e-\x1f\x7f-\x84\x86-\x9f\ud800-\udfff\ufffe\uffff]"
)

def sanitize(v: str) -> str:
    return _ILLEGAL.sub("", v)
```

Strips all characters that are illegal in Excel cells using a compiled regex. This is applied to every incoming serial reading before it is stored in the buffer or written to a file. Without this, `openpyxl` raises an `IllegalCharacterError` when the serial stream contains control characters (common in raw binary/protocol traffic).

---

#### `csv_to_excel(csv_path, xlsx_path) → int`

Converts the intermediate CSV file produced by continuous save into a formatted Excel workbook. Returns the number of data rows written.

The CSV has 6 columns internally: `index | full_timestamp | date | time | value | phase_flag`

The `phase_flag` column is used to determine row colour:

```python
FILLS = {
    "pre":  PatternFill("solid", fgColor="C8E6C9"),  # green
    "post": PatternFill("solid", fgColor="FFE0B2"),  # orange
    "1":    PatternFill("solid", fgColor="FFE0B2"),  # orange (simple window)
    "alt":  PatternFill("solid", fgColor="D6E4F0"),  # alternating blue
    "":     PatternFill(),                           # white
}
```

This function runs in a background thread when you press **■ Stop & Save Excel**, so the UI remains responsive during large file conversions.

---

### SerialLogger Class

The `SerialLogger` class is the engine of the application. It is completely decoupled from the GUI — it communicates back via callbacks.

#### `__init__`

Initialises all state variables. Key attributes:

| Attribute | Type | Description |
|---|---|---|
| `serial_port` | `serial.Serial` | The open serial port object |
| `is_running` | `bool` | Controls the read loop thread |
| `data_buffer` | `list[(datetime, str)]` | Rolling in-memory buffer of all readings |
| `lock` | `threading.Lock` | Protects `data_buffer` from concurrent access |
| `_cs_active` | `bool` | Whether continuous save is currently running |
| `_cs_gate` | `bool` | True when Mode B gate is open (readings should be saved) |
| `_cs_in_post` | `bool` | True when Mode C is inside its post-window |
| `_cs_stop_evt` | `threading.Event` | Set to signal scheduler threads to stop immediately |
| `_nth_step` | `int` | Decimation factor (1 = no decimation) |
| `_nth_counter` | `int` | Counts gate-passing readings since last save |

---

#### `connect(port, baud)` / `disconnect()`

```python
def connect(self, port, baud):
    self.serial_port = serial.Serial(port, baud, timeout=1)
```

Opens the serial port with a 1-second read timeout. The timeout ensures `readline()` does not block indefinitely if the device stops sending data.

`disconnect()` sets `is_running = False` (which exits the read loop on its next iteration), calls `stop_continuous_save()`, then closes the port.

---

#### `start_reading(interval_s, on_data_cb)` / `_read_loop()`

```python
def start_reading(self, interval_s, on_data_cb):
    self.is_running = True
    threading.Thread(
        target=self._read_loop, args=(interval_s, on_data_cb), daemon=True
    ).start()
```

Launches a daemon thread that polls the serial buffer at the configured sampling interval. The loop uses a time-tracking variable (`nxt`) to maintain accurate intervals without drifting:

```python
nxt = time.time()
while self.is_running:
    now = time.time()
    if now >= nxt:
        nxt = now + interval_s
        # ... read and process
    time.sleep(0.005)   # 5ms sleep to avoid 100% CPU usage
```

For each non-empty reading:
1. The raw bytes are decoded as UTF-8 (with error replacement)
2. `sanitize()` strips illegal characters
3. The clean value is appended to `data_buffer` (thread-safe)
4. `_cs_write()` is called — returns `True` if the reading was saved to CSV
5. `on_data_cb(ts, clean, saved)` notifies the GUI

---

#### `save_snapshot_to_excel(filepath, interval_s)`

Reads the last `interval_s` seconds from `data_buffer` (under lock), creates a new `openpyxl` workbook, writes the data rows with alternating blue fill, and saves a Summary sheet. This runs on the main thread (triggered by button click) so it's safe to call without any special threading consideration.

---

#### `start_continuous_save(csv_path, window, period, pre, post, on_gate_cb)`

```python
def start_continuous_save(self, csv_path, window=0.0, period=0.0,
                           pre=0.0, post=0.0, on_gate_cb=None):
```

Opens the CSV file immediately (so data is never lost even if the app crashes), sets all CS state variables, then picks the appropriate scheduler thread:

```python
if pre > 0 or post > 0:
    t = threading.Thread(target=self._scheduler_prepost, daemon=True)
elif window > 0 and period > 0:
    t = threading.Thread(target=self._scheduler_window, daemon=True)
else:
    return   # Mode A: no scheduler needed — _cs_write handles everything
```

---

#### `stop_continuous_save() → str`

```python
def stop_continuous_save(self) -> str:
    with self._cs_lock:
        self._cs_active = False
        self._cs_stop_evt.set()   # wake scheduler if sleeping
        # ... flush and close file
        return path
```

Sets `_cs_active = False` and signals `_cs_stop_evt` so any `stop.wait(timeout=...)` calls in the scheduler threads return immediately rather than waiting for the next timeout to expire. Returns the CSV path so the caller can convert it to Excel.

---

#### `_cs_write(ts, val, flag) → bool`

The core write gate. Called for every incoming reading. Returns `True` if the row was actually written.

```python
def _cs_write(self, ts, val, flag="") -> bool:
    with self._cs_lock:
        if not self._cs_active or self._cs_writer is None:
            return False

        # Step 1: Gate / mode check
        if self._cs_pre > 0 or self._cs_post > 0:     # Mode C
            if not self._cs_in_post:
                return False
            flag = "post"
        elif not self._cs_gate:                         # Mode B (gate closed)
            return False
        elif self._cs_window > 0:
            flag = "1"

        # Step 2: Nth-sample decimation
        if self._nth_step > 1:
            self._nth_counter += 1
            if self._nth_counter < self._nth_step:
                return False
            self._nth_counter = 0

        # Write the row
        self._cs_writer.writerow([index, full_ts, date, time, val, flag])
        self._cs_file.flush()   # flush immediately — no data loss on crash
```

The two-step design is intentional: **gate first, then decimate**. This ensures the N counter only ticks on readings that would actually be saved, making the decimation behaviour predictable regardless of which mode is active.

---

#### `_scheduler_window()`

```python
def _scheduler_window(self):
    stop = self._cs_stop_evt
    while True:
        self._cs_gate = True                         # open
        if self._cs_on_gate_cb: self._cs_on_gate_cb("open")
        if stop.wait(timeout=self._cs_window): break # sleep, or exit if stopped

        self._cs_gate = False                        # close
        if self._cs_on_gate_cb: self._cs_on_gate_cb("closed")
        gap = self._cs_period - self._cs_window
        if gap > 0:
            if stop.wait(timeout=gap): break

        if not self._cs_active: break
```

Uses `threading.Event.wait(timeout=N)` instead of `time.sleep(N)`. The critical advantage: when `stop_continuous_save()` calls `self._cs_stop_evt.set()`, the `wait()` returns `True` immediately — the thread exits within milliseconds rather than waiting for the current sleep phase to expire.

---

#### `_scheduler_prepost()`

```python
def _scheduler_prepost(self):
    stop = self._cs_stop_evt
    # Wait for first trigger
    if stop.wait(timeout=self._cs_period): return

    while self._cs_active:
        # 1. Flush pre-buffer
        if pre > 0:
            cutoff = now - timedelta(seconds=pre)
            pre_rows = [r for r in data_buffer if r.ts >= cutoff]
            for ts, v in pre_rows:
                self._cs_write_row_direct(ts, v, "pre")

        # 2. Open post-window
        self._cs_in_post = True
        if stop.wait(timeout=post): break

        # 3. Close and sleep
        self._cs_in_post = False
        gap = self._cs_period - post
        if stop.wait(timeout=gap): break
```

Pre-buffer rows are written with `_cs_write_row_direct()` which bypasses the gate and decimation checks — the full pre-buffer is always preserved intact.

---

### App Class (GUI)

The `App` class inherits from `tk.Tk` and builds the entire interface in `_build_ui()`. It communicates with `SerialLogger` exclusively through method calls and callbacks.

#### Thread-safe UI updates

The serial reading thread cannot directly update tkinter widgets (tkinter is not thread-safe). All UI updates from background threads are marshalled to the main thread using `self.after(0, callback)`:

```python
def _on_data(self, ts, raw, saved):
    # Called from the read thread
    self.after(0, self._append_live_row, ts, raw, saved)
    #          ↑ schedules _append_live_row on the main event loop
```

#### `_start_continuous_save()`

Reads and validates all UI settings, then calls `self.logger.start_continuous_save()`. Also applies the N-decimation setting:

```python
nth = int(self.nth_var.get())
self.logger._nth_step = max(1, nth)
self.logger._nth_counter = 0
```

#### `_stop_continuous_save()`

Calls `self.logger.stop_continuous_save()` to get the CSV path, resets decimation state, then launches a background thread to run `csv_to_excel()`:

```python
def _cvt():
    n = csv_to_excel(csv_path, xlsx_path)
    os.remove(csv_path)             # clean up intermediate CSV
    self.after(0, self._cs_done, xlsx_path, n)

threading.Thread(target=_cvt, daemon=True).start()
```

---

## Threading Model

The application uses up to four threads simultaneously:

```
Main thread (tkinter event loop)
│
├── Read thread (daemon)          — _read_loop()
│   Polls serial port at sampling_interval_s.
│   Writes to data_buffer (protected by self.lock).
│   Calls _cs_write() for CSV output.
│   Calls self.after(0, ...) for UI updates.
│
├── Scheduler thread (daemon)     — _scheduler_window() or _scheduler_prepost()
│   Sleeps with threading.Event.wait() so it can be woken instantly.
│   Sets _cs_gate / _cs_in_post flags read by the Read thread.
│   Calls on_gate_cb() → self.after(0, _show_gate) for indicator updates.
│
└── Conversion thread (daemon)    — csv_to_excel()
    Launched when Stop is pressed.
    Reads the CSV and writes Excel — CPU-bound, runs in background.
    Calls self.after(0, _cs_done) when finished.
```

**Lock usage:**
- `self.lock` — protects `data_buffer` (read thread writes, main thread reads for snapshot/pre-flush)
- `self._cs_lock` — protects all `_cs_*` file/writer state (read thread and scheduler both access)

---

## Troubleshooting

| Problem | Likely Cause | Solution |
|---|---|---|
| Port not appearing in dropdown | Device not connected or driver not installed | Connect device, install driver (e.g. CH340, CP2102), click ↻ |
| `Permission denied` on port | Another app has the port open, or no permission (Linux) | Close other apps; on Linux: `sudo usermod -a -G dialout $USER` |
| Garbled data in live view | Wrong baud rate | Match baud rate to your device's configuration |
| `IllegalCharacterError` in Excel | Binary/control characters in serial data | Already handled by `sanitize()` — if still occurring, check device is sending plain text |
| EXE shows SmartScreen warning | Unsigned PyInstaller binary | Click "More info → Run anyway" — this is a false positive |
| EXE is slow to start | PyInstaller extraction on first run | Normal — the EXE extracts itself to a temp folder on first launch |
| No data saved after pressing Start | N=10 set but device sends < 10 readings | Reduce N or wait for more data to arrive |
| Pre-buffer is empty in Mode ③ | Session just started, buffer not populated yet | Wait at least `pre` seconds after connecting before the first trigger fires |
| Auto-save creates empty files | Snapshot Interval longer than time connected | Set Snapshot Interval ≤ time since connecting |

---

## License

MIT License — free to use, modify, and distribute with attribution.

```
Made by Emad
```
