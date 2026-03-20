<div align="center">

# 🔬 Sea-Bird Scientific Digital Sensor Workbench

**SBS DSW** — Professional bench workflow tool for digital sensor testing

![Version](https://img.shields.io/badge/version-1.5.0-00d4aa?style=flat-square)
![Python](https://img.shields.io/badge/python-3.9+-1e3a5f?style=flat-square)
![Platform](https://img.shields.io/badge/platform-Windows-0078d4?style=flat-square)

</div>

---

## 📖 Table of Contents

| Section | Description |
|---------|-------------|
| [Quick Start](#-quick-start) | Get running in 5 minutes |
| [Features](#-features) | What's new in v1.5.0 |
| [User Interface](#-user-interface) | Layout, panels, and tab views |
| [Serial Sniffer](#-serial-port-sniffer) | Traffic monitoring & bridge modes |
| [Sample Parser](#-sample-parser-configuration) | Configure field parsing |
| [Session Management](#-session-management) | Save, load, compare sessions |
| [Troubleshooting](#-troubleshooting) | Common issues & fixes |
| [Developer Guide](#-developer-guide) | Build from source |

---

## 🚀 Quick Start

### First-Time Setup

1. **Launch** the application (`sbs_dsw.exe` or run from source)
2. **Connect** your sensor via USB-to-serial adapter
3. **Configure** in ~60 seconds, then start testing

### 5-Minute Workflow

```
┌─────────────────────────────────────────────────────────────────┐
│  1. CONNECTION                                                  │
│     • Click "Refresh" → Select COM port → Set baud rate        │
│     • Click "▶ Connect"                                         │
├─────────────────────────────────────────────────────────────────┤
│  2. TEST SETUP                                                  │
│     • Enter Operator name (required)                            │
│     • Set Samples count (minimum 20)                            │
│     • Verify Results Root folder                                │
├─────────────────────────────────────────────────────────────────┤
│  3. RUN TEST                                                    │
│     • Click "▶ Run Test" in Actions                             │
│     • Watch Live Plot tab (auto-opens)                          │
│     • Results appear when complete (auto-switches)              │
├─────────────────────────────────────────────────────────────────┤
│  4. SAVE                                                        │
│     • Click "Save" to export session JSON                       │
│     • Sample CSVs saved automatically per-unit                  │
└─────────────────────────────────────────────────────────────────┘
```

---

## ✨ Features

### What's New in v1.5.0

| Feature | Description |
|---------|-------------|
| **Visual Documentation** | Professional SVG wireframe diagrams showing UI layout and workflows |
| **Enhanced README** | Embedded diagrams for Live Plot, Console, Setup, and Sniffer modes |

### Core Features (v1.4)

| Feature | Description |
|---------|-------------|
| **Modern Dark Theme** | Refined styling with teal accents, consistent button hierarchy |
| **Smart Layouts** | Auto-collapse panels during tests for maximum plot visibility |
| **Max View Mode** | One-click toggle to hide all controls and maximize workspace |
| **Serial Sniffer** | View raw hex/ASCII data from any COM port |
| **Mirror Mode** | Sniff traffic while routing through internal handlers |
| **com0com Bridge** | Intercept traffic between external apps and hardware |
| **Clearer UI** | All buttons use readable text labels instead of unicode symbols |

### Core Capabilities

- **Multi-Port Testing** — Connect up to 10 COM ports simultaneously
- **Parallel Batch Runs** — Execute tests across all ports with configurable delays
- **Live Plotting** — Real-time visualization during active runs
- **Flexible Parser** — Configure delimiters, regex, scaling, and derived fields
- **Session History** — Compare current results against reference sessions
- **Debug Console** — Per-port serial terminal with HEX/ASCII/DEC/BIN modes
- **Auto-Update** — Built-in updater checks your hosted manifest URL

---

## 🖥️ User Interface

### Layout Overview

<div align="center">
<img src="docs/images/dashboard-layout.svg" alt="Dashboard Layout" width="700">
</div>

<details>
<summary><strong>Text Diagram (fallback)</strong></summary>

```
┌────────────────────────────────────────────────────────────────────────┐
│  HEADER BAR                                                            │
│  [Help] [About] [Update] [Results] [Live] [Config] [Reset]             │
├────────────────────────────────────────────────────────────────────────┤
│  CONNECTION PANEL                    ┌─ PORT STATION ─────────────────┐│
│  COM Port [▼] Baud [▼] [Refresh]     │ Slot 1: COM5 - CONNECTED      ││
│  [▶ Connect] [Reconnect] [Disconnect]│ Slot 2: COM6 - RUNNING        ││
│  [▶ All] [◼ All] [≡ Ports]           │ Slot 3: --                    ││
├──────────────────────────────────────┴────────────────────────────────┤
│  TEST SETUP (collapsible)                                              │
│  Operator: [________] Bath ID: [________] Notes: [________]            │
│  Bath Temp: [____] Salinity: [____] Samples: [____] Results Root: [...] │
├────────────────────────────────────────────────────────────────────────┤
│  ACTIONS                                                               │
│  [▶ Run Test] [Reset] [Save]  |  [Live] [Console] [Detach]             │
│  [Plot] [Load] [Reload] [CSV Col]  |  [□ Max View]                     │
│  Runs: [__] Delay: [__]s  |  🌙 Dark  |  Mode: ○ Prod ○ Dev            │
├────────────────────────────────────────────────────────────────────────┤
│  ┌─ TABS ─────────────────────────────────────────────────────────────┐│
│  │ 📋 Results │ 📈 Live │ ⌨ Console │ 🔍 Sniffer │ 📜 Log │ ⚙ Setup │ ││
│  ├─────────────────────────────────────────────────────────────────────┤│
│  │                                                                     ││
│  │                    [Tab Content Area]                               ││
│  │                                                                     ││
│  └─────────────────────────────────────────────────────────────────────┘│
└────────────────────────────────────────────────────────────────────────┘
```

</details>

### Panel Reference

<details>
<summary><strong>Connection Panel</strong></summary>

| Control | Action |
|---------|--------|
| **Refresh** | Scan for available COM ports |
| **▶ Connect** | Open selected port at chosen baud rate |
| **Reconnect** | Reconnect selected port (useful after baud change) |
| **Disconnect** | Close selected port |
| **▶ All** | Connect all detected ports |
| **◼ All** | Disconnect all ports |
| **≡ Ports** | Toggle Port Station visibility |

</details>

<details>
<summary><strong>Test Setup Fields</strong></summary>

| Field | Required | Description |
|-------|----------|-------------|
| **Operator** | ✅ Yes | Your name/ID for traceability |
| **Bath ID** | No | Test fixture identifier |
| **Notes** | No | Free-form run notes |
| **Bath Temp (C)** | No | Environmental metadata |
| **Salinity (PSU)** | No | Environmental metadata |
| **Samples** | ✅ Yes | Samples per run (minimum 20) |
| **Results Root** | ✅ Yes | Output folder for all files |

</details>

<details>
<summary><strong>Action Buttons</strong></summary>

| Button | Description |
|--------|-------------|
| **▶ Run Test** | Start parallel test on all connected ports |
| **Reset** | Clear current session and start fresh |
| **Save** | Export session to JSON file |
| **Live** | Switch to Live Plot tab |
| **Console** | Switch to Serial Console tab |
| **Detach** | Pop out console to separate window |
| **Plot** | Open session comparison plot |
| **Load** | Load a saved session JSON |
| **Reload** | Refresh current session data |
| **CSV Col** | Toggle sample CSV path column visibility |
| **□ Max View** | Collapse all panels for maximum tab space |

</details>

<details>
<summary><strong>Run Controls</strong></summary>

| Control | Description |
|---------|-------------|
| **Runs** | Number of consecutive runs per port (1-50) |
| **Delay (s)** | Wait time between runs in a batch |
| **🌙 Dark** | Toggle dark/light theme |
| **Mode** | Switch between Production/Development modes |
| **Units tested** | Count of unique serial numbers this session (max 10) |
| **Runs left** | Remaining runs in current batch |

</details>

### Tab Views

<details>
<summary><strong>📈 Live Plot Tab</strong></summary>

Real-time visualization during active runs. Auto-opens when test starts.

<div align="center">
<img src="docs/images/live-plot.svg" alt="Live Plot Tab" width="650">
</div>

| Control | Description |
|---------|-------------|
| **Auto Y** | Auto-scale Y axis to fit data |
| **Points** | Number of visible data points (default: 100) |
| **Filter Ports** | Show/hide specific port traces |
| **Pause** | Freeze current view while run continues |

</details>

<details>
<summary><strong>⌨ Console Tab</strong></summary>

Raw serial I/O with timestamped TX/RX logging.

<div align="center">
<img src="docs/images/console-tab.svg" alt="Console Tab" width="650">
</div>

| Control | Description |
|---------|-------------|
| **Clear** | Clear console buffer |
| **Copy Text** | Copy console text to clipboard |
| **Timestamps** | Toggle timestamp display |
| **Show Hex** | Show hexadecimal representation |
| **Command** | Type and send commands to connected device |

</details>

---

## 🔍 Serial Port Sniffer

The **Sniffer** tab provides powerful tools for monitoring and intercepting serial traffic.

<div align="center">
<img src="docs/images/sniffer-modes.svg" alt="Sniffer Modes" width="700">
</div>

### Direct Sniffing Mode

Monitor raw data from any COM port:

1. Select a COM port from the dropdown
2. Click **Start Sniffing**
3. View data in your preferred format (HEX, ASCII, DEC, BIN)

| Option | Description |
|--------|-------------|
| **Display mode** | Choose HEX, ASCII, DEC, or BIN output |
| **Show timestamps** | Prepend receive time to each line |
| **Auto scroll** | Keep newest data visible |
| **Clear** | Clear the capture buffer |
| **Export** | Save captured data to file |
| **Copy** | Copy captured data to clipboard |

### Mirror Mode

Route data through your application while viewing traffic:

```
┌──────────┐     ┌─────────────────┐     ┌──────────┐
│  Sensor  │ ──► │  SBS DSW App    │ ──► │  Parser  │
│          │     │  (displays all) │     │          │
└──────────┘     └─────────────────┘     └──────────┘
```

1. Select the sensor's COM port
2. Enable **Mirror Mode** checkbox
3. Click **Start Sniffing**

Data flows to internal handlers AND appears in the sniffer display.

### com0com Bridge Mode

Intercept traffic between external applications and hardware:

```
┌────────────────┐     ┌───────────┐     ┌─────────────────┐     ┌──────────┐
│ External App   │ ──► │   COM10   │ ──► │   SBS DSW       │ ──► │  Sensor  │
│ (uses COM10)   │ ◄── │ (virtual) │ ◄── │   Bridge        │ ◄── │  (COM5)  │
└────────────────┘     └───────────┘     │   (displays)    │     └──────────┘
                                         └─────────────────┘
```

**Prerequisites:**
- Install [com0com](https://sourceforge.net/projects/com0com/) virtual null-modem driver
- Create a port pair (e.g., COM10 ↔ COM11)

**Setup Steps:**

1. Configure your external application to use COM10 (virtual)
2. Connect your sensor hardware to COM5 (physical)
3. In the Sniffer tab:
   - Set **Virtual Port (app-facing)**: COM10
   - Set **Physical Port (sensor)**: COM5
   - Click **Start Bridge**

**Buttons:**
- **Detect com0com** — Auto-discover virtual port pairs
- **?** — Show inline help dialog

---

## ⚙ Sample Parser Configuration

Open the **Setup** tab or click **Config** in the header.

<div align="center">
<img src="docs/images/setup-parser.svg" alt="Setup Tab" width="700">
</div>

### Quick Setup

```
1. Paste a sample line:    1234,56.78,90.12,OK
2. Set delimiter:          ,
3. Click:                  [Quick Setup + Plot]
```

### Parser Controls

| Setting | Description |
|---------|-------------|
| **Sample Cmd** | Command sent to request each sample (default: `tsr`) |
| **Delimiter** | Character separating fields (comma, space, tab, etc.) |
| **Trim Prefix** | Remove fixed prefix before parsing |
| **Start Token** | Skip first N tokens |
| **Regex** | Extract payload with regular expression |

### Field Editor

Each parsed field can be configured:

| Column | Description |
|--------|-------------|
| **Field Key** | Internal identifier |
| **Description** | Human-readable name |
| **Unit** | Display unit (ns, mV, °C, etc.) |
| **Scale** | Multiplier (raw, milli, micro, kilo) |
| **Min / Max** | Validation thresholds |
| **StuckN** | Flag if value unchanged for N samples |
| **Derived Expr** | Calculate from other fields |
| **Live** | Show in live plot |
| **Session** | Include in session stats |

### Profile Management

| Button | Action |
|--------|--------|
| **Quick Setup + Plot** | Auto-configure from sample line |
| **Load From Example** | Populate fields from example |
| **Apply Measureands** | Apply current field configuration |
| **Save Profile** | Save configuration to JSON file |
| **Load Profile** | Load configuration from JSON file |
| **Reset Default** | Restore factory defaults |

---

## 📊 Session Management

### Session Workflow

```
Test Execution ──► Results Table ──► Save JSON ──► Load & Compare Later
```

### Files Generated

Each test run creates:

| File | Contents |
|------|----------|
| `sbe83_session_<id>.json` | Complete session with all runs |
| `sbe83_session_<id>.csv` | Tabular session summary |
| `SBS83_SN<serial>_<timestamp>_samples.csv` | Raw sample data |
| `SBS83_SN<serial>_<timestamp>.log` | DS/DC output and metadata |
| `SBS83_SN<serial>_<timestamp>_summary.json` | Per-run metrics |

### Folder Structure

```
Results Root/
├── sessions/
│   └── PreCalTest/
│       ├── sbe83_session_*.csv
│       ├── sbe83_session_*.json
│       └── profiles/
│           └── *.json
└── <serial>/
    └── PreCalTest/
        ├── SBS83_SN*_samples.csv
        ├── SBS83_SN*_summary.json
        └── SBS83_SN*.log
```

### Session Comparison Plot

Compare current results against a reference session:

1. Click **Plot** in Actions
2. Select a metric to compare
3. Click **Load Reference Session** to load a saved JSON
4. Filter by serial number in either session
5. Use **Pause Plot** to freeze updates

### Severity Classifications

| Level | Criteria |
|-------|----------|
| **PASS** | Red noise ≤ 10 ns AND Blue noise ≤ 10 ns |
| **WARN** | Either noise > 10 ns AND ≤ 20 ns |
| **FAIL** | Either noise > 20 ns |
| **UNKNOWN** | Insufficient numeric data |

---

## 🔧 Troubleshooting

<details>
<summary><strong>"Run Test" button is disabled</strong></summary>

**Possible causes:**
- No port connected → Connect at least one COM port
- Missing required field → Fill in **Operator** field
- Run already active → Wait for current run to complete

</details>

<details>
<summary><strong>No samples collected during test</strong></summary>

**Check:**
- **Sample Cmd** matches your sensor's command
- **Baud rate** is correct for your device
- **Cable/adapter** is properly connected
- **Delimiter** and **Regex** settings match your data format

</details>

<details>
<summary><strong>Plot shows flat line or no data</strong></summary>

**Verify:**
- Selected **Plot field** exists in parser configuration
- **Unit** and **Scale** settings are appropriate
- Check **Run Log** for parse errors or timeouts

</details>

<details>
<summary><strong>Console commands get no response</strong></summary>

**Try:**
- Toggle **CR** and/or **LF** line endings
- Click **Read** after **Send** (if not streaming)
- Confirm port is connected and not busy with a test

</details>

<details>
<summary><strong>com0com bridge doesn't work</strong></summary>

**Verify:**
- com0com driver is installed and running
- Port pair is created (check Device Manager)
- External app is configured to use the virtual port
- Physical port is correctly selected for sensor side

</details>

---

## 👨‍💻 Developer Guide

### Run from Source

```powershell
# Clone and setup
git clone <repository>
cd git_entry

# Create virtual environment
python -m venv .venv
.venv\Scripts\Activate.ps1

# Install dependencies
pip install -e .

# Run application
python -m sbs_dsw.app
```

### Build Executable

```powershell
# Install PyInstaller
pip install pyinstaller

# Build (from project root)
pyinstaller app.spec

# Output: dist/sbs_dsw.exe
```

### Project Structure

```
git_entry/
├── src/sbs_dsw/
│   ├── app.py          # Main application
│   ├── styles.py       # Theme and styling
│   └── *_config.json   # Runtime configuration
├── docs/
│   └── QUICKSTART_ONE_PAGE.md
├── tools/update_server/
│   ├── publish_update.py
│   └── serve_updates.py
├── assets/
│   └── *.ico, *.png
└── README.md           # This file
```

### Web Updater Setup

Host your own update server:

1. **Publish an update:**
   ```powershell
   python tools/update_server/publish_update.py --exe dist/sbs_dsw.exe --version 1.4.0
   ```

2. **Start server:**
   ```powershell
   python tools/update_server/serve_updates.py --host 0.0.0.0 --port 8080
   ```

3. **Configure app:**
   Set **Update Feed URL** to `http://<your-server>:8080/manifest.json`

**Manifest format:**
```json
{
  "version": "1.4.0",
  "url": "https://your-server.com/sbs_dsw.exe",
  "sha256": "optional_64_char_hash",
  "notes": "Release notes shown to user"
}
```

---

## 📸 Screenshots

> **Note:** Add screenshots to `docs/images/` folder for visual documentation.

Recommended screenshots:
1. Main dashboard with connection panel and port station
2. Live Plot during active test run  
3. Sample Setup tab with parser configuration
4. Session comparison plot with reference loaded
5. Serial Sniffer in bridge mode
6. Console tab with hex display mode

To add screenshots, create `docs/images/` folder and reference them like:
```markdown
![Main Dashboard](docs/images/dashboard.png)
```

---

## 📄 License

See [LICENSE.txt](LICENSE.txt) for details.

---

<div align="center">

**Sea-Bird Scientific Digital Sensor Workbench** — Built for precision testing

*v1.4.0 • © Sea-Bird Scientific*

</div>
]]>
