# Sea-Bird Scientific Digital Sensor Workbench (SBS DSW)

Version: `v1.2.0`

This file is the primary end-user guide and is the source opened by the in-app `Help` link.

## Start Here

- New user: go to [Fast Path](docs/QUICKSTART_ONE_PAGE.md#fast-path)
- Need a specific control name: go to [Panel Reference](#panel-reference)

## Quick Navigation

| I need to... | Go here |
|---|---|
| connect ports and start a test | [Fast Path](docs/QUICKSTART_ONE_PAGE.md#fast-path) |
| understand every button in the app | [Panel Reference](#panel-reference) |
| change parser fields and units | [Generic Sample Format (Parser + Mapping)](#generic-sample-format-parser-mapping) |
| compare current run to an old session | [Session Plot and Comparison](#session-plot-and-comparison) |
| find where files are saved | [Output Files and Folder Structure](#output-files-and-folder-structure) |
| troubleshoot run issues | [Troubleshooting by Symptom](#troubleshooting-by-symptom) |

## FAQ (Quick Answers)

| Question | Short answer | Details |
|---|---|---|
| How do I start a run fast? | Connect port, fill `Operator`, set `Samples`, click `Run Test`. | [Fast Path](docs/QUICKSTART_ONE_PAGE.md#fast-path) |
| Why is `Run Test` disabled? | No connected port, missing `Operator`, or run already active. | [Troubleshooting by Symptom](#troubleshooting-by-symptom) |
| Where do I change parser settings? | Open the `Sample Setup` tab or click `Config`. | [Generic Sample Format (Parser + Mapping)](#generic-sample-format-parser-mapping) |
| How do I compare to a previous session? | Open `Session Plot`, then load a reference session JSON. | [Session Plot and Comparison](#session-plot-and-comparison) |
| Where are output files saved? | Under `Results Root` in session and per-serial folders. | [Output Files and Folder Structure](#output-files-and-folder-structure) |
| Can I test multiple ports at once? | Yes, runs execute in parallel across connected ports. | [Actions](#actions) |
| How do I export session results? | Click `Save JSON` in `Actions`. | [Actions](#actions) |
| How do I export serial console traffic? | Use `Export Console CSV` in `Serial Consoles`. | [Serial Consoles](#serial-consoles) |
| How do I hide or show sample CSV paths? | Click `CSV Column` in `Actions`. | [Actions](#actions) |
| What do `PASS`, `WARN`, `FAIL` mean? | They are noise-based severity levels. | [Severity Logic](#severity-logic) |

## What This Application Does

SBS DSW is a bench workflow tool for digital sensors. It combines:

- multi-port serial connection management (up to 10 COM ports)
- repeatable sample-based test execution
- live plotting during runs
- configurable sample parsing and field mapping
- per-port serial debug consoles
- session history, comparison plotting, and exportable artifacts

## What Is New In `v1.2.0`

- Help link navigation fixed in packaged app (`Help` no longer resolves section links to `_MEI...` directory index pages).

- Header quick actions: `Results`, `Live`, `Config`, `Reset`
- Config Mode for parser setup-first workflow
- Generic Sample Format editor:
  - quick setup from one sample line
  - parser profiles (`Save Profile` / `Load Profile`)
  - per-field units and scaling
  - derived expressions
  - min/max/stuck-run rule flags
- Parallel multi-port batch runs with `Runs`, `Delay (s)`, and `Runs left`
- Session comparison plot with reference JSON loading
- Detachable serial console and display modes (`ASCII`, `HEX`, `DEC`, `BIN`)

## 5-Minute Quickstart

1. Launch `sbs_dsw.exe`.
2. In `Connection`, click `Refresh Ports`, choose COM port and baud, then click `Connect Selected` (or `Connect All`).
3. In `Test Setup`, enter `Operator`, set `Samples` (minimum 20), and confirm `Results Root`.
4. In `Actions`, set `Runs` and optional `Delay (s)`, then click `Run Test`.
5. Monitor `Live Plot`, `Run Log`, and `Port Station View`.
6. Review completed rows in `Test Results`.
7. Click `Save JSON` when done.

## Command Examples

### Run from source checkout

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install .
python sbs_dsw.py
```

### Run packaged executable

```powershell
.\dist\sbs_dsw.exe
```

### Inspect latest session outputs

```powershell
Get-ChildItem .\SBE83\sessions\PreCalTest | Sort-Object LastWriteTime -Descending | Select-Object -First 10
```

## UI Map

- Header links: `Help`, `About`, `Results`, `Live`, `Config`, `Reset`
- Main cards: `Connection`, `Port Station View`, `Test Setup`, `Actions`
- Tabs: `Test Results`, `Live Plot`, `Run Log`, `Serial Consoles`

## Core Workflow

1. Connect ports in `Connection`.
2. Enter run metadata in `Test Setup`.
3. Start run from `Actions`.
4. Watch live behavior in `Live Plot` and status in `Run Log`.
5. Review metrics in `Test Results`.
6. Save session and exports.

## Panel Reference

### Connection

- `Refresh Ports`: re-scan available COM ports
- `Connect Selected`: open selected COM port
- `Reconnect @ Baud`: reconnect selected port at current baud
- `Disconnect Selected`: close selected COM port
- `Connect All`: connect all detected ports (max 10)
- `Disconnect All`: close all connected ports
- `Hide/Show Port Station`: collapse or expand card grid

Status row:

- connected count
- connected list

### Port Station View

Each card shows:

- slot
- COM port
- serial (`SN:`)
- state badge

Typical status flow: `CONNECTED` -> `RUNNING` -> `PASS/WARN/FAIL` -> `COMPLETE`.

### Test Setup

| Field | Purpose | Required |
|---|---|---|
| `Operator` | Operator identifier for traceability | Yes |
| `Bath ID` | Optional setup ID | No |
| `Notes` | Run notes | No |
| `Bath Temp (C)` | Environmental metadata | No |
| `Salinity (PSU)` | Environmental metadata | No |
| `Samples` | samples per run (minimum 20) | Yes |
| `Results Root` | top folder where output folders/files are written | Yes |

### Actions

Buttons:

- `Run Test`: starts parallel runs on all connected ports
- `Save JSON`: saves current session rows to session JSON
- `Reset`: starts a new session and clears table
- `Live Plot`: jumps to Live Plot tab
- `Console`: jumps to Serial Consoles tab
- `Detach`: undock/dock console tab
- `Session Plot`: plot current session
- `Load Session`: load a different session JSON and plot
- `Reload JSON`: reload current session JSON for plotting
- `CSV Column`: show/hide `sample_csv` column in results table

Run controls:

- `Runs`: batch count per connected port
- `Delay (s)`: wait time between runs in a batch
- `Dark Mode`: theme toggle
- `Units tested`: unique serial counter (session limit is 10)
- `Runs left`: active batch progress by port

### Test Results

Table includes timestamp, port, serial, noise/voltage stats, flags, and sample CSV path.

Tips:

- click column headers to sort
- use `CSV Column` to show/hide sample CSV path
- inspect `flags` for limit and data-quality issues

### Live Plot

Controls:

- `Plot field`, `Refresh Plot`
- `Auto Y`, `Ymin`, `Ymax`
- `Points`
- `Filter Ports`, `Visible Ports`
- `X Start`, `X End`
- `Std Dev`, `Samples`

Right pane shows recent parsed sample lines during active runs.

### Serial Consoles

Per-port controls:

- `Send`, `Read`, `Stream`

Global console controls:

- `Clear Selected Debug Tab`
- `Export Console CSV`
- `CR`, `LF`
- display mode: `ASCII`, `HEX`, `DEC`, `BIN`

## Generic Sample Format (Parser + Mapping)

Open the `Sample Setup` tab, or click `Config` in the header.

Quick setup:

1. Paste one line into `Example sample`.
2. Set `Delimiter`.
3. Click `Quick Setup + Plot`.

Parser controls:

- `Sample Cmd`: command sent for each sample (default `tsr`)
- `Trim Prefix`: remove fixed prefix before split
- `Start Token`: ignore first N tokens
- `Regex`: extract payload before split

Field editor columns:

- `Field Key`, `Description`
- `Unit`, `Scale` (raw/milli/micro/kilo)
- `Min`, `Max`, `StuckN`
- `Derived Expr`
- `Live`, `Session`, `Default`

Profiles and apply:

- `Save Profile`, `Load Profile`
- `Apply Measureands`
- `Reset Default`

## Session Plot and Comparison

`Session Plot` opens a dedicated plot window.

Use it to:

- choose a numeric metric
- filter current session by serial
- load a reference session JSON
- filter reference serials
- pause/resume redraws

## Output Files and Folder Structure

Default output root for packaged exe:

- `<exe folder>/SBE83`

Within `Results Root`:

- `sessions/PreCalTest/sbe83_session_<session_id>.csv`
- `sessions/PreCalTest/sbe83_session_<session_id>.json`
- `sessions/PreCalTest/profiles/*.json`
- `<serial>/PreCalTest/SBS83_SN<serial>_<timestamp>_samples.csv`
- `<serial>/PreCalTest/SBS83_SN<serial>_<timestamp>.log`
- `<serial>/PreCalTest/SBS83_SN<serial>_<timestamp>_summary.json`

## Severity Logic

- `PASS`: both red/blue noise <= 10 ns
- `WARN`: either noise > 10 ns and <= 20 ns
- `FAIL`: either noise > 20 ns
- `UNKNOWN`: insufficient numeric data

## Troubleshooting by Symptom

### `Run Test` is disabled

- connect at least one port
- fill `Operator`
- make sure no run is currently active

### No samples collected

- verify `Sample Cmd`
- verify baud and cable path
- verify delimiter/regex/start token settings

### Plot looks flat or incorrect

- verify selected `Plot field` exists in parser config
- verify unit/scale and min/max settings
- check `Run Log` for parse or timeout issues

### Console command has no response

- adjust `CR` and `LF`
- use `Read` after `Send`
- confirm the port is connected and idle

## Documentation Map

## Screenshot Guidance (Optional)

If you want visual callouts, add images under `docs/images/` for:

1. main dashboard with Connection + Actions
2. Live Plot during active run
3. Sample Setup tab with Generic Sample Format
4. Session Plot with reference session loaded
5. Serial Console with stream and display mode controls

## Soon To Be Expanded

- Standalone quickstart addendum draft: `docs/QUICKSTART_ONE_PAGE.md`
