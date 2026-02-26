# Quickstart One-Page Addendum

Use this as a fast checklist for new users and time-critical bench runs.

## Quick Navigation

| Need | Go here |
|---|---|
| shortest path to first run | [Fast Path](#fast-path) |
| full setup checklist | [Full One-Page Checklist](#full-one-page-checklist) |
| pass/fail meaning | [Pass/Warn/Fail Rules](#passwarnfail-rules) |
| immediate unblock steps | [If Blocked](#if-blocked) |

## Fast Path

1. Launch `sbs_dsw.exe`.
2. `Refresh Ports`.
3. Connect ports.
4. Fill `Operator`.
5. Set `Samples` (>=20).
6. Click `Run Test`.
7. Review `Test Results`.
8. Click `Save JSON`.

## Full One-Page Checklist

### 1. Launch and Connect

1. Open `sbs_dsw.exe`.
2. Click `Refresh Ports`.
3. Set COM port and baud.
4. Click `Connect Selected` or `Connect All`.

Success check:

- connected ports are listed
- Port Station cards show `CONNECTED`

### 2. Configure Run

1. In `Test Setup`, enter `Operator`.
2. Set `Samples` (minimum 20).
3. Confirm `Results Root`.
4. In `Actions`, set:
- `Runs` (start with `1`)
- `Delay (s)` (start with `0`)

### 3. Execute

1. Click `Run Test`.
2. Watch:
- `Live Plot` for signal behavior
- `Run Log` for progress
- Port statuses (`RUNNING`, then `PASS/WARN/FAIL`)

### 4. Review

1. Open `Test Results`.
2. Check:
- `red_ns`
- `blue_ns`
- voltage std/avg columns
- `flags`

### 5. Save

1. Click `Save JSON`.
2. Optional: `Export Console CSV` in `Serial Consoles`.

## Quick Commands (Optional)

### Run packaged app

```powershell
.\dist\sbs_dsw.exe
```

### Run from source

```powershell
python sbs_dsw.py
```

## Pass/Warn/Fail Rules

- `PASS`: red and blue noise <= 10 ns
- `WARN`: either noise > 10 ns and <= 20 ns
- `FAIL`: either noise > 20 ns

## If Blocked

- `Run Test` disabled: connect a port and fill `Operator`.
- No data: verify baud and `Sample Cmd`.
- Parsing wrong: use `Config` and re-apply parser settings.
