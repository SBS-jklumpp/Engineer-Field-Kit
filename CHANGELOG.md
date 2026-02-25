# Changelog

All notable changes to this project will be documented here.

## Unreleased
- Initial public scaffolding
- App packaging + baseline UI

## 1.1.0 (2026-02-25)
- Reworked session analysis plotting from baseline-delta mode to two-session comparison with per-session sensor filtering.
- Added hover tooltips and pause/resume controls to session plot windows for easier run-to-run inspection.
- Improved test setup workflow: renamed Station to Notes, removed salt-bath-only wording, relaxed required fields, and added visible results-root path with browse selection.
- Added live batch progress visibility via runs-left counter.
- Expanded serial console controls: CR/LF send options and display modes for ASCII, HEX, DEC, and BIN (showing non-printable bytes explicitly).
- Replaced tab detachment behavior with console-only detach/dock logic and strengthened dock recovery.
- Changed packaged app behavior to run windowed (no terminal console window).
- Updated packaged-run storage behavior so result folders are created beside the executable.
- Updated About metadata author to Justin Klumpp.

## 0.1.0 (2026-02-25)
- Project initialized
