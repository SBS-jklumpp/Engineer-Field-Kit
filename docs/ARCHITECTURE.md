# Architecture

Engineer’s Field Kit – Multitool is structured as a small Python package with a GUI entrypoint.

## Goals
- Fast iteration on “bench utilities”
- Minimal, maintainable dependencies
- Clear separation between UI and analysis logic

## Suggested module boundaries
- `app.py`: GUI entrypoint + wiring
- `styles.py`: UI theme + reusable style helpers
- `config.json`: defaults for ports, logging, plot settings
- `analysis/`: (future) reusable math / stats / parsers
- `io/`: (future) serial, file logging, import/export

## Principles
- Keep parsing deterministic and testable.
- Prefer pure functions for analysis logic.
- Push UI-specific code to the UI layer.
