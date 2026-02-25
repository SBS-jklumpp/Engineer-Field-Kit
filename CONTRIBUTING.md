# Contributing to Engineer’s Field Kit – Multitool

Thanks for wanting to contribute. This project is meant to stay **practical**, **engineer-friendly**, and **easy to maintain**.

## Ground rules
- Keep features focused on real engineering workflows.
- Prefer simple, readable code over “clever” code.
- Avoid heavy dependencies unless the value is very clear.
- Preserve the dark UI look/feel and consistency.

## Development setup
1. Create a virtual environment
2. Install requirements
3. Run the app

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m engineers_field_kit_multitool.app
```

## Branching
- `main` is always releasable.
- Use feature branches: `feature/<short-name>` or `fix/<short-name>`.

## Commit messages
Use concise, descriptive messages:
- `Add serial plot smoothing`
- `Fix config load path on Windows`
- `Refactor plot widget for readability`

## Tests
Add tests when:
- you change parsing logic
- you change analysis math
- you add new utilities that can be validated deterministically

Run:
```bash
pytest -q
```

## Pull requests
- Keep PRs small and scoped.
- Include screenshots for UI changes.
- Include example data (redacted) when useful.
