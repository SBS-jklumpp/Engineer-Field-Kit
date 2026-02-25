# Releasing

## Versioning
Use semantic-ish versioning:
- 0.x while rapidly iterating
- 1.0 once the core UX and file formats stabilize

## Release checklist
- Update `CHANGELOG.md`
- Tag the version (e.g., `v0.1.1`)
- Build with PyInstaller (if distributing binaries)
- Attach binaries to GitHub Releases
