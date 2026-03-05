## What's New in v1.5.0

### Official Rebrand
- Renamed from **Engineers Field Kit Multitool** to **Seabird Digital Sensor Workbench**
- Updated all application references, window titles, and branding

### New Features
- **Instrument Serial Number Display**: Automatically queries and displays the instrument serial number when a COM port is connected
- Background async query prevents UI blocking during connection

### Improvements
- **Help Page Navigation**: Fixed table of contents links - clicking headings now scrolls to the correct section
- **Header UI Refresh**: New logo display (64px), consistent purple theme, improved version badge styling
- **Cleaner Help Rendering**: Removed CDATA artifacts and raw badge markup from help page display

### Build
- Windows executable: sbs_dsw.exe
- Built with PyInstaller
