# FWC26 TX Schedule

Broadcast operations dashboard for the FIFA World Cup 2026 IBC (International Broadcast Centre). Displays a horizontal timeline for each match day showing every match, every broadcast service (Digital Content Live, FCO, RIS, Datatainment, Match Buddy, Automated Content Creation, Closed Captions), and every feed within each service. Includes MD-1 days with team training and press conference times.

The entire application lives in a single HTML file (`fwc26_tx_schedule_v4.html`) with SheetJS inlined. No server, no build step, no dependencies — open it directly in a browser. It can be shared via email, network drive, or USB stick and works fully offline.

Schedule data is loaded from `data/FWC26_TX_Input_v3.xlsx` via the "Load Excel" button in the UI. Parsed data is stored in localStorage and survives page reloads. The HTML ships with real v3 data pre-loaded as a fallback, so it works out of the box without loading any Excel file.

The `scripts/` folder contains Python utilities: `generate_excel.py` builds the input Excel template, `extract_xl_data.py` parses the Excel into JS literals, and `embed_data.py` injects that data directly into the HTML. Detailed technical documentation covering the data model, rendering pipeline, zoom system, tooltip logic, and stack migration notes is in `docs/engineering.md`.

All times are in IBC Time (CDT / UTC−5). The timeline supports 3h and 6h zoom modes with drag-to-scroll and a scroll mirror bar at the bottom of the screen. Dark and light modes are available via the toggle in the nav bar.
