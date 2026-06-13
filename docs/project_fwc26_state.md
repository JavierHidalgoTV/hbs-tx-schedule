---
name: project-fwc26-state
description: "Current implementation state of the FWC26 TX Schedule tool — what's done and what's pending"
metadata: 
  node_type: memory
  type: project
  originSessionId: b4f5de29-ca7e-4fe8-9d62-5780b7e9a8d4
---

Working file: `fwc26_tx_schedule_v4.html`. Source data: `FWC26_TX_Input_v3.xlsx`.

**Done — Excel parsing**
- All sheets parsed positionally (by column index, NOT by name) — avoids CRLF/LF issues in shared strings headers
- `sheetData(name)` helper: reads raw arrays, skips banner (row 0) + header (row 1), returns data from row 2+
- Matches: also positional, auto-detects data start by finding first row where col 0 is a positive integer (M#)
- Services/Feeds/Subtasks/Settings all positional with fixed column order (user guarantees order)
- Falls back to `DEMO_SERVICES` if no Services sheet present
- `FWC26_TX_Input_v3.xlsx` has 6 sheets: Services (8 svcs), Feeds (~30), Subtasks (6), Settings, Matches (104), Overrides (empty)
- Overrides sheet is currently ignored

**Done — Sticky nav**
- Entire nav-bar is sticky: title/clock/live-btn row + calendar strip + controls row (View/Zoom/Collapse/Filter) + filter panel (when open)
- All controls are now inside `#nav-bar` so they stay visible on scroll

**Done — Now-line ticker**
- `#now-label` (time label on now-line) pins just below nav-bar when the time axis scrolls off screen, instead of disappearing

**Pending**
- P3: MD-1 tooltips with exact IBC time
- P4: ISO 10/11 overlap handling in compressed views
- P5: Architecture split (single HTML vs multi-file)
- Overrides sheet support (currently empty in v3, but sheet exists)

**Key facts**
- IBC Time = CDT = UTC−5
- All parsing is positional — if user changes column ORDER in Excel, parser breaks. User guarantees order.
- localStorage keys: `fwc26_services`, `fwc26_matches`
- CRLF issue: Excel shared strings store multiline headers as `\r\n` but Python xml parser normalizes to `\n` — SheetJS in browser preserves `\r\n`. Positional parsing sidesteps this entirely.

**Why:** User runs this at IBC for FWC26 broadcast ops.
