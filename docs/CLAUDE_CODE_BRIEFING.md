# FWC26 Transmission Schedule — Claude Code Briefing

## What this is
A broadcast operations tool for FWC26 (FIFA World Cup 2026) built as a single-file HTML app. Used by operators at IBC (International Broadcast Center) to monitor live transmission schedules, services, and feeds across multiple matches per day.

---

## Current state
- Single file: `fwc26_tx_schedule_v4.html` (~1500 lines, vanilla JS + CSS, no framework)
- Fully functional prototype, data is hardcoded in JS arrays
- All data currently covers MD1–MD8 (10–18 June 2026)

---

## Data hierarchy

```
Match (e.g. MEX vs RSA, Match A, 11 Jun)
└── Service (DCL, DCL AntiPiracy, FCO, RIS, Datatainment, Match Buddy, ACC, Closed Captions)
    └── Feed (Team A, WF, ESF, Fan & Reaction, ISO 1, ISO 6...)
        ├── TX Start offset (min from KO, e.g. -150)
        ├── TX End offset (min from KO, e.g. +65)
        └── Subtask (name + start offset from KO, e.g. "SRT available to MP's" at -45min)

MD-1 (day before match — separate TX chain, NOT part of match hierarchy)
└── MD-1 A / MD-1 B / MD-1 C / MD-1 D  (one per match channel)
    ├── Training TX chain
    │   ├── Team A Training (absolute IBC time from dataset)
    │   └── Team B Training (absolute IBC time from dataset)
    └── Press Conference TX chain
        ├── Team A PC (absolute IBC time from dataset)
        └── Team B PC (absolute IBC time from dataset)
```

---

## Key data structures in JS

### Matches array
```js
{
  m: 1,                    // match number
  md: 1,                   // matchday number
  date: '2026-06-11',      // YYYY-MM-DD
  dateLabel: '11 June 2026',
  ch: 'Match A',           // Match A/B/C/D/E/F
  teamA: 'MEX', teamB: 'RSA',
  fullA: 'Mexico', fullB: 'South Africa',
  koIBC: 14,               // KO hour in IBC Time (CDT = UTC-5), decimal (e.g. 14.5 = 14:30)
  city: 'Mexico City',
  localKO: '13:00', utcKO: '19:00', tz: 'UTC-6',
  dur: 105,                // match duration in minutes (estimated)
  trainA: '12:00 IBC',     // MD-1 Training Team A (absolute IBC time string)
  pcA: '15:30 IBC',        // MD-1 Press Conference Team A
  trainB: '18:00 IBC',     // MD-1 Training Team B
  pcB: '16:30 IBC',        // MD-1 Press Conference Team B
}
```

### Services array
```js
{
  id: 'dcl',
  name: 'Digital Content Live',
  start: -150,   // service start offset from KO in minutes
  stop: 65,      // service stop offset from KO in minutes
  feeds: [
    {
      name: 'Team A',
      start: -150,      // TX Start offset from KO (min)
      stop: 65,         // TX End offset from KO (min)
      subtasks: [
        { name: 'SRT available to MP\'s', offset: -45 }
      ]
    },
    {
      name: 'MD-1 PC',
      individual: true,   // uses absolute time from match.pcA/pcB, not offset
      subtasks: []
    }
  ]
}
```

### Channel colors
```js
CH_COLORS = {
  'Match A': { full:'#1a3a5c', svc:'#122840', feed:'#0d1e30', text:'#7dbfff', pip:'#7dbfff', ko:'rgba(125,191,255,0.45)' },
  'Match B': { full:'#0f3d2c', svc:'#0a2a1e', feed:'#071a12', text:'#5dd4a0', pip:'#5dd4a0', ko:'rgba(93,212,160,0.45)' },
  'Match C': { full:'#2e1f4a', svc:'#1e1432', feed:'#140d22', text:'#b89cff', pip:'#b89cff', ko:'rgba(184,156,255,0.45)' },
  'Match D': { full:'#3d1f10', svc:'#2a1408', feed:'#1a0d04', text:'#ff9e6a', pip:'#ff9e6a', ko:'rgba(255,158,106,0.45)' },
}
```

---

## Features implemented

### Navigation
- **Sticky nav bar** with IBC Time clock (CDT = UTC-5, live updating every second)
- **⬤ Live Now** button — jumps to today's date, centers timeline on current time
- **Calendar strip** (10 Jun – 19 Jul) with dots for match days and MD-1 days
- Click day → show only that day; click again → back to Live Now

### Timeline
- **Dynamic bounds per day** — axis start/end calculated from earliest feed start to latest feed end (KO ± max offsets), with 15min padding
- **Pixels per hour**: 160px/h in 1h mode, fit-to-container in All mode
- **Drag to scroll** horizontally (mouse + touch)
- **Sticky scroll mirror** — fixed scrollbar at bottom of viewport, synced with visible timeline
- **Now line** — white vertical line tracking current IBC time, with HH:MM:SS label above axis

### Hierarchy expansion
- **Match row** → click to expand/collapse services panel
- **Service row** → click to expand/collapse feeds
- **MD-1 row** → same as match row, shows DCL + ACC services with Training/PC feeds

### View modes (toolbar)
- **All** — full hierarchy: match → service → feeds (user collapses individually)
- **Svc** — shows service rows only, feeds hidden
- **Feed** — flat deduped feed list (unique feeds sorted by TX start, no service layer). MD-1 shows Training and PC as chains with one block per team

### Visual details
- **Feed block** (outer) = TX window of the feed
- **Service window** (inner, semitransparent) = when the service actually starts within the feed block. Service start time shown as label at left edge of inner block
- **Subtask markers** = thin vertical lines inside feed block at their offset time. Tooltip on hover shows name + IBC time
- **KO line** = colored vertical line at kickoff time on every row

### Filters (toolbar ⚙)
- **Channel filter** — Match A/B/C/D + MD-1 A/B/C/D chips with channel colors
- **Service filter** — chip per service to show/hide across all matches

### Other controls
- **Zoom**: 1h (160px/h, scrollable) | All (fit to container)
- **Collapse/Expand all** button
- **Light/Dark mode** toggle

---

## What needs to be built next

### Priority 1 — Data architecture (Excel input files)
Replace hardcoded JS arrays with two Excel input files:

**File 1: `config.xlsx`** — tournament configuration (static, changes rarely)
- Sheet `Services`: service name, start offset, stop offset
- Sheet `Feeds`: service, feed name, TX start offset, TX end offset, active_from date (for conditional feeds like ISO 10/11)
- Sheet `Subtasks`: service, feed, subtask name, start offset
- Sheet `Settings`: IBC timezone offset, tournament start date, tournament end date, px_per_hour, match duration estimate

**File 2: `matches.xlsx`** — match data (updated as tournament progresses)
- Sheet `Matches`: M#, MD#, date, channel, team A trigram, team B trigram, full name A, full name B, city, local KO, UTC KO, timezone, IBC KO
- Sheet `MD1_Times`: M#, MD-1 train A (IBC), MD-1 PC A (IBC), MD-1 train B (IBC), MD-1 PC B (IBC)
- Sheet `Overrides`: M#, channel, service, feed, override TX start (IBC), override TX end (IBC), notes

### Priority 2 — Conditional feeds by date
Feeds like ISO 10 Player A and ISO 11 Player B should only appear from a configurable start date (e.g. `13/06/2026`). Field in Feeds sheet: `active_from` (DD/MM/YYYY, blank = always active).

### Priority 3 — MD-1 view improvements
- In Feed view, MD-1 Training and PC chains show one block per team (already implemented)
- Still needed: tooltip on hover of each team block shows exact IBC time

### Priority 4 — ISO 10/11 display in compressed views
When ISO 10 Player A and ISO 11 Player B are both visible in All/Feed mode, they overlap because both are -10/+10min from KO. Consider staggering or grouping them.

### Priority 5 — Architecture migration
Consider splitting the single HTML into:
- `index.html` — layout shell
- `schedule.js` — data loading + rendering
- `config.js` or loaded from Excel via SheetJS
- `style.css`

Or keep as single file but load data from external Excel/JSON files via fetch.

---

## Tech stack
- Vanilla JS, no framework
- CSS custom properties for theming (dark/light)
- `openpyxl` (Python) for generating Excel templates
- SheetJS (`xlsx` npm/CDN) is the recommended library for reading Excel in browser

---

## Files
- `fwc26_tx_schedule_v4.html` — current working prototype
- `FWC26_TX_Input_v2.xlsx` — Excel input template (Structure / Matches / Services / Feeds / Subtasks / Overrides sheets)

---

## Known issues / quirks
- IBC Time is hardcoded as UTC-5 (CDT). During winter it would be UTC-6 (CST) — make configurable
- `koIBC` is stored as decimal hours (e.g. `14` = 14:00, `14.5` = 14:30). Post-midnight KOs are stored as >24 (e.g. `25` = 01:00 next day) to simplify timeline math
- MD-1 date sections are auto-generated as `prevDateKey(match.date)` — the calendar strip now starts on 10 Jun to include the first MD-1 day
- The `services-panel` uses `.open` class for visibility; `feeds-panel` uses `style.display` directly — should be unified
- `applyFilters()` detects MD-1 rows by checking if match-id text starts with "MD-1" — fragile, should use a data attribute
