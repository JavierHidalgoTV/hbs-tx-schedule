# FWC26 Transmission Schedule — Engineering Reference

> Source file: `fwc26_tx_schedule_v4.html`
> Last updated: 2026-06-12

---

## 1. App Overview

**What it does.** A read-only broadcast operations dashboard for the FIFA World Cup 2026 IBC (International Broadcast Centre). It renders a horizontal timeline for each match day, showing:

- Every match on that day (one row per match, grouped by channel)
- Every broadcast service active around each match (Digital Content Live, DCL AntiPiracy, FCO, RIS, Datatainment, Monitoring, Automated Content Creation, Closed Captions)
- Every feed within each service, with optional subtask markers
- The day before each match (MD-1), showing Training and Press Conference times for both teams

**Who uses it.** Broadcast operations staff at the IBC during the tournament. The UI is monitor-sized, density-first, keyboard-free.

**Single-file architecture rationale.** The entire app — SheetJS (minified, ~500 KB inlined), all CSS, all HTML, and all JS — lives in one `.html` file. This is intentional:

- Zero dependencies to install or serve; open locally in a browser
- Can be emailed, shared on a network drive, or loaded from a USB stick
- No build step; edits are live after save+reload
- SheetJS is inlined rather than CDN-loaded to guarantee offline operation

The tradeoff is a large file (~1 MB). All logic is in vanilla JS/CSS; no framework, no bundler.

---

## 2. Data Model

### 2.1 `match` object

```
{
  m:         number,    // match number (M1–M104, unique)
  md:        number,    // match day number (1–34)
  date:      string,    // "YYYY-MM-DD" in local (IBC) time
  dateLabel: string,    // "11 June 2026" (display)
  ch:        string,    // "Match A" | "Match B" | "Match C" | "Match D" | "Match E" | "Match F"
  teamA:     string,    // 3-letter code, e.g. "MEX"
  teamB:     string,
  fullA:     string,    // full country name, e.g. "Mexico"
  fullB:     string,
  city:      string,    // venue city, e.g. "Mexico City"
  koIBC:     number,    // DECIMAL HOURS in IBC time (UTC−5 / CDT), e.g. 14.0 = 14:00 IBC
  dur:       number,    // scheduled match duration in MINUTES (typically 105)
  trainA:    string|null, // Team A training time as "HH:MM IBC" string, e.g. "12:00 IBC"
  pcA:       string|null, // Team A press conference time as "HH:MM IBC" string
  trainB:    string|null,
  pcB:       string|null,
}
```

**Critical type distinction:**

- `match.koIBC` is a **decimal number** (hours). `14.5` means 14:30 IBC. Used in all time arithmetic.
- `match.trainA`, `match.pcA`, `match.trainB`, `match.pcB` are **strings** in `"HH:MM IBC"` format. They must be parsed with `parseIBCTime()` before any arithmetic.
- There is no `match.ko` field. Always use `match.koIBC`.
- There is no `match.venue` field. The field is named `match.city`.

### 2.2 `service` object

```
{
  id:     string,    // slugified stable ID, e.g. "digitalcon_0"
  name:   string,    // display name, e.g. "Digital Content Live"
  start:  number,    // TX start offset in MINUTES relative to KO (negative = before KO)
  stop:   number,    // TX stop offset in MINUTES relative to Final Whistle (FW = KO + dur)
  feeds:  Feed[],
}
```

### 2.3 `feed` object

```
{
  name:       string,       // display name, e.g. "WF", "Team A", "Fan & Reaction"
  start:      number,       // TX start offset in MINUTES relative to KO
  stop:       number,       // TX stop offset in MINUTES relative to FW
  subtasks:   Subtask[],    // array, may be empty
  individual: boolean?,     // if true: this feed is driven by MD-1 times, not KO offset
  activeFrom: string?,      // "DD/MM/YYYY" — feed is skipped for match dates before this
}
```

### 2.4 `subtask` object

```
{
  name:   string,  // display name, e.g. "Frontend Available to FIFA"
  offset: number,  // minutes relative to KO (negative = before KO)
}
```

### 2.5 `settings` object (localStorage key `fwc26_settings`)

A flat key-value dict. The only key consumed by the current JS is `match_duration_min`, `md1_training_dur_min`, and `md1_pc_dur_min` (see Section 12 for full list).

---

## 3. Data Loading Pipeline

### 3.1 localStorage keys

| Key | Content |
|---|---|
| `fwc26_services` | JSON array of service objects |
| `fwc26_matches` | JSON array of match objects |
| `fwc26_settings` | JSON object, key→value |

### 3.2 Startup resolution

```js
function _tryParseLS(key){ try{ const v=localStorage.getItem(key); return v?JSON.parse(v):null; }catch(e){ return null; } }
let SERVICES = _tryParseLS('fwc26_services') || DEMO_SERVICES;
let matches   = _tryParseLS('fwc26_matches')  || DEMO_MATCHES;
```

If either key is absent or unparseable, the corresponding `DEMO_*` constant is used. Settings are loaded separately in an IIFE after the schedule builder constants.

### 3.3 Excel parsing (SheetJS)

The user clicks "Load Excel" to open an overlay. They can drag-and-drop or browse for `.xlsx`, `.xls`, `.csv`, or `.tsv` files. The file is read with `FileReader`:

- CSV/TSV: `XLSX.read(text, {type:'string', raw:true, FS:delim})`
- Binary: `XLSX.read(new Uint8Array(e.target.result), {type:'array'})`

Then `parseWorkbook(wb)` is called.

**Sheet detection:**

- `Settings`: optional, looked up by name `'Settings'`
- `Services`: optional, looked up by name `'Services'`; falls back to `DEMO_SERVICES` if absent
- `Matches`: primary data; looked up by name `'Matches'`, falls back to `wb.SheetNames[0]`

**Row skipping.** Each sheet is parsed with `XLSX.utils.sheet_to_json(ws, {header:1, defval:null})`. The first two rows (index 0 and 1) are always skipped — row 0 is treated as a banner, row 1 as a header. The function `sheetData(name)` encapsulates this: `rows.slice(2).filter(r=>r&&r.some(v=>v!==null&&v!==''))`.

**Matches sheet — positional column mapping (column ORDER matters, not names):**

| Column index | Field |
|---|---|
| 0 | M# (match number, integer) |
| 1 | MD (match day number) |
| 2 | Date ("DD/MM/YYYY" string OR Excel date serial) |
| 3 | Channel ("Match A" … "Match F") |
| 4 | Team A (3-letter code) |
| 5 | Team B (3-letter code) |
| 6 | Full Name A |
| 7 | Full Name B |
| 8 | City |
| 9 | IBC KO time ("HH:MM" string OR Excel time serial 0.0–1.0) |
| 10 | Duration (minutes; falls back to `match_duration_min` setting or 105) |
| 11 | Train A ("HH:MM" / "HH:MM:SS" / Excel serial / "HH:MM IBC") |
| 12 | PC A |
| 13 | Train B |
| 14 | PC B |

`dataStart` is autodetected: the parser scans rows 0–4 for a row whose column 0 is a positive integer.

**Services sheet — positional columns:**

| Column | Field |
|---|---|
| 0 | Service name |
| 1 | TX start offset (min from KO) |
| 2 | TX stop offset (min from FW) |
| 3 | Notes (ignored) |

**Feeds sheet — positional columns:**

| Column | Field |
|---|---|
| 0 | Service name (must match a row in Services) |
| 1 | Feed name |
| 2 | TX start offset (min from KO) |
| 3 | TX stop offset (min from FW) |
| 4 | `individual` flag ("YES" = true) |
| 5 | Active from ("DD/MM/YYYY") |
| 6 | Notes (ignored) |

**Subtasks sheet — positional columns:**

| Column | Field |
|---|---|
| 0 | Service name |
| 1 | Feed name |
| 2 | Subtask name |
| 3 | Offset (min from KO) |
| 4 | Notes (ignored) |

**Settings sheet — positional columns:**

| Column | Field |
|---|---|
| 0 | Key (string) |
| 1 | Value |
| 2 | Description (ignored) |

### 3.4 Apply and reload

After parsing, `_parsedXlData` holds `{services, matches, settings}`. Clicking "Apply" calls `applyLoadedData()`:

```js
localStorage.setItem('fwc26_services',  JSON.stringify(_parsedXlData.services));
localStorage.setItem('fwc26_matches',   JSON.stringify(_parsedXlData.matches));
localStorage.setItem('fwc26_settings',  JSON.stringify(_parsedXlData.settings||{}));
window.location.reload();
```

The page fully reloads so that `SERVICES` and `matches` are re-initialized from the new localStorage values. "Use Demo" calls `useDemoData()` which clears all three keys then reloads.

### 3.5 DEMO data

`DEMO_SERVICES` (8 services) and `DEMO_MATCHES` (104 matches, MD1–MD34) are hardcoded in the file as real v3 FWC26 schedule data. They are the fallback when no Excel file has been loaded.

---

## 4. Rendering Pipeline

### 4.1 Day enumeration

All match dates and their MD-1 predecessor dates are collected into `allDatesMap` (a `Map<string, DayRecord>`). Days are sorted chronologically and iterated:

```js
const allDays = [...allDatesMap.values()].sort((a,b)=>a.date.localeCompare(b.date));
```

For each day:

- `dm` = matches on that day (sorted by channel A→F)
- `md1Matches` = matches whose date is the day after this day (i.e., this is MD-1 for those matches)

### 4.2 `calcDayBounds(dayMatches, md1Matches)` → `{startH, endH, span}`

Computes the timeline window for the day:

1. For each match in `dayMatches` ∪ `md1Matches`, compute `earliest = koIBC − MAX_EARLY_MIN/60` (150 min = 2.5 h before KO) and `latest = koIBC + (dur + max(MAX_LATE_MIN=65, POST_FW_DISPLAY_MIN=15))/60 + 0.25`.
2. For each `md1Matches` entry, also include `trainA/pcA/trainB/pcB` parsed times (extending window to `h + 1.5` hours).
3. Floor start to nearest 30 min with 15 min padding; cap end at 26 (allows post-midnight display); clamp start ≥ 0.
4. If no matches at all, fall back to `GLOBAL_START_H=10` to `GLOBAL_END_H=24`.

### 4.3 `makePct(startH, span)` and `pct(h)`

Returns a closure:

```js
function makePct(startH, span){ return h => ((h - startH) / span) * 100; }
```

Used throughout HTML generation to convert a decimal-hour time into a CSS `left` percentage within the track. Example: if `startH=8, span=16`, then `pct(14.0)` returns `37.5` (%).

### 4.4 `makeTicks(startH, endH)` → `number[]`

Returns hourly tick positions. If `span > 10`, step is 2 (every 2 hours), otherwise 1. First tick is aligned to the step. Used to generate `.gl` (grid line) divs and `.time-tick` labels.

### 4.5 Initial track width (build time)

At build time a constant `PX_PER_HOUR = 80` is used to compute `trackW = Math.round(SP * PX_PER_HOUR)`. This is set as `--track-w` on both the `.timeline-container` and the `.time-axis` via `style`. The zoom system overrides this at runtime.

### 4.6 DOM structure per day section

```
.day-section[data-date, data-startH, data-span]
  .day-sticky
    .day-header
  .time-axis[--track-w]
    .time-axis-label
    .time-axis-track-clip
      .time-axis-track
        .time-tick (×N)
  .timeline-container[--track-w, --tc-max-w]
    .match-group (×N)
      .match-row  ← click to expand
        .row-label
        .track
          .grid-overlay  (.gl×N, .now-line)
          .ko-line
          .match-block
      .services-panel
        .svc-group (×N per service)
          .service-row  ← click to expand
            .svc-label
            .svc-track
              .grid-overlay
              .ko-line
              .svc-block
          .svc-feeds-body
            .feed-row (×N)
              .feed-label
              .feed-track
                .grid-overlay
                .ko-line
                .feed-block
                  .subtask-marker (×N)
      .feeds-panel (view mode: feed)
```

For MD-1 days the same `.match-group` structure is used, but the `.match-row` has class `md1-row` and the `.track` contains `match-block` elements per activity (one block per Training/PC event, not a single wide block — see Section 5).

---

## 5. MD-1 Logic

### 5.1 What MD-1 days are

For every match, the day before that match's date is a "MD-1 day" (Match Day Minus 1). That day holds the team trainings and press conferences. `allDaysMap` includes MD-1 entries with `isMatchDay: false`.

MD-1 rows appear in the `.timeline-container` of the MD-1 day section, mixed with any other matches on that same day that happen to share it.

### 5.2 `md1Matches`

For a given day, `md1Matches` is:

```js
const md1Matches = sortByChannel(matches.filter(m => prevDateKey(m.date) === day.date));
```

These are matches whose match date is `day.date + 1`.

### 5.3 MD-1 first-level blocks (`.track`)

The top-level `.match-row.md1-row` track renders **one block per activity** (not one wide block spanning the whole MD-1 window). Specifically, `md1Feeds` contains up to four entries:

```
{ name: `${teamA} Training`, startH: trainAH, dur: MD1_TRAIN_DUR }
{ name: `${teamA} PC`,       startH: pcAH,    dur: MD1_PC_DUR   }
{ name: `${teamB} Training`, startH: trainBH, dur: MD1_TRAIN_DUR }
{ name: `${teamB} PC`,       startH: pcBH,    dur: MD1_PC_DUR   }
```

Entries where `startH === null` are filtered out. Each is rendered as a separate `.match-block` div positioned individually within the `.track`.

### 5.4 `MD1_TRAIN_DUR` / `MD1_PC_DUR`

Declared as module-level `let` variables, initialized to defaults, then overridden from `fwc26_settings` localStorage:

```js
let MD1_TRAIN_DUR = 15/60;  // default: 15 min = 0.25 h
let MD1_PC_DUR    = 20/60;  // default: 20 min = 0.333 h
```

Override logic reads `settings['md1_training_dur_min']` and `settings['md1_pc_dur_min']` from localStorage.

### 5.5 MD-1 services panel

When a MD-1 row is expanded, `md1ServicesHtml` shows two fixed services: "Digital Content Live" and "Automated Content Creation (WSC)". Each service's `.svc-track` shows the same `md1Feeds` entries as individual `svc-block.md1-svc-block` divs (class `md1-svc-block`).

The feeds panel (`md1DedupFeedsHtml`) shows "Training" and "Press Conference" chains — two rows, each spanning from the earliest of TeamA/TeamB event to the latest, with subtask markers at each team's individual time.

---

## 6. Sticky Label Mechanism

### 6.1 Why CSS `position:sticky` fails

Each row (`.match-row`, `.service-row`, `.feed-row`) is a flex container. Its label (`.row-label`, `.svc-label`, `.feed-label`) is 240px wide, and the track occupies the remaining width. The track's width is set by `--track-w` (e.g., several thousand pixels when zoomed in), making the containing `.timeline-container` horizontally scrollable with `overflow-x:auto`.

CSS `position:sticky; left:0` is supposed to pin the label to the left edge of the scroll container. However, it does not work reliably in flex containers where the item's natural width prevents the sticky constraint from taking effect. Different browser engines handle this inconsistently, and the label visually scrolls away with the content.

### 6.2 JS-based `translateX(scrollLeft)` approach

The solution is to simulate sticky positioning in JavaScript. On every scroll event, each label is translated rightward by the scroll offset:

```js
function applyLabelSticky(tc){
  const sl = tc.scrollLeft;
  tc.querySelectorAll('.row-label,.svc-label,.feed-label').forEach(el => {
    el.style.transform = `translateX(${sl}px)`;
  });
  // Also sync the time-axis label and track
  const axisLabel = sec.querySelector('.time-axis-label');
  const axisTrack = sec.querySelector('.time-axis-track');
  if(axisLabel) axisLabel.style.transform = `translateX(${sl}px)`;
  if(axisTrack) axisTrack.style.transform = `translateX(-${sl}px)`;
}
```

Scroll events do not bubble, so capture phase is used:

```js
document.addEventListener('scroll', e => {
  if(e.target.classList?.contains('timeline-container')) applyLabelSticky(e.target);
}, {capture:true, passive:true});
```

`syncAllStickyLabels()` iterates all visible timeline containers and calls `applyLabelSticky` on each. It is called after programmatic `scrollLeft` changes (zoom, goLive).

### 6.3 Why `align-self:stretch` + `border-bottom` on labels matters

Labels use `align-self:stretch` so they grow to the full row height (which may vary). The `border-bottom:1px solid var(--border2)` is applied to the label itself so it always spans the full row height, even when the label has `height: auto`. This is essential for visual row separation — see Section 7.

---

## 7. Row Border Rendering

### 7.1 The compositing layer problem

All label elements (`.row-label`, `.svc-label`, `.feed-label`) have `transform: translateX(...)` applied to them by the sticky mechanism. Applying `transform` promotes an element to its own compositing layer. A composited element is painted independently and composited on top of adjacent content, which means its `border-bottom` or `border-right` may be clipped, misaligned, or rendered at 1px off depending on subpixel rounding.

If borders were applied only to label elements that are in a `transform` layer, they would not align perfectly with borders on track elements that are in the normal layer.

### 7.2 Border placement strategy

Row borders are applied to the track elements (`.track`, `.svc-track`, `.feed-track`), which are always in the normal flow and always visible in the viewport (they form the scrollable content). These elements use:

```css
.track      { border-bottom: 1px solid var(--border2); box-sizing: border-box; }
.svc-track  { border-bottom: 1px solid var(--border2); box-sizing: border-box; }
.feed-track { border-bottom: 1px solid var(--border2); box-sizing: border-box; }
```

`box-sizing: border-box` ensures the border is included in the declared height, preventing height mismatch between the label column (which has `align-self:stretch`) and the track column.

### 7.3 `--border` vs `--border2`

```css
/* Dark mode */
--border:  rgba(255,255,255,0.28);   /* lighter, used for structural separators */
--border2: rgba(255,255,255,0.55);   /* stronger, used for row-level grid lines */

/* Light mode */
--border:  rgba(10,30,80,0.09);
--border2: rgba(10,30,80,0.16);
```

`--border` is used for the outer timeline container, service body grouping, and grid overlays. `--border2` is used for per-row bottom borders within the scrollable area (the denser visual grid).

---

## 8. Zoom System

### 8.1 State

```js
let currentZoom = 'all';  // 'all' | '3h' | '6h'
const PX_PER_HOUR_PRESETS = {'1h':null,'3h':null,'6h':null,'all':null};  // placeholder
```

Note: the preset dict is declared but not populated; the actual logic is inline in `applyZoom`.

### 8.2 `setZoom(mode, btn)`

Called from the "3h" / "6h" / "All" buttons. Sets `currentZoom`, deactivates all zoom buttons, activates `btn`, then calls `applyZoom()`.

### 8.3 `zoomWindowH()` → `number | null`

```js
function zoomWindowH(){
  return currentZoom==='1h' ? 2 : currentZoom==='3h' ? 3 : currentZoom==='6h' ? 6 : null;
}
```

Returns the visible time window in hours, or `null` for "All" (fit-to-width) mode.

### 8.4 `applyZoom()`

For each visible `.timeline-container`:

- If `windowH === null` (All mode): `trackW = max(containerWidth - 240, 300)`. Scroll position reset to 0.
- If `windowH` is 3 or 6: `pxh = max(60, round(visibleTrackW / windowH))`. Then `trackW = round(span * pxh)`. Scroll is centered on current IBC time (or start of day if outside range).

`--track-w` CSS variable is set on both the `.timeline-container` and the corresponding `.time-axis` element. `--tc-max-w` is set to `none` in all cases.

After zoom, `updateMirror()` and `syncAllStickyLabels()` are called via `setTimeout`.

### 8.5 Scroll mirror (`#scroll-mirror-wrap`)

A fixed 14px bar at the bottom of the viewport. Contains `#scroll-mirror-inner` whose width is programmatically set to match the active timeline container's total scrollable width (`240 + trackW`).

`updateMirror()` finds the `.timeline-container` that is most visible in the viewport (by pixel overlap) and listens to its scroll events. Scrolling the mirror scrolls the active TC; scrolling the TC syncs the mirror. A `syncingFromMirror`/`syncingFromTC` flag prevents scroll event loops.

### 8.6 Drag-to-scroll

Implemented with `mousedown`/`mousemove`/`mouseup` on `document`. Active only on elements matching `DRAG_TARGETS = '.track,.svc-track,.feed-track,.grid-overlay,.time-axis-track,.gl,.now-line,.ko-line,.match-block,.svc-block,.feed-block,.time-axis'`. Drag threshold is 4px. A `moved` flag (cleared after 50 ms) blocks the subsequent `click` event to prevent accidental row expansion.

### 8.7 Touch scroll

Separate `touchstart`/`touchmove`/`touchend` handlers mirror the drag logic for touch devices.

---

## 9. Tooltip System

### 9.1 Architecture

A single `#tooltip` div (position:fixed, z-index:9999). `showTooltip(html)` sets its `innerHTML` and `display:block`. `hideTooltip()` sets `display:none`.

Position is updated on `mousemove`: tooltip follows the cursor, clamped to viewport edges.

Event delegation is used (single `mouseover` listener on `document`) rather than per-element listeners. The `mouseout` handler hides the tooltip when leaving the target unless the relatedTarget is inside the same element.

### 9.2 Priority order of target checks (`mouseover` handler)

1. `.subtask-marker` — shows subtask name, IBC time, offset-from-KO, feed name
2. `.md1-block` — shows MD-1 activity label, start/end IBC, duration, match teams, date, KO
3. `.feed-block:not(.md1-block)` — shows feed name, TX start/end IBC, service start/end, offset strings; optionally MD-1 section
4. `.svc-block:not(.md1-svc-block)` — shows service name, TX start/end IBC, offset strings
5. `.match-row.md1-row` — shows "MD-1 A" label, teams, channel, city, Training/PC times for both teams, match date, KO
6. `.match-row:not(.md1-row)` — shows match teams, channel, city, KO IBC, duration

### 9.3 Tooltip CSS classes

| Class | Role |
|---|---|
| `.tt-head` | Bold header line (match name, service name, etc.) |
| `.tt-row` | Key-value row: `display:flex; justify-content:space-between` |
| `.tt-k` | Label/key text (`color:#c0d0ec`) |
| `.tt-v` | Value text (`color:#ffffff`) |
| `.tt-div` | Horizontal rule separator |
| `.tt-ch-dot` | Small colored square matching the channel color |

---

## 10. Time Helpers

### `hhmm(h)` → `"HH:MM"` string

Converts decimal hours to `"HH:MM"` **without wrapping at 24**. So `25.5` → `"25:30"`. Supports negatives: `-0.5` → `"-00:30"`. Used for raw time display and tooltip values where broadcast context requires showing times past midnight as-is.

```js
function hhmm(h){
  if(h<0){ return '-'+hhmm(-h); }
  const hh=Math.floor(h), mm=Math.round((h-hh)*60);
  return String(hh).padStart(2,'0')+':'+String(mm).padStart(2,'0');
}
```

### `hhmmLabel(h)` → `"HH:MM"` or `"HH:MM +1"` string

Used for **end-time display** labels. Wraps at 24: if `h >= 24`, computes `w = h % 24` and appends `" +1"` to indicate next calendar day. Returns `"—"` for null/NaN. Used in tooltips and block labels for stop times that may cross midnight.

```js
function hhmmLabel(h){
  if(h===null||h===undefined||isNaN(h)) return '—';
  if(h>=24){ const w=h%24; return hhmm(w)+' +1'; }
  return hhmm(h);
}
```

### `parseIBCTime(s)` → `number | null`

Parses `"HH:MM IBC"` strings (and plain `"HH:MM"`) to decimal hours. Regex: `/(\d+):(\d+)/`.

### `hhmmToDecimal(v)` → `number | null`

Used inside `parseWorkbook`. Handles two input formats:
- Excel time serial (number 0–1, fraction of day): multiplied by 24
- String `"HH:MM"` or `"HH:MM:SS"`: parsed with regex

### `toIBCTimeStr(v)` → `"HH:MM IBC" | null`

Converts any Excel time value (serial, "HH:MM:SS", "HH:MM") to the canonical `"HH:MM IBC"` string stored in match objects. Strips seconds from `HH:MM:SS`. Returns existing value unchanged if it already ends with `"IBC"`.

### `getIBCDateKey()` → `"YYYY-MM-DD"`

Computes the current IBC date (UTC−5) from the browser clock:

```js
const ibcMs = d.getTime() - 5*3600000;
const ibcDate = new Date(ibcMs);
```

**IBC = CDT = UTC−5.** All times in the app are in IBC/CDT. The nav bar label reads "IBC Time (CDT)".

---

## 11. Color System

### `CH_COLORS` map

```js
const CH_COLORS = {
  'Match A': { full:'#1a3a5c', svc:'#122840', feed:'#0d1e30', svcDim:'#1a3a5c99', feedDim:'#1a3a5c77', text:'#7dbfff', pip:'#7dbfff', ko:'rgba(125,191,255,0.45)' },
  'Match B': { full:'#0f3d2c', svc:'#0a2a1e', feed:'#071a12', svcDim:'#0f3d2c99', feedDim:'#0f3d2c77', text:'#5dd4a0', pip:'#5dd4a0', ko:'rgba(93,212,160,0.45)'  },
  'Match C': { full:'#2e1f4a', svc:'#1e1432', feed:'#140d22', svcDim:'#2e1f4a99', feedDim:'#2e1f4a77', text:'#b89cff', pip:'#b89cff', ko:'rgba(184,156,255,0.45)' },
  'Match D': { full:'#3d1f10', svc:'#2a1408', feed:'#1a0d04', svcDim:'#3d1f1099', feedDim:'#3d1f1077', text:'#ff9e6a', pip:'#ff9e6a', ko:'rgba(255,158,106,0.45)' },
  'Match E': { full:'#1a3d38', svc:'#122a26', feed:'#0d1a18', svcDim:'#1a3d3899', feedDim:'#1a3d3877', text:'#5de0cf', pip:'#5de0cf', ko:'rgba(93,224,207,0.45)'  },
  'Match F': { full:'#3d1a2e', svc:'#2a1020', feed:'#1a0814', svcDim:'#3d1a2e99', feedDim:'#3d1a2e77', text:'#ff7abf', pip:'#ff7abf', ko:'rgba(255,122,191,0.45)' },
};
```

| Field | Used on |
|---|---|
| `full` | Match-level blocks (`.match-block`), feed-level blocks (`.feed-block`), MD-1 service blocks |
| `svc` | Service-level blocks (`.svc-block`) |
| `feed` | (reserved/available; currently `full` used on feed-blocks) |
| `svcDim` | Dimmed version with 60% alpha — not currently used in rendering |
| `feedDim` | Dimmed version with 47% alpha — not currently used in rendering |
| `text` | Text color inside blocks, label accents |
| `pip` | Color dots/stripes in labels, subtask markers, toggle arrows |
| `ko` | KO vertical line color (`.ko-line`) |

### Dark/Light mode

The `CH_COLORS` values are hardcoded for dark mode. In light mode, because inline `style` attributes (generated from CH_COLORS) take precedence over class rules, the light mode CSS overrides use `!important`:

```css
body.light-mode .row-label  { background:var(--s1)  !important; }
body.light-mode .svc-label  { background:var(--s2)  !important; }
body.light-mode .feed-label { background:var(--s3)  !important; }
```

The channel block colors are left as-is in light mode (the deep navy/teal/purple palettes are still readable on the light backgrounds).

---

## 12. Settings Sheet

The `Settings` sheet (columns: 0=Key, 1=Value, 2=Description/ignored) is optional. Currently recognized keys:

| Key | Type | Default | Consumed by |
|---|---|---|---|
| `match_duration_min` | number | 105 | `parseWorkbook`: fallback duration for match rows without column 10 value |
| `md1_training_dur_min` | number | 15 | `MD1_TRAIN_DUR` IIFE after schedule constants |
| `md1_pc_dur_min` | number | 20 | `MD1_PC_DUR` IIFE after schedule constants |

Keys `ibc_offset`, `tour_start`, `tour_end`, and `px_per_hour` are **not currently consumed** by the JS, though they may appear in Excel templates as documentation/future hooks. The app hardcodes IBC = UTC−5, tour range 2026-06-10 to 2026-07-19, and `PX_PER_HOUR=80` at build time.

Settings are stored in `fwc26_settings` localStorage as a flat object. They are read once at page load in the MD1 duration IIFE; a page reload is required for settings changes to take effect.

---

## 13. Known Gotchas / Non-Obvious Decisions

**`match.koIBC` not `match.ko`.** The field is specifically named `koIBC` to make the timezone context explicit. There is no `ko` field. All code that references the KO time must use `match.koIBC`.

**`match.city` not `match.venue`.** The venue location is stored in `city` (e.g., `"Los Angeles, CA"`, `"New York, NJ"`). There is no `venue` field.

**Positional Excel parsing — column ORDER matters, not names.** The parser uses `XLSX.utils.sheet_to_json(ws, {header:1})` which gives plain arrays (no column-name keys). Column positions are fixed (see Section 3). If the user adds or reorders columns in their Excel file, the parser will silently misread data. The only defense is the `koIBC===null` guard (rows with no parseable KO are skipped).

**CRLF / shared-strings issue.** `XLSX.read` with `{type:'array'}` handles CRLF in shared strings cells automatically in recent SheetJS versions. However, if a string cell value has a trailing `\r`, the `str(v).trim()` helper removes it. For CSV input, `{type:'string'}` is used which also trims internally.

**`box-sizing:border-box` on track elements.** `.track`, `.svc-track`, `.feed-track` all have `box-sizing:border-box` and a `border-bottom:1px solid`. Without this, the declared height (e.g., `height:40px`) would not include the border, causing the track to be 1px taller than the adjacent label. This 1px gap would accumulate down the column.

**`.subtask-marker` with `will-change:transform`.** Subtask markers are 2px-wide absolutely-positioned divs. At sub-pixel positions, a 2px div can render as 1px on non-retina displays due to rounding. `will-change:transform` promotes the element to its own compositor layer, ensuring it is always rendered at exactly 2 device pixels regardless of its fractional CSS `left` position.

**MD-1 blocks in match row track are individual, not one wide block.** On MD-1 days the top-level `.track` for each match channel row contains multiple `.match-block` elements (one per Training/PC activity), each individually positioned by `pct()`. There is no single wide "MD-1" block. This allows the scheduler to see gaps between activities.

**`individual: true` feeds.** Feeds marked `individual:true` in services are NOT rendered with a block in the match-day view. Instead they show an arrow `"→ see MD-1 day"` label at the KO line position. The actual rendering appears in the MD-1 day's service panel.

**`activeFrom` date format is `"DD/MM/YYYY"`.** When comparing `f.activeFrom` against `match.date` (which is `"YYYY-MM-DD"`), the code converts `activeFrom` inline:
```js
const[dd,mm,yyyy]=f.activeFrom.split('/');
if(match.date < `${yyyy}-${mm.padStart(2,'0')}-${dd.padStart(2,'0')}`) return;
```

**Light mode forces `!important` overrides on labels.** Because block background colors are applied as inline `style` attributes (generated from CH_COLORS), they have higher specificity than class selectors. Light mode label overrides therefore use `!important` on `.row-label`, `.svc-label`, `.feed-label`.

**`time-axis-track` translates in the negative direction.** When the timeline-container scrolls right by `sl` pixels, the time-axis label translates by `+sl` (sticking to left) and the time-axis track itself translates by `-sl` (keeping tick labels aligned with their positions in the scrolled track below).

---

## 14. Stack Migration Notes

The following notes cover what would change when migrating to React, Vue, or a similar component framework.

**SheetJS can stay.** `parseWorkbook` is a pure function: `(wb) → {services, matches, settings}`. It can be extracted as-is into a utility module with zero framework dependencies.

**Rendering maps cleanly to components.** The DOM hierarchy translates directly:

| Current element | Proposed component |
|---|---|
| `day-section` iteration | `<DaySection day={day} matches={dm} md1Matches={md1Matches} />` |
| `match-group` | `<MatchGroup match={match} services={services} />` |
| `service-row` + feeds | `<ServiceRow service={svc} match={match} pct={pct} />` |
| `feed-row` | `<FeedRow feed={f} match={match} pct={pct} />` |

**`pct` function.** `makePct(startH, span)` returns a closure. In React this would be a `useMemo`-derived function passed via prop or context.

**localStorage strategy.** The current on-load resolution (`_tryParseLS || DEMO_*`) maps cleanly to a Zustand/Jotai store or React context initialized at app root. The demo fallback logic remains identical.

**CSS variables remain.** `--track-w`, `--nav-h`, `--border`, `--border2`, `--accent`, etc. all live in `:root`/`body.light-mode` and are set programmatically with `setProperty`. This approach works identically in any framework; no change needed.

**The sticky-label hack would be replaced.** In a virtualized table (react-virtualized, TanStack Virtual, etc.), the label column is a separate sticky column, not a flex sibling of the scrollable track. The `translateX(scrollLeft)` approach would be replaced by native CSS sticky on the label column of the virtual table, which works correctly when the table has its own independent scroll container with `overflow:auto`.

**Tooltip via event delegation → React synthetic events or Floating UI.** The current `document.addEventListener('mouseover', ...)` delegation pattern would become per-element `onMouseEnter`/`onMouseLeave` props in JSX, or be replaced with a library like Floating UI for positioning.

**Drag-to-scroll → library or pointer events API.** The current implementation can be extracted into a React hook `useDragScroll(ref)` with minimal changes.

**The `will-change:transform` on subtask markers should be kept** regardless of framework, as it addresses a genuine sub-pixel rendering issue with 2px-wide absolutely-positioned elements.
