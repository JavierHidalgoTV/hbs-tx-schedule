# FWC26 TX Schedule — Domain Glossary

## Match Duration (`dur`)

Time from KO to Final Whistle, in minutes. Used as the reference point for all TX End offsets.

**Default:** 105 minutes (45′ first half + 15′ half-time + 45′ second half).
**Override:** `Dur (min)` column in Matches Excel sheet — used for knockout rounds (potential extra time ~150 min).

The visual match block on the timeline extends to `KO + dur + 15 min` (a fixed 15-minute post-FW display buffer). The right edge of the block is always dashed + labeled "est." to indicate the duration is estimated.

---

## IBC Time

IBC = International Broadcast Centre (CDT = UTC-5). All times in this tool are expressed in IBC Time unless labelled otherwise. KO times, feed offsets, subtask offsets, and MD-1 absolute times are all in IBC Time.

---

## KO (Kickoff)

The moment the match begins in the field. TX Start offsets and subtask offsets are expressed relative to KO (negative = before KO).

---

## Final Whistle (FW)

The estimated end of the match. `FW = KO + dur`. All TX End offsets are expressed relative to FW.

---

## Feed TX Window

The scheduled transmission window for a specific feed-within-service:

- **TX Start** = minutes from KO (always negative; e.g., `-90` = 90 min before KO)
- **TX End** = minutes from Final Whistle (negative = before FW, positive = after FW)

Examples with default `dur = 105`:

| Feed | Service | TX End (from FW) | Absolute end |
|---|---|---|---|
| WF | DCL | −90 | KO + 15 min |
| WF | FCO / RIS / etc. | +15 | KO + 120 min |
| Team A/B | DCL / ACC | −40 | KO + 65 min |
| Fan & Reaction | any | −45 | KO + 60 min |

Services have the same start/stop reference convention as feeds.

---

## Match Day (MD)

A numbered matchday in the tournament (MD1–MD8 for the group stage). Not a calendar day — multiple matches may share the same MD number and the same calendar date.

---

## MD-1

The day before a match. Used for Training and Press Conference transmissions (absolute IBC times, not KO-relative offsets).
