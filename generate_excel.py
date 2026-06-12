"""Generate FWC26_TX_Input_v3.xlsx with all 28 matches, active_from column, and full team names."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

wb = openpyxl.Workbook()
wb.remove(wb.active)

H_FILL = PatternFill("solid", fgColor="1E2235")
H_FONT = Font(bold=True, color="DDE1ED", name="Calibri", size=10)
D_FONT  = Font(name="Calibri", size=10)
AF_FILL = PatternFill("solid", fgColor="1A2E0A")
AF_FONT = Font(color="8BC34A", name="Calibri", size=10, bold=True)

def style_header(ws, row=2):
    for cell in ws[row]:
        if cell.value:
            cell.font = H_FONT
            cell.fill = H_FILL
            cell.alignment = Alignment(wrap_text=True, vertical="center")

def set_widths(ws, widths):
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

# ─── Services ────────────────────────────────────────────────────────────────
ws = wb.create_sheet("Services")
ws.append(["SERVICES — Start/Stop offsets in minutes from KO. Applied to all matches unless overridden."])
ws.append(["Service Name", "Start Offset\n(min from KO)", "Stop Offset\n(min from KO)", "Notes"])
style_header(ws)
for row in [
    ("Digital Content Live",             -150, 65,  None),
    ("DCL AntiPiracy",                    -90, 15,  None),
    ("FCO",                               -45, 15,  None),
    ("RIS",                               -45, 60,  None),
    ("Datatainment",                      -45, 15,  None),
    ("Match Buddy",                       -90, 15,  None),
    ("Automated Content Creation (WSC)", -150, 65,  None),
    ("Closed Captions",                   -90, 15,  None),
]:
    ws.append(list(row))
set_widths(ws, [36, 14, 14, 30])

# ─── Feeds ───────────────────────────────────────────────────────────────────
ws = wb.create_sheet("Feeds")
ws.append(["FEEDS — TX Start/Stop offsets in minutes from KO.  Individual=YES → uses absolute MD-1 time.  active_from (DD/MM/YYYY) = blank means always shown."])
ws.append(["Service", "Feed Name", "TX Start\n(min from KO)", "TX End\n(min from KO)", "Individual\nSchedule", "active_from\n(DD/MM/YYYY)", "Notes"])
style_header(ws)
feeds = [
    ("Digital Content Live", "Team A",        -150, 65,  None, None, None),
    ("Digital Content Live", "Team B",        -150, 65,  None, None, None),
    ("Digital Content Live", "Fan & Reaction",-120, 60,  None, None, None),
    ("Digital Content Live", "WF",             -90, 15,  None, None, None),
    ("Digital Content Live", "ESF",            -90, 15,  None, None, None),
    ("Digital Content Live", "MD-1 PC",        None,None,"YES",None, None),
    ("Digital Content Live", "MD-1 Training",  None,None,"YES",None, None),
    ("DCL AntiPiracy",       "ESF",            -90, 15,  None, None, None),
    ("DCL AntiPiracy",       "WF",             -90, 15,  None, None, None),
    ("FCO",                  "WF",             -90, 15,  None, None, None),
    ("RIS",                  "WF",             -90, 15,  None, None, None),
    ("RIS",                  "Fan & Reaction",-120, 60,  None, None, None),
    ("Datatainment",         "Datatainment",   -90, 15,  None, None, None),
    ("Match Buddy",          "MP UNI",         -60, 15,  None, None, None),
    ("Match Buddy",          "WF",             -90, 15,  None, None, None),
    ("Match Buddy",          "ISO 9 Tactical Main", -60, 20, None, None, None),
    ("Match Buddy",          "ISO 10 Player A",     -10, 10, None, "13/06/2026", "Active from MD3 onwards"),
    ("Match Buddy",          "ISO 11 Player B",     -10, 10, None, "13/06/2026", "Active from MD3 onwards"),
    ("Match Buddy",          "Fan & Reaction",-120, 60,  None, None, None),
    ("Match Buddy",          "Team A",        -150, 65,  None, None, None),
    ("Match Buddy",          "Team B",        -150, 65,  None, None, None),
    ("Match Buddy",          "ISO 1",          -30, 20,  None, None, None),
    ("Match Buddy",          "ISO 6 Cable Cam",-70, 20,  None, None, None),
    ("Automated Content Creation (WSC)", "Team A",        -150, 65,  None, None, None),
    ("Automated Content Creation (WSC)", "Team B",        -150, 65,  None, None, None),
    ("Automated Content Creation (WSC)", "Fan & Reaction",-120, 60,  None, None, None),
    ("Automated Content Creation (WSC)", "WF",             -90, 15,  None, None, None),
    ("Automated Content Creation (WSC)", "MD-1 PC",        None,None,"YES",None, None),
    ("Automated Content Creation (WSC)", "MD-1 Training",  None,None,"YES",None, None),
    ("Closed Captions",      "WF",             -90, 15,  None, None, None),
]
for f in feeds:
    r = ws.append(list(f))
# Highlight active_from column (col 6) for rows with a value
for row in ws.iter_rows(min_row=3):
    if row[5].value:
        row[5].fill = AF_FILL
        row[5].font = AF_FONT
set_widths(ws, [36, 22, 14, 10, 12, 14, 28])

# ─── Subtasks ────────────────────────────────────────────────────────────────
ws = wb.create_sheet("Subtasks")
ws.append(["SUBTASKS — One row per subtask per feed. Start Offset = minutes from KO (negative = before KO)."])
ws.append(["Service", "Feed", "Subtask Name", "Start Offset\n(min from KO)", "Notes"])
style_header(ws)
for row in [
    ("FCO",  "WF",             "Frontend Available to FIFA",  -15, None),
    ("FCO",  "WF",             "SRT stream available",        -45, None),
    ("RIS",  "WF",             "WebRTC available to MP's",    -15, None),
    ("RIS",  "WF",             "SRT available to MP's",       -45, None),
    ("RIS",  "Fan & Reaction", "WebRTC available to MP's",    -60, None),
    ("RIS",  "Fan & Reaction", "SRT available to MP's",       -45, None),
]:
    ws.append(list(row))
set_widths(ws, [36, 22, 32, 14, 30])

# ─── Settings ────────────────────────────────────────────────────────────────
ws = wb.create_sheet("Settings")
ws.append(["SETTINGS — App configuration. Do not change Key names."])
ws.append(["Key", "Value", "Description"])
style_header(ws)
for row in [
    ("ibc_offset",         -5,            "IBC Time offset from UTC. CDT=-5. Change to -6 for CST."),
    ("tour_start",         "10/06/2026",  "Calendar start date (DD/MM/YYYY)"),
    ("tour_end",           "19/07/2026",  "Calendar end date (DD/MM/YYYY)"),
    ("match_duration_min", 105,           "Default estimated match duration in minutes"),
    ("md1_train_dur_min",  90,            "MD-1 Training block duration estimate (minutes)"),
    ("md1_pc_dur_min",     45,            "MD-1 Press Conference block duration (minutes)"),
    ("px_per_hour",        160,           "Pixels per hour in 1h zoom mode"),
]:
    ws.append(list(row))
set_widths(ws, [22, 14, 50])

# ─── Matches ─────────────────────────────────────────────────────────────────
ws = wb.create_sheet("Matches")
ws.append(["MATCHES — One row per match. IBC KO = kickoff in IBC Time (CDT=UTC-5). MD-1 times are absolute IBC times (HH:MM)."])
ws.append([
    "M#", "MD#", "Date\n(DD/MM/YYYY)", "Channel\n(Match A-F)",
    "Team A\n(trigram)", "Team B\n(trigram)", "Full Name A", "Full Name B",
    "City", "Local KO\n(HH:MM)", "UTC KO\n(HH:MM)", "Time Zone",
    "IBC KO\n(HH:MM)", "MD-1\nTrain A (IBC)", "MD-1\nPC A (IBC)",
    "MD-1\nTrain B (IBC)", "MD-1\nPC B (IBC)",
])
style_header(ws)
matches = [
    # MD1 — 11 Jun
    ( 1,1,"11/06/2026","Match A","MEX","RSA","Mexico",            "South Africa",            "Mexico City",       "13:00","19:00",   "UTC-6","14:00","12:00","15:30","18:00","16:30"),
    ( 2,1,"11/06/2026","Match B","KOR","CZE","Korea Republic",    "Czechia",                 "Guadalajara",       "20:00","02:00+1", "UTC-6","21:00","17:30","15:30","18:45","16:30"),
    # MD2 — 12 Jun
    ( 3,2,"12/06/2026","Match A","CAN","BIH","Canada",            "Bosnia and Herzegovina",  "Toronto",           "15:00","19:00",   "UTC-4","14:00","14:00","18:45","10:00","17:45"),
    ( 4,2,"12/06/2026","Match B","USA","PAR","United States",     "Paraguay",                "Los Angeles",       "18:00","01:00+1", "UTC-7","20:00","13:00","17:45","11:00","20:45"),
    # MD3 — 13 Jun
    ( 5,3,"13/06/2026","Match C","HAI","SCO","Haiti",             "Scotland",                "Boston, MA",        "21:00","01:00+1", "UTC-4","20:00","17:00","16:00","09:00","10:30"),
    ( 6,3,"13/06/2026","Match D","AUS","TUR","Australia",         "Türkiye",                 "Vancouver",         "21:00","04:00+1", "UTC-7","23:00","21:00","19:45","17:00","20:30"),
    ( 7,3,"13/06/2026","Match B","BRA","MAR","Brazil",            "Morocco",                 "New York, NJ",      "18:00","22:00",   "UTC-4","17:00","09:00","13:30","17:00","14:30"),
    ( 8,3,"13/06/2026","Match A","QAT","SUI","Qatar",             "Switzerland",             "San Francisco, CA", "12:00","19:00",   "UTC-7","14:00","12:00","17:45","12:15","20:45"),
    # MD4 — 14 Jun
    ( 9,4,"14/06/2026","Match C","CIV","ECU","Côte d'Ivoire",    "Ecuador",                 "Philadelphia, PA",  "19:00","23:00",   "UTC-4","18:00","16:00","14:45","08:00","17:45"),
    (10,4,"14/06/2026","Match A","GER","CUW","Germany",           "Curaçao",                 "Houston, TX",       "12:00","17:00",   "UTC-5","12:00","10:00","18:45","09:00","19:45"),
    (11,4,"14/06/2026","Match B","NED","JPN","Netherlands",       "Japan",                   "Dallas, TX",        "15:00","20:00",   "UTC-5","15:00","10:30","18:45","10:00","15:45"),
    (12,4,"14/06/2026","Match D","SWE","TUN","Sweden",            "Tunisia",                 "Monterrey",         "20:00","02:00+1", "UTC-6","21:00","19:00","15:30","21:00","16:45"),
    # MD5 — 15 Jun
    (13,5,"15/06/2026","Match C","KSA","URU","Saudi Arabia",      "Uruguay",                 "Miami, FL",         "18:00","22:00",   "UTC-4","17:00","17:00","14:45","09:00","17:45"),
    (14,5,"15/06/2026","Match A","ESP","CPV","Spain",             "Cabo Verde",              "Atlanta, GA",       "12:00","16:00",   "UTC-4","11:00","09:30","14:45","13:00","17:45"),
    (15,5,"15/06/2026","Match D","IRN","NZL","IR Iran",           "New Zealand",             "Los Angeles, CA",   "18:00","01:00+1", "UTC-7","20:00","19:30","17:45","12:45","20:45"),
    (16,5,"15/06/2026","Match B","BEL","EGY","Belgium",           "Egypt",                   "Seattle, WA",       "12:00","19:00",   "UTC-7","14:00","13:00","16:30","12:00","17:30"),
    # MD6 — 16 Jun
    (17,6,"16/06/2026","Match A","FRA","SEN","France",            "Senegal",                 "New York, NJ",      "15:00","19:00",   "UTC-4","14:00","15:00","09:30","16:00","13:30"),
    (18,6,"16/06/2026","Match B","IRQ","NOR","Iraq",              "Norway",                  "Boston, MA",        "18:00","22:00",   "UTC-4","17:00","17:30","19:00","14:30","16:00"),
    (19,6,"16/06/2026","Match C","ARG","ALG","Argentina",         "Algeria",                 "Kansas City, MO",   "20:00","01:00+1", "UTC-5","20:00","18:00","14:30","20:00","15:30"),
    (20,6,"16/06/2026","Match D","AUT","JOR","Austria",           "Jordan",                  "San Francisco, CA", "21:00","04:00+1", "UTC-7","23:00","13:00","20:45","22:00","17:45"),
    # MD7 — 17 Jun
    (21,7,"17/06/2026","Match C","GHA","PAN","Ghana",             "Panama",                  "Toronto",           "19:00","23:00",   "UTC-4","18:00","18:00","13:30","08:00","14:30"),
    (22,7,"17/06/2026","Match B","ENG","CRO","England",           "Croatia",                 "Dallas, TX",        "15:00","20:00",   "UTC-5","15:00","11:00","18:45","15:00","19:45"),
    (23,7,"17/06/2026","Match A","POR","COD","Portugal",          "Congo DR",                "Houston, TX",       "12:00","17:00",   "UTC-5","12:00","09:30","18:45","13:00","15:45"),
    (24,7,"17/06/2026","Match D","UZB","COL","Uzbekistan",        "Colombia",                "Mexico City",       "20:00","02:00+1", "UTC-6","21:00","09:30","19:45","18:30","16:45"),
    # MD8 — 18 Jun
    (25,8,"18/06/2026","Match A","CZE","RSA","Czechia",           "South Africa",            "Atlanta, GA",       "12:00","16:00",   "UTC-4","11:00","09:00","13:30","16:00","14:30"),
    (26,8,"18/06/2026","Match B","SUI","BIH","Switzerland",       "Bosnia and Herzegovina",  "Los Angeles, CA",   "12:00","19:00",   "UTC-7","14:00","12:45","20:45","11:00","21:45"),
    (27,8,"18/06/2026","Match C","CAN","QAT","Canada",            "Qatar",                   "Vancouver",         "15:00","22:00",   "UTC-7","17:00","17:00","20:45","13:00","17:30"),
    (28,8,"18/06/2026","Match D","MEX","KOR","Mexico",            "Korea Republic",          "Guadalajara",       "19:00","01:00+1", "UTC-6","20:00","19:00","15:30","17:30","16:30"),
]
for m in matches:
    ws.append(list(m))
set_widths(ws, [5,5,14,12,8,8,22,24,20,10,10,9,10,12,12,12,12])

# ─── Overrides ───────────────────────────────────────────────────────────────
ws = wb.create_sheet("Overrides")
ws.append(["OVERRIDES — Absolute IBC times for specific match+service/feed. Overrides offset defaults. Leave blank to use offsets."])
ws.append(["M#", "Channel", "Service", "Feed", "Override TX Start\n(IBC HH:MM)", "Override TX End\n(IBC HH:MM)", "Reason / Notes"])
style_header(ws)
set_widths(ws, [5, 12, 36, 22, 14, 14, 40])

out = "c:/Users/HBMC Op 01/.vscode/hbs-tx-schedule/hbs-tx-schedule/FWC26_TX_Input_v3.xlsx"
wb.save(out)
print(f"Saved: {out}")
print(f"Sheets: {wb.sheetnames}")
