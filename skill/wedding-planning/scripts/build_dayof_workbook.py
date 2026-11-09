#!/usr/bin/env python3
"""Build an actual-day wedding workbook (Excel) from a JSON file.

Usage:
    python build_dayof_workbook.py input.json output.xlsx
    python build_dayof_workbook.py --template output.xlsx     # blank workbook with one example row per sheet

The JSON may contain any of these keys; missing keys simply skip that sheet:
    title, date_text, run_sheet, tea_ceremony, emcee_script, music_av,
    angbaos, packing_list, vendors, guests, meal_options
See assets/example_dayof.json for the exact shape. Contains no personal data.

Design notes
- Run Sheet times are text (e.g. "10.00am") so they read exactly as planned.
- Tea Ceremony start times, the Angbao total, guest counts per table/meal and
  tea-ceremony total pax / end time are FORMULAS, so edits flow through.
- Status colours on the Run Sheet: green = NEW, yellow = UPDATED, pink = TO CONFIRM.
"""
import json
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

FONT = "Arial"
BROWN = "7A4E3B"
BAND = "F3E3D3"
BAND_TXT = "5A3626"
FILLS = {"new": "E3F1DF", "updated": "FFF4CC", "confirm": "FDE2E1"}
THIN = Side(style="thin", color="D9C7B8")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CHECK_OPTIONS = '"☐,☑"'


def f(bold=False, color="000000", size=10, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(hex_):
    return PatternFill("solid", fgColor=hex_)


def header(ws, row, labels, widths=None):
    for i, text in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=text)
        c.font = f(True, "FFFFFF", 10)
        c.fill = fill(BROWN)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 24


def band(ws, row, text, ncols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = f(True, BAND_TXT, 11)
    c.fill = fill(BAND)
    c.alignment = Alignment(vertical="center")
    for col in range(1, ncols + 1):
        ws.cell(row=row, column=col).fill = fill(BAND)
    ws.row_dimensions[row].height = 22


def body(ws, row, values, wrap_cols=(), center_cols=()):
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = f()
        c.border = BORDER
        c.alignment = Alignment(
            vertical="top",
            wrap_text=True,
            horizontal="center" if i in center_cols else "left",
        )


def page_setup(ws, header_row=1, landscape=True):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{header_row}:{header_row}"
    ws.page_margins.left = ws.page_margins.right = 0.4


def check_validation(ws, col_letter, first, last):
    dv = DataValidation(type="list", formula1=CHECK_OPTIONS, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}{first}:{col_letter}{last}")


def grey_when_done(ws, col_letter, first, last, ncols):
    rng = f"A{first}:{get_column_letter(ncols)}{last}"
    ws.conditional_formatting.add(
        rng,
        FormulaRule(
            formula=[f'${col_letter}{first}="☑"'],
            font=Font(name=FONT, color="808080", strike=True),
        ),
    )


# ---------------------------------------------------------------- sheets
def sheet_run(wb, items, title):
    ws = wb.active
    ws.title = "Run Sheet"
    header(ws, 1, ["Time", "Duration", "Event / Details", "Lead (PIC)",
                   "People Involved", "Remarks", "Review"],
           [14, 12, 52, 22, 32, 40, 14])
    ws["A2"] = "Legend:"
    ws["A2"].font = f(True)
    ws["C2"] = "Green = NEW   |   Yellow = UPDATED   |   Pink = TO CONFIRM   |   Edit only the Time / Event / Lead / People / Remarks / Review columns"
    ws["C2"].font = f(italic=True)
    ws["C2"].fill = fill(FILLS["updated"])
    ws["C2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("C2:G2")
    ws.row_dimensions[2].height = 30
    r = 3
    for it in items:
        if "section" in it:
            band(ws, r, it["section"].upper(), 7)
        else:
            body(ws, r, [it.get("time"), it.get("duration"), it.get("event"),
                         it.get("lead"), it.get("people"), it.get("remarks"),
                         it.get("review")], center_cols=(1, 2))
            st = (it.get("status") or "").lower()
            if st in FILLS:
                for col in range(1, 8):
                    ws.cell(row=r, column=col).fill = fill(FILLS[st])
        r += 1
    ws.freeze_panes = "A2"
    page_setup(ws, 1)
    return ws


def sheet_tea(wb, rows):
    ws = wb.create_sheet("Tea Ceremony")
    ws["A1"] = "Tea Ceremony"
    ws["A1"].font = f(True, size=14)
    ws["A2"] = ("Start time (col B) is a formula: previous start + previous duration. "
                "Change the first start time or any duration (minutes) and the rest updates.")
    ws["A2"].font = f(italic=True)
    header(ws, 4, ["#", "Start", "Duration (min)", "Side", "Relative(s) receiving tea",
                   "Address term", "Pax", "Notes", "Called by", "Done"],
           [5, 9, 14, 20, 30, 20, 7, 30, 14, 8])
    first = 5
    r = first
    n = 0
    prev = None  # row number of the previous ceremony step
    for row in rows:
        if "section" in row:
            band(ws, r, row["section"], 10)
        else:
            n += 1
            # first step takes a typed start time; later steps = previous start + previous duration
            start = row.get("start", 0.4166666666666667) if prev is None else f"=B{prev}+C{prev}/1440"
            body(ws, r, [n, start, row.get("duration_min"), row.get("side"),
                         row.get("relatives"), row.get("address"), row.get("pax"),
                         row.get("notes"), row.get("called_by"), "☐"],
                 center_cols=(1, 2, 3, 7, 10))
            ws.cell(row=r, column=2).number_format = "h:mm"
            prev = r
        r += 1
    last = r - 1
    while ws.cell(row=last, column=2).value in (None, "") and last > first:
        last -= 1  # ignore trailing section bands
    ws.cell(row=last + 2, column=6, value="Total pax").font = f(True)
    ws.cell(row=last + 2, column=7, value=f"=SUM(G{first}:G{last})").font = f(True)
    ws.cell(row=last + 3, column=6, value="Total minutes").font = f(True)
    ws.cell(row=last + 3, column=7, value=f"=SUM(C{first}:C{last})").font = f(True)
    ws.cell(row=last + 4, column=6, value="Ends at (approx.)").font = f(True)
    ws.cell(row=last + 4, column=7, value=f"=B{last}+C{last}/1440").number_format = "h:mm"
    ws.cell(row=last + 4, column=7).font = f(True)
    check_validation(ws, "J", first, last)
    grey_when_done(ws, "J", first, last, 10)
    ws.freeze_panes = "A5"
    page_setup(ws, 4)


def sheet_emcee(wb, rows):
    ws = wb.create_sheet("Emcee Script")
    header(ws, 1, ["Cue", "Time", "Segment", "Trigger / When to speak",
                   "Script (what the emcee says)", "Second-language version",
                   "Music / AV cue", "Stand-by before speaking", "Est. length",
                   "Notes", "Done"],
           [7, 10, 20, 28, 55, 40, 24, 22, 11, 28, 8])
    r = 2
    first = 2
    for row in rows:
        if "section" in row:
            band(ws, r, row["section"].upper(), 11)
        else:
            body(ws, r, [row.get("cue"), row.get("time"), row.get("segment"),
                         row.get("trigger"), row.get("script"), row.get("script_2"),
                         row.get("av_cue"), row.get("standby"), row.get("length"),
                         row.get("notes"), "☐"], center_cols=(1, 2, 9, 11))
        r += 1
    check_validation(ws, "K", first, r)
    grey_when_done(ws, "K", first, r, 11)
    ws.freeze_panes = "A2"
    page_setup(ws, 1)


def sheet_music(wb, rows):
    ws = wb.create_sheet("Music and AV")
    header(ws, 1, ["Cue", "Time", "Moment", "Song / file", "Start at (mm:ss)",
                   "Fade / stop when", "Volume", "Tested"],
           [7, 11, 26, 36, 16, 30, 14, 9])
    r = 2
    for row in rows:
        body(ws, r, [row.get("cue"), row.get("time"), row.get("moment"),
                     row.get("song"), row.get("start_at"), row.get("stop_when"),
                     row.get("volume"), "☐"], center_cols=(1, 2, 5, 7, 8))
        r += 1
    check_validation(ws, "H", 2, r)
    ws.freeze_panes = "A2"
    page_setup(ws, 1)


def sheet_angbao(wb, rows):
    ws = wb.create_sheet("Angbaos")
    ws["A1"] = "Angbaos / Red packets / Tips"
    ws["A1"].font = f(True, size=14)
    ws["A2"] = "Only edit Name, Amount, Hand over by, When, Done. The total updates automatically."
    ws["A2"].font = f(italic=True)
    header(ws, 4, ["S/N", "Name", "Amount", "Hand over by", "When", "Done"],
           [6, 40, 12, 18, 26, 8])
    first = 5
    r = first
    for i, row in enumerate(rows, start=1):
        body(ws, r, [i, row.get("name"), row.get("amount"), row.get("hand_over_by"),
                     row.get("when"), "☐"], center_cols=(1, 3, 6))
        ws.cell(row=r, column=3).number_format = "#,##0.00"
        r += 1
    last = r - 1
    ws.cell(row=r + 1, column=2, value="Total").font = f(True)
    c = ws.cell(row=r + 1, column=3, value=f"=SUM(C{first}:C{last})")
    c.font = f(True)
    c.number_format = "#,##0.00"
    ws.cell(row=r + 2, column=2, value="Still to hand over").font = f(True)
    c = ws.cell(row=r + 2, column=3, value=f'=SUMIF(F{first}:F{last},"<>☑",C{first}:C{last})')
    c.font = f(True)
    c.number_format = "#,##0.00"
    check_validation(ws, "F", first, last)
    grey_when_done(ws, "F", first, last, 6)
    ws.freeze_panes = "A5"
    page_setup(ws, 4, landscape=False)


def sheet_packing(wb, rows):
    ws = wb.create_sheet("Packing List")
    header(ws, 1, ["Station", "Item", "Qty", "Provided by", "Packed", "Set up", "Remarks"],
           [22, 40, 7, 18, 9, 9, 40])
    r = 2
    first = 2
    for row in rows:
        if "section" in row:
            band(ws, r, row["section"], 7)
        else:
            body(ws, r, [row.get("station"), row.get("item"), row.get("qty"),
                         row.get("provided_by"), "☐", "☐", row.get("remarks")],
                 center_cols=(3, 5, 6))
        r += 1
    check_validation(ws, "E", first, r)
    check_validation(ws, "F", first, r)
    ws.cell(row=r + 1, column=2, value="Items packed").font = f(True)
    ws.cell(row=r + 1, column=3,
            value=f'=COUNTIF(E{first}:E{r-1},"☑")&" / "&COUNTA(B{first}:B{r-1})').font = f(True)
    ws.freeze_panes = "A2"
    page_setup(ws, 1)


def sheet_vendors(wb, rows):
    ws = wb.create_sheet("Vendor List")
    header(ws, 1, ["S/N", "Category", "Vendor", "Items / scope", "Contact", "Arrival time", "Remarks"],
           [6, 18, 28, 36, 26, 14, 36])
    r = 2
    for i, row in enumerate(rows, start=1):
        body(ws, r, [i, row.get("category"), row.get("vendor"), row.get("items"),
                     row.get("contact"), row.get("arrival"), row.get("remarks")],
             center_cols=(1, 6))
        r += 1
    ws.auto_filter.ref = f"A1:G{r-1}"
    ws.freeze_panes = "A2"
    page_setup(ws, 1)


def sheet_guests(wb, rows, meals):
    ws = wb.create_sheet("Seating and Meals")
    header(ws, 1, ["Table", "Seat", "Name", "Meal", "Dietary / remarks"],
           [22, 8, 28, 28, 36])
    r = 2
    for row in rows:
        body(ws, r, [row.get("table"), row.get("seat"), row.get("name"),
                     row.get("meal"), row.get("remarks")], center_cols=(2,))
        r += 1
    last = max(r - 1, 2)
    ws.auto_filter.ref = f"A1:E{last}"
    # Summary block to the right
    ws["G1"], ws["H1"] = "Meal", "Count"
    for c in (ws["G1"], ws["H1"]):
        c.font = f(True, "FFFFFF")
        c.fill = fill(BROWN)
        c.alignment = Alignment(horizontal="center")
    ws.column_dimensions["G"].width = 34
    ws.column_dimensions["H"].width = 9
    rr = 2
    for m in meals:
        ws.cell(row=rr, column=7, value=m).font = f()
        ws.cell(row=rr, column=8, value=f'=COUNTIF($D$2:$D${last},G{rr})').font = f()
        rr += 1
    ws.cell(row=rr, column=7, value="Total guests").font = f(True)
    ws.cell(row=rr, column=8, value=f'=COUNTA($C$2:$C${last})').font = f(True)
    ws.cell(row=rr + 1, column=7, value="With dietary remarks").font = f(True)
    ws.cell(row=rr + 1, column=8, value=f'=COUNTA($E$2:$E${last})').font = f(True)
    ws.freeze_panes = "A2"
    page_setup(ws, 1)


# ---------------------------------------------------------------- main
def build(data, out):
    wb = Workbook()
    sheet_run(wb, data.get("run_sheet", []), data.get("title", "Wedding Day"))
    if data.get("tea_ceremony"):
        sheet_tea(wb, data["tea_ceremony"])
    if data.get("emcee_script"):
        sheet_emcee(wb, data["emcee_script"])
    if data.get("music_av"):
        sheet_music(wb, data["music_av"])
    if data.get("angbaos"):
        sheet_angbao(wb, data["angbaos"])
    if data.get("packing_list"):
        sheet_packing(wb, data["packing_list"])
    if data.get("vendors"):
        sheet_vendors(wb, data["vendors"])
    if data.get("guests"):
        sheet_guests(wb, data["guests"], data.get("meal_options", []))
    wb.save(out)


TEMPLATE = {
    "title": "EXAMPLE - replace with your details",
    "run_sheet": [
        {"section": "Day before"},
        {"time": "8.00pm", "duration": "2h", "event": "EXAMPLE: Venue setup walk-through",
         "lead": "Day-of Captain", "people": "Couple, venue", "remarks": "Replace this row", "status": "confirm"},
        {"section": "Wedding day - morning"},
        {"time": "9.00am", "duration": "30 mins", "event": "EXAMPLE: Makeup artist arrives",
         "lead": "Makeup artist", "people": "Bride", "remarks": "", "status": "new"},
    ],
    "tea_ceremony": [
        {"section": "Side A"},
        {"start": 0.4166666666666667, "duration_min": 2, "side": "Grandparent", "relatives": "EXAMPLE",
         "address": "EXAMPLE", "pax": 1, "called_by": "Caller"},
    ],
    "emcee_script": [{"section": "Pre-event"},
                     {"cue": "A1", "time": "9.30am", "segment": "EXAMPLE", "script": "EXAMPLE script"}],
    "music_av": [{"cue": "S1", "time": "11.00am", "moment": "EXAMPLE entrance",
                  "song": "[Song TBC]", "volume": "Medium"}],
    "angbaos": [{"name": "EXAMPLE vendor", "amount": 20, "hand_over_by": "Groom", "when": "Morning"}],
    "packing_list": [{"section": "Reception"},
                     {"station": "Reception", "item": "EXAMPLE guest book", "qty": 1, "provided_by": "Own"}],
    "vendors": [{"category": "Photographer", "vendor": "EXAMPLE", "items": "Actual day",
                 "contact": "Name (phone)", "arrival": "7.00am"}],
    "guests": [{"table": "Table 1", "seat": "1", "name": "EXAMPLE guest", "meal": "Chicken", "remarks": ""}],
    "meal_options": ["Chicken", "Fish", "Vegetarian", "Kids meal"],
}


if __name__ == "__main__":
    args = sys.argv[1:]
    if len(args) == 2 and args[0] == "--template":
        build(TEMPLATE, args[1])
    elif len(args) == 2:
        with open(args[0], encoding="utf-8") as fh:
            build(json.load(fh), args[1])
    else:
        print(__doc__)
        sys.exit(1)
    print("Wrote", args[-1])
