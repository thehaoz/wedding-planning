# Actual-Day Excel Workbook

Turn the planning workspace into one printable, phone-friendly Excel file for the day itself. This is the "single source the helpers open on the morning". It is built from information already planned elsewhere (Notion, Sheets, notes); the skill's job is to gather it, structure it and generate the file.

## What the workbook contains

| Sheet | Purpose | Built from (planning source) | Live formulas |
|---|---|---|---|
| Run Sheet | Master timeline from 2 nights before to pack-up, with owner (PIC), people involved, remarks, review column and colour legend | Programme flow / master run sheet, bridal-party duty page | none (text times) |
| Tea Ceremony | Serving order with start time, duration, side, address term, pax, caller, done tick | Family customs page | start times chain from durations; total pax; total minutes; end time |
| Emcee Script | Cue-by-cue script with trigger, optional second-language version, music/AV cue, stand-by, length | Programme flow plus the couple's story | none |
| Music and AV | Song/file per moment with start time, stop cue, volume, tested tick | Programme flow AV notes | none |
| Angbaos | Every red packet or tip: who, how much, who hands it over, when | Budget ledger (tips/helper lines) and helper list | total, amount still to hand over |
| Packing List | Items by station with quantity, provider, packed, set-up | Items checklist and décor/vendor pages | packed count |
| Vendor List | Contact, scope, arrival time, remarks (invoice/ref numbers) | Vendors database | none (filterable) |
| Seating and Meals | Table, seat, name, meal, dietary remarks, plus meal counts | RSVP/guest database | meal counts, total guests, dietary-remark count |

Optional extras seen in practice: a venue layout image sheet, and extra seating views per segment (ceremony vs lunch). Offer these only if the couple has the material.

## Sheet conventions (match these)

- Header row: dark brown fill, white bold text. Section bands (for example "ONE NIGHT BEFORE"): light tan fill, bold brown text.
- Run Sheet legend: green = NEW, yellow = UPDATED, pink = TO CONFIRM. Keep a Review column for the Captain's notes.
- Freeze the header row, wrap text, landscape fit-to-width printing with repeating header.
- Checkbox columns use a drop-down of an empty and a ticked box; ticked rows grey out and strike through.
- Written in Arial; no personal data is baked into the generator.

## Workflow

1. **Collect.** Ask the couple which planning source to use (Notion pages, a spreadsheet, or pasted notes). Read it all before writing anything. For Notion: the Programme Flow page, Bridal Party Duty page, tea ceremony/family customs page, RSVP and guest database, Vendors database, Expenses (tips/ang bao lines), Items checklist.
2. **Map to JSON.** Fill `assets/example_dayof.json`'s shape:
   - `run_sheet`: list of `{section}` and `{time, duration, event, lead, people, remarks, status}` rows.
   - `tea_ceremony`: `{start, duration_min, side, relatives, address, pax, notes, called_by}`. Only the first row needs `start`; durations must be numbers in minutes.
   - `emcee_script`, `music_av`, `angbaos`, `packing_list`, `vendors`, `guests`, `meal_options` as in the example.
3. **Resolve gaps before generating.** List every placeholder (missing phone number, unnamed helper, "TBC" song, unconfirmed vendor) and mark those rows `"status": "confirm"` so they show pink. Do not invent names, numbers or amounts.
4. **Generate.** `python scripts/build_dayof_workbook.py data.json DayOf.xlsx` (needs `openpyxl`). Use `--template out.xlsx` for a blank file with one example row per sheet.
5. **Recalculate and check.** Open in Excel/LibreOffice (or run the platform's recalc step) and confirm tea-ceremony end time, ang bao total, and meal counts match the source. Check that guest totals equal the RSVP count.
6. **Hand over.** Save the workbook beside the planning files. Tell the couple which rows are pink (still to confirm), and suggest sharing a read-only copy with helpers and printing the Run Sheet and Packing List.

## Guardrails

- Never put real guest names or phone numbers into the published skill. Those live only in the couple's own generated workbook.
- One source of truth: if a number differs between the planning source and the workbook (guest count, timing), flag it rather than choosing.
- Keep the Run Sheet to what a helper needs at a glance: time, what, who leads, remarks. Longer scripts belong in the Emcee Script sheet.
- After the wedding, keep the workbook as the record of what actually happened (update "Review" notes).
