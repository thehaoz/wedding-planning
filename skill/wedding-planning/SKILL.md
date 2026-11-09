---
name: wedding-planning
description: Plan a wedding end to end for engaged couples. Builds a backward timeline from the wedding date, a budget with planned/committed/actual tracking, a vendor and task system, and a day-of run sheet. Use when someone mentions their wedding, engagement, venue, banquet, solemnisation, budget, guest list, vendors, or run sheet.
---

# Wedding Planning

Help a couple go from "we're engaged" to "everything is booked, paid and rehearsed". This skill is distilled from one real 7-8 month wedding plan (outdoor solemnisation, tea ceremony, lunch reception for roughly 100 guests) and the mistakes that plan revealed. Treat the examples as one data point, not a standard: localise to the couple's country, culture, religion and budget. Percentages and lead times in the references are starting points, not targets.

## Before anything else: intake (ask, don't assume)

Ask these in one batch, then proceed with sensible defaults for anything unanswered:

1. Wedding date (or window) and today's date, so you know the weeks remaining.
2. Total budget ceiling and who is paying (couple, parents, split).
3. Guest count (firm vs rough) and the largest family constraint (elders, dietary, travel).
4. Location and culture: legal registration needs, religious or traditional ceremonies (tea ceremony, betrothal gifts, etc.).
5. Format: single day or multi-event; indoor, outdoor, or both.
6. Top three priorities (e.g. food, photos, venue) so budget follows what matters.
7. Where they already track things (Notion, Sheets, paper). Use that tool rather than introducing another.

If the date is under 6 months away, say so plainly and use the compressed timeline in `references/timeline.md`.

## Workflow

1. **Timeline.** Build a backward plan from the date using `references/timeline.md`. Output as dated tasks with owners.
2. **Budget.** Set the ceiling first, then allocate by category using `references/budget.md`. Always include a contingency line of 5-10%. Track three numbers per line: Planned, Committed (signed quote), Actual (paid).
3. **Vendors.** One row per vendor with quote, deposit, balance, due dates, contact and day-of arrival time. Payments are separate rows linked to the vendor.
4. **Tasks.** Every task gets an owner, a due date and a category. A task with no date is a wish.
5. **Guests.** Single guest database: RSVP, meal choice, dietary flags, table, side of family. Everything else (run sheet, caterer brief, seating) links to it; never retype counts.
6. **Day-of.** Build the run sheet from `references/dayof-runsheet.md` at least 8 weeks out, rehearse timings, fill every contact by T-4 weeks. Then, about 2-4 weeks out, generate the **actual-day Excel workbook** (run sheet, tea ceremony, emcee script, music/AV, angbaos, packing list, vendors, seating and meals) from the planning data using `references/dayof-workbook.md` and `scripts/build_dayof_workbook.py`.
7. **Monthly review.** Reconcile spending against bank/receipts, tick off finished tasks, list what is overdue. See `references/lessons.md` for the audit checklist.
8. **After the wedding.** Run the 30-day wrap-up list in `references/lessons.md`.

## Rules that prevent the most common failures

- **Never present a budget without a ceiling.** If the user has none, help them set one before allocating.
- **Reconcile any number that appears in two places.** If a vendor page and the ledger disagree, flag it and ask for the invoice. Do not pick one.
- **Count ceremony groups and time them early.** Tea ceremonies, family photos, speeches and processions run longer than estimated; add 25-30% buffer.
- **Confirm venue clear-out time in writing** whenever the plan shifts the schedule.
- **Rain plan in writing** for any outdoor element, including the indoor backup space.
- **Verify licences and legal steps** (solemniser licence, registration slot, witnesses' ID) against the official government source, never only review sites. Say when you couldn't verify.
- **Don't invent prices.** Give ranges only when sourced, label them as guidance, and recommend 2-3 quotes.
- **Never fabricate completion.** If you can't see whether a task was done, ask.

## Output style

Match the couple's tool. For the day itself, always offer the Excel workbook (printable, works offline). In Notion: databases for Budget, Payments, Vendors, Tasks, Guests, plus pages for Run sheet and Lessons. Elsewhere: a spreadsheet with the same tabs. Keep answers short; deliver the timeline, budget and run sheet as files or databases, not walls of chat.

## Reference files

- `references/timeline.md`: backward timeline (12 months, 6 months, 3 months, compressed)
- `references/budget.md`: allocation framework, payment schedule, ledger columns, monthly reconciliation
- `references/dayof-runsheet.md`: run sheet template, bridal-party roles, contingency kit, items checklist
- `references/dayof-workbook.md`: how to turn planning data into the actual-day Excel workbook (sheet list, mapping, workflow)
- `scripts/build_dayof_workbook.py`: generator (`python build_dayof_workbook.py data.json out.xlsx`, or `--template out.xlsx`); needs `openpyxl`
- `assets/example_dayof.json`: input shape with fictional example data
- `references/lessons.md`: audit checklist, mistakes to avoid, post-wedding wrap-up
- `references/culture-modules.md`: optional modules (Chinese tea ceremony, betrothal gifts, civil solemnisation example)
