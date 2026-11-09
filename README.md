# Wedding Planning Folder

Anonymised case study and reusable skill, distilled from a real wedding plan.

- `wedding-retrospective.md`: overview of what was done, the timeline, budget analysis, and improvements. Start here.
- `skill/wedding-planning/`: a reusable Claude skill for engaged couples.
  - `SKILL.md`: the skill itself (intake questions, workflow, rules)
  - `references/timeline.md`, `budget.md`, `dayof-runsheet.md`, `dayof-workbook.md`, `lessons.md`, `culture-modules.md`
  - `scripts/build_dayof_workbook.py`: generates the actual-day Excel workbook from JSON
  - `assets/example_dayof.json`: fictional example input

The skill contains no personal data (no guest names, phone numbers or addresses). Budget figures appear only as rounded percentages and dates only as time-before-wedding.

To use it, upload the `skill/wedding-planning` folder as a custom skill, or keep it as a reference playbook.

## Notion template

A ready-to-duplicate Notion workspace (Budget, Payments, Vendors, Tasks, Guests, Run sheet) is available here:

[Wedding planning template](https://app.notion.com/p/Wedding-planning-template-362b750e50cd80aba148dd093ec106a6?source=copy_link)

Duplicate it into your own Notion account before editing — it pairs with the skill above, which can read from and write to a Notion workspace in this shape.
