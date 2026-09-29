# Ticket Notes to Excel

A small desktop tool (Windows + macOS + Linux) that reads your rough daily
ticket notes (`.txt`) and adds them as rows to your `Ticketing Sheet` Excel file.

**Duplicate Ticket IDs are never removed.** One ticket with three comments gives
three rows.

## Run it

Requires Python 3.9+ with Tk (included in the python.org installers).

```bash
git clone https://github.com/<you>/ticket-notes-to-excel.git
cd ticket-notes-to-excel
python -m pip install -r requirements.txt
python main.py            # Windows: py main.py
```

macOS note: if `import tkinter` fails, install Python from python.org, or
`brew install python-tk`. Linux: `sudo apt install python3-tk`.

## Using it

1. **Notes file** - Browse to your `.txt`. It is parsed immediately.
2. **Preview** - check every row. Double-click a cell to edit it, select rows and
   press *Remove selected rows* to leave them out. Yellow rows have no priority.
3. **Excel file** - Browse to today's workbook, or *Create new...* to make one with
   the same structure as your template. The last file used is remembered.
4. **Add rows to Excel** - rows are written below the last used row.
   Existing rows are not touched.

Safety: a timestamped copy is saved in a `backups/` folder next to the workbook
before writing (can be switched off). If rows identical to the ones you are about
to add already exist, you get a "Possible repeat" prompt (protects against running
the same notes twice) - it never skips anything on its own. Close the workbook in
Excel before writing.

## Excel layout

Sheet `Ticketing Sheet`, columns in this order:

| ticket_no | ticket_priority | ticket_status | ticket_sub | action_taken | time_taken |
|---|---|---|---|---|---|
| Ticket ID | P1-P5 | In Progress / Assigned / Complete | Title | Comment / action taken | left blank |

Formatting (Calibri 11, red bold header, centered wrapped cells, thin borders,
column widths, 30 pt rows) follows the template. Rows with long text get a taller
row (up to Excel's 409 pt limit) so the text is visible; untick *Grow row height*
to keep 30 pt.

## Notes format understood

```text
1. T20260929.0043 - Alert: CPU on SBS is Failed    | P2 | Assigned
================================================
Body text: whatever you did / commented (any number of lines)
_____________________________________________________________
```

* `N.` numbering, `| P# | Status`, and the `=====` line are all optional.
* An entry ends at `_____`, at the next header line, or at end of file.
* Several header lines in a row (no body between them) share the body below
  them, and inherit the priority/status of the first one that has them.
* A missing priority is filled from another entry of the **same ticket**.
  Status is never inherited (it changes between updates).
* `_____XXXXX_____` divider lines are ignored.
* Pasted ticket-system comments starting with `Recent Comments` and
  `2026-09-29 12:13:57 EDT - Name ...` lines have those two label lines removed
  (checkbox in Step 1).
* Title whitespace is normalised; description line breaks are kept.

If your format changes, the patterns are at the top of `ticket_tool/parser.py`.

## Tests

```bash
pip install -r requirements-dev.txt
python -m pytest
```

## Build a double-click app (optional)

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --name TicketNotesToExcel main.py
```
Build on each OS separately (the output is in `dist/`).

## Publish to GitHub

```bash
git init
git add .
git commit -m "Ticket Notes to Excel"
git branch -M main
git remote add origin https://github.com/<you>/ticket-notes-to-excel.git
git push -u origin main
```
Create the empty repository on github.com first (no README/licence). The
`.gitignore` blocks `.xlsx` and `.txt` files so real ticket data and customer
names are not uploaded; only `samples/sample_notes.txt` (fake data) is included.
