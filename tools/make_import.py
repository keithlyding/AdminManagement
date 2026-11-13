"""Turn the Action Items workbook into the private setup file for the Work Hub lists.

    python tools/make_import.py "Action Items List.xlsm" --me "Your Name" -o private/workhub-setup-private.js

It reads the All Tasks, Recurring and Drop Downs tabs and writes one JavaScript file: your
data followed by setup/create-lists.js. Paste that file into the browser console on your
SharePoint site (see setup/01-create-the-lists.md).

The output holds your work data. Keep it out of this repository: the private/ folder and
spreadsheets are in .gitignore.
"""
import argparse
import calendar
import datetime as dt
import json
import pathlib
import re
import sys

import openpyxl

ROOT = pathlib.Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "setup" / "create-lists.js"


def text(v):
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def iso(v):
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    return None


def find_header(ws, wanted, max_row=12):
    """Return (row, {header: column}) for the first row holding all wanted headers."""
    for r in range(1, max_row + 1):
        row = {text(c.value): c.column for c in ws[r] if text(c.value)}
        if all(w in row for w in wanted):
            return r, row
    raise SystemExit("Could not find the headers %s on the %s tab." % (wanted, ws.title))


def read_tasks(ws):
    r0, h = find_header(ws, ["Task", "Status", "Due Date", "Requester"])
    col = {
        "assigned": "Date Asg'd", "requester": "Requester", "assignedTo": "Assigned To:",
        "task": "Task", "description": "Description", "notes": "Hurdles/Needs/Notes",
        "status": "Status", "due": "Due Date", "completed": "Completion Date",
        "sheetScore": "Priority Score", "ease": "Ease", "cost": "Cost", "time": "Time",
        "category": "Category", "necessity": "Necessity",
    }
    tasks = []
    for r in range(r0 + 1, ws.max_row + 1):
        get = lambda name: ws.cell(r, h[col[name]]).value if col[name] in h else None
        if not text(get("task")):
            continue
        t = {k: (iso(get(k)) if k in ("assigned", "due", "completed") else text(get(k))) for k in col}
        t["sheetScore"] = get("sheetScore")
        tasks.append(t)
    return tasks


def column_under(ws, header):
    """Values listed below a header cell on the Drop Downs tab."""
    for row in ws.iter_rows():
        for c in row:
            if text(c.value) == header:
                out = []
                for r in range(c.row + 1, ws.max_row + 1):
                    v = text(ws.cell(r, c.column).value)
                    if not v:
                        break
                    out.append(v)
                return out
    return []


def read_people(ws, tasks):
    requesters = column_under(ws, "Requester")
    assignees = column_under(ws, "Assigned To")
    managers = column_under(ws, "Management Requester")
    names = []
    for n in requesters + assignees + managers:
        if n not in names:
            names.append(n)
    return [{"name": n, "email": None, "management": n in managers,
             "requester": n in requesters, "assignee": n in assignees} for n in names]


def nth_workday(year, month, n):
    days = [dt.date(year, month, d) for d in range(1, calendar.monthrange(year, month)[1] + 1)]
    work = [d for d in days if d.weekday() < 5]
    return work[n - 1] if n <= len(work) else work[-1]


def next_nth_workday(n, today):
    y, m = today.year, today.month
    while True:
        d = nth_workday(y, m, n)
        if d >= today:
            return d
        m += 1
        if m == 13:
            y, m = y + 1, 1


MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_name) if m}


def read_recurring(ws, today):
    r0, h = find_header(ws, ["Recurring Work", "Timing / Date Rule", "Next Due"])
    get = lambda r, name: ws.cell(r, h[name]).value if name in h else None
    last = ws.max_row
    if "tblRecurring" in ws.tables:  # stop at the end of the table (a legend sits below it)
        last = openpyxl.utils.cell.range_boundaries(ws.tables["tblRecurring"].ref)[3]
    out = []
    for r in range(r0 + 1, last + 1):
        title = text(get(r, "Recurring Work"))
        if not title:
            continue
        rule = text(get(r, "Timing / Date Rule")) or ""
        item = {
            "title": title, "workstream": text(get(r, "Workstream")),
            "deliverable": text(get(r, "Actual Task / Deliverable")), "owners": text(get(r, "Owner(s)")),
            "frequency": text(get(r, "Frequency")), "rule": rule, "source": text(get(r, "Source")),
            "notes": text(get(r, "Action / Notes")), "leadDays": 5,
            "ruleType": "Manual (set Next Due by hand)", "ruleN": None, "nextDue": None,
        }
        m = re.search(r"(\d+)(?:st|nd|rd|th) workday", rule, re.I)
        a = re.search(r"due (%s) (\d{1,2}) each year" % "|".join(MONTHS), rule, re.I)
        if m and (item["frequency"] or "").lower().startswith("monthly"):
            item["ruleType"], item["ruleN"] = "Nth workday of month", int(m.group(1))
            item["nextDue"] = next_nth_workday(item["ruleN"], today).isoformat()
            item["leadDays"] = 3
        else:
            if a:
                item["ruleType"] = "Annual date"
                month, day = MONTHS[a.group(1).lower()], int(a.group(2))
                d = dt.date(today.year, month, day)
                item["nextDue"] = (d if d >= today else dt.date(today.year + 1, month, day)).isoformat()
                item["leadDays"] = 45
            else:
                # Keep a date from the sheet only if it is still ahead of us.
                for name in ("Next Due", "Next Window End"):
                    d = iso(get(r, name))
                    if d and d >= today.isoformat():
                        item["nextDue"] = d
                        break
        out.append(item)
    return out


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("workbook")
    p.add_argument("--me", required=True, help="your name exactly as it appears in Assigned To")
    p.add_argument("--site", help="SharePoint site address (optional; found from the page otherwise)")
    p.add_argument("--today", help="yyyy-mm-dd (default: today)")
    p.add_argument("-o", "--out", default=str(ROOT / "private" / "workhub-setup-private.js"))
    p.add_argument("--json", help="also write the data as JSON to this path (for checking)")
    a = p.parse_args(argv)

    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    wb = openpyxl.load_workbook(a.workbook, data_only=True, keep_vba=False)
    tasks = read_tasks(wb["All Tasks"])
    data = {
        "config": {"myName": a.me, "siteUrl": a.site},
        "tasks": tasks,
        "people": read_people(wb["Drop Downs"], tasks),
        "recurring": read_recurring(wb["Recurring"], today) if "Recurring" in wb.sheetnames else [],
    }
    if a.json:
        pathlib.Path(a.json).parent.mkdir(parents=True, exist_ok=True)
        pathlib.Path(a.json).write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
    out = pathlib.Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("// PRIVATE: holds your work data. Do not commit or share.\n"
                   "window.WORKHUB_DATA = " + json.dumps(data, ensure_ascii=False) + ";\n\n"
                   + SCRIPT.read_text(encoding="utf-8"), encoding="utf-8")
    print("%d tasks, %d people, %d recurring items -> %s" % (
        len(tasks), len(data["people"]), len(data["recurring"]), out))


if __name__ == "__main__":
    sys.exit(main())
