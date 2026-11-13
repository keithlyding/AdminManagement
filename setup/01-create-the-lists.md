# Step 1: Create the lists

You get three lists, already filled with your data:

- **Action Items**: every task, with the spreadsheet's columns, six views and the Priority
  Score.
- **People**: the requesters and assignees from the Drop Downs tab; tick *Management* for
  the people whose requests add 3 to the score.
- **Recurring Work**: the Recurring tab.

## 1. Pick a home for them

The lists need a SharePoint site that only you (or your team, if you choose) can see.

- **Best: a private site of your own.** In SharePoint (from the Microsoft 365 app launcher)
  choose **+ Create site > Team site**, name it for example *Work Hub*, set **Privacy:
  Private**. Making a private Team in Teams gives you one as well. Its address looks like
  `https://<company>.sharepoint.com/sites/WorkHub`.
- **If your company blocks creating sites:** your OneDrive works too. Its address looks like
  `https://<company>-my.sharepoint.com/personal/<you>`. Open it from OneDrive, then use that
  page in step 3.

## 2. Get the setup file

- **With your data:** use the private file you were sent (`workhub-setup-private.js`). It
  holds your tasks, so keep it to yourself and delete it when done.
- **Remaking it from a newer spreadsheet:** on any computer with Python and openpyxl:
  `python tools/make_import.py "Action Items List.xlsm" --me "Your Name"`. It writes
  `private/workhub-setup-private.js`. `--me` must be your name exactly as it appears in the
  *Assigned To* column.
- **Empty lists only:** `setup/create-lists.js` from this repository.

## 3. Run it

1. On your work PC, open the site (or OneDrive) from step 1 in **Edge** or **Chrome**.
2. Press **F12** (or right-click the page > **Inspect**), and click the **Console** tab.
3. Open the setup file in Notepad, **Ctrl+A**, **Ctrl+C**.
4. Click in the Console, **Ctrl+V**, **Enter**. The browser may ask you to type
   `allow pasting` first; that is its standard warning about pasting code. This script
   only talks to your own SharePoint site, signed in as you.
5. Wait for the message *Work Hub setup finished* (about a minute; progress shows in the
   Console in green).
6. Open **Site contents** (gear icon or left menu): you'll see *Action Items*, *People* and
   *Recurring Work*.

Running it again is safe: existing lists, columns and views are kept, and data is only
loaded into a list that is still empty.

**If it stops with an error:** copy the red line and send it to me. Nothing is half-broken:
fix the cause and run it again; it picks up where it stopped. (The script has been tested
against a stand-in for SharePoint, not against your real site, so a first-run surprise is
possible.)

**If F12 is blocked on your PC:** your IT may have turned the developer tools off. Then
create the lists by hand from the column table below, or tell me and I'll write a Power
Automate flow that does the same job.

## 4. Requester and Assigned To drop-downs

Both are drop-downs filled from the **People** list, like the Drop Downs tab of the
spreadsheet. A name that isn't there can still be typed in.

- **Adding someone:** add them to *People* (tick *Requester* and/or *Can be assigned*), then
  run `setup/sync-dropdowns.js` the same way as the setup (F12 > Console > paste > Enter).
  It refreshes both drop-downs in a few seconds.
- **Lists made before 13 Nov 2026** have the two columns as text boxes: run
  `setup/sync-dropdowns.js` once to turn them into drop-downs. It keeps every item's values
  (it checks each one afterwards and puts back any that changed).

## 5. Check a few things

- **People**: tick *Management* for anyone else whose requests should outrank others, and
  fill in **Email** for the people who email you requests. Flow 1 uses the Email to put
  their name (instead of their address) in *Requester*.
- **Action Items > My Open Tasks** is the default view. The views:

  | View | Shows |
  |---|---|
  | My Open Tasks | Open items with your name in *Assigned To* (shared assignments too), highest score first |
  | Others' Open Tasks | Open items assigned to anyone else, or shared |
  | All Open Tasks | Everything not Complete |
  | Due in 7 Days or Past Due | What's coming, earliest first |
  | To Sort (no Category or Necessity) | New items from email or meetings that still need a Category and Necessity to be scored. Empty this every day. |
  | Completed | Done, newest first |

- **Recurring Work**: check each **Next Due**. The two monthly capital updates were worked
  out from their "Nth workday" rule; the June 30 plan from its date. The rest came from the
  spreadsheet or are blank: their rules ("before monthly close", "notice-driven") need a
  date from you. A duty with no Next Due never turns into a task.

## The columns (for flows, and for creating the lists by hand)

The internal name never changes and is what the flows use. You can rename the display name
in Lists any time.

**Action Items**

| Display name | Internal name | Type |
|---|---|---|
| Task | Title | Single line of text |
| Description | Details | Multiple lines of text |
| Hurdles / Needs / Notes | HurdlesNotes | Multiple lines of text |
| Requester | Requester | Choice (drop-down, typing a new name allowed): the People ticked *Requester* |
| Assigned To | AssigneeNames | Choice (drop-down, typing a new name allowed): the People ticked *Can be assigned* (default: your name) |
| Status | TaskStatus | Choice: Not Started, In Progress, Waiting / Blocked, Ready for Testing, On Hold, Complete (default Not Started) |
| Date Assigned | DateAssigned | Date only (default today) |
| Due Date | TaskDueDate | Date only |
| Completion Date | TaskCompletedDate | Date only |
| Priority Score | PriorityScore | Number, 0 decimals |
| Category | TaskCategory | Choice: the 12 categories |
| Necessity | Necessity | Choice: Low, Medium, High, Urgent |
| Ease | Ease | Choice: Difficult, Moderate, Easy |
| Cost | CostLevel | Choice: High, Medium, Low |
| Time | EffortTime | Choice: > 1 Week, 2-3 Days, 1 Day, <1 Day |
| Quick-Win Score | QuickWin | Number, 0 decimals |
| Source | ItemSource | Choice: Manual, Email, Meeting, L2L, OneNote, Recurring, Imported (default Manual) |
| Source ID | ItemSourceId | Single line of text (the email's ID, so a flag never makes two items) |
| Recurring ID | RecurringId | Number (which Recurring Work item made it) |

**People**: Name (Title), Email (PersonEmail, text), Management (IsManagement, yes/no),
Requester (IsRequester, yes/no), Can be assigned (IsAssignee, yes/no).

**Recurring Work**: Recurring Work (Title), Workstream, Actual Task / Deliverable
(Deliverable, multiple lines), Owner(s) (Owners), Frequency, Timing / Date Rule
(TimingRule), Schedule Type (RuleType, choice: *Nth workday of month*, *Annual date*,
*Manual (set Next Due by hand)*), Workday Number (N) (RuleN, number), Create Task This Many
Days Early (LeadDays, number, default 5), Next Due (NextDue, date only), Active (IsActive,
yes/no), Source (RecSource), Action / Notes (RecNotes, multiple lines).
