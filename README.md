# Work Hub

One place for every action item: the requests in email, what comes out of meetings, L2L
dispatches, notes and the recurring duties, all in one Microsoft List that you work from on
your PC, iPad or phone.

It uses only what Microsoft 365 already gives you, with standard (not Premium) connectors:

| Piece | What it does |
|---|---|
| **Microsoft Lists**: *Action Items*, *People*, *Recurring Work* | Where everything lives. The same fields, drop-downs and Priority Score as the Action Items spreadsheet. |
| **Flow 1: Flagged email** | Flag an email in Outlook and it becomes an action item within a few minutes. |
| **Flow 2: Priority refresh** | Recalculates every open item's Priority Score 4 times each weekday, so the list stays sorted by what matters, and stamps the Completion Date when you mark something Complete. |
| **Flow 3: Recurring tasks** | Turns each recurring duty into an action item a few days before it is due, then works out the next due date. |
| **Flow 4: Daily review** | Each afternoon, one item listing today's meetings, a reminder to move their action items in from your notes. |

## Set it up (about an hour, all on your work PC)

1. [Create the lists](setup/01-create-the-lists.md) (10 minutes)
2. [Flow 1: Flagged email](flows/01-flagged-email.md) (10 minutes)
3. [Flow 2: Priority refresh](flows/02-priority-refresh.md) (20 minutes)
4. [Flow 3: Recurring tasks](flows/03-recurring-tasks.md) (20 minutes, optional)
5. [Flow 4: Daily review](flows/04-daily-review.md) (10 minutes, optional)
6. [Use it on the iPad](docs/ipad.md)

Then read [the daily routine](docs/daily-routine.md): the lists only help if everything goes
into them.

Coming later: [roadmap](docs/roadmap.md) (L2L dispatches, OneNote, a Power Apps front end).

## Priority Score

The same formula as the spreadsheet:

**Priority Score = Category + 3 × Necessity + Management Bonus + Due Date Points**

| Category | Score | | Necessity | Score | | Due date | Points |
|---|---|---|---|---|---|---|---|
| Other | 1 | | Low | 1 | | Past due | 10 |
| Training / People Development | 2 | | Medium | 2 | | Due today | 8 |
| Ergonomics / Operator Support | 3 | | High | 3 | | Due tomorrow | 6 |
| Controls / Automation / IT | 4 | | Urgent | 4 | | Within 3 days | 4 |
| Capital Project / Infrastructure | 5 | | | | | Within 7 days | 2 |
| Energy / Utilities / Sustainability | 6 | | | | | Later or no date | 0 |
| Productivity / Waste Reduction | 7 | | | | | | |
| Quality | 8 | | | | | | |
| Reliability / Downtime Reduction | 9 | | | | | | |
| Customer Impact | 10 | | | | | | |
| Compliance / Regulatory | 11 | | | | | | |
| Safety | 12 | | | | | | |

Management Bonus: 3 when the Requester is ticked as *Management* in the People list.
Complete items score -1. Ties are broken by the **Quick-Win Score** (Ease + Cost + Time:
Easy, Low cost and short jobs first). A score of 22 or more shows in bold red.

## Row colors

| Row | Color |
|---|---|
| Past due | red |
| Due today | yellow |
| Priority Score 22 or more | orange, bold |
| Everything else | every other row light gray, for easier reading |
| Complete | gray text |

First match wins: a past-due high-priority item is red. To put the colors on a view you
made yourself, run `setup/apply-formatting.js` (F12 > Console > paste > Enter).

## Your data stays out of this repository

This repository is public. It holds only the setup script, the flow instructions and the
converter. Your tasks go from your spreadsheet straight into your list:
`tools/make_import.py` writes a private file (into `private/`, which git ignores) that you
paste into your browser once. Never commit a spreadsheet, the private file, or a key.
