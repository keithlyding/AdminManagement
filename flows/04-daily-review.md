# Flow 4: Daily review

**What it does:** every weekday at 3:30 pm it adds one action item, *Process meeting notes
for <day>*, due today, listing that day's meetings from your Outlook calendar. It's the
nudge to go through your OneNote pages and move each action item into the list before you
leave. Mark it Complete when done. No meetings that day, no item.

See [Flow 1](01-flagged-email.md#conventions-used-in-all-four-flow-guides) for how to
rename actions and paste expressions.

## Build it

1. **Create > Scheduled cloud flow.** Name: *Work Hub - Daily review*. Repeat every **1
   Week**. Trigger: Time zone Eastern, On these days Monday-Friday, At these hours **15**,
   At these minutes **30**.

2. **Compose** **Today**:
   `formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'yyyy-MM-dd')`

3. **Office 365 Outlook > Get calendar view of events (V3).** Rename **TodaysEvents**.
   - Calendar id: **Calendar**
   - Start Time: `convertToUtc(concat(outputs('Today'), 'T00:00:00'), 'Eastern Standard Time')`
   - End Time: `convertToUtc(concat(outputs('Today'), 'T23:59:59'), 'Eastern Standard Time')`

4. **Filter array** **Meetings**. From: `body('TodaysEvents')?['value']`. Advanced mode:
   `@not(equals(item()?['isAllDay'], true))`

5. **Condition**: `length(body('Meetings'))` is greater than `0`. In **True**:

   a. **Select** **MeetingLines**. From `body('Meetings')`. Map in text mode:
      `concat(formatDateTime(convertFromUtc(item()?['start'], 'Eastern Standard Time'), 'h:mm tt'), '  ', item()?['subject'])`

   b. **SharePoint > Create item** (Action Items):

   | Field | Value |
   |---|---|
   | Task | `concat('Process meeting notes for ', formatDateTime(outputs('Today'), 'ddd MM/dd'))` |
   | Description | `join(body('MeetingLines'), decodeUriComponent('%0A'))` |
   | Hurdles / Needs / Notes | *type*: For each meeting: open its notes, add every action item to this list with Source = Meeting, then mark this Complete. |
   | Requester Value | *Enter custom value*, type Daily review |
   | Assigned To Value | *choose* your name |
   | Status Value | Not Started |
   | Date Assigned | `outputs('Today')` |
   | Due Date | `outputs('Today')` |
   | Category Value | Other |
   | Necessity Value | High |
   | Source Value | Meeting |

6. **Save** and **Run** to test.

If the meeting times show in UTC, the V3 action already gave local times: use
`formatDateTime(item()?['start'], 'h:mm tt')` instead.
