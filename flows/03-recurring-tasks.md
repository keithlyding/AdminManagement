# Flow 3: Recurring tasks

**What it does:** every morning it looks at *Recurring Work*. When a duty's **Next Due** is
within its **Create Task This Many Days Early** (lead days), it adds an action item for it
(Source = Recurring, Due Date = Next Due), unless one is still open. Then it moves Next Due
forward:

| Schedule Type | Next Due becomes |
|---|---|
| Nth workday of month | The Nth workday (Mon-Fri) of the following month. Company holidays are not skipped, same as the spreadsheet. |
| Annual date | The same date next year |
| Manual (set Next Due by hand) | Empty. Type the next date yourself when you know it (for example when the close calendar or a notice arrives). |

See [Flow 1](01-flagged-email.md#conventions-used-in-all-four-flow-guides) for how to
rename actions and paste expressions.

## Build it

1. **Create > Scheduled cloud flow.** Name: *Work Hub - Recurring tasks*. Repeat every **1
   Day**. In the trigger: Time zone Eastern, At these hours **5**, At these minutes **30**.

2. **Compose** **Today**:
   `formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'yyyy-MM-dd')`

3. **SharePoint > Get items** **ActiveDuties**. List: **Recurring Work**. Filter Query:
   `IsActive eq 1 and NextDue ne null`

4. **Apply to each** **EachDuty** over `body('ActiveDuties')?['value']`. Inside:

   a. **Condition** **IsTime**: left, expression:
      `lessOrEquals(ticks(formatDateTime(items('EachDuty')?['NextDue'], 'yyyy-MM-dd')), ticks(addDays(outputs('Today'), int(coalesce(items('EachDuty')?['LeadDays'], 5)), 'yyyy-MM-dd')))`
      / is equal to / expression `true`

   b. In **True**: **SharePoint > Get items** **OpenCopy**. List: **Action Items**. Filter
      Query: type `RecurringId eq ` then expression `items('EachDuty')?['ID']` then type
      ` and TaskStatus ne 'Complete'`

   c. Still in True: **Condition**: `length(body('OpenCopy')?['value'])` is equal to `0`.
      In its **True**:

      i. **SharePoint > Create item** (Action Items):

      | Field | Value |
      |---|---|
      | Task | `items('EachDuty')?['Title']` |
      | Description | `items('EachDuty')?['Deliverable']` |
      | Hurdles / Needs / Notes | `items('EachDuty')?['RecNotes']` |
      | Requester Value | *Enter custom value*, type Recurring |
      | Assigned To Value | *Enter custom value*, then `items('EachDuty')?['Owners']` |
      | Status Value | *choose* Not Started |
      | Date Assigned | `outputs('Today')` |
      | Due Date | `formatDateTime(items('EachDuty')?['NextDue'], 'yyyy-MM-dd')` |
      | Source Value | *choose* Recurring |
      | Recurring ID | `items('EachDuty')?['ID']` |

      ii. **Control > Switch.** On: `items('EachDuty')?['RuleType']?['Value']`

      - **Case** equals `Nth workday of month`:
        1. **Compose** **NextMonthStart**:
           `startOfMonth(addToTime(startOfMonth(formatDateTime(items('EachDuty')?['NextDue'], 'yyyy-MM-dd')), 1, 'Month'), 'yyyy-MM-dd')`
        2. **Select** **MonthDays**. From: `range(0, 31)`. Map in text mode:
           `addDays(outputs('NextMonthStart'), item(), 'yyyy-MM-dd')`
        3. **Filter array** **Workdays**. From: `body('MonthDays')`. Choose **Edit in
           advanced mode** and paste:
           `@and(equals(substring(item(), 0, 7), substring(outputs('NextMonthStart'), 0, 7)), greater(dayOfWeek(item()), 0), less(dayOfWeek(item()), 6))`
        4. **Update item** (Recurring Work): Id `items('EachDuty')?['ID']`, Recurring Work
           `items('EachDuty')?['Title']`, Next Due:
           `body('Workdays')?[sub(min(int(coalesce(items('EachDuty')?['RuleN'], 1)), length(body('Workdays'))), 1)]`
      - **Case** equals `Annual date`: **Update item** (Recurring Work): Id, Recurring Work
        as above, Next Due:
        `addToTime(formatDateTime(items('EachDuty')?['NextDue'], 'yyyy-MM-dd'), 1, 'Year', 'yyyy-MM-dd')`
      - **Default**: **Update item** (Recurring Work): Id, Recurring Work as above, Next
        Due: expression `null`

5. **Save** and **Run** once. With the imported data, the monthly capital forecast (10th
   workday) should appear right away. **One duplicate to delete by hand:** the spreadsheet
   already has a dated *November Capital Forecast* task; delete one of the two.

## Notes

- A duty with an open item is skipped until that item is Complete. It isn't created twice.
- To pause a duty, untick **Active**.
- **If Default can't clear the date** (the Update item rejects `null`): set Manual duties'
  Next Due to a far-future date instead, by changing the Default branch's value to
  `'2099-12-31'`, and ignore them in the *Active by Next Due* view.
- Nth workday and quarterly rules beyond these three types: tell me which and I'll add them.
