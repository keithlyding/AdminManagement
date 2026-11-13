# Flow 2: Priority refresh

**What it does:** 4 times every weekday (5:00, 10:00, 13:00, 16:00) it recalculates the
Priority Score and Quick-Win Score of every open item, so the list stays sorted as due dates
come closer and as you fill in Category and Necessity. It also fills in the **Completion
Date** of items you just marked Complete. Want it now? Open the flow and press **Run**.

Why not every few minutes: a Microsoft 365 license allows about 6,000 flow actions a day
([Microsoft: platform limits](https://learn.microsoft.com/en-au/training/modules/power-platform-architecture/6-limits)).
With ~70 open items one run is ~450 actions, so 4 runs leave plenty for the other flows.
If your list grows past ~250 open items, drop to 3 runs a day.

See [Flow 1](01-flagged-email.md#conventions-used-in-all-four-flow-guides) for how to
rename actions and paste expressions.

## Build it

1. **Create > Scheduled cloud flow.** Name: *Work Hub - Priority refresh*. Repeat every
   **1 Week**. **Create**.
2. Open the **Recurrence** trigger: **Time zone** (UTC-05:00) Eastern Time, **On these
   days** Monday-Friday, **At these hours** 5, 10, 13, 16, **At these minutes** 0.

3. **+ > Data Operation > Compose.** Rename **Today**. Inputs, expression:
   `formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'yyyy-MM-dd')`

4. **Compose.** Rename **Scores**. Inputs: paste this text (not as an expression):

   ```json
   {"Category": {"Other": 1, "Training / People Development": 2, "Ergonomics / Operator Support": 3, "Controls / Automation / IT": 4, "Capital Project / Infrastructure": 5, "Energy / Utilities / Sustainability": 6, "Productivity / Waste Reduction": 7, "Quality": 8, "Reliability / Downtime Reduction": 9, "Customer Impact": 10, "Compliance / Regulatory": 11, "Safety": 12},
    "Necessity": {"Low": 1, "Medium": 2, "High": 3, "Urgent": 4},
    "Ease": {"Difficult": 1, "Moderate": 2, "Easy": 3},
    "Cost": {"High": 1, "Medium": 2, "Low": 3},
    "Time": {"> 1 Week": 1, "2-3 Days": 2, "1 Day": 3, "<1 Day": 4}}
   ```

   If you change a category's score in the spreadsheet sense, change it here.

5. **SharePoint > Get items.** Rename **Managers**. List: **People**. Filter Query:
   `IsManagement eq 1`

6. **Data Operation > Select.** Rename **ManagerNames**. From, expression:
   `body('Managers')?['value']`. Click the small **T** (switch to text mode) next to Map and
   enter the expression `item()?['Title']`

7. **SharePoint > Get items.** Rename **OpenItems**. List: **Action Items**. Filter Query:
   `TaskStatus ne 'Complete'`. Top Count: `5000`. Then in the action's **Settings**, turn
   **Pagination** on with Threshold `5000`.

8. **Control > Apply to each.** Rename **EachItem**. Select an output, expression:
   `body('OpenItems')?['value']`. Inside it, add:

   a. **Compose** **DaysLeft**:
      `if(empty(items('EachItem')?['TaskDueDate']), 999, div(sub(ticks(formatDateTime(items('EachItem')?['TaskDueDate'], 'yyyy-MM-dd')), ticks(outputs('Today'))), 864000000000))`

   b. **Compose** **DuePoints**:
      `if(less(outputs('DaysLeft'), 0), 10, if(equals(outputs('DaysLeft'), 0), 8, if(equals(outputs('DaysLeft'), 1), 6, if(lessOrEquals(outputs('DaysLeft'), 3), 4, if(lessOrEquals(outputs('DaysLeft'), 7), 2, 0)))))`

   c. **Compose** **NewScore**:
      `add(add(add(coalesce(outputs('Scores')?['Category']?[coalesce(items('EachItem')?['TaskCategory']?['Value'], '')], 0), mul(3, coalesce(outputs('Scores')?['Necessity']?[coalesce(items('EachItem')?['Necessity']?['Value'], '')], 0))), if(contains(body('ManagerNames'), coalesce(items('EachItem')?['Requester'], '')), 3, 0)), outputs('DuePoints'))`

   d. **Compose** **NewQuickWin**:
      `add(add(coalesce(outputs('Scores')?['Ease']?[coalesce(items('EachItem')?['Ease']?['Value'], '')], 0), coalesce(outputs('Scores')?['Cost']?[coalesce(items('EachItem')?['CostLevel']?['Value'], '')], 0)), coalesce(outputs('Scores')?['Time']?[coalesce(items('EachItem')?['EffortTime']?['Value'], '')], 0))`

   e. **Condition**: left, expression:
      `or(not(equals(float(coalesce(items('EachItem')?['PriorityScore'], -99)), float(outputs('NewScore')))), not(equals(float(coalesce(items('EachItem')?['QuickWin'], -99)), float(outputs('NewQuickWin')))))`
      / **is equal to** / right: expression `true`

   f. In **True**: **SharePoint > Update item**. List: **Action Items**.
      - Id: `items('EachItem')?['ID']`
      - Task: `items('EachItem')?['Title']`
      - Priority Score: `outputs('NewScore')`
      - Quick-Win Score: `outputs('NewQuickWin')`
      - Leave every other field **empty** (empty means "don't change").

9. After the Apply to each (outside it): **SharePoint > Get items.** Rename **JustDone**.
   List: **Action Items**. Filter Query:
   `TaskStatus eq 'Complete' and TaskCompletedDate eq null`

10. **Apply to each** **EachDone** over `body('JustDone')?['value']`, inside it **Update
    item** (Action Items):
    - Id: `items('EachDone')?['ID']`
    - Task: `items('EachDone')?['Title']`
    - Completion Date: `outputs('Today')`
    - Priority Score: `-1`

11. **Save**, then **Run** it once. Afterwards the Priority Scores in *My Open Tasks* should
    match what the spreadsheet would show today.

## If something doesn't work

- **"Scores" errors with "property … doesn't exist on type String":** the Compose kept your
  text as text. Replace `outputs('Scores')` in the four expressions with
  `json(string(outputs('Scores')))`.
- **Every item updates on every run:** harmless, only uses more of the daily actions. Tell
  me and we'll find out why.
- **Some scores look one day off:** the time zone. Check the Recurrence trigger and the two
  `Eastern Standard Time` texts.
