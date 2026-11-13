# Flow 1: Flagged email becomes an action item

**What it does:** you flag an email in Outlook (PC, iPad or phone) and a few minutes later
it is in *Action Items*. Source = Email, Task = the subject, Description = the start of the
email, Requester = the sender. It then shows in **To Sort** until you give it a Category and
Necessity. Flagging the same email again, or changing it, never makes a second item.

No forwarding rule, nothing leaves Microsoft 365.

## Conventions used in all four flow guides

- Build in **make.powerautomate.com** on your PC.
- **Rename each action** to the name given in **bold** (click the action, click its title at
  the top of the panel). The expressions refer to actions by these names; with other names
  they fail.
- **Expressions**: click in the box, choose the **fx** (expression) button, paste the
  expression, **Add**. Paste them exactly. In the guides, `Eastern Standard Time` is the
  time zone (it switches to daylight time by itself). If you are in another zone, use its
  Windows name, for example `Central Standard Time`.
- **Site Address**: pick your Work Hub site (or OneDrive: choose *Enter custom value* and
  paste its address).

## Build it

1. **Create > Automated cloud flow.** Name: *Work Hub - Flagged email*. Trigger: search
   "flagged", choose **When an email is flagged (V3)** (Office 365 Outlook). **Create**.
2. Trigger settings: **Folder: Inbox**. Leave the rest as is.

3. **+ > SharePoint > Get items.** Rename to **FindExisting**.
   - Site Address: your site. List Name: **Action Items**.
   - Advanced parameters > **Filter Query**: type `ItemSourceId eq '` then insert the
     expression `triggerOutputs()?['body/id']` then type `'`
   - **Top Count**: `1`

4. **+ > Control > Condition.** Left box, expression:
   `length(body('FindExisting')?['value'])` / **is equal to** / right box: `0`

5. In **True**:

   a. **SharePoint > Get items.** Rename to **FindSender**. List: **People**.
      - Filter Query: type `PersonEmail eq '` then the expression
        `replace(triggerOutputs()?['body/from'], '''', '''''')` then type `'`
      - Top Count: `1`

   b. **SharePoint > Create item.** List: **Action Items**. Fill in:

   | Field | Value (expression unless it says *type*) |
   |---|---|
   | Task | `if(greater(length(coalesce(triggerOutputs()?['body/subject'], '')), 250), substring(triggerOutputs()?['body/subject'], 0, 250), coalesce(triggerOutputs()?['body/subject'], '(no subject)'))` |
   | Description | `triggerOutputs()?['body/bodyPreview']` |
   | Hurdles / Needs / Notes | `concat('Email: ', coalesce(triggerOutputs()?['body/subject'], ''), ' - from ', triggerOutputs()?['body/from'], ', ', formatDateTime(convertFromUtc(triggerOutputs()?['body/receivedDateTime'], 'Eastern Standard Time'), 'MM/dd/yyyy'))` |
   | Requester | `if(empty(body('FindSender')?['value']), triggerOutputs()?['body/from'], first(body('FindSender')?['value'])?['Title'])` |
   | Assigned To | *type* your name, exactly as in the list |
   | Status Value | *choose* Not Started |
   | Date Assigned | `formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'yyyy-MM-dd')` |
   | Source Value | *choose* Email |
   | Source ID | `triggerOutputs()?['body/id']` |

   Leave **False** empty (the item already exists).

6. **Save.** Flag any email in Outlook, wait 1-5 minutes, and look in *Action Items > To
   Sort*. Flows show their runs under **My flows > (the flow) > Run history**.

## If something doesn't work

- **An expression box says the property doesn't exist:** instead of the expression, use
  **Dynamic content** from the trigger (Subject, From, Body Preview, Received Time, Message
  Id). The names above are my best knowledge of the V3 trigger's outputs, not checked
  against your tenant.
- **It ran but Requester is an address:** add that person's Email to *People*.
- **It runs twice for one flag:** that's the trigger firing on a later change; the
  FindExisting step stops the duplicate, so nothing to do.
- Only emails in the **Inbox** count. To flag from other folders, make a copy of the flow
  per folder.
