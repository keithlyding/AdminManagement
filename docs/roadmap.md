# Roadmap

## Done (version 1)

- Lists: Action Items, People, Recurring Work, with the spreadsheet's data and scoring.
- Flows: flagged email, priority refresh, recurring tasks, daily review.

## Next

### L2L dispatches
L2L has a REST API with a dispatches endpoint
([L2L API docs](https://support.leading2lean.com/hc/en-us/articles/218236297)). Calling it
from Power Automate needs the HTTP connector, which is Premium. Two ways without Premium,
to check:
1. **L2L's notification emails.** If L2L emails you when a dispatch is assigned, a flow
   on those emails (like Flow 1, filtered on the sender) can create items with Source = L2L.
   The dispatch number goes in Source ID so it's never added twice.
2. **An L2L export or report** sent on a schedule to your mailbox or OneDrive.

Needs: a sample of an L2L dispatch email (with anything sensitive blanked out).

### OneNote
The OneNote connector can read pages, but handwritten ink may not come through as text.
Test first: a flow that lists the pages of one section, to see what a pencil-written page
gives back. Until then, the Daily review item covers OneNote.

### Power Apps front end (version 2)
A tablet app on the same lists: a **Today** screen (today's meetings beside the top
tasks), one-tap capture with Source preset, and a review screen. It runs inside the Power
Apps app from the App Store. Worth building once version 1 has been in daily use for a few
weeks and it's clear what the Lists app is missing.
