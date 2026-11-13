/*
 * Work Hub: creates the three Microsoft Lists (Action Items, People, Recurring Work)
 * on a SharePoint site, with their columns, views and color rules, and optionally
 * loads your data into them.
 *
 * How to run it (see setup/01-create-the-lists.md for the full steps):
 *   1. Open your SharePoint site in Edge or Chrome on your work PC.
 *   2. Press F12, open the Console tab.
 *   3. Paste the whole private file you were given (it holds your data followed by
 *      this script) and press Enter. Or paste only this file to create empty lists.
 *
 * It is safe to run again: lists, columns and views that already exist are left alone,
 * and data is only loaded into a list that is still empty.
 *
 * This file holds no personal data. Your data comes from window.WORKHUB_DATA, which the
 * private file sets before this script runs.
 */
(async function workHubSetup() {
  "use strict";

  const DATA = window.WORKHUB_DATA || { config: {}, tasks: [], people: [], recurring: [] };
  const CONFIG = DATA.config || {};

  // ---------------------------------------------------------------- choice lists
  const STATUS = ["Not Started", "In Progress", "Waiting / Blocked", "Ready for Testing", "On Hold", "Complete"];
  const CATEGORY = [ // score = position (1..12), same as the Drop Downs tab
    "Other", "Training / People Development", "Ergonomics / Operator Support",
    "Controls / Automation / IT", "Capital Project / Infrastructure",
    "Energy / Utilities / Sustainability", "Productivity / Waste Reduction", "Quality",
    "Reliability / Downtime Reduction", "Customer Impact", "Compliance / Regulatory", "Safety"];
  const NECESSITY = ["Low", "Medium", "High", "Urgent"];   // 1..4, counted 3 times
  const EASE = ["Difficult", "Moderate", "Easy"];          // 1..3, tie-break only
  const COST = ["High", "Medium", "Low"];                  // 1..3, tie-break only
  const TIME = ["> 1 Week", "2-3 Days", "1 Day", "<1 Day"]; // 1..4, tie-break only
  const SOURCE = ["Manual", "Email", "Meeting", "L2L", "OneNote", "Recurring", "Imported"];
  const RULE_TYPE = ["Nth workday of month", "Annual date", "Manual (set Next Due by hand)"];

  // ---------------------------------------------------------------- scoring
  // Priority Score = Category + 3 x Necessity + Management Bonus + Due Date Points.
  // Exactly the formula of the spreadsheet's Priority Score column. The
  // "Priority refresh" flow keeps it up to date every hour; this only fills it in
  // for the imported rows so the list is sorted correctly from the first minute.
  function dayNumber(isoDate) { // days since 1970 for a yyyy-mm-dd date
    const [y, m, d] = isoDate.slice(0, 10).split("-").map(Number);
    return Math.round(Date.UTC(y, m - 1, d) / 86400000);
  }
  function localToday() {
    const n = new Date();
    return n.getFullYear() + "-" + String(n.getMonth() + 1).padStart(2, "0") + "-" + String(n.getDate()).padStart(2, "0");
  }
  function duePoints(due, today) {
    if (!due) return 0;
    const days = dayNumber(due) - dayNumber(today);
    if (days < 0) return 10;
    if (days === 0) return 8;
    if (days === 1) return 6;
    if (days <= 3) return 4;
    if (days <= 7) return 2;
    return 0;
  }
  function score(list, value) { const i = list.indexOf(value); return i < 0 ? 0 : i + 1; }
  function priorityScore(t, managers, today) {
    if (t.status === "Complete") return -1;
    return score(CATEGORY, t.category) + 3 * score(NECESSITY, t.necessity)
      + (managers.has(t.requester) ? 3 : 0) + duePoints(t.due, today);
  }
  function quickWin(t) { return score(EASE, t.ease) + score(COST, t.cost) + score(TIME, t.time); }
  // Exposed so the numbers can be checked against the spreadsheet.
  window.WORKHUB_SCORING = { priorityScore, quickWin, duePoints };
  if (DATA.scoreOnly) return;

  // ---------------------------------------------------------------- SharePoint REST
  function findSiteUrl() {
    if (CONFIG.siteUrl) return CONFIG.siteUrl.replace(/\/+$/, "");
    const ctx = window._spPageContextInfo;
    if (ctx && ctx.webAbsoluteUrl) return ctx.webAbsoluteUrl;
    const m = location.pathname.match(/^\/(sites|teams|personal)\/[^\/]+/i);
    if (m) return location.origin + m[0];
    const typed = prompt("Paste the address of your SharePoint site (for example https://company.sharepoint.com/sites/WorkHub):");
    if (!typed) throw new Error("No site address, nothing was changed.");
    return typed.replace(/\/+$/, "");
  }
  const SITE = findSiteUrl();
  const API = SITE + "/_api";
  let digest = null;

  async function getDigest() {
    const r = await fetch(API + "/contextinfo", { method: "POST", headers: { Accept: "application/json;odata=verbose" }, credentials: "include" });
    if (!r.ok) throw new Error("Could not sign in to " + SITE + " (" + r.status + "). Are you on a page of that site?");
    digest = (await r.json()).d.GetContextWebInformation.FormDigestValue;
  }
  async function sp(method, path, body, extra) {
    const headers = Object.assign({ Accept: "application/json;odata=verbose" }, extra || {});
    if (method !== "GET") {
      headers["X-RequestDigest"] = digest;
      headers["Content-Type"] = "application/json;odata=verbose";
    }
    const r = await fetch(API + path, { method: method === "GET" ? "GET" : "POST", headers, credentials: "include", body: body === undefined ? undefined : JSON.stringify(body) });
    if (r.status === 404 && method === "GET") return null;
    if (!r.ok) {
      const text = await r.text();
      throw new Error(method + " " + path + " failed (" + r.status + "): " + text.slice(0, 400));
    }
    if (r.status === 204) return null;
    const text = await r.text();
    return text ? JSON.parse(text).d : null;
  }
  const merge = (path, body) => sp("POST", path, body, { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" });
  // SharePoint keeps view and column formatting inside XML, so <, > and & are sent as
  // JSON \u escapes (the same JSON, but nothing for the XML reader to trip on).
  const formatJson = o => JSON.stringify(o).replace(/</g, "\\u003c").replace(/>/g, "\\u003e").replace(/&/g, "\\u0026");
  const xml = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
  const q = s => String(s).replace(/'/g, "''"); // inside '...' in a REST URL
  const log = (...a) => console.log("%c[Work Hub]", "color:#0a7d55;font-weight:bold", ...a);

  // ---------------------------------------------------------------- column definitions
  // [internal name, display name, type, options]
  // Internal names never change; flows and color rules use them. Display names can be
  // renamed in Lists at any time without breaking anything.
  const withExtra = (base, values) => base.concat([...new Set(values.filter(v => v && !base.includes(v)))]);
  const tasks = DATA.tasks || [];
  const ACTION_FIELDS = [
    ["Details", "Description", "Note"],
    ["HurdlesNotes", "Hurdles / Needs / Notes", "Note"],
    ["Requester", "Requester", "Text"],
    ["AssigneeNames", "Assigned To", "Text", { default: CONFIG.myName || "" }],
    ["TaskStatus", "Status", "Choice", { choices: withExtra(STATUS, tasks.map(t => t.status)), default: "Not Started" }],
    ["DateAssigned", "Date Assigned", "Date", { today: true }],
    ["TaskDueDate", "Due Date", "Date"],
    ["TaskCompletedDate", "Completion Date", "Date"],
    ["PriorityScore", "Priority Score", "Number"],
    ["TaskCategory", "Category", "Choice", { choices: withExtra(CATEGORY, tasks.map(t => t.category)) }],
    ["Necessity", "Necessity", "Choice", { choices: withExtra(NECESSITY, tasks.map(t => t.necessity)) }],
    ["Ease", "Ease", "Choice", { choices: withExtra(EASE, tasks.map(t => t.ease)) }],
    ["CostLevel", "Cost", "Choice", { choices: withExtra(COST, tasks.map(t => t.cost)) }],
    ["EffortTime", "Time", "Choice", { choices: withExtra(TIME, tasks.map(t => t.time)) }],
    ["QuickWin", "Quick-Win Score", "Number"],
    ["ItemSource", "Source", "Choice", { choices: SOURCE, default: "Manual" }],
    ["ItemSourceId", "Source ID", "Text"],
    ["RecurringId", "Recurring ID", "Number"],
  ];
  const PEOPLE_FIELDS = [
    ["PersonEmail", "Email", "Text"],
    ["IsManagement", "Management (adds 3 to priority)", "Boolean"],
    ["IsRequester", "Requester", "Boolean", { default: "1" }],
    ["IsAssignee", "Can be assigned", "Boolean"],
  ];
  const RECURRING_FIELDS = [
    ["Workstream", "Workstream", "Text"],
    ["Deliverable", "Actual Task / Deliverable", "Note"],
    ["Owners", "Owner(s)", "Text", { default: CONFIG.myName || "" }],
    ["Frequency", "Frequency", "Text"],
    ["TimingRule", "Timing / Date Rule", "Text"],
    ["RuleType", "Schedule Type", "Choice", { choices: RULE_TYPE, default: RULE_TYPE[2] }],
    ["RuleN", "Workday Number (N)", "Number"],
    ["LeadDays", "Create Task This Many Days Early", "Number", { default: "5" }],
    ["NextDue", "Next Due", "Date"],
    ["IsActive", "Active", "Boolean", { default: "1" }],
    ["RecSource", "Source", "Text"],
    ["RecNotes", "Action / Notes", "Note"],
  ];

  function fieldXml([name, , type, o = {}]) {
    const base = `Name="${name}" StaticName="${name}" DisplayName="${name}"`;
    if (type === "Note") return `<Field Type="Note" ${base} NumLines="6" RichText="FALSE" UnlimitedLengthInDocumentLibrary="FALSE" />`;
    if (type === "Text") return `<Field Type="Text" ${base} MaxLength="255">${o.default ? `<Default>${xml(o.default)}</Default>` : ""}</Field>`;
    if (type === "Number") return `<Field Type="Number" ${base} Decimals="0">${o.default ? `<Default>${xml(o.default)}</Default>` : ""}</Field>`;
    if (type === "Boolean") return `<Field Type="Boolean" ${base}><Default>${o.default || "0"}</Default></Field>`;
    if (type === "Date") return `<Field Type="DateTime" ${base} Format="DateOnly">${o.today ? "<Default>[today]</Default>" : ""}</Field>`;
    if (type === "Choice") return `<Field Type="Choice" ${base} Format="Dropdown" FillInChoice="FALSE">${o.default ? `<Default>${xml(o.default)}</Default>` : ""}<CHOICES>${o.choices.map(c => `<CHOICE>${xml(c)}</CHOICE>`).join("")}</CHOICES></Field>`;
    throw new Error("Unknown column type " + type);
  }

  // ---------------------------------------------------------------- color rules
  // Row colors in the task views: red = past due, yellow = due today (not for completed
  // items, which are grayed). Priority Score is bold red at 22 or more, like the sheet.
  const ROW_FORMAT = {
    "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/row-formatting.schema.json",
    "additionalRowClass": "=if([$TaskStatus] == 'Complete', 'ms-fontColor-neutralTertiary', if([$TaskDueDate] == '', '', if(toLocaleDateString([$TaskDueDate]) == toLocaleDateString(@now), 'sp-field-severity--warning', if(@now > [$TaskDueDate], 'sp-field-severity--severeWarning', ''))))"
  };
  const SCORE_FORMAT = {
    "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
    "elmType": "div",
    "txtContent": "@currentField",
    "style": {
      "font-weight": "=if(@currentField >= 22, 'bold', 'normal')",
      "color": "=if(@currentField >= 22, '#a4262c', '')"
    }
  };
  const STATUS_FORMAT = {
    "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
    "elmType": "div",
    "txtContent": "@currentField",
    "attributes": {
      "class": "=if(@currentField == 'Complete', 'sp-field-severity--good', if(@currentField == 'In Progress', 'sp-field-severity--low', if(@currentField == 'Waiting / Blocked' || @currentField == 'On Hold', 'sp-field-severity--blocked', '')))"
    },
    "style": { "padding": "0 6px", "border-radius": "10px" }
  };

  // ---------------------------------------------------------------- views
  const me = CONFIG.myName ? xml(CONFIG.myName) : "";
  const OPEN = "<Neq><FieldRef Name='TaskStatus'/><Value Type='Choice'>Complete</Value></Neq>";
  const ORDER = "<OrderBy><FieldRef Name='PriorityScore' Ascending='FALSE'/><FieldRef Name='QuickWin' Ascending='FALSE'/><FieldRef Name='TaskDueDate' Ascending='TRUE'/></OrderBy>";
  const TASK_COLUMNS = ["LinkTitle", "AssigneeNames", "TaskStatus", "TaskDueDate", "PriorityScore", "Requester", "Details", "HurdlesNotes", "ItemSource"];
  const ACTION_VIEWS = [
    me && { title: "My Open Tasks", isDefault: true, columns: TASK_COLUMNS,
      query: `<Where><And><Contains><FieldRef Name='AssigneeNames'/><Value Type='Text'>${me}</Value></Contains>${OPEN}</And></Where>${ORDER}` },
    me && { title: "Others' Open Tasks", columns: TASK_COLUMNS,
      query: `<Where><And><Neq><FieldRef Name='AssigneeNames'/><Value Type='Text'>${me}</Value></Neq>${OPEN}</And></Where>${ORDER}` },
    { title: "All Open Tasks", isDefault: !me, columns: TASK_COLUMNS, query: `<Where>${OPEN}</Where>${ORDER}` },
    { title: "Due in 7 Days or Past Due", columns: TASK_COLUMNS,
      query: `<Where><And><Leq><FieldRef Name='TaskDueDate'/><Value Type='DateTime'><Today OffsetDays='7'/></Value></Leq>${OPEN}</And></Where><OrderBy><FieldRef Name='TaskDueDate' Ascending='TRUE'/><FieldRef Name='PriorityScore' Ascending='FALSE'/></OrderBy>` },
    { title: "To Sort (no Category or Necessity)", columns: ["LinkTitle", "ItemSource", "Requester", "AssigneeNames", "TaskCategory", "Necessity", "TaskDueDate", "Details"],
      query: `<Where><And><Or><IsNull><FieldRef Name='TaskCategory'/></IsNull><IsNull><FieldRef Name='Necessity'/></IsNull></Or>${OPEN}</And></Where><OrderBy><FieldRef Name='Created' Ascending='FALSE'/></OrderBy>` },
    { title: "Completed", columns: ["LinkTitle", "AssigneeNames", "TaskCompletedDate", "Requester", "Details"],
      query: `<Where><Eq><FieldRef Name='TaskStatus'/><Value Type='Choice'>Complete</Value></Eq></Where><OrderBy><FieldRef Name='TaskCompletedDate' Ascending='FALSE'/></OrderBy>` },
  ].filter(Boolean);

  // ---------------------------------------------------------------- building blocks
  async function ensureList(urlName, title, description) {
    let list = await sp("GET", `/web/lists/getbytitle('${q(title)}')?$select=Id,Title,ItemCount`);
    if (!list) list = await sp("GET", `/web/lists/getbytitle('${q(urlName)}')?$select=Id,Title,ItemCount`);
    if (list) { log(`List "${title}" already exists, keeping it.`); return list; }
    // Created under its short name so its web address has no spaces, then renamed.
    list = await sp("POST", "/web/lists", { __metadata: { type: "SP.List" }, BaseTemplate: 100, Title: urlName, Description: description, AllowContentTypes: false });
    await merge(`/web/lists(guid'${list.Id}')`, { __metadata: { type: "SP.List" }, Title: title });
    log(`Created list "${title}".`);
    return sp("GET", `/web/lists(guid'${list.Id}')?$select=Id,Title,ItemCount`);
  }
  async function ensureFields(list, defs, titleDisplay) {
    const L = `/web/lists(guid'${list.Id}')`;
    const existing = await sp("GET", `${L}/fields?$select=InternalName&$filter=Hidden eq false`);
    const have = new Set(existing.results.map(f => f.InternalName));
    for (const def of defs) {
      if (have.has(def[0])) continue;
      // Options 8 = use the Name attribute as the internal name.
      await sp("POST", `${L}/fields/createfieldasxml`, { parameters: { __metadata: { type: "SP.XmlSchemaFieldCreationInformation" }, SchemaXml: fieldXml(def), Options: 8 } });
      await merge(`${L}/fields/getbyinternalnameortitle('${def[0]}')`, { __metadata: { type: "SP.Field" }, Title: def[1] });
    }
    if (titleDisplay) await merge(`${L}/fields/getbyinternalnameortitle('Title')`, { __metadata: { type: "SP.Field" }, Title: titleDisplay });
  }
  async function setColumnFormat(list, field, format) {
    await merge(`/web/lists(guid'${list.Id}')/fields/getbyinternalnameortitle('${field}')`, { __metadata: { type: "SP.Field" }, CustomFormatter: formatJson(format) });
  }
  async function setViewColumns(L, viewId, columns) {
    const V = `${L}/views(guid'${viewId}')`;
    await sp("POST", `${V}/viewfields/removeallviewfields`);
    for (const c of columns) await sp("POST", `${V}/viewfields/addviewfield('${c}')`);
  }
  async function ensureViews(list, views, rowFormat) {
    const L = `/web/lists(guid'${list.Id}')`;
    const existing = await sp("GET", `${L}/views?$select=Id,Title`);
    const have = new Map(existing.results.map(v => [v.Title, v.Id]));
    for (const v of views) {
      let viewId = have.get(v.title);
      if (!viewId) {
        const made = await sp("POST", `${L}/views`, { __metadata: { type: "SP.View" }, Title: v.title, PersonalView: false, ViewQuery: v.query, RowLimit: 100, Paged: true });
        viewId = made.Id;
        await setViewColumns(L, viewId, v.columns);
        log(`  view "${v.title}"`);
      }
      // Applied to existing views too, so a run that stopped half-way is finished off.
      const extra = { __metadata: { type: "SP.View" } };
      if (rowFormat) extra.CustomFormatter = formatJson(rowFormat);
      if (v.isDefault) extra.DefaultView = true;
      if (Object.keys(extra).length > 1) await merge(`${L}/views(guid'${viewId}')`, extra);
    }
  }
  async function setDefaultViewColumns(list, columns) {
    const L = `/web/lists(guid'${list.Id}')`;
    const all = await sp("GET", `${L}/views/getbytitle('All Items')?$select=Id`);
    if (all) await setViewColumns(L, all.Id, columns);
  }
  async function addItems(list, rows, label) {
    if (!rows.length) return [];
    const fresh = await sp("GET", `/web/lists(guid'${list.Id}')?$select=ItemCount,ListItemEntityTypeFullName`);
    if (fresh.ItemCount > 0) { log(`"${list.Title}" already has ${fresh.ItemCount} items, not loading ${label} again.`); return null; }
    const made = [];
    for (let i = 0; i < rows.length; i++) {
      const body = Object.assign({ __metadata: { type: fresh.ListItemEntityTypeFullName } }, rows[i]);
      made.push(await sp("POST", `/web/lists(guid'${list.Id}')/items`, body));
      if ((i + 1) % 20 === 0) log(`  ${i + 1} of ${rows.length} ${label}`);
    }
    log(`Loaded ${rows.length} ${label}.`);
    return made;
  }
  // Date-only values are sent at noon UTC so the same calendar day shows in every
  // US time zone.
  const asDate = d => d ? d.slice(0, 10) + "T12:00:00Z" : null;
  const clean = o => { for (const k of Object.keys(o)) if (o[k] === null || o[k] === undefined || o[k] === "") delete o[k]; return o; };

  // ---------------------------------------------------------------- run
  log("Setting up on " + SITE);
  await getDigest();

  const people = await ensureList("People", "People", "Requesters and assignees. Tick Management for people whose requests add 3 to the priority score.");
  await ensureFields(people, PEOPLE_FIELDS, "Name");
  await setDefaultViewColumns(people, ["LinkTitle", "PersonEmail", "IsManagement", "IsRequester", "IsAssignee"]);

  const actions = await ensureList("ActionItems", "Action Items", "Every action item from email, meetings, L2L, notes and requests.");
  await ensureFields(actions, ACTION_FIELDS, "Task");
  await setColumnFormat(actions, "PriorityScore", SCORE_FORMAT);
  await setColumnFormat(actions, "TaskStatus", STATUS_FORMAT);
  await ensureViews(actions, ACTION_VIEWS, ROW_FORMAT);

  const recurring = await ensureList("RecurringWork", "Recurring Work", "Repeating duties. The Recurring tasks flow turns each one into an action item before it is due.");
  await ensureFields(recurring, RECURRING_FIELDS, "Recurring Work");
  await ensureViews(recurring, [{ title: "Active by Next Due", isDefault: true,
    columns: ["LinkTitle", "Owners", "Frequency", "TimingRule", "RuleType", "RuleN", "NextDue", "LeadDays", "IsActive"],
    query: "<Where><Eq><FieldRef Name='IsActive'/><Value Type='Boolean'>1</Value></Eq></Where><OrderBy><FieldRef Name='NextDue' Ascending='TRUE'/></OrderBy>" }]);

  // Data
  const today = localToday();
  const managers = new Set((DATA.people || []).filter(p => p.management).map(p => p.name));
  await addItems(people, (DATA.people || []).map(p => clean({
    Title: p.name, PersonEmail: p.email, IsManagement: !!p.management, IsRequester: !!p.requester, IsAssignee: !!p.assignee })), "people");
  await addItems(actions, tasks.map(t => clean({
    Title: t.task.slice(0, 255), Details: t.description, HurdlesNotes: t.notes, Requester: t.requester,
    AssigneeNames: t.assignedTo, TaskStatus: t.status || "Not Started", DateAssigned: asDate(t.assigned),
    TaskDueDate: asDate(t.due), TaskCompletedDate: asDate(t.completed),
    PriorityScore: priorityScore(t, managers, today), QuickWin: quickWin(t),
    TaskCategory: t.category, Necessity: t.necessity, Ease: t.ease, CostLevel: t.cost, EffortTime: t.time,
    ItemSource: "Imported" })), "action items");
  await addItems(recurring, (DATA.recurring || []).map(r => clean({
    Title: r.title.slice(0, 255), Workstream: r.workstream, Deliverable: r.deliverable, Owners: r.owners,
    Frequency: r.frequency, TimingRule: r.rule && r.rule.slice(0, 255), RuleType: r.ruleType, RuleN: r.ruleN,
    LeadDays: r.leadDays, NextDue: asDate(r.nextDue), IsActive: true, RecSource: r.source && r.source.slice(0, 255),
    RecNotes: r.notes })), "recurring items");

  log("Done. Open " + SITE + "/Lists/ActionItems to see your action items.");
  alert("Work Hub setup finished. Open Site contents to see Action Items, People and Recurring Work.");
})().catch(e => { console.error("[Work Hub] Setup stopped:", e); alert("Work Hub setup stopped: " + e.message); });
