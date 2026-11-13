/*
 * Work Hub: drop-downs for Requester and Assigned To.
 *
 * Makes the Action Items list's Requester and Assigned To columns drop-downs, filled from
 * the People list (people ticked "Requester" and "Can be assigned", plus every name already
 * used in the list). A name that isn't in the list can still be typed in.
 *
 * Run it:
 *   - once, to turn the two columns from text boxes into drop-downs (lists made with the
 *     first version of create-lists.js), and
 *   - again whenever you add people to the People list, to put them in the drop-downs.
 *
 * Same way as the setup: open the SharePoint site with the lists, F12 > Console, paste this
 * whole file, Enter. It holds no personal data.
 *
 * Safety: it copies every item's Requester and Assigned To first, checks them after the
 * change, and puts back any value that didn't survive.
 */
(async function workHubSyncDropdowns() {
  "use strict";
  const CONFIG = (window.WORKHUB_DATA && window.WORKHUB_DATA.config) || {};

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
  const log = (...a) => console.log("%c[Work Hub]", "color:#0a7d55;font-weight:bold", ...a);

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
    if (!r.ok) throw new Error(method + " " + path + " failed (" + r.status + "): " + (await r.text()).slice(0, 400));
    if (r.status === 204) return null;
    const text = await r.text();
    return text ? JSON.parse(text).d : null;
  }
  const merge = (path, body) => sp("POST", path, body, { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" });
  const xml = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
  const sortNames = arr => [...new Set(arr.map(v => (v || "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));

  // Turns a field's schema into a fill-in drop-down with these choices, keeping its ID,
  // names and default.
  function choiceSchema(schema, choices) {
    let s = schema
      .replace(/\sType="[^"]*"/, ' Type="Choice"')
      .replace(/\sMaxLength="[^"]*"/, "")
      .replace(/\sFormat="[^"]*"/, "")
      .replace(/\sFillInChoice="[^"]*"/, "")
      .replace(/<Field\b/, '<Field Format="Dropdown" FillInChoice="TRUE"');
    s = s.replace(/<CHOICES>[\s\S]*?<\/CHOICES>/, "").replace(/<CHOICES\s*\/>/, "");
    const block = "<CHOICES>" + choices.map(c => "<CHOICE>" + xml(c) + "</CHOICE>").join("") + "</CHOICES>";
    if (/\/>\s*$/.test(s)) return s.replace(/\s*\/>\s*$/, ">" + block + "</Field>");
    return s.replace(/<\/Field>\s*$/, block + "</Field>");
  }

  log("Drop-downs on " + SITE);
  await getDigest();
  const A = "/web/lists/getbytitle('Action Items')";
  const list = await sp("GET", A + "?$select=ListItemEntityTypeFullName");

  // 1. Remember every item's values.
  const items = [];
  let next = A + "/items?$select=Id,Requester,AssigneeNames&$top=500";
  while (next) {
    const page = await sp("GET", next);
    items.push(...page.results);
    next = page.__next ? page.__next.replace(API, "") : null;
  }
  const valueOf = v => (v && typeof v === "object" ? v.Value : v) || "";

  // 2. The choices: People ticked for the role, plus every name already in use.
  const people = (await sp("GET", "/web/lists/getbytitle('People')/items?$select=Title,IsRequester,IsAssignee&$top=5000")).results;
  const wanted = {
    Requester: sortNames(people.filter(p => p.IsRequester).map(p => p.Title).concat(items.map(i => valueOf(i.Requester)))),
    AssigneeNames: sortNames(people.filter(p => p.IsAssignee).map(p => p.Title).concat(items.map(i => valueOf(i.AssigneeNames)), [CONFIG.myName])),
  };

  // 3. Change the columns.
  for (const name of ["Requester", "AssigneeNames"]) {
    const f = await sp("GET", `${A}/fields/getbyinternalnameortitle('${name}')?$select=SchemaXml,TypeAsString`);
    await merge(`${A}/fields/getbyinternalnameortitle('${name}')`, { __metadata: { type: "SP.Field" }, SchemaXml: choiceSchema(f.SchemaXml, wanted[name]) });
    log(`${name === "Requester" ? "Requester" : "Assigned To"}: drop-down with ${wanted[name].length} names` + (f.TypeAsString === "Choice" ? " (refreshed)" : " (was a text box)"));
  }

  // 4. Check that no value was lost; put back any that was.
  let fixed = 0;
  for (const before of items) {
    const after = await sp("GET", `${A}/items(${before.Id})?$select=Id,Requester,AssigneeNames`);
    const patch = { __metadata: { type: list.ListItemEntityTypeFullName } };
    for (const k of ["Requester", "AssigneeNames"]) if (valueOf(after[k]) !== valueOf(before[k])) patch[k] = valueOf(before[k]) || null;
    if (Object.keys(patch).length > 1) { await merge(`${A}/items(${before.Id})`, patch); fixed++; }
  }
  log(`Checked ${items.length} items` + (fixed ? `, put back the values of ${fixed}.` : ", all values kept."));
  alert("Work Hub: Requester and Assigned To are now drop-downs (" + items.length + " items checked).");
})().catch(e => { console.error("[Work Hub] Stopped:", e); alert("Work Hub stopped: " + e.message); });
