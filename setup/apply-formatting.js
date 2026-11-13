/*
 * Work Hub: row colors.
 *
 * Applies the row colors to every view of your lists (including views you made yourself):
 *   Action Items, first match wins:
 *     past due ............................ red
 *     due today ........................... yellow
 *     Priority Score 22 or more ........... orange, bold
 *     everything else ..................... every other row light gray
 *     Complete ............................ gray text
 *   Recurring Work and People: every other row light gray.
 *
 * Run it on your SharePoint site: F12 > Console > paste this whole file > Enter. Run it
 * again after you make a new view. It holds no personal data and changes no items.
 */
(async function workHubApplyFormatting() {
  "use strict";
  // Keep these identical to setup/create-lists.js (tests/formatting.test.js checks it).
  const ZEBRA = "if(@rowIndex % 2 == 1, 'ms-bgColor-neutralLighter', '')";
  const ROW_FORMAT = {
    "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/row-formatting.schema.json",
    "additionalRowClass": "=if([$TaskStatus] == 'Complete', 'ms-fontColor-neutralTertiary ' + " + ZEBRA + ", if([$TaskDueDate] != '' && toLocaleDateString([$TaskDueDate]) == toLocaleDateString(@now), 'sp-field-severity--warning', if([$TaskDueDate] != '' && @now > [$TaskDueDate], 'sp-field-severity--blocked', if([$PriorityScore] >= 22, 'sp-field-severity--severeWarning ms-fontWeight-semibold', " + ZEBRA + "))))"
  };
  const STRIPES = {
    "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/row-formatting.schema.json",
    "additionalRowClass": "=" + ZEBRA
  };
  window.WORKHUB_FORMATS = { ROW_FORMAT, STRIPES };
  if (window.WORKHUB_FORMATS_ONLY) return;

  const ctx = window._spPageContextInfo;
  const m = location.pathname.match(/^\/(sites|teams|personal)\/[^\/]+/i);
  const SITE = (ctx && ctx.webAbsoluteUrl) || (m ? location.origin + m[0] : null);
  if (!SITE) throw new Error("Open a page of your Work Hub SharePoint site first.");
  const API = SITE + "/_api";
  const log = (...a) => console.log("%c[Work Hub]", "color:#0a7d55;font-weight:bold", ...a);
  // SharePoint keeps formatting inside XML: send <, > and & as JSON \u escapes.
  const formatJson = o => JSON.stringify(o).replace(/</g, "\\u003c").replace(/>/g, "\\u003e").replace(/&/g, "\\u0026");

  const r = await fetch(API + "/contextinfo", { method: "POST", headers: { Accept: "application/json;odata=verbose" }, credentials: "include" });
  if (!r.ok) throw new Error("Could not sign in to " + SITE + " (" + r.status + ").");
  const digest = (await r.json()).d.GetContextWebInformation.FormDigestValue;
  async function sp(method, path, body, extra) {
    const headers = Object.assign({ Accept: "application/json;odata=verbose" }, extra || {});
    if (method !== "GET") { headers["X-RequestDigest"] = digest; headers["Content-Type"] = "application/json;odata=verbose"; }
    const res = await fetch(API + path, { method: method === "GET" ? "GET" : "POST", headers, credentials: "include", body: body === undefined ? undefined : JSON.stringify(body) });
    if (res.status === 404 && method === "GET") return null;
    if (!res.ok) throw new Error(method + " " + path + " failed (" + res.status + "): " + (await res.text()).slice(0, 400));
    const text = res.status === 204 ? "" : await res.text();
    return text ? JSON.parse(text).d : null;
  }

  let total = 0;
  for (const [title, format] of [["Action Items", ROW_FORMAT], ["Recurring Work", STRIPES], ["People", STRIPES]]) {
    const L = `/web/lists/getbytitle('${title}')`;
    const views = await sp("GET", `${L}/views?$select=Id,Title,Hidden,PersonalView`);
    if (!views) { log(`No "${title}" list here, skipped.`); continue; }
    for (const v of views.results.filter(v => !v.Hidden && !v.PersonalView)) {
      await sp("POST", `${L}/views(guid'${v.Id}')`, { __metadata: { type: "SP.View" }, CustomFormatter: formatJson(format) }, { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" });
      total++;
    }
    log(`${title}: colors set on ${views.results.filter(v => !v.Hidden && !v.PersonalView).length} views.`);
  }
  alert("Work Hub: row colors set on " + total + " views. Refresh the list page to see them.");
})().catch(e => { console.error("[Work Hub] Stopped:", e); alert("Work Hub stopped: " + e.message); });
