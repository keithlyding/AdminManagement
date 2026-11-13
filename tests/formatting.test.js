// Checks the row-color rules:
//  1. setup/create-lists.js and setup/apply-formatting.js hold the same rules,
//  2. the Action Items rule gives each kind of row the intended color (the rule is run
//     with a small translation of SharePoint's formula syntax into JavaScript),
//  3. apply-formatting.js sets them on every shared view, escaped for SharePoint's XML.
//   node tests/formatting.test.js
const fs = require("fs"), path = require("path"), vm = require("vm"), assert = require("assert");
const read = f => fs.readFileSync(path.join(__dirname, "..", "setup", f), "utf8");
const create = read("create-lists.js"), apply = read("apply-formatting.js");

// 1. Same rules in both scripts.
const ruleLines = src => src.split("\n").filter(l => /const ZEBRA =|"additionalRowClass":/.test(l)).map(l => l.trim());
assert.deepStrictEqual(ruleLines(apply), ruleLines(create), "row rules differ between the two scripts");

// 2. Evaluate the rule.
const ctx = { window: {}, location: { pathname: "/sites/Hub/x", origin: "https://x" } };
ctx.window = ctx; ctx.WORKHUB_FORMATS_ONLY = true; vm.createContext(ctx);
vm.runInContext(apply, ctx);
const rule = ctx.WORKHUB_FORMATS.ROW_FORMAT.additionalRowClass;
function classFor(row, rowIndex, now) {
  const js = rule.slice(1)
    .replace(/\[\$(\w+)\]/g, "r.$1").replace(/@rowIndex/g, "rowIndex").replace(/@now/g, "now")
    .replace(/\bif\(/g, "iff(");
  return new Function("r", "rowIndex", "now", "iff", "toLocaleDateString",
    "return " + js)(row, rowIndex, now, (c, a, b) => c ? a : b, d => new Date(d).toDateString()).trim();
}
const now = new Date(2026, 10, 13, 15, 0);
const day = n => new Date(2026, 10, 13 + n); // midnight, like a date-only column
const open = { TaskStatus: "In Progress", TaskDueDate: "", PriorityScore: 5 };
const cases = [
  [{ ...open, TaskDueDate: day(-1) }, 0, "sp-field-severity--blocked"],
  [{ ...open, TaskDueDate: day(-1), PriorityScore: 30 }, 1, "sp-field-severity--blocked"],
  [{ ...open, TaskDueDate: day(0) }, 0, "sp-field-severity--warning"],
  [{ ...open, TaskDueDate: day(3), PriorityScore: 22 }, 0, "sp-field-severity--severeWarning ms-fontWeight-semibold"],
  [{ ...open, PriorityScore: 21 }, 0, ""],
  [{ ...open, PriorityScore: 21 }, 1, "ms-bgColor-neutralLighter"],
  [{ ...open, TaskDueDate: day(5) }, 1, "ms-bgColor-neutralLighter"],
  [{ ...open, TaskStatus: "Complete", TaskDueDate: day(-9), PriorityScore: 40 }, 0, "ms-fontColor-neutralTertiary"],
  [{ ...open, TaskStatus: "Complete" }, 1, "ms-fontColor-neutralTertiary ms-bgColor-neutralLighter"],
];
for (const [row, i, want] of cases) assert.strictEqual(classFor(row, i, now), want, JSON.stringify(row) + " row " + i);

// 3. Apply to every shared view, escaped.
const views = { "Action Items": [{ Id: "a1", Title: "All Items" }, { Id: "a2", Title: "Mine" }, { Id: "a3", Title: "Private", PersonalView: true }],
  "Recurring Work": [{ Id: "r1", Title: "All Items" }], "People": [{ Id: "p1", Title: "All Items" }] };
const set = {};
const res = (s, d) => ({ ok: s < 400, status: s, text: async () => d === undefined ? "" : JSON.stringify({ d }), json: async () => ({ d }) });
const run = { window: {}, console: { log: () => {}, error: (...a) => { console.error(...a); process.exitCode = 1; } }, alert: () => {},
  location: { pathname: "/sites/Hub/Lists/ActionItems/AllItems.aspx", origin: "https://x.sharepoint.com" },
  fetch: async (url, opt) => {
    const p = decodeURIComponent(url.replace("https://x.sharepoint.com/sites/Hub/_api", ""));
    if (p === "/contextinfo") return res(200, { GetContextWebInformation: { FormDigestValue: "d" } });
    const t = p.match(/getbytitle\('([^']+)'\)/)[1];
    if (p.endsWith("PersonalView")) return res(200, { results: views[t] });
    const body = JSON.parse(opt.body);
    assert(!/[<>&]/.test(body.CustomFormatter), "formatter must be escaped for XML");
    set[t + "/" + p.match(/views\(guid'(\w+)'\)/)[1]] = JSON.parse(body.CustomFormatter);
    return res(204);
  } };
run.window = run; vm.createContext(run);
vm.runInContext(apply, run);
setTimeout(() => {
  assert.deepStrictEqual(Object.keys(set).sort(), ["Action Items/a1", "Action Items/a2", "People/p1", "Recurring Work/r1"]);
  assert.strictEqual(set["Action Items/a2"].additionalRowClass, rule);
  assert.strictEqual(set["People/p1"].additionalRowClass, "=if(@rowIndex % 2 == 1, 'ms-bgColor-neutralLighter', '')");
  console.log("formatting.test.js: all checks passed");
}, 100);
