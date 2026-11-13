// Runs setup/sync-dropdowns.js against a stand-in for SharePoint holding the two columns as
// text boxes (as the first version made them), with made-up data.
//   node tests/sync-dropdowns.test.js
const vm = require("vm"), fs = require("fs"), path = require("path"), assert = require("assert");
const API = "https://x.sharepoint.com/sites/Hub/_api";
const fields = {
  Requester: { SchemaXml: '<Field Type="Text" DisplayName="Requester" Name="Requester" StaticName="Requester" MaxLength="255" ID="{1}" SourceID="{9}" />', TypeAsString: "Text" },
  AssigneeNames: { SchemaXml: '<Field Type="Text" DisplayName="Assigned To" Name="AssigneeNames" StaticName="AssigneeNames" MaxLength="255" ID="{2}" SourceID="{9}"><Default>Pat Example</Default></Field>', TypeAsString: "Text" },
};
const items = Array.from({ length: 7 }, (_, i) => ({ Id: i + 1, Requester: ["Boss Person", "Vendor & Co", null][i % 3], AssigneeNames: i % 2 ? "Pat Example" : "Sam Other" }));
const people = [{ Title: "Boss Person", IsRequester: true, IsAssignee: false }, { Title: "New Hire", IsRequester: false, IsAssignee: true }];
let loseOne = true; // pretend SharePoint drops one value in the conversion, to test the put-back
const res = (status, d) => ({ ok: status < 400, status, text: async () => d === undefined ? "" : JSON.stringify({ d }), json: async () => ({ d }) });
const fetch = async (url, opt) => {
  const p = decodeURIComponent(url.replace(API, "")); const verb = opt.headers["X-HTTP-Method"] || opt.method;
  const body = opt.body ? JSON.parse(opt.body) : undefined;
  if (p === "/contextinfo") return res(200, { GetContextWebInformation: { FormDigestValue: "d" } });
  if (p.startsWith("/web/lists/getbytitle('Action Items')?")) return res(200, { ListItemEntityTypeFullName: "SP.Data.ActionItemsListItem" });
  let m = p.match(/fields\/getbyinternalnameortitle\('(\w+)'\)/);
  if (m && verb === "GET") return res(200, fields[m[1]]);
  if (m && verb === "MERGE") {
    const x = body.SchemaXml;
    assert(/^<Field [^>]*Type="Choice"[^>]*>.*<\/Field>$/.test(x), "valid choice schema: " + x);
    assert(x.includes('ID="{') && x.includes('FillInChoice="TRUE"') && !x.includes("MaxLength"));
    assert(!/&(?!amp;|lt;|gt;|quot;|apos;)/.test(x), "ampersands escaped");
    fields[m[1]] = { SchemaXml: x, TypeAsString: "Choice" };
    if (loseOne && m[1] === "Requester") { items[0].Requester = null; loseOne = false; }
    return res(204);
  }
  if (p.includes("/items?$select=Id,Requester")) return res(200, { results: items.map(i => ({ ...i })) });
  if (p.startsWith("/web/lists/getbytitle('People')/items")) return res(200, { results: people });
  m = p.match(/items\((\d+)\)/);
  if (m && verb === "GET") { const i = items[m[1] - 1]; return res(200, { ...i, Requester: fields.Requester.TypeAsString === "Choice" && i.Requester ? i.Requester : i.Requester }); }
  if (m && verb === "MERGE") { Object.assign(items[m[1] - 1], body); delete items[m[1] - 1].__metadata; return res(204); }
  throw new Error("unmocked " + verb + " " + p);
};
const alerts = [];
const ctx = { console: { log: () => {}, error: (...a) => { console.error(...a); process.exitCode = 1; } }, alert: m => alerts.push(m), prompt: () => null, fetch,
  location: { pathname: "/sites/Hub/Lists/ActionItems/AllItems.aspx", origin: "https://x.sharepoint.com" }, WORKHUB_DATA: { config: { myName: "Pat Example" } } };
ctx.window = ctx; vm.createContext(ctx);
const src = fs.readFileSync(path.join(__dirname, "..", "setup", "sync-dropdowns.js"), "utf8");
(async () => {
  vm.runInContext(src, ctx); await new Promise(r => setTimeout(r, 200));
  const choices = n => [...fields[n].SchemaXml.matchAll(/<CHOICE>(.*?)<\/CHOICE>/g)].map(m => m[1]);
  assert.deepStrictEqual(choices("Requester"), ["Boss Person", "Vendor &amp; Co"]);
  assert.deepStrictEqual(choices("AssigneeNames"), ["New Hire", "Pat Example", "Sam Other"]);
  assert(fields.AssigneeNames.SchemaXml.includes("<Default>Pat Example</Default>"), "default kept");
  assert.strictEqual(items[0].Requester, "Boss Person", "lost value put back");
  // Run again after adding a person: refreshes the choices, no duplicate CHOICES block.
  people.push({ Title: "Another Boss", IsRequester: true });
  vm.runInContext(src, ctx); await new Promise(r => setTimeout(r, 200));
  assert.deepStrictEqual(choices("Requester"), ["Another Boss", "Boss Person", "Vendor &amp; Co"]);
  assert.strictEqual((fields.Requester.SchemaXml.match(/<CHOICES>/g) || []).length, 1);
  assert(alerts.every(a => !/stopped/i.test(a)), alerts.join(" | "));
  console.log("sync-dropdowns.test.js: all checks passed");
})().catch(e => { console.error(e); process.exit(1); });
