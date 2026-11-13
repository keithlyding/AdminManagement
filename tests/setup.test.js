// Runs setup/create-lists.js against an in-memory stand-in for SharePoint's REST API,
// with made-up data, and checks the lists, columns, views, items and scores it makes.
//   node tests/setup.test.js
// It catches mistakes in the script's own logic; it can't prove that real SharePoint
// accepts every call the same way.
const vm=require('vm'),fs=require('fs');
const lists={}; let id=0; const calls=[];
const guid=()=>'00000000-0000-0000-0000-'+String(++id).padStart(12,'0');
function find(path){ let m=path.match(/lists\(guid'([^']+)'\)/); if(m) return lists[m[1]];
  m=path.match(/getbytitle\('((?:[^']|'')+)'\)/); if(m){const t=m[1].replace(/''/g,"'"); return Object.values(lists).find(l=>l.Title===t);} }
function res(status,d){return {ok:status<400,status,text:async()=>d===undefined?'':JSON.stringify({d}),json:async()=>({d})};}
global.fetch=async(url,opt)=>{
  const path=url.replace('https://x.sharepoint.com/sites/Hub/_api',''); const body=opt.body?JSON.parse(opt.body):undefined;
  const verb=(opt.headers['X-HTTP-Method']||opt.method); calls.push(verb+' '+path);
  if(path==='/contextinfo') return res(200,{GetContextWebInformation:{FormDigestValue:'dig'}});
  if(path==='/web/lists' && verb==='POST'){const g=guid(); lists[g]={Id:g,Title:body.Title,fields:new Map([['Title',{Title:'Title'}]]),views:new Map([['All Items',{Id:guid(),fields:[]}]]),items:[]}; return res(201,{Id:g});}
  const L=find(path);
  if(path.startsWith('/web/lists') && !L) return res(404);
  if(/\/web\/lists(\(guid'[^']+'\)|\/getbytitle\('[^']+'\))\?/.test(path)) return res(200,{Id:L.Id,Title:L.Title,ItemCount:L.items.length,ListItemEntityTypeFullName:'SP.Data.'+L.Id+'Item'});
  if(verb==='MERGE' && /lists\(guid'[^']+'\)$/.test(path)){L.Title=body.Title;return res(204);}
  if(path.includes('/fields?')) return res(200,{results:[...L.fields.keys()].map(k=>({InternalName:k}))});
  if(path.endsWith('/fields/createfieldasxml')){const x=body.parameters.SchemaXml; const n=x.match(/Name="([^"]+)"/)[1]; if(L.fields.has(n)) throw new Error('dup field '+n);
     if(!/^<Field [^>]*(\/>|>.*<\/Field>)$/.test(x)) throw new Error('bad xml '+x); L.fields.set(n,{xml:x}); return res(201,{});}
  let m=path.match(/fields\/getbyinternalnameortitle\('([^']+)'\)$/);
  if(m){ if(!L.fields.has(m[1])) return res(404); if(/[<>&]/.test(body.CustomFormatter||'')) return res(500,{error:'XmlException'}); Object.assign(L.fields.get(m[1]),body); return res(204);}
  if(path.endsWith('/views?$select=Id,Title')) return res(200,{results:[...L.views].map(([t,v])=>({Title:t,Id:v.Id}))});
  if(path.endsWith('/views') && verb==='POST'){const g=guid(); L.views.set(body.Title,{Id:g,fields:[],query:body.ViewQuery}); return res(201,{Id:g});}
  m=path.match(/views\/getbytitle\('([^']+)'\)/); if(m){const v=L.views.get(m[1]); return v?res(200,{Id:v.Id}):res(404);}
  m=path.match(/views\(guid'([^']+)'\)(.*)$/);
  if(m){const v=[...L.views.values()].find(v=>v.Id===m[1]); if(!v) return res(404);
     if(m[2]==='/viewfields/removeallviewfields'){v.fields=[];return res(200,{});}
     const f=m[2].match(/addviewfield\('([^']+)'\)/); if(f){ if(f[1]!=='LinkTitle'&&!L.fields.has(f[1])) throw new Error('view field missing '+f[1]); v.fields.push(f[1]); return res(200,{});}
     if(verb==='MERGE'){ if(/[<>&]/.test(body.CustomFormatter||'')) return res(500,{error:'XmlException: raw < > or & in CustomFormatter'});
       Object.assign(v,body);return res(204);} }
  if(path.endsWith('/items') && verb==='POST'){ for(const k of Object.keys(body)) if(k!=='__metadata'&&k!=='Title'&&!L.fields.has(k)) throw new Error('item field missing '+k);
     L.items.push(body); return res(201,{Id:L.items.length});}
  throw new Error('unmocked '+verb+' '+path);
};
const ctx={window:{},console:{log:()=>{},error:(...a)=>{console.error(...a);process.exitCode=1}},alert:m=>console.log('ALERT:',m),prompt:()=>null,fetch:global.fetch,location:{pathname:'/sites/Hub/SitePages/Home.aspx',origin:'https://x.sharepoint.com'},JSON,Object,Set,Map,String,Math,Date,Number,Error,Promise,Array,RegExp};
ctx.window=ctx; vm.createContext(ctx);
const assert=require('assert');
const today=new Date(); const iso=d=>d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
const plus=n=>{const d=new Date(today);d.setDate(d.getDate()+n);return iso(d);};
ctx.WORKHUB_DATA={config:{myName:"Pat Example"},
 people:[{name:"Boss Person",management:true,requester:true},{name:"Pat Example",assignee:true,requester:true}],
 tasks:[
  {task:"Overdue safety fix",requester:"Boss Person",assignedTo:"Pat Example",status:"In Progress",due:plus(-2),category:"Safety",necessity:"High",ease:"Easy",cost:"Low",time:"<1 Day"},
  {task:"Due today, quality",requester:"Pat Example",assignedTo:"Pat Example; Sam Other",due:plus(0),category:"Quality",necessity:"Low"},
  {task:"No date, no category",requester:"Someone",assignedTo:"Sam Other"},
  {task:"Finished",status:"Complete",category:"Safety",necessity:"Urgent",due:plus(-30)}],
 recurring:[{title:"Monthly thing",ruleType:"Nth workday of month",ruleN:5,leadDays:3,nextDue:plus(10)}]};
const fs2=require('path'); const src=fs.readFileSync(fs2.join(__dirname,'..','setup','create-lists.js'),'utf8');
(async()=>{
 vm.runInContext(src,ctx); await new Promise(r=>setTimeout(r,300));
 const byTitle=t=>Object.values(lists).find(l=>l.Title===t);
 const a=byTitle('Action Items'), p=byTitle('People'), r=byTitle('Recurring Work');
 assert(a&&p&&r,'three lists');
 assert.strictEqual(a.items.length,4); assert.strictEqual(p.items.length,2); assert.strictEqual(r.items.length,1);
 const s=Object.fromEntries(a.items.map(i=>[i.Title,i.PriorityScore]));
 assert.strictEqual(s["Overdue safety fix"],12+9+3+10); // category + 3*necessity + management + overdue
 assert.strictEqual(s["Due today, quality"],8+3+8);
 assert.strictEqual(s["No date, no category"],0);
 assert.strictEqual(s["Finished"],-1);
 assert.strictEqual(a.items[0].QuickWin,3+3+4);
 assert.strictEqual(a.items[2].TaskStatus,"Not Started");
 assert(a.views.get('My Open Tasks').DefaultView,'My Open Tasks is the default view');
 assert(a.views.get('My Open Tasks').query.includes('Pat Example'));
 const rule=JSON.parse(a.views.get('My Open Tasks').CustomFormatter).additionalRowClass;
 assert(rule.includes('@now > [$TaskDueDate]'),'row rule survives the escaping');
 // A run that stopped after creating a view (as on the first real run) is finished off by the next one.
 delete a.views.get('Completed').CustomFormatter; a.views.get('My Open Tasks').DefaultView=false;
 assert.strictEqual(a.fields.get('Title').Title,'Task');
 const before=calls.length; vm.runInContext(src,ctx); await new Promise(r=>setTimeout(r,300));
 assert.strictEqual(calls.slice(before).filter(c=>/createfield|POST \/web\/lists$|\/items$/.test(c)).length,0,'second run creates nothing');
 assert.strictEqual(a.items.length,4);
 assert(a.views.get('Completed').CustomFormatter && a.views.get('My Open Tasks').DefaultView,'rerun finishes views');
 console.log('setup.test.js: all checks passed ('+calls.length+' REST calls)');
})().catch(e=>{console.error(e);process.exit(1);});
