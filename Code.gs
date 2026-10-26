/**
 * Hamsa Social Groups - Backend (Google Apps Script)
 * الاستخدام: أضف الملفين Code.gs و index.html إلى مشروع Apps Script المرتبط بالشيت.
 * شغّل setup() مرة واحدة، ثم عيّن كلمة مرور المدير عبر setAdminPassword() وانشر التطبيق.
 */
const ADMIN_NAME = 'المدير';

const SCHEMA = {
  Groups: ['id','name','platform','url','gov_id','area_id','sub_area','category','description','added_by','created_at','updated_at','verification','status','notes','members','is_demo'],
  Governorates: ['id','name_ar','name_en','active'],
  Areas: ['id','gov_id','parent_id','name_ar','name_en','type','active'],
  Employees: ['id','name','role','username','code_hash','active','phone','email','job_title','notes'],
  Categories: ['id','name_ar'],
  ActivityLog: ['time','employee','action','target','details'],
  Settings: ['key','value'],
  TaskPool: ['id','type','title','target','details','active','created_at'],
  DailyTasks: ['id','date','employee_id','employee','pool_id','type','title','target','details','status','result','created_at','updated_at'],
  TaskArchives: ['quarter','file_id','count','created_at']
};

const GOVS = [['Cairo','القاهرة'],['Giza','الجيزة'],['Alexandria','الإسكندرية'],['Qalyubia','القليوبية'],['Sharqia','الشرقية'],['Dakahlia','الدقهلية'],['Beheira','البحيرة'],['Gharbia','الغربية'],['Monufia','المنوفية'],['Kafr El Sheikh','كفر الشيخ'],['Damietta','دمياط'],['Port Said','بورسعيد'],['Ismailia','الإسماعيلية'],['Suez','السويس'],['North Sinai','شمال سيناء'],['South Sinai','جنوب سيناء'],['Fayoum','الفيوم'],['Beni Suef','بني سويف'],['Minya','المنيا'],['Assiut','أسيوط'],['Sohag','سوهاج'],['Qena','قنا'],['Luxor','الأقصر'],['Aswan','أسوان'],['Red Sea','البحر الأحمر'],['New Valley','الوادي الجديد'],['Matrouh','مطروح']];
const CATEGORIES = ['سياحة','سفر وحجز','مجتمع محلي','تجارة وبيع','وظائف','تعليم','عام'];
// بيانات مناطق مبدئية (ناقصة عن قصد - وسّعها بـ importAreas أو بالإضافة من الشيت)
const AREAS_SEED = {
  'الإسكندرية': [['حي شرق','حي',['سموحة','سيدي جابر','كليوباترا','زيزينيا']],['حي المنتزه','حي',[]],['حي وسط','حي',[]],['حي الجمرك','حي',[]],['حي غرب','حي',[]],['حي العجمي','حي',[]],['حي العطارين','حي',[]],['برج العرب','مدينة',[]]],
  'القاهرة': [['مدينة نصر','حي',[]],['مصر الجديدة','حي',[]],['المعادي','حي',[]],['حلوان','حي',[]],['شبرا','حي',[]],['الزمالك','حي',[]],['التجمع الخامس','حي',[]]],
  'الجيزة': [['الدقي','حي',[]],['العجوزة','حي',[]],['الهرم','حي',[]],['6 أكتوبر','مدينة',[]],['الشيخ زايد','مدينة',[]]]
};

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SCHEMA).forEach(n => {
    let sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (sh.getLastRow() === 0) { sh.getRange(1,1,1,SCHEMA[n].length).setValues([SCHEMA[n]]).setFontWeight('bold'); sh.setFrozenRows(1); }
    else if(n==='Employees') {
      const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
      SCHEMA.Employees.forEach(col=>{if(h.indexOf(col)<0){sh.getRange(1,sh.getLastColumn()+1).setValue(col);h.push(col);}});
    }
  });
  const d = ss.getSheetByName('Sheet1'); if (d && d.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(d);
  const g = ss.getSheetByName('Governorates');
  if (g.getLastRow() < 2) g.getRange(2,1,GOVS.length,4).setValues(GOVS.map((x,i)=>['g'+pad(i+1),x[1],x[0],true]));
  const c = ss.getSheetByName('Categories');
  if (c.getLastRow() < 2) c.getRange(2,1,CATEGORIES.length,2).setValues(CATEGORIES.map((x,i)=>['c'+(i+1),x]));
  const a = ss.getSheetByName('Areas');
  if (a.getLastRow() < 2) {
    const govs = rows('Governorates'), out = []; let n = 1;
    Object.keys(AREAS_SEED).forEach(gn => {
      const gv = govs.find(x => x.name_ar === gn);
      AREAS_SEED[gn].forEach(ar => { const id = 'a'+(n++); out.push([id,gv.id,'',ar[0],'',ar[1],true]);
        ar[2].forEach(s => out.push(['a'+(n++),gv.id,id,s,'','حي فرعي',true])); });
    });
    a.getRange(2,1,out.length,7).setValues(out);
  }
  const e = ss.getSheetByName('Employees');
  if (e.getLastRow() < 2) { const h=e.getRange(1,1,1,e.getLastColumn()).getValues()[0],o={id:'e1',name:ADMIN_NAME,role:'admin',username:'admin',code_hash:'',active:true};e.appendRow(h.map(k=>o[k]??'')); }
  else { const eh=e.getRange(1,1,1,e.getLastColumn()).getValues()[0], ui=eh.indexOf('username'), ri=eh.indexOf('role');
    if(ui>=0&&ri>=0){const ev=e.getDataRange().getValues(),idi=eh.indexOf('id'); for(let i=1;i<ev.length;i++)if(!ev[i][ui])e.getRange(i+1,ui+1).setValue(ev[i][ri]==='admin'?'admin':String(ev[i][idi]||('staff'+i)).toLowerCase());} }
  const s = ss.getSheetByName('Settings');
  if (s.getLastRow() < 2) s.appendRow(['geo_dataset','بيانات جغرافية مبدئية وغير مكتملة - وسّعها بالاستيراد']);
  ensureDailyTaskTrigger();
  ensureArchiveTrigger();
}

/* ---------- Helpers ---------- */
function pad(n){ return ('0'+n).slice(-2); }
function hexDigest(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8).map(b=>('0'+(b&255).toString(16)).slice(-2)).join('');}
function passwordHash(password){const salt=Utilities.getUuid().replace(/-/g,'');let value=salt+':'+password;for(let i=0;i<4096;i++)value=hexDigest(value);return salt+':'+value;}
function verifyPassword(password,stored){stored=String(stored||'');if(stored.indexOf(':')<0)return stored===hexDigest(password);const p=stored.split(':');let value=p[0]+':'+password;for(let i=0;i<4096;i++)value=hexDigest(value);return value===p[1];}
function sheet(n){ return SpreadsheetApp.getActive().getSheetByName(n); }
function rows(n){ const v = sheet(n).getDataRange().getValues(), h = v.shift(); return v.map(r => { const o = {}; h.forEach((k,i)=>o[k]=r[i]); return o; }); }
function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function log(emp, action, target, details){ sheet('ActivityLog').appendRow([new Date().toISOString(), emp, action, target, details||'']); }
function recentActivity(){const sh=sheet('ActivityLog'),last=sh.getLastRow();if(last<2)return [];const h=sh.getRange(1,1,1,5).getValues()[0],start=Math.max(2,last-199),data=sh.getRange(start,1,last-start+1,5).getValues();return data.reverse().map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i]);return o;});}
function detectPlatform(u){
  const h = (u.match(/^https?:\/\/([^\/?#]+)/i)||[])[1]; if (!h) return null; const x = h.toLowerCase();
  if (/(^|\.)(facebook\.com|fb\.com|fb\.me)$/.test(x)) return 'facebook';
  if (/(^|\.)(chat\.whatsapp\.com|wa\.me|whatsapp\.com)$/.test(x)) return 'whatsapp';
  if (/(^|\.)(t\.me|telegram\.me|telegram\.org)$/.test(x)) return 'telegram';
  if (/(^|\.)instagram\.com$/.test(x)) return 'instagram';
  if (/(^|\.)tiktok\.com$/.test(x)) return 'tiktok';
  if (/(^|\.)(youtube\.com|youtu\.be)$/.test(x)) return 'youtube';
  return 'other';
}
function norm(u){ return u.trim().toLowerCase().replace(/^https?:\/\/(www\.|m\.)?/,'').replace(/[?#].*$/,'').replace(/\/+$/,''); }

/* ---------- API ---------- */
function doGet(){ return HtmlService.createHtmlOutputFromFile('index').setTitle('همسة للسياحة - مهامي اليومية'); }
function doPost(e) { let req; try { req=JSON.parse(e.postData.contents); } catch(x){return out({ok:false,error:'طلب غير صالح'});} return out(apiRequest(req)); }
function apiCall(req){return JSON.parse(JSON.stringify(apiRequest(req)));}
function apiRequest(req) {
  req=req||{};
  const all=rows('Employees'); let emp;
  if(req.action==='login'){
    const supplied=String(req.password||'').trim(),username=String(req.username||'').trim().toLowerCase();
    emp=all.find(x=>String(x.active).toUpperCase()==='TRUE'&&verifyPassword(supplied,x.code_hash)&&String(x.username||'').trim().toLowerCase()===username);
    if(!emp)return {ok:false,error:'اسم المستخدم أو كلمة المرور غير صحيحة، أو الحساب موقوف',auth:false};
    if(String(emp.code_hash||'').indexOf(':')<0){const es=sheet('Employees'),eh=es.getRange(1,1,1,es.getLastColumn()).getValues()[0],ei=all.findIndex(x=>x.id===emp.id);es.getRange(ei+1,eh.indexOf('code_hash')+1).setValue(passwordHash(supplied));}
    const token=Utilities.getUuid()+Utilities.getUuid();CacheService.getScriptCache().put('login:'+token,emp.id,21600);
    return {ok:true,sessionToken:token,user:{id:emp.id,name:emp.name,role:emp.role}};
  }
  const userId=CacheService.getScriptCache().get('login:'+String(req.token||''));
  emp=all.find(x=>x.id===userId&&String(x.active).toUpperCase()==='TRUE');
  if (!emp) return {ok:false,error:'اسم المستخدم أو كلمة المرور غير صحيحة، أو الحساب موقوف',auth:false};
  const isAdmin = emp.role === 'admin';
  const lock = LockService.getScriptLock();
  try {
    switch (req.action) {
      case 'bootstrap': return {ok:true, user:{id:emp.id,name:emp.name,role:emp.role},
        govs: rows('Governorates').filter(g=>String(g.active).toUpperCase()==='TRUE'), areas: rows('Areas').filter(a=>String(a.active).toUpperCase()==='TRUE'), cats: rows('Categories'), stats: stats()};
      case 'groups': return Object.assign({ok:true}, queryGroups(req));
      case 'bulkAdd': {if(!isAdmin)return {ok:false,error:'مش مسموح'}; lock.waitLock(30000); const r=bulkAddGroups(req.data||{},emp); lock.releaseLock(); return r;}
      case 'addGroup': { lock.waitLock(15000); const r = addGroup(req.data, emp); lock.releaseLock(); return r; }
      case 'setStatus': { lock.waitLock(15000); const r = updateGroup(req.id, req.patch, emp, isAdmin); lock.releaseLock(); return r; }
      case 'reportBroken': { lock.waitLock(15000); const r = updateGroup(req.id, {status:'broken'}, emp, true); lock.releaseLock(); return r; }
      case 'deleteGroup': { if (!isAdmin) return {ok:false,error:'مش مسموح'}; lock.waitLock(15000); const r = deleteGroup(req.id, emp); lock.releaseLock(); return r; }
      case 'addArea': { if (!isAdmin) return {ok:false,error:'مش مسموح'}; lock.waitLock(15000);
        const d = req.data; if (!d.name_ar || !d.gov_id) { lock.releaseLock(); return {ok:false,error:'الاسم والمحافظة مطلوبين'}; }
        const sh = sheet('Areas'); const id = 'a'+Utilities.getUuid().slice(0,8);
        sh.appendRow([id,d.gov_id,d.parent_id||'',d.name_ar,d.name_en||'',d.type||'حي',true]); log(emp.name,'addArea',id,d.name_ar); lock.releaseLock(); return {ok:true,id}; }
      case 'importAreas': { if (!isAdmin) return {ok:false,error:'مش مسموح'}; lock.waitLock(30000);
        // rows: [{gov_id,parent_id,name_ar,name_en,type}] - بيرجع تقرير بالمرفوض
        const ex = rows('Areas'), add = [], bad = [];
        (req.rows||[]).forEach((r,i) => { if (!r.name_ar || !r.gov_id) return bad.push({row:i+1,reason:'ناقص اسم أو محافظة'});
          if (ex.some(a=>a.gov_id===r.gov_id && a.name_ar===r.name_ar && String(a.parent_id)===String(r.parent_id||''))) return bad.push({row:i+1,reason:'مكرر'});
          add.push(['a'+Utilities.getUuid().slice(0,8),r.gov_id,r.parent_id||'',r.name_ar,r.name_en||'',r.type||'حي',true]); });
        if (add.length) sheet('Areas').getRange(sheet('Areas').getLastRow()+1,1,add.length,7).setValues(add);
        log(emp.name,'importAreas','',add.length+' added'); lock.releaseLock(); return {ok:true,added:add.length,rejected:bad}; }
      case 'activity': { if (!isAdmin) return {ok:false,error:'مش مسموح'}; return {ok:true, items: recentActivity()}; }
      case 'getTasks': { if (isAdmin) return {ok:false,error:'المهام اليومية مخصصة للموظفين'}; generateDailyTasks(); return {ok:true,items:taskRowsFor(emp,false)}; }
      case 'getTeamTasks': { if (!isAdmin) return {ok:false,error:'مش مسموح'}; generateDailyTasks(); return {ok:true,items:taskRowsFor(emp,true)}; }
      case 'getTaskReport': {if(!isAdmin)return {ok:false,error:'مش مسموح'};return taskReport(req.from||taskDate(),req.to||req.from||taskDate(),req.employee_id||'',req.type||'');}
      case 'listEmployees': { if (!isAdmin) return {ok:false,error:'مش مسموح'}; return {ok:true,items:rows('Employees').map(e=>({id:e.id,name:e.name,username:e.username||'',role:e.role,active:String(e.active).toUpperCase()==='TRUE',phone:e.phone||'',email:e.email||'',job_title:e.job_title||'',notes:e.notes||''}))}; }
      case 'saveEmployee': { if (!isAdmin) return {ok:false,error:'مش مسموح'}; return saveEmployee(req.data||{}); }
      case 'resetEmployeePassword': { if (!isAdmin) return {ok:false,error:'مش مسموح'}; return resetEmployeePassword(req.id,req.password); }
      case 'addTaskPool': { if (!isAdmin) return {ok:false,error:'مش مسموح'}; const d=req.data||{}; if(!d.type||!d.title||!d.target)return {ok:false,error:'النوع والعنوان والرقم أو الرابط مطلوبين'}; const id='P'+Utilities.getUuid().slice(0,8); sheet('TaskPool').appendRow([id,d.type,d.title,d.target,d.details||'',true,new Date().toISOString()]); log(emp.name,'addTaskPool',id,d.title); return {ok:true,id}; }
      case 'bulkAddTasks': {if(!isAdmin)return {ok:false,error:'مش مسموح'};return bulkAddTasks(req.rows||[],emp);}
      case 'returnTaskToPool': {if(!isAdmin)return {ok:false,error:'مش مسموح'};return returnTaskToPool(req.id,emp);}
      case 'requestMoreCalls': { if(isAdmin)return {ok:false,error:'الطلب للموظفين فقط'}; lock.waitLock(15000);const r={ok:true,assigned:assignExtraCalls(emp,10)};lock.releaseLock();return r; }
      case 'updateTask': { lock.waitLock(15000); const r=updateDailyTask(req.id,req.patch||{},emp,isAdmin); lock.releaseLock(); return r; }
      default: return {ok:false,error:'أمر غير معروف'};
    }
  } catch (err) { try{lock.releaseLock();}catch(_){} return {ok:false,error:String(err)}; }
}

function addGroup(d, emp) {
  if (!d || !d.name || !d.url || !d.gov_id) return {ok:false, error:'الاسم والرابط والمحافظة مطلوبين'};
  if (!/^https?:\/\/\S+\.\S+/i.test(d.url)) return {ok:false, error:'الرابط غير صالح'};
  const platform = detectPlatform(d.url);
  if (d.platform && d.platform !== platform && platform !== 'other') return {ok:false, error:'الرابط مش تابع للمنصة المختارة (المنصة المكتشفة: '+platform+')'};
  const key = norm(d.url), gs=sheet('Groups'), gl=gs.getLastRow(), urls=gl>1?gs.getRange(2,4,gl-1,1).getValues():[];
  if (urls.some(r => norm(String(r[0]||'')) === key)) return {ok:false, duplicate:true, error:'الرابط ده متسجل قبل كده'};
  const now = new Date().toISOString(), id = 'G'+Utilities.getUuid().slice(0,8).toUpperCase();
  sheet('Groups').appendRow([id,d.name,d.platform||platform,d.url.trim(),d.gov_id,d.area_id||'',d.sub_area||'',d.category||'',d.description||'',emp.name,now,now,'pending','active',d.notes||'',d.members||'',false]);
  log(emp.name,'create',id,d.name);
  return {ok:true, id, stats: stats()};
}

function bulkAddGroups(d, emp) {
  const sh=sheet('Groups'), last=sh.getLastRow(), known={};if(last>1)sh.getRange(2,4,last-1,1).getValues().forEach(r=>known[norm(String(r[0]||''))]=true);
  const now=new Date().toISOString(), add=[]; let duplicates=0,bad=0;
  (d.rows||[]).forEach(x=>{const url=String(x.url||'').trim(); if(!/^https?:\/\/\S+\.\S+/i.test(url)){bad++;return;} const k=norm(url); if(known[k]){duplicates++;return;} known[k]=true;
    const platform=detectPlatform(url); add.push(['G'+Utilities.getUuid().slice(0,8).toUpperCase(),x.name||url,d.platform||platform,url,d.gov_id||'',d.area_id||'', '',d.category||'','',emp.name,now,now,'pending','active','', '',false]);});
  if(add.length)sh.getRange(sh.getLastRow()+1,1,add.length,add[0].length).setValues(add);
  log(emp.name,'bulkAdd','',add.length+' groups'); return {ok:true,added:add.length,duplicates,bad};
}

function updateGroup(id, patch, emp, canEdit) {
  const sh = sheet('Groups'), v = sh.getDataRange().getValues(), h = v[0];
  const idx = v.findIndex((r,i)=>i>0 && r[0]===id); if (idx<0) return {ok:false,error:'السجل مش موجود'};
  const allowed = ['status','verification','name','description','notes','category','members'];
  Object.keys(patch||{}).forEach(k => {
    if (allowed.indexOf(k)<0) return;
    if (k==='verification' && emp.role!=='admin') return;       // التوثيق للأدمن فقط
    if (k!=='status' && !canEdit && v[idx][h.indexOf('added_by')]!==emp.name) return;
    sh.getRange(idx+1,h.indexOf(k)+1).setValue(patch[k]);
  });
  sh.getRange(idx+1,h.indexOf('updated_at')+1).setValue(new Date().toISOString());
  log(emp.name,'update',id,JSON.stringify(patch));
  return {ok:true, stats: stats()};
}

function deleteGroup(id, emp) {
  const sh = sheet('Groups'), v = sh.getDataRange().getValues(), idx = v.findIndex((r,i)=>i>0&&r[0]===id);
  if (idx<0) return {ok:false,error:'السجل مش موجود'}; sh.deleteRow(idx+1); log(emp.name,'delete',id); return {ok:true, stats: stats()};
}

function queryGroups(q) {
  const sh=sheet('Groups'),last=sh.getLastRow(),size=24,page=Math.max(1,Number(q.page)||1),hasFilters=!!((q.q||'').trim()||q.gov_id||q.area_id||q.platform||q.status||q.verification||q.added_by),sort=q.sort||'newest';
  if(!hasFilters&&sort==='newest'){
    const total=Math.max(0,last-1),end=last-(page-1)*size,start=Math.max(2,end-size+1),count=Math.max(0,end-start+1),h=SCHEMA.Groups;
    const data=count?sh.getRange(start,1,count,h.length).getValues():[];
    return {total,page,size,items:data.reverse().map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i]);return o;})};
  }
  let list = rows('Groups');
  const t = (q.q||'').trim().toLowerCase();
  if (t) list = list.filter(g => [g.name,g.url,g.category,g.description,g.sub_area].join(' ').toLowerCase().indexOf(t) >= 0);
  ['gov_id','area_id','platform','status','verification','added_by'].forEach(k => { if (q[k]) list = list.filter(g => String(g[k])===q[k]); });
  const s = q.sort||'newest';
  list.sort((a,b)=> s==='oldest' ? (a.created_at>b.created_at?1:-1) : s==='name' ? String(a.name).localeCompare(String(b.name),'ar') : s==='updated' ? (a.updated_at<b.updated_at?1:-1) : (a.created_at<b.created_at?1:-1));
  const total = list.length;
  return {total, page, size, items: list.slice((page-1)*size, page*size)};
}

function stats() {
  const cache=CacheService.getScriptCache(),cached=cache.get('group_stats_v1');if(cached)return JSON.parse(cached);
  const g = rows('Groups'), today = taskDate(), by = {}, per = {};
  g.forEach(x => { by[x.platform]=(by[x.platform]||0)+1; per[x.added_by]=(per[x.added_by]||0)+1; });
  const result={total:g.length, today:g.filter(x=>String(x.created_at).slice(0,10)===today).length, byPlatform:by, byEmployee:per,
    pending:g.filter(x=>x.verification==='pending').length, broken:g.filter(x=>x.status==='broken').length,
    areas: rows('Areas').length, govs: 27,
    recent: g.slice(-5).reverse().map(x=>({id:x.id,name:x.name,platform:x.platform,added_by:x.added_by}))};cache.put('group_stats_v1',JSON.stringify(result),60);return result;
}

function taskDate(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function taskDateValue(v){ return v instanceof Date ? Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd') : String(v); }
function ensureDailyTaskTrigger(){
  const exists=ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='generateDailyTasks');
  if(!exists) ScriptApp.newTrigger('generateDailyTasks').timeBased().everyDays(1).atHour(5).create();
}
function ensureArchiveTrigger(){const exists=ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='archiveOldTasks');if(!exists)ScriptApp.newTrigger('archiveOldTasks').timeBased().everyDays(1).atHour(6).create();}
function compactTaskPool(){
  const sh=sheet('TaskPool'),last=sh.getLastRow();if(last<2)return;const active=sh.getRange(2,6,last-1,1).getValues();let inactive=0;active.forEach(r=>{if(String(r[0]).toUpperCase()!=='TRUE')inactive++;});if(inactive<50000)return;
  const h=sh.getRange(1,1,1,7).getValues()[0],data=sh.getRange(2,1,last-1,7).getValues().filter(r=>String(r[5]).toUpperCase()==='TRUE');
  sh.clearContents();sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold');for(let i=0;i<data.length;i+=5000){const b=data.slice(i,i+5000);sh.getRange(sh.getLastRow()+1,1,b.length,h.length).setValues(b);}const newLast=data.length+1;if(last>newLast)sh.deleteRows(newLast+1,last-newLast);
}
function quarterStart(d){const y=Number(d.slice(0,4)),m=Number(d.slice(5,7)),qm=Math.floor((m-1)/3)*3+1;return y+'-'+pad(qm)+'-01';}
function quarterLabel(d){return d.slice(0,4)+'-Q'+(Math.floor((Number(d.slice(5,7))-1)/3)+1);}
function archiveOldTasks(){
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const sh=sheet('DailyTasks'),last=sh.getLastRow();if(last<2)return;
    const h=sh.getRange(1,1,1,SCHEMA.DailyTasks.length).getValues()[0],data=sh.getRange(2,1,last-1,h.length).getValues(),current=quarterStart(taskDate()),groups={},keep=[];
    data.forEach(r=>{const date=taskDateValue(r[1]);if(date<current)(groups[quarterLabel(date)]||(groups[quarterLabel(date)]=[])).push(r);else keep.push(r);});
    const meta=sheet('TaskArchives');
    Object.keys(groups).forEach(q=>{
      let rec=rows('TaskArchives').find(x=>x.quarter===q),file;
      if(rec)file=SpreadsheetApp.openById(rec.file_id);else{file=SpreadsheetApp.create('Hamsa Task History '+q);const first=file.getSheets()[0];first.setName('TaskHistory');first.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold');meta.appendRow([q,file.getId(),0,new Date().toISOString()]);rec=rows('TaskArchives').find(x=>x.quarter===q);}
      const ah=file.getSheetByName('TaskHistory'),al=ah.getLastRow(),known=new Set(al>1?ah.getRange(2,1,al-1,1).getValues().flat().map(String):[]),fresh=groups[q].filter(r=>!known.has(String(r[0])));
      for(let i=0;i<fresh.length;i+=5000){const batch=fresh.slice(i,i+5000);ah.getRange(ah.getLastRow()+1,1,batch.length,h.length).setValues(batch);}
      if(rec){const rr=meta.getDataRange().getValues(),qi=rr.findIndex((x,i)=>i>0&&x[0]===q);if(qi>0)meta.getRange(qi+1,3).setValue(Math.max(0,ah.getLastRow()-1));}
    });
    if(groups && Object.keys(groups).length){sh.clearContents();sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold');for(let i=0;i<keep.length;i+=5000){const b=keep.slice(i,i+5000);sh.getRange(sh.getLastRow()+1,1,b.length,h.length).setValues(b);}}
  }finally{lock.releaseLock();}
}
function generateDailyTasks(){
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const today=taskDate(), dh=sheet('DailyTasks');
    const last=dh.getLastRow(), todayExists=last>1&&dh.getRange(2,2,last-1,1).createTextFinder(today).matchEntireCell(true).findNext();
    if(todayExists)return;
    const employees=rows('Employees').filter(e=>String(e.active).toUpperCase()==='TRUE'&&e.role!=='admin');
    const pool=rows('TaskPool').filter(p=>String(p.active).toUpperCase()==='TRUE');
    if(!employees.length||!pool.length)return;
    const counts={call:50,post:15,find_page:15}, byType={call:[],post:[],find_page:[]};
    pool.forEach(p=>{if(byType[p.type])byType[p.type].push(p);});
    const batch=[], used=[]; const now=new Date().toISOString();
    Object.keys(counts).forEach(type=>{let cursor=0;for(let round=0;round<counts[type]&&cursor<byType[type].length;round++)employees.forEach(e=>{if(cursor>=byType[type].length)return;const p=byType[type][cursor++];batch.push(['T'+Utilities.getUuid().slice(0,8),today,e.id,e.name,p.id,p.type,p.title,p.target,p.details,'pending','',now,now]);used.push(p.id);});});
    if(batch.length)dh.getRange(dh.getLastRow()+1,1,batch.length,batch[0].length).setValues(batch);
    if(used.length){const psh=sheet('TaskPool'),pn=psh.getLastRow(),ids=psh.getRange(2,1,pn-1,1).getValues(),active=psh.getRange(2,6,pn-1,1),av=active.getValues(),set=new Set(used);for(let i=0;i<ids.length;i++)if(set.has(ids[i][0]))av[i][0]=false;active.setValues(av);compactTaskPool();}
  } finally { lock.releaseLock(); }
}
function taskRowsFor(emp,isAdmin){
  const today=taskDate();
  const sh=sheet('DailyTasks'),last=sh.getLastRow(); if(last<2)return [];
  const matches=sh.getRange(2,2,last-1,1).createTextFinder(today).matchEntireCell(true).findAll();if(!matches.length)return [];
  const start=Math.min.apply(null,matches.map(x=>x.getRow())),end=Math.max.apply(null,matches.map(x=>x.getRow())),v=sh.getRange(start,1,end-start+1,SCHEMA.DailyTasks.length).getValues();
  return v.filter(r=>taskDateValue(r[1])===today&&(isAdmin||r[2]===emp.id)).map(r=>({id:r[0],date:taskDateValue(r[1]),employee:r[3],employee_id:r[2],type:r[5],title:r[6],target:r[7],details:r[8],status:r[9],result:r[10]}));
}
function taskReport(from,to,employeeId,type){
  const items=[],addRows=data=>data.forEach(r=>{const date=taskDateValue(r[1]);if(date<from||date>to||(employeeId&&r[2]!==employeeId)||(type&&r[5]!==type))return;items.push({id:r[0],date,employee_id:r[2],employee:r[3],pool_id:r[4],type:r[5],title:r[6],target:r[7],status:r[9],result:r[10]});});
  const readRange=(s)=>{const n=s.getLastRow();if(n<2)return;const dates=s.getRange(2,2,n-1,1).getValues();let lo=-1,hi=-1;dates.forEach((x,i)=>{const d=taskDateValue(x[0]);if(d>=from&&d<=to){if(lo<0)lo=i;hi=i;}});if(lo>=0)addRows(s.getRange(lo+2,1,hi-lo+1,SCHEMA.DailyTasks.length).getValues());};
  readRange(sheet('DailyTasks'));
  const ars=rows('TaskArchives'),startQ=quarterLabel(from),endQ=quarterLabel(to);ars.forEach(a=>{if(a.quarter<startQ||a.quarter>endQ)return;const f=SpreadsheetApp.openById(a.file_id),as=f.getSheetByName('TaskHistory');readRange(as);});
  const done=items.filter(x=>x.status==='done').length;return {ok:true,items:items.slice(-1500).reverse(),summary:{total:items.length,done,percent:items.length?Math.round(done*100/items.length):0}};
}
function returnTaskToPool(taskId,emp){
  const sh=sheet('DailyTasks'),last=sh.getLastRow();let row=null;
  if(last>1){const cell=sh.getRange(2,1,last-1,1).createTextFinder(taskId).matchEntireCell(true).findNext();if(cell)row=sh.getRange(cell.getRow(),1,1,SCHEMA.DailyTasks.length).getValues()[0];}
  if(!row){rows('TaskArchives').some(a=>{const f=SpreadsheetApp.openById(a.file_id),as=f.getSheetByName('TaskHistory'),n=as.getLastRow(),cell=n>1?as.getRange(2,1,n-1,1).createTextFinder(taskId).matchEntireCell(true).findNext():null;if(cell){row=as.getRange(cell.getRow(),1,1,SCHEMA.DailyTasks.length).getValues()[0];return true;}return false;});}
  if(!row)return {ok:false,error:'المهمة غير موجودة'};
  const poolId=row[4],ps=sheet('TaskPool'),pn=ps.getLastRow(),cell=pn>1?ps.getRange(2,1,pn-1,1).createTextFinder(poolId).matchEntireCell(true).findNext():null;
  if(cell)ps.getRange(cell.getRow(),6).setValue(true);else ps.appendRow([poolId||'P'+Utilities.getUuid().slice(0,8),row[5],row[6],row[7],row[8],true,new Date().toISOString()]);
  log(emp.name,'returnTaskToPool',taskId,row[6]);return {ok:true};
}
function assignExtraCalls(emp,amount){
  const todayItems=taskRowsFor(emp,false),completed=todayItems.filter(t=>t.type==='call'&&t.status==='done').length;if(completed<50)return 0;
  const today=taskDate(), pool=rows('TaskPool').filter(p=>String(p.active).toUpperCase()==='TRUE'&&p.type==='call').slice(0,amount);
  if(!pool.length)return 0; const sh=sheet('DailyTasks'),now=new Date().toISOString(),batch=pool.map(p=>['T'+Utilities.getUuid().slice(0,8),today,emp.id,emp.name,p.id,p.type,p.title,p.target,p.details,'pending','',now,now]);
  sh.getRange(sh.getLastRow()+1,1,batch.length,batch[0].length).setValues(batch); const ps=sheet('TaskPool'),n=ps.getLastRow(),ids=ps.getRange(2,1,n-1,1).getValues(),active=ps.getRange(2,6,n-1,1),av=active.getValues(),wanted=new Set(pool.map(p=>p.id));for(let i=0;i<ids.length;i++)if(wanted.has(ids[i][0]))av[i][0]=false;active.setValues(av);compactTaskPool();return pool.length;
}
function updateDailyTask(id,patch,emp,isAdmin){
  const sh=sheet('DailyTasks'),last=sh.getLastRow(),cell=last>1?sh.getRange(2,1,last-1,1).createTextFinder(id).matchEntireCell(true).findNext():null;
  if(!cell)return {ok:false,error:'المهمة غير موجودة'};
  if(isAdmin)return {ok:false,error:'المدير يراجع المهام فقط'};
  const rowNum=cell.getRow(),v=sh.getRange(rowNum,1,1,SCHEMA.DailyTasks.length).getValues()[0],h=SCHEMA.DailyTasks;
  if(v[h.indexOf('employee_id')]!==emp.id)return {ok:false,error:'المهمة مش مخصصة ليك'};
  const status=String(patch.status||''); if(['pending','in_progress','done','blocked'].indexOf(status)<0)return {ok:false,error:'حالة المهمة غير صحيحة'};
  const type=v[h.indexOf('type')],result=String(patch.result||'').slice(0,1000);
  if(type==='call'&&status==='done'&&['answered','no_answer','interested','wrong_number'].indexOf(result)<0)return {ok:false,error:'اختار نتيجة المكالمة الأول'};
  sh.getRange(rowNum,h.indexOf('status')+1).setValue(status);
  sh.getRange(rowNum,h.indexOf('result')+1).setValue(result);
  sh.getRange(rowNum,h.indexOf('updated_at')+1).setValue(new Date().toISOString());
  return {ok:true};
}
function bulkAddTasks(items,emp){
  const allowed=['call','post','find_page'], sh=sheet('TaskPool'), now=new Date().toISOString(), outRows=[];let bad=0;
  items.slice(0,5000).forEach(x=>{const type=String(x.type||''),title=String(x.title||'').trim(),target=String(x.target||'').trim();if(allowed.indexOf(type)<0||!title||!target){bad++;return;}outRows.push(['P'+Utilities.getUuid().slice(0,8),type,title,target,String(x.details||'').slice(0,1000),true,now]);});
  if(outRows.length)sh.getRange(sh.getLastRow()+1,1,outRows.length,7).setValues(outRows);log(emp.name,'bulkAddTasks','',outRows.length+' tasks');return {ok:true,added:outRows.length,bad};
}
function saveEmployee(d){
  const sh=sheet('Employees'),v=sh.getDataRange().getValues(),h=v[0],name=String(d.name||'').trim();
  if(!name)return {ok:false,error:'اسم الموظف مطلوب'};
  const username=String(d.username||'').trim().toLowerCase();if(!/^[a-z0-9._-]{3,30}$/.test(username))return {ok:false,error:'اسم المستخدم لازم يكون 3 أحرف أو أكثر بالإنجليزي والأرقام'};
  let idx=d.id?v.findIndex((r,i)=>i>0&&r[0]===d.id):-1, id;
  if(d.id&&idx<1)return {ok:false,error:'الموظف غير موجود'};
  if(v.slice(1).some((r,i)=>i+1!==idx&&String(r[h.indexOf('username')]||'').trim().toLowerCase()===username))return {ok:false,error:'اسم المستخدم مستخدم بالفعل'};
  if(!d.id){const password=String(d.password||'').trim();if(password.length<8)return {ok:false,error:'كلمة المرور لازم تكون 8 أحرف على الأقل'};id='e'+Utilities.getUuid().slice(0,8);const obj={id,name,role:'employee',username,code_hash:passwordHash(password),active:true,phone:String(d.phone||'').trim(),email:String(d.email||'').trim(),job_title:String(d.job_title||'').trim(),notes:String(d.notes||'').trim()};sh.appendRow(h.map(k=>obj[k]??''));log('المدير','createEmployee',id,name);return {ok:true,id};}
  id=d.id; const fields={name,username,phone:String(d.phone||'').trim(),email:String(d.email||'').trim(),job_title:String(d.job_title||'').trim(),notes:String(d.notes||'').trim(),active:d.active===false?false:true};
  Object.keys(fields).forEach(k=>{const c=h.indexOf(k);if(c>=0)sh.getRange(idx+1,c+1).setValue(fields[k]);});log('المدير','updateEmployee',id,name);return {ok:true,id};
}
function resetEmployeePassword(id,password){
  const p=String(password||'').trim();if(p.length<8)return {ok:false,error:'كلمة المرور لازم تكون 8 أحرف على الأقل'};
  const sh=sheet('Employees'),v=sh.getDataRange().getValues(),h=v[0],idx=v.findIndex((r,i)=>i>0&&r[0]===id);
  if(idx<1||v[idx][h.indexOf('role')]==='admin')return {ok:false,error:'الموظف غير موجود'};
  sh.getRange(idx+1,h.indexOf('code_hash')+1).setValue(passwordHash(p));log('المدير','resetEmployeePassword',id);return {ok:true};
}

/* ---------- تعيين كلمة مرور المدير من نافذة آمنة ---------- */
function setAdminPassword() {
  const ui=SpreadsheetApp.getUi(),answer=ui.prompt('كلمة مرور مدير النظام','اكتب كلمة مرور لا تقل عن 8 أحرف. لن تُحفظ في ملف الكود.',ui.ButtonSet.OK_CANCEL);
  if(answer.getSelectedButton()!==ui.Button.OK)return;
  const password=answer.getResponseText().trim();if(password.length<8){ui.alert('كلمة المرور لازم تكون 8 أحرف على الأقل. شغّل الدالة مرة أخرى.');return;}
  const sh=sheet('Employees'),v=sh.getDataRange().getValues(),h=v[0],i=v.findIndex((r,n)=>n>0&&r[h.indexOf('role')]==='admin');
  if(i<0){const o={id:'e1',name:ADMIN_NAME,role:'admin',username:'admin',code_hash:passwordHash(password),active:true};sh.appendRow(h.map(k=>o[k]??''));}
  else{sh.getRange(i+1,h.indexOf('username')+1).setValue('admin');sh.getRange(i+1,h.indexOf('code_hash')+1).setValue(passwordHash(password));sh.getRange(i+1,h.indexOf('active')+1).setValue(true);}
  ui.alert('تم تعيين كلمة مرور المدير بنجاح. اسم المستخدم هو admin.');
}

/* ---------- تشخيص بسيط بدون طباعة أي كلمات مرور أو بصمات ---------- */
function diagnose() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Employees');
  if (!sh) { Logger.log('❌ مفيش شيت Employees - شغّل setup'); return; }
  const list = rows('Employees');
  Logger.log('عدد الموظفين: ' + list.length);
  list.forEach(x => Logger.log(x.name + ' | role=' + x.role + ' | active=' + x.active + ' | اسم مستخدم موجود=' + !!x.username + ' | كلمة مرور مُعدة=' + !!x.code_hash));
}

/* ================= مناطق مصر (مدن / مراكز / أحياء) - شغّل seedAllAreas مرة واحدة ================= */
const GEO = {
  'القاهرة': {t:'حي', a:['المرج','السلام أول','السلام ثان','المطرية','عين شمس','النزهة','مصر الجديدة','مدينة نصر','الزيتون','الأميرية','حدائق القبة','الشرابية','شبرا','الساحل','روض الفرج','الوايلي','الظاهر','الأزبكية','عابدين','الموسكي','باب الشعرية','الجمالية','بولاق','الزمالك','وسط البلد','السيدة زينب','مصر القديمة','الخليفة','المقطم','المعادي','دار السلام','البساتين','طرة','المعصرة','حلوان','15 مايو','التبين','التجمع الأول','التجمع الثالث','التجمع الخامس','القطامية','الرحاب','مدينتي','الشروق','بدر','العاصمة الإدارية الجديدة','الأسمرات']},
  'الجيزة': {t:'مركز', a:['الجيزة','الدقي','العجوزة','إمبابة','بولاق الدكرور','الهرم','فيصل','العمرانية','الطالبية','المنيب','ساقية مكي','المريوطية','كرداسة','الوراق','أوسيم','الصف','أطفيح','العياط','الواحات البحرية','منشأة القناطر','البدرشين','الحوامدية','أبو النمرس','6 أكتوبر','الشيخ زايد','حدائق أكتوبر','كفر طهرمس','ناهيا']},
  'الإسكندرية': {t:'حي', a:['حي شرق','حي المنتزه أول','حي المنتزه ثان','حي وسط','حي الجمرك','حي غرب','حي العجمي','حي العطارين','حي اللبان','حي العامرية','برج العرب','برج العرب الجديدة']},
  'القليوبية': {t:'مركز', a:['بنها','قليوب','شبرا الخيمة','القناطر الخيرية','الخانكة','كفر شكر','طوخ','شبين القناطر','الخصوص','العبور','قها','القلج','سرياقوس','بهتيم','مسطرد','أبو زعبل']},
  'الشرقية': {t:'مركز', a:['الزقازيق','العاشر من رمضان','بلبيس','منيا القمح','أبو حماد','أبو كبير','فاقوس','ههيا','ديرب نجم','الحسينية','كفر صقر','الإبراهيمية','مشتول السوق','القرين','أولاد صقر','الصالحية الجديدة','القنايات','صان الحجر']},
  'الدقهلية': {t:'مركز', a:['المنصورة','المنصورة الجديدة','طلخا','ميت غمر','دكرنس','أجا','منية النصر','السنبلاوين','الجمالية','المطرية','بلقاس','شربين','المنزلة','تمي الأمديد','نبروه','ميت سلسيل','بني عبيد','جمصة','الكردي']},
  'البحيرة': {t:'مركز', a:['دمنهور','كفر الدوار','رشيد','إدكو','أبو المطامير','أبو حمص','الدلنجات','المحمودية','الرحمانية','إيتاي البارود','حوش عيسى','شبراخيت','كوم حمادة','بدر','وادي النطرون','النوبارية الجديدة','النوبارية','السادات']},
  'الغربية': {t:'مركز', a:['طنطا','المحلة الكبرى','كفر الزيات','زفتى','السنطة','قطور','بسيون','سمنود']},
  'المنوفية': {t:'مركز', a:['شبين الكوم','مدينة السادات','منوف','أشمون','الباجور','قويسنا','بركة السبع','تلا','الشهداء','سرس الليان']},
  'كفر الشيخ': {t:'مركز', a:['كفر الشيخ','دسوق','فوه','مطوبس','بيلا','الحامول','بلطيم','الرياض','سيدي سالم','قلين','برج البرلس']},
  'دمياط': {t:'مركز', a:['دمياط','دمياط الجديدة','رأس البر','فارسكور','الزرقا','كفر سعد','كفر البطيخ','عزبة البرج']},
  'بورسعيد': {t:'حي', a:['حي الشرق','حي العرب','حي المناخ','حي الضواحي','حي الزهور','بورفؤاد','حي الجنوب']},
  'الإسماعيلية': {t:'مركز', a:['الإسماعيلية','التل الكبير','فايد','القنطرة شرق','القنطرة غرب','القصاصين','أبو صوير']},
  'السويس': {t:'حي', a:['حي السويس','حي الأربعين','حي فيصل','حي الجناين','حي عتاقة']},
  'شمال سيناء': {t:'مدينة', a:['العريش','الشيخ زويد','رفح','بئر العبد','الحسنة','نخل']},
  'جنوب سيناء': {t:'مدينة', a:['شرم الشيخ','دهب','نويبع','طابا','سانت كاترين','الطور','رأس سدر','أبو رديس','أبو زنيمة']},
  'الفيوم': {t:'مركز', a:['الفيوم','سنورس','إطسا','طامية','يوسف الصديق','إبشواي','الفيوم الجديدة']},
  'بني سويف': {t:'مركز', a:['بني سويف','الواسطى','ناصر','إهناسيا','ببا','سمسطا','الفشن','بني سويف الجديدة']},
  'المنيا': {t:'مركز', a:['المنيا','العدوة','مغاغة','بني مزار','مطاي','سمالوط','أبو قرقاص','ملوي','دير مواس','المنيا الجديدة']},
  'أسيوط': {t:'مركز', a:['أسيوط','ديروط','القوصية','أبنوب','منفلوط','أبو تيج','الغنايم','ساحل سليم','البداري','صدفا','الفتح','أسيوط الجديدة']},
  'سوهاج': {t:'مركز', a:['سوهاج','أخميم','البلينا','المراغة','المنشاة','دار السلام','جرجا','جهينة','ساقلتة','طما','طهطا','العسيرات','سوهاج الجديدة']},
  'قنا': {t:'مركز', a:['قنا','أبوتشت','نجع حمادي','دشنا','الوقف','قفط','نقادة','فرشوط','قوص']},
  'الأقصر': {t:'مركز', a:['الأقصر','الزينية','البياضية','القرنة','أرمنت','الطود','إسنا']},
  'أسوان': {t:'مركز', a:['أسوان','دراو','كوم أمبو','نصر النوبة','إدفو','أبو سمبل','كلابشة','البصيلية','أسوان الجديدة']},
  'البحر الأحمر': {t:'مدينة', a:['الغردقة','رأس غارب','سفاجا','القصير','مرسى علم','الشلاتين','حلايب']},
  'الوادي الجديد': {t:'مركز', a:['الخارجة','الداخلة','الفرافرة','باريس','بلاط']},
  'مطروح': {t:'مدينة', a:['مرسى مطروح','الحمام','العلمين','الضبعة','النجيلة','سيدي براني','السلوم','سيوة']}
};
const KIDS = {
  /* ---------- القاهرة ---------- */
  'القاهرة|مدينة نصر': ['الحي الأول','الحي الثاني','الحي الثالث','الحي الرابع','الحي السادس','الحي السابع','الحي الثامن','الحي التاسع','الحي العاشر','مكرم عبيد','عباس العقاد','زهراء مدينة نصر'],
  'القاهرة|مصر الجديدة': ['هليوبوليس','الكوربة','روكسي','ألماظة','الحجاز','سانت فاتيما','ميريلاند','كلية البنات','قصر البارون'],
  'القاهرة|المعادي': ['دجلة','المعادي الجديدة','زهراء المعادي','سرايات المعادي','كورنيش المعادي','ثكنات المعادي','المعراج','عرب المعادي'],
  'القاهرة|حلوان': ['حلوان البلد','حدائق حلوان','عين حلوان','عزبة الوالدة'],
  'القاهرة|شبرا': ['الخلفاوي','سانت تيريز','الترعة البولاقية','مسرة','دوران شبرا'],
  'القاهرة|وسط البلد': ['التحرير','طلعت حرب','باب اللوق','التوفيقية','جاردن سيتي'],
  'القاهرة|السيدة زينب': ['قصر العيني','المنيرة','الحلمية'],
  'القاهرة|مصر القديمة': ['المنيل','الفسطاط','الملك الصالح','عين الصيرة','فم الخليج'],
  'القاهرة|المقطم': ['الهضبة الوسطى','الهضبة العليا','الأسمرات'],
  'القاهرة|عين شمس': ['غمرة','الألف مسكن','عزبة النخل','حلمية الزيتون','عين شمس الغربية','عين شمس الشرقية'],
  'القاهرة|المطرية': ['عزبة حمادة','المطرية البلد'],
  'القاهرة|التجمع الخامس': ['الحي الأول','الحي الثاني','الحي الثالث','الحي الرابع','الحي الخامس','النرجس','البنفسج','الياسمين','الأندلس','اللوتس','بيت الوطن','الدبلوماسيين','الشويفات','جاردينيا','المستثمرين الشمالية','المستثمرين الجنوبية'],
  'القاهرة|العاصمة الإدارية الجديدة': ['الحي الحكومي','الحي المالي','الحي الدبلوماسي','الحي السكني الأول','الحي السكني الثاني','الحي السكني الثالث','الحي السكني الخامس'],
  'القاهرة|الشروق': ['الحي الأول','الحي الثاني','الحي الثالث','الحي الرابع','الحي الخامس','الحي السادس','الحي السابع','الحي الثامن','الحي التاسع','الحي العاشر'],
  /* ---------- الجيزة ---------- */
  'الجيزة|الجيزة': ['المنيب','ساقية مكي','الجيزة البلد','بين السرايات','ميدان الجيزة'],
  'الجيزة|العجوزة': ['المهندسين','ميت عقبة','الكيت كات','أرض اللواء','العجوزة البلد','الصحفيين'],
  'الجيزة|الدقي': ['المساحة','ميدان الدقي','الجامعة'],
  'الجيزة|الهرم': ['المريوطية','الطالبية','ترسا','كفر غطاطي','نزلة السمان','المنصورية','شارع الهرم','الرماية'],
  'الجيزة|فيصل': ['الطوابق','كفر طهرمس','المشروع','العشرين'],
  'الجيزة|إمبابة': ['الكيت كات','وردان','المنيرة الغربية','أرض الجمعية','ناهيا','بشتيل'],
  'الجيزة|6 أكتوبر': ['الحي الأول','الحي الثاني','الحي الثالث','الحي الرابع','الحي الخامس','الحي السادس','الحي السابع','الحي الثامن','الحي التاسع','الحي العاشر','الحي الحادي عشر','الحي الثاني عشر','الحي المتميز','المحور المركزي','دريم لاند','الواحة'],
  'الجيزة|الشيخ زايد': ['الحي الأول','الحي الثاني','الحي الثالث','الحي الرابع','الحي الخامس','الحي السادس','الحي السابع','الحي الثامن','الحي التاسع','الحي العاشر','بيفرلي هيلز','الحي المتميز'],
  /* ---------- الإسكندرية ---------- */
  'الإسكندرية|حي شرق': ['سموحة','سيدي جابر','سان ستيفانو','كليوباترا','زيزينيا','جليم','رشدي','بولكلي','لوران','فلمنج','كامب شيزار','الإبراهيمية','جناكليس'],
  'الإسكندرية|حي وسط': ['محطة الرمل','محرم بك','كوم الدكة','الشاطبي'],
  'الإسكندرية|حي المنتزه أول': ['العصافرة','ميامي','المندرة','سيدي بشر','خالد بن الوليد'],
  'الإسكندرية|حي المنتزه ثان': ['المعمورة','أبو قير','بيانكي'],
  'الإسكندرية|حي غرب': ['الورديان','القباري','المكس','الدخيلة'],
  'الإسكندرية|حي الجمرك': ['المنشية','الأنفوشي','رأس التين','بحري','كرموز'],
  'الإسكندرية|حي العطارين': ['الأزاريطة','الشلالات','العطارين البلد'],
  'الإسكندرية|حي العجمي': ['الهانوفيل','البيطاش','الكيلو 21','أبيس','العجمي البلد','بنجر السكر'],
  'الإسكندرية|حي العامرية': ['العامرية البلد','الكينج مريوط','الكيلو 26'],
  'الإسكندرية|برج العرب': ['برج العرب البلد','القرية السياحية','الحي الأول','الحي الثاني','الحي الثالث'],
  /* ---------- محافظات أخرى ---------- */
  'القليوبية|شبرا الخيمة': ['شبرا الخيمة أول','شبرا الخيمة ثان','بهتيم','مسطرد'],
  'القليوبية|بنها': ['بنها البلد','كفر الجزار','مدينة بنها الجديدة'],
  'الدقهلية|المنصورة': ['توريل','حي الجامعة','المشاية','جديلة','سندوب'],
  'الغربية|طنطا': ['سيجر','الجلاء','حي أول طنطا','حي ثان طنطا'],
  'الشرقية|الزقازيق': ['حي أول الزقازيق','حي ثان الزقازيق','حي ثالث الزقازيق','القومية','المحطة'],
  'جنوب سيناء|شرم الشيخ': ['نعمة باي','خليج نبق','الهضبة','رأس أم سيد','شارك باي'],
  'البحر الأحمر|الغردقة': ['الدهار','السقالة','الكوثر','سهل حشيش','الجونة','مكادي','الممشى السياحي'],
  'الأقصر|الأقصر': ['الكرنك','البر الغربي','الطود'],
  'أسوان|أسوان': ['السيل','الصداقة','كورنيش أسوان']
};

function seedAllAreas() {
  const govs = rows('Governorates'), ex = rows('Areas');
  const seen = new Set(ex.map(a => a.gov_id + '|' + (a.parent_id||'') + '|' + a.name_ar));
  const idMap = {}; ex.forEach(a => { if (!a.parent_id) idMap[a.gov_id + '|' + a.name_ar] = a.id; });
  const add = [], newId = () => 'a' + Utilities.getUuid().slice(0, 8);
  Object.keys(GEO).forEach(gn => {
    const gv = govs.find(x => x.name_ar === gn);
    if (!gv) { Logger.log('محافظة مش موجودة: ' + gn); return; }
    GEO[gn].a.forEach(name => {
      let id = idMap[gv.id + '|' + name];
      if (!id) { id = newId(); idMap[gv.id + '|' + name] = id; add.push([id, gv.id, '', name, '', GEO[gn].t, true]); }
      (KIDS[gn + '|' + name] || []).forEach(kid => {
        const k = gv.id + '|' + id + '|' + kid;
        if (!seen.has(k)) { seen.add(k); add.push([newId(), gv.id, id, kid, '', 'حي فرعي', true]); }
      });
    });
  });
  const sh = sheet('Areas');
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 7).setValues(add);
  const v = sh.getDataRange().getValues();
  v.forEach((r, i) => { if (i > 0 && r[3] === 'حي المنتزه') sh.getRange(i + 1, 7).setValue(false); });
  Logger.log('تمت إضافة ' + add.length + ' منطقة');
}

/* ================= استيراد بيانات الجروبات المبدئية (63 جروب) - شغّل importStarterGroups مرة واحدة =================
   ترتيب التشغيل: setup  >  seedAllAreas  >  importStarterGroups
   كل الجروبات بتتسجل verification=pending و status=active، والمكرر (نفس الرابط) بيتتخطى. */
const STARTER_GROUPS = [
  ["كراكيب المعادي", "القاهرة", "المعادي", "https://www.facebook.com/groups/1625148161045894/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["فتكات", "مصر / غير محدد", "غير محدد", "https://www.facebook.com/groups/210847405917307/", "مجتمع / إعلانات محتملة", "رابط وارد في دليل منشور؛ المنطقة غير محددة"],
  ["أوليكس للبيع والشراء", "مصر / غير محدد", "غير محدد", "https://www.facebook.com/groups/262120607288356/", "بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["إعلن وبيع واشتري", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/adweny/", "إعلانات / بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["بيع واشتري واعلن منتجاتك", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/382751902093975/", "إعلانات / بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق القليوبية", "القليوبية", "عام", "https://www.facebook.com/groups/1883901455260783/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["كراكيب أوليكس كفر الشيخ", "كفر الشيخ", "عام", "https://www.facebook.com/groups/1102410609845588/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["تجار جملة", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/598870316925505/", "تجارة", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق سوهاج التجاري", "سوهاج", "عام", "https://www.facebook.com/groups/152820108754405/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق منتجات مشروعك", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/179971252211422/", "تجارة / تسويق", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق تجار الصيني والمصري", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/185283255152328/", "تجارة", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["أون لاين للبيع والإعلان", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/144829436273417/", "إعلانات / بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق الزقازيق", "الشرقية", "الزقازيق", "https://www.facebook.com/groups/1762962243777852/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["بيع واشتري كل حاجة", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/407247869463816/", "بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق أوليكس مصر للجميع", "مصر", "عام", "https://www.facebook.com/groups/526602297526229/", "بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["أوليكس مصر", "مصر", "عام", "https://www.facebook.com/groups/BMWGOOL/", "بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["تسويق", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/873556886017116/", "تسويق", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق الموسكي للتجارة", "القاهرة", "الموسكي", "https://www.facebook.com/groups/souk.elmosky/", "تجارة / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["تسويق المنصورة", "الدقهلية", "المنصورة", "https://www.facebook.com/groups/1691529584411086/", "تسويق / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق طنطا المحلة المنصورة", "الغربية / الدقهلية", "طنطا والمحلة والمنصورة", "https://www.facebook.com/groups/1564428440537057/", "بيع وشراء / محلي", "يغطي أكثر من مدينة بحسب الاسم المنشور"],
  ["سوق يا إسكندرية", "الإسكندرية", "عام", "https://www.facebook.com/groups/1695513353836141/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق الهرم حدائق الأهرام", "الجيزة", "الهرم وحدائق الأهرام", "https://www.facebook.com/groups/skan.harm/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق أهل إسكندرية", "الإسكندرية", "عام", "https://www.facebook.com/groups/942571062426613/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["اشتري وبيع أي حاجة أون لاين", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/811624625639679/", "بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق الجمعة بيع واشتري", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/120363981758387/", "بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["تجمع تجار الجملة 1", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/1111072832272523/", "تجارة جملة", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["تجمع تجار الجملة 2", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/150778865281525/", "تجارة جملة", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق بني سويف", "بني سويف", "عام", "https://www.facebook.com/groups/BNS.SOUQ/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق مصر الإلكتروني", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/samir.abas/", "بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق أوليكس مصر", "مصر", "عام", "https://www.facebook.com/groups/474216842726399/", "بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["جروب إعلانات عام 1", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/239137823209262/", "إعلانات", "الرابط مدرج في قائمة منشورة بدون اسم واضح"],
  ["جروب إعلانات عام 2", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/406797559743961/", "إعلانات", "الرابط مدرج في قائمة منشورة بدون اسم واضح"],
  ["جروب إعلانات عام 3", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/224825131218208/", "إعلانات", "الرابط مدرج في قائمة منشورة بدون اسم واضح"],
  ["Internet Egypt", "مصر", "عام", "https://www.facebook.com/groups/internetegypt/", "مجتمع / خدمات", "الرابط وارد في قائمة منشورة؛ نوع المحتوى الحالي غير متحقق"],
  ["جروب عام 5", "مصر / غير محدد", "غير محدد", "https://www.facebook.com/groups/133599668682722/", "غير محدد", "الرابط مدرج في قائمة منشورة بدون اسم واضح"],
  ["جروب عام 6", "مصر / غير محدد", "غير محدد", "https://www.facebook.com/groups/146715845344731/", "غير محدد", "الرابط مدرج في قائمة منشورة بدون اسم واضح"],
  ["جروب عام 7", "مصر / غير محدد", "غير محدد", "https://www.facebook.com/groups/712568422494483/", "غير محدد", "الرابط مدرج في قائمة منشورة بدون اسم واضح"],
  ["Jobs for U", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/jobsforu/", "وظائف", "قد لا يكون مناسبًا لتسويق الحج والعمرة؛ أدرج للمراجعة فقط"],
  ["بيع وشراء المستعمل والجديد بأسعار مناسبة", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/649909865889263/", "بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["إعلانات مجانية 1", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/613115032066442/", "إعلانات", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["إعلانات مجانية 2", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/2268042063412560/", "إعلانات", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["بيزنس", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/394302064561431/", "أعمال / تسويق", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["دعايا وإعلان مصر", "مصر", "عام", "https://www.facebook.com/groups/1789837277990259/", "إعلانات / تسويق", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["تسويق أون لاين أعلن مجانًا", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/715643595171454/", "تسويق", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوقلي", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/187282956673564/", "بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["هنبيعلك", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/1763518130587203/", "بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق مصر المفتوح", "مصر", "عام", "https://www.facebook.com/groups/3798394020386518/", "بيع وشراء", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["إعلانات لكل شيء مجانًا", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/482234708779611/", "إعلانات", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["يلا ماركتنج", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/1844532272460793/", "تسويق", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["فكرة مشروع", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/1421363801228663/", "مشروعات / أعمال", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["بيع وشراء القاهرة والجيزة", "القاهرة / الجيزة", "عام", "https://www.facebook.com/groups/327927701337608/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["بيع وشراء مدينة الشروق", "القاهرة", "مدينة الشروق", "https://www.facebook.com/groups/971602732866271/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["إعلانات تسويق العجمي", "الإسكندرية", "العجمي", "https://www.facebook.com/groups/Shopping.Alex/", "إعلانات / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["كراكيب شبرا", "القاهرة", "شبرا", "https://www.facebook.com/groups/1217482124999178/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق قليوب", "القليوبية", "قليوب", "https://www.facebook.com/groups/641790932990793/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق المنيب", "الجيزة", "المنيب", "https://www.facebook.com/groups/272161676578428/", "بيع وشراء / محلي", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["تسويق أون لاين مجانًا", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/170477763476279/", "تسويق", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["سوق لعب الأطفال", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/212816050361394/", "بيع وشراء متخصص", "غير مستهدف مباشرة لتسويق الحج والعمرة"],
  ["سوق الجملة", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/343233270817257/", "تجارة جملة", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["إعلانات سوق مصر", "مصر", "عام", "https://www.facebook.com/groups/3691631060852577/", "إعلانات", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["إعلانات 4You", "مصر / غير محدد", "عام", "https://www.facebook.com/groups/ads.4you/", "إعلانات", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["إعلانات تجار مصر", "مصر", "عام", "https://www.facebook.com/groups/516084155072544/", "تجارة / إعلانات", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"],
  ["الإعلانات في مصر", "مصر", "عام", "https://www.facebook.com/groups/239513831623854/", "إعلانات", "رابط وارد في دليل منشور؛ لم يتم التحقق من فتحه الآن"]
];

function importStarterGroups() {
  const ss = SpreadsheetApp.getActive();
  // محافظة "عام" للجروبات اللي مش مرتبطة بمحافظة محددة
  const GEN = 'مصر - عام (غير محدد)';
  let govs = rows('Governorates');
  if (!govs.some(g => g.name_ar === GEN)) {
    sheet('Governorates').appendRow(['g' + pad(govs.length + 1), GEN, 'Egypt-General', true]);
    govs = rows('Governorates');
  }
  const areas = rows('Areas');
  const cats = rows('Categories').map(c => c.name_ar);
  const seen = new Set(rows('Groups').map(g => norm(String(g.url))));
  const now = new Date().toISOString(), out = [], skipped = [];

  const mapCat = t => {
    if (/وظائف/.test(t)) return 'وظائف';
    if (/مجتمع/.test(t) && !/بيع/.test(t)) return 'مجتمع محلي';
    if (/بيع|تجارة|إعلان|تسويق|أعمال|مشروع/.test(t)) return 'تجارة وبيع';
    return 'عام';
  };

  STARTER_GROUPS.forEach(r => {
    const name = r[0], govTxt = r[1], areaTxt = r[2], url = r[3], type = r[4], note = r[5];
    const key = norm(url);
    if (seen.has(key)) { skipped.push(name + ' (مكرر)'); return; }
    seen.add(key);
    // المحافظة: أول اسم محافظة مذكور، وإلا "عام"
    const parts = govTxt.split('/').map(s => s.trim());
    let gov = null;
    for (const p of parts) { gov = govs.find(g => g.name_ar === p); if (gov) break; }
    if (!gov) gov = govs.find(g => g.name_ar === GEN);
    const extraGov = parts.length > 1 && gov.name_ar !== GEN ? ' | يغطي أيضًا: ' + parts.filter(p => p !== gov.name_ar).join('، ') : '';
    // المنطقة: أطول اسم منطقة (بدون parent) داخل المحافظة موجود في نص المنطقة
    let areaId = '', sub = '';
    if (areaTxt && areaTxt !== 'عام' && areaTxt !== 'غير محدد') {
      sub = areaTxt;
      const hit = areas.filter(a => a.gov_id === gov.id && !a.parent_id && areaTxt.indexOf(a.name_ar) >= 0)
                       .sort((a, b) => String(b.name_ar).length - String(a.name_ar).length)[0];
      if (hit) areaId = hit.id;
    }
    const cat = mapCat(type);
    const notes = ('النوع: ' + type + ' | السماح بالإعلانات: غير معروف | ' + note + extraGov).slice(0, 500);
    out.push(['G' + Utilities.getUuid().slice(0, 8).toUpperCase(), name, 'facebook', url, gov.id, areaId, sub, cat, '', 'استيراد', now, now, 'pending', 'active', notes, '', false]);
  });

  if (out.length) { const sh = sheet('Groups'); sh.getRange(sh.getLastRow() + 1, 1, out.length, out[0].length).setValues(out); }
  log('استيراد', 'importStarterGroups', '', out.length + ' added, ' + skipped.length + ' skipped');
  Logger.log('تمت إضافة ' + out.length + ' جروب، وتخطي ' + skipped.length + (skipped.length ? ': ' + skipped.join(' ، ') : ''));
}
