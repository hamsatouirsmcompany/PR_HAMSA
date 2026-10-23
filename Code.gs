/**
 * Hamsa Social Groups - Backend (Google Apps Script)
 * الاستخدام: افتح Google Sheet > Extensions > Apps Script > الصق الملف ده
 * 1) غيّر ADMIN_NAME و ADMIN_CODE تحت  2) شغّل setup() مرة واحدة  3) Deploy > Web app
 */
const ADMIN_NAME = 'المدير';
const ADMIN_CODE = 'Hamsa2026Admin'; // كود دخول الأدمن (غيّره - بيتخزن مشفّر SHA-256)

const SCHEMA = {
  Groups: ['id','name','platform','url','gov_id','area_id','sub_area','category','description','added_by','created_at','updated_at','verification','status','notes','members','is_demo'],
  Governorates: ['id','name_ar','name_en','active'],
  Areas: ['id','gov_id','parent_id','name_ar','name_en','type','active'],
  Employees: ['id','name','role','code_hash','active'],
  Categories: ['id','name_ar'],
  ActivityLog: ['time','employee','action','target','details'],
  Settings: ['key','value']
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
  if (e.getLastRow() < 2) e.appendRow(['e1',ADMIN_NAME,'admin',hash(ADMIN_CODE),true]);
  const s = ss.getSheetByName('Settings');
  if (s.getLastRow() < 2) s.appendRow(['geo_dataset','بيانات جغرافية مبدئية وغير مكتملة - وسّعها بالاستيراد']);
}

// إضافة موظف: غيّر القيم وشغّل الدالة من المحرر
function addEmployee() {
  const name = 'اسم الموظف', role = 'employee', code = 'كود-الموظف';
  const sh = SpreadsheetApp.getActive().getSheetByName('Employees');
  sh.appendRow(['e'+(sh.getLastRow()+1), name, role, hash(code), true]);
}

/* ---------- Helpers ---------- */
function pad(n){ return ('0'+n).slice(-2); }
function hash(s){ return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(b => ('0'+(b&255).toString(16)).slice(-2)).join(''); }
function sheet(n){ return SpreadsheetApp.getActive().getSheetByName(n); }
function rows(n){ const v = sheet(n).getDataRange().getValues(), h = v.shift(); return v.map(r => { const o = {}; h.forEach((k,i)=>o[k]=r[i]); return o; }); }
function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function log(emp, action, target, details){ sheet('ActivityLog').appendRow([new Date().toISOString(), emp, action, target, details||'']); }
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
function doGet(){ return out({ok:true, msg:'Hamsa Social Groups API'}); }

function doPost(e) {
  let req; try { req = JSON.parse(e.postData.contents); } catch(x){ return out({ok:false,error:'طلب غير صالح'}); }
  const emp = rows('Employees').find(x => String(x.active).toUpperCase() === 'TRUE' && String(x.code_hash).trim() === hash(String(req.code||'').trim()));
  if (!emp) return out({ok:false, error:'كود الدخول غير صحيح', auth:false}); // التحقق على السيرفر
  const isAdmin = emp.role === 'admin';
  const lock = LockService.getScriptLock();
  try {
    switch (req.action) {
      case 'login': return out({ok:true, user:{name:emp.name, role:emp.role}});
      case 'bootstrap': return out({ok:true, user:{name:emp.name, role:emp.role},
        govs: rows('Governorates').filter(g=>String(g.active).toUpperCase()==='TRUE'), areas: rows('Areas').filter(a=>String(a.active).toUpperCase()==='TRUE'), cats: rows('Categories'), stats: stats()});
      case 'groups': return out(Object.assign({ok:true}, queryGroups(req)));
      case 'addGroup': { lock.waitLock(15000); const r = addGroup(req.data, emp); lock.releaseLock(); return out(r); }
      case 'setStatus': { lock.waitLock(15000); const r = updateGroup(req.id, req.patch, emp, isAdmin); lock.releaseLock(); return out(r); }
      case 'reportBroken': { lock.waitLock(15000); const r = updateGroup(req.id, {status:'broken'}, emp, true); lock.releaseLock(); return out(r); }
      case 'deleteGroup': { if (!isAdmin) return out({ok:false,error:'مش مسموح'}); lock.waitLock(15000); const r = deleteGroup(req.id, emp); lock.releaseLock(); return out(r); }
      case 'addArea': { if (!isAdmin) return out({ok:false,error:'مش مسموح'}); lock.waitLock(15000);
        const d = req.data; if (!d.name_ar || !d.gov_id) { lock.releaseLock(); return out({ok:false,error:'الاسم والمحافظة مطلوبين'}); }
        const sh = sheet('Areas'); const id = 'a'+Utilities.getUuid().slice(0,8);
        sh.appendRow([id,d.gov_id,d.parent_id||'',d.name_ar,d.name_en||'',d.type||'حي',true]); log(emp.name,'addArea',id,d.name_ar); lock.releaseLock(); return out({ok:true,id}); }
      case 'importAreas': { if (!isAdmin) return out({ok:false,error:'مش مسموح'}); lock.waitLock(30000);
        // rows: [{gov_id,parent_id,name_ar,name_en,type}] - بيرجع تقرير بالمرفوض
        const ex = rows('Areas'), add = [], bad = [];
        (req.rows||[]).forEach((r,i) => { if (!r.name_ar || !r.gov_id) return bad.push({row:i+1,reason:'ناقص اسم أو محافظة'});
          if (ex.some(a=>a.gov_id===r.gov_id && a.name_ar===r.name_ar && String(a.parent_id)===String(r.parent_id||''))) return bad.push({row:i+1,reason:'مكرر'});
          add.push(['a'+Utilities.getUuid().slice(0,8),r.gov_id,r.parent_id||'',r.name_ar,r.name_en||'',r.type||'حي',true]); });
        if (add.length) sheet('Areas').getRange(sheet('Areas').getLastRow()+1,1,add.length,7).setValues(add);
        log(emp.name,'importAreas','',add.length+' added'); lock.releaseLock(); return out({ok:true,added:add.length,rejected:bad}); }
      case 'activity': { if (!isAdmin) return out({ok:false,error:'مش مسموح'}); return out({ok:true, items: rows('ActivityLog').slice(-200).reverse()}); }
      default: return out({ok:false,error:'أمر غير معروف'});
    }
  } catch (err) { try{lock.releaseLock();}catch(_){} return out({ok:false,error:String(err)}); }
}

function addGroup(d, emp) {
  if (!d || !d.name || !d.url || !d.gov_id) return {ok:false, error:'الاسم والرابط والمحافظة مطلوبين'};
  if (!/^https?:\/\/\S+\.\S+/i.test(d.url)) return {ok:false, error:'الرابط غير صالح'};
  const platform = detectPlatform(d.url);
  if (d.platform && d.platform !== platform && platform !== 'other') return {ok:false, error:'الرابط مش تابع للمنصة المختارة (المنصة المكتشفة: '+platform+')'};
  const key = norm(d.url);
  if (rows('Groups').some(g => norm(String(g.url)) === key)) return {ok:false, duplicate:true, error:'الرابط ده متسجل قبل كده'};
  const now = new Date().toISOString(), id = 'G'+Utilities.getUuid().slice(0,8).toUpperCase();
  sheet('Groups').appendRow([id,d.name,d.platform||platform,d.url.trim(),d.gov_id,d.area_id||'',d.sub_area||'',d.category||'',d.description||'',emp.name,now,now,'pending','active',d.notes||'',d.members||'',false]);
  log(emp.name,'create',id,d.name);
  return {ok:true, id, stats: stats()};
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
  let list = rows('Groups');
  const t = (q.q||'').trim().toLowerCase();
  if (t) list = list.filter(g => [g.name,g.url,g.category,g.description,g.sub_area].join(' ').toLowerCase().indexOf(t) >= 0);
  ['gov_id','area_id','platform','status','verification','added_by'].forEach(k => { if (q[k]) list = list.filter(g => String(g[k])===q[k]); });
  const s = q.sort||'newest';
  list.sort((a,b)=> s==='oldest' ? (a.created_at>b.created_at?1:-1) : s==='name' ? String(a.name).localeCompare(String(b.name),'ar') : s==='updated' ? (a.updated_at<b.updated_at?1:-1) : (a.created_at<b.created_at?1:-1));
  const size = 24, page = Math.max(1, q.page||1), total = list.length;
  return {total, page, size, items: list.slice((page-1)*size, page*size)};
}

function stats() {
  const g = rows('Groups'), today = new Date().toISOString().slice(0,10), by = {}, per = {};
  g.forEach(x => { by[x.platform]=(by[x.platform]||0)+1; per[x.added_by]=(per[x.added_by]||0)+1; });
  return {total:g.length, today:g.filter(x=>String(x.created_at).slice(0,10)===today).length, byPlatform:by, byEmployee:per,
    pending:g.filter(x=>x.verification==='pending').length, broken:g.filter(x=>x.status==='broken').length,
    areas: rows('Areas').length, govs: 27,
    recent: g.slice(-5).reverse().map(x=>({id:x.id,name:x.name,platform:x.platform,added_by:x.added_by}))};
}

/* ---------- لو الكود بيطلع غلط: غيّر الكود هنا وشغّل الدالة دي ---------- */
function resetAdminCode() {
  const code = 'Hamsa2026Admin';   // اكتب الكود الجديد هنا (يفضل حروف إنجليزي وأرقام)
  const sh = sheet('Employees');
  const v = sh.getDataRange().getValues();
  const i = v.findIndex((r, n) => n > 0 && r[2] === 'admin');
  if (i < 0) { sh.appendRow(['e1', 'المدير', 'admin', hash(code), true]); }
  else { sh.getRange(i + 1, 4).setValue(hash(code)); sh.getRange(i + 1, 5).setValue(true); }
}

/* ---------- تشخيص: شغّل الدالة دي وبعدين افتح Execution log ---------- */
function diagnose() {
  const code = 'Hamsa2026Admin';  // نفس الكود اللي بتكتبه في الموقع
  const sh = SpreadsheetApp.getActive().getSheetByName('Employees');
  if (!sh) { Logger.log('❌ مفيش شيت Employees - شغّل setup'); return; }
  const list = rows('Employees');
  Logger.log('عدد الموظفين: ' + list.length);
  list.forEach(x => Logger.log(x.name + ' | role=' + x.role + ' | active=' + x.active + ' | hash=' + String(x.code_hash).slice(0,12) + '...'));
  Logger.log('الهاش المتوقع للكود: ' + hash(code).slice(0,12) + '...');
  const ok = list.some(x => String(x.active).toUpperCase()==='TRUE' && String(x.code_hash).trim()===hash(code));
  Logger.log(ok ? '✅ الكود صحيح في الشيت - المشكلة في الـ Deploy أو رابط الـ API' : '❌ الكود مش مطابق - شغّل resetAdminCode');
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
