/**
 * نظام إقفال صناديق الكاشير
 * الخادم: المستخدمون والصلاحيات وحفظ البيانات في جدول Google
 */
const TZ = 'Asia/Riyadh';
const DATA = '_data', USERS = '_users', SETTINGS = '_settings', LOG = 'سجل الإقفالات';
const NETS = [['gulf', 'شبكة خليجية'], ['master', 'ماستركارد'], ['mada', 'شبكة سعودية (مدى)'], ['visa', 'فيزا'], ['amex', 'أمريكان إكسبرس']];
const DAYS = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
const ROLE_NAMES = { acc: 'محاسب الفرع', mgr: 'مدير الفرع', sup: 'مسؤول الكاشيرات', lead: 'مشرفة الكاشيرات', own: 'المالك' };
/* what each role may be given, and what it gets by default */
const ROLE_PERMS = {
  acc: { allowed: ['boxes', 'summary', 'reports', 'commission', 'att', 'log', 'cashbox', 'petty', 'branches', 'settings'], def: { tabs: ['boxes', 'summary', 'reports', 'commission', 'att', 'log', 'cashbox', 'petty', 'branches', 'settings'], xlsx: true, wa: false, days: 0 } },
  mgr: { allowed: ['short', 'sales', 'boxes', 'summary', 'reports', 'commission', 'att', 'log', 'cashbox', 'petty'], def: { tabs: ['short', 'sales'], xlsx: false, wa: false, days: 7 } },
  sup: { allowed: ['boxes', 'summary', 'short'], def: { tabs: ['boxes', 'summary'], xlsx: true, wa: false, days: 7 } },
  lead: { allowed: ['cash', 'short'], def: { tabs: ['cash', 'short'], xlsx: false, wa: true, days: 7 } },
  own: { allowed: ['branches', 'summary', 'reports', 'sales', 'short', 'boxes', 'commission', 'att', 'log', 'cashbox', 'petty'], def: { tabs: ['branches', 'summary', 'reports', 'sales', 'short'], xlsx: true, wa: false, days: 0 } }
};
const DAY_CHOICES = [2, 7, 14, 30, 60, 90];
function perms_(u) {
  const R = ROLE_PERMS[u.role] || ROLE_PERMS.sup, d = R.def;
  if (u.role === 'acc') return JSON.parse(JSON.stringify(d));
  let o = {}; try { o = u.perms ? (typeof u.perms === 'string' ? JSON.parse(u.perms) : u.perms) : {}; } catch (e) { o = {}; }
  let tabs = R.allowed.filter(function (k) { return (Array.isArray(o.tabs) ? o.tabs : d.tabs).indexOf(k) >= 0; });
  if (!tabs.length) tabs = d.tabs.slice();
  // sending to staff is the lead's by default; the accountant may also give it to a cashier supervisor
  const wa = (u.role === 'lead' || u.role === 'sup') ? (typeof o.wa === 'boolean' ? o.wa : d.wa) : false;
  if (wa && tabs.indexOf('short') < 0) tabs.push('short');
  // an owner limited to totals sees the branch report only: cash, networks and their total
  const totals = u.role === 'own' && o.totals === true;
  if (totals) tabs = ['branches'];
  return {
    totals: u.role === 'own' ? totals : undefined,
    // the cashier supervisor's entries either wait for the accountant's approval or count in the reports at once; unset follows the company default
    approve: u.role === 'sup' && typeof o.approve === 'boolean' ? o.approve : undefined,
    tabs: tabs,
    xlsx: typeof o.xlsx === 'boolean' ? o.xlsx : d.xlsx,
    wa: wa,
    days: d.days ? (DAY_CHOICES.indexOf(Number(o.days)) >= 0 ? Number(o.days) : d.days) : 0,
    branches: u.role === 'own' ? (Array.isArray(o.branches) ? o.branches.map(String).slice(0, 50) : '*') : undefined
  };
}
function windowDates_(t, n) { const o = []; for (let k = 0; k < n; k++) o.push(addDays_(t, -k)); return o; }
const SENT = 'سجل إشعارات العجز';
/* ---------- branches: each branch keeps its own sheets on the same spreadsheet; the first branch keeps the original names ---------- */
let BR = '';
function brCode_(id) { return 'ف' + String(id).replace(/^f/, ''); }
function bname_(name) { if (!BR || name === USERS || name === LOGO || name === WMS) return name; return name + ' - ' + brCode_(BR); }
function SS_() {
  const ss = SpreadsheetApp.getActive();
  return { getSheetByName: function (n) { return ss.getSheetByName(bname_(n)); }, insertSheet: function (n, i) { return i == null ? ss.insertSheet(bname_(n)) : ss.insertSheet(bname_(n), i); },
    getSheets: function () { return ss.getSheets(); }, deleteSheet: function (x) { return ss.deleteSheet(x); }, getUrl: function () { return ss.getUrl(); } };
}
function withBr_(id, fn) { const k = BR; BR = id; try { return fn(); } finally { BR = k; } }
const COMPANY_KEYS = ['store', 'logoV', 'logoPrint', 'logoWm', 'wmV', 'wmLevel', 'wmMode', 'branches'];
function mainSettings_() { return withBr_('', getSettings_); }
function branchIds_() { return ['main'].concat((mainSettings_().branches || []).map(function (b) { return b.id; })); }
function branchName_(id) {
  const m = mainSettings_();
  if (!id || id === 'main') return m.branch || 'الفرع الرئيسي';
  const b = (m.branches || []).filter(function (x) { return x.id === id; })[0]; return b ? b.name : id;
}
function isAdmin_(u) { return u.role === 'acc' && (!u.branch || u.branch === '*'); }
/* the system manager is the first all-branches accountant (the one who set the system up); company identity, Excel import, users and branches are his, or of an accountant he gives full permissions */
let SUPER_ = null;
function superName_() { if (SUPER_ === null) { const f = users_().find(function (x) { return isAdmin_(x); }); SUPER_ = f ? f.username : ''; } return SUPER_; }
function isSuper_(u) { return isAdmin_(u) && u.username === superName_(); }
function isTop_(u) { return isAdmin_(u) && (isSuper_(u) || String(u.perms) === 'FULL'); }
function userBranches_(u) {
  const all = branchIds_();
  if (isAdmin_(u)) return all;
  if (u.role === 'own') { const pb = perms_(u).branches; if (pb === '*' || !Array.isArray(pb)) return all; const L = all.filter(function (id) { return pb.indexOf(id) >= 0; }); return L.length ? L : ['main']; }
  return [all.indexOf(u.branch) >= 0 ? u.branch : 'main'];
}
function pickBranch_(u, want) { const al = userBranches_(u), id = al.indexOf(String(want || '')) >= 0 ? String(want) : al[0]; return id === 'main' ? '' : id; }
function delKey_() { return BR ? 'del_' + BR : 'del'; }
const MAX_BOXES = 60, DEF_BOXES = 20;
function nBoxes_(st) { return Math.max(1, Math.min(MAX_BOXES, parseInt(st && st.boxSlots, 10) || DEF_BOXES)); }
/* networks shown in sheets: the enabled ones, plus any hidden one that still holds an amount */
function netsShown_(st, boxesList) {
  const off = (st && st.netsOff) || [];
  return NETS.filter(function (x) { return off.indexOf(x[0]) < 0 || (boxesList || []).some(function (b) { return b && Number(b[x[0]]); }); });
}
const SESSION_MS = 365 * 24 * 3600 * 1000;
/* foreign currencies a box may receive: code, name, a starting exchange rate to the riyal (editable in settings) */
const FX = [['USD', 'دولار أمريكي', 3.75], ['AED', 'درهم إماراتي', 1.02], ['KWD', 'دينار كويتي', 12.2], ['BHD', 'دينار بحريني', 9.95], ['QAR', 'ريال قطري', 1.03], ['OMR', 'ريال عماني', 9.74]];
/* the cash of a box = riyals counted + foreign cash at the rate it was received; the rate stays with the box */
function fxIn_(o, b, pb, st) {
  const n = function (v) { return Math.round((Number(v) || 0) * 100) / 100; };
  const F = st.fx || {}, d = {}, r = {}; let sum = 0, any = false;
  FX.forEach(function (x) {
    const c = x[0], a = n((b.fxd || {})[c]); if (!(a > 0) || a > 10000000) return;
    const rate = Number(((pb && pb.fxr) || {})[c]) || Number(((F.cur || {})[c] || {}).rate) || x[2];
    d[c] = a; r[c] = rate; sum += a * rate; any = true;
  });
  if (any) { o.cashSar = n(b.cashSar != null ? b.cashSar : b.cash); o.fxd = d; o.fxr = r; o.cash = n(o.cashSar + sum); }
  else if (b.cashSar != null) o.cash = n(b.cashSar);
}
const DEFAULT_SETTINGS = {
  store: '', branch: '',
  cashiers: [], shifts: ['صباحية', 'مسائية'],
  signA: 'توقيع أمين الصندوق', signB: 'توقيع مسؤول الكاشيرات',
  rate: 0.75, rateNet: true, supPrev: false, supHours: 12, boxStart: 1,
  boxSlots: 20, boxNos: [], boxLabels: [], boxOff: [], netsOff: ['amex'], logoV: '', logoPrint: true, logoWm: true, wmV: '', wmLevel: 1, wmMode: 'cover', supApprove: true, allowFuture: false, modCashbox: true, modPetty: false,
  boxNames: ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  staff: {}
};

const APP_VERSION = '4.5';
/* the accountant may change a day for this many days; older days become view-only */
const ACC_EDIT_DAYS = 30;
/* the visible per-day sheets are kept this many months (all data stays in the system permanently) */
const KEEP_MONTHS = 18;
function periodOpen_(date) { return date >= addDays_(today_(), -ACC_EDIT_DAYS); }
const FUTURE_MSG = 'لا يمكن الإدخال بتاريخ لاحق لتاريخ اليوم. يقدر المحاسب يسمح بذلك من الإعدادات ← قواعد التشغيل';
const CLOSED_MSG = 'هذا اليوم ضمن فترة مغلقة (أقدم من ' + ACC_EDIT_DAYS + ' يوماً) ومتاح للاطلاع فقط';

/* the cashier lead sees cash and shortages for this many days (today included) */
const LEAD_DAYS = 7;
function leadDates_(t) { const o = []; for (let k = 0; k < LEAD_DAYS; k++) o.push(addDays_(t, -k)); return o; }

/* company logo: kept as an image data URL in a hidden sheet, split into cells */
const LOGO = '_logo';
/* images kept as data URLs in hidden sheets: the company logo and the watermark picture */
function img_(sheet, key) {
  const s = SS_().getSheetByName(sheet), v = mainSettings_()[key] || '';
  if (!s || !v || s.getLastRow() < 1) return { v: '', data: '' };
  return { v: v, data: s.getRange(1, 1, s.getLastRow(), 1).getValues().map(function (r) { return String(r[0]); }).join('') };
}
function saveImg_(sheet, key, p, label) {
  const d = String(p.data || '');
  if (d && (!/^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+\/=]+$/.test(d) || d.length > 400000)) throw new Error('صورة ' + label + ' غير صالحة أو كبيرة جداً');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const s = hiddenSheet_(sheet); s.clear();
    if (d) {
      const parts = []; for (let i = 0; i < d.length; i += 40000) parts.push([d.slice(i, i + 40000)]);
      ensureRows_(s, parts.length); s.getRange(1, 1, parts.length, 1).setNumberFormat('@').setValues(parts);
    }
    const v = d ? String(Date.now()) : '';
    withBr_('', function () { const st = getSettings_(); st[key] = v; hiddenSheet_(SETTINGS).getRange(1, 1).setValue(JSON.stringify(st)); });
    bump_();
    return { v: v };
  } finally { lock.releaseLock(); }
}
const WMS = '_watermark';
function logo_() { return img_(LOGO, 'logoV'); }
function saveLogo_(p) { return saveImg_(LOGO, 'logoV', p, 'الشعار'); }
function wm_() { return img_(WMS, 'wmV'); }
function saveWm_(p) { return saveImg_(WMS, 'wmV', p, 'العلامة المائية'); }
/* the picture used as watermark: the chosen one, or the company logo */
function wmSrc_() { const st = getSettings_(); if (st.logoWm === false) return ''; return (st.wmV && wm_().data) || logo_().data || ''; }

function jsSum_(t) {
  let h = 0;
  for (let i = 0; i < t.length; i++) h = (h * 31 + t.charCodeAt(i)) % 1000000007;
  return h;
}

function doGet(e) {
  BR = '';
  if (e && e.parameter && e.parameter.n) return noticePage_(String(e.parameter.n));
  const problems = [], parts = [];
  [['Js1', typeof APP_JS_1 !== 'undefined' ? APP_JS_1 : null], ['Js2', typeof APP_JS_2 !== 'undefined' ? APP_JS_2 : null], ['Js3', typeof APP_JS_3 !== 'undefined' ? APP_JS_3 : null]]
    .forEach(function (x, k) {
      if (!x[1]) problems.push('الملف <b>' + x[0] + '.gs</b> غير موجود في المشروع، أو اسمه مكتوب بشكل مختلف.');
      else {
        const mk = String(x[1][x[1].length - 1]).split(':'), body = x[1].slice(0, -1).join('\n');
        if (mk[0] !== '//@@END' + (k + 1)) problems.push('الملف <b>' + x[0] + '.gs</b> لم يُلصق كاملاً، احذف محتواه والصقه من جديد.');
        else if (mk.length === 3 && (Number(mk[1]) !== body.length || Number(mk[2]) !== jsSum_(body)))
          problems.push('الملف <b>' + x[0] + '.gs</b> تغيّر أثناء النسخ (حرف ناقص أو زائد)، احذف محتواه والصقه من جديد من الملف الأصلي.');
        else parts.push(body);
      }
    });
  let html = '';
  try { html = HtmlService.createHtmlOutputFromFile('Index').getContent(); } catch (e) { problems.push('الملف <b>Index</b> غير موجود في المشروع.'); }
  if (html && html.indexOf('appSlot') < 0) problems.push('الملف <b>Index</b> قديم أو ناقص، الصق النسخة الجديدة منه.');
  if (problems.length) {
    return HtmlService.createHtmlOutput(
      '<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;max-width:560px;margin:30px auto;padding:20px;line-height:1.9;font-size:16px">' +
      '<h2 style="color:#b3261e;margin-top:0">النظام يحتاج إكمال التركيب</h2>' +
      problems.map(function (x) { return '<p>• ' + x + '</p>'; }).join('') +
      '<p>الحل: افتح الملف المذكور في محرر Apps Script، امسح محتواه كاملاً، والصق محتوى الملف بنفس الاسم من جديد، ثم احفظ وانشر إصداراً جديداً.</p>' +
      '<p style="color:#666;font-size:13px">الإصدار ' + APP_VERSION + '</p></div>')
      .setTitle('إقفال صناديق الكاشير').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // the program travels Base64-encoded so nothing between the server and the browser can alter it
  const b64 = Utilities.base64Encode(parts.join('\n'), Utilities.Charset.UTF_8);
  const script = '<script>(function(){try{var b=atob("' + b64 + '"),u=new Uint8Array(b.length);for(var i=0;i<b.length;i++)u[i]=b.charCodeAt(i);' +
    'var s=document.createElement("script");s.text="(function(){\\n"+new TextDecoder("utf-8").decode(u)+"\\n})();";document.body.appendChild(s);}' +
    'catch(e){if(window.__bootErr)__bootErr("تعذّر تحميل البرنامج: "+e.message,true)}})();</script>';
  const k = html.lastIndexOf('</body>');
  html = k >= 0 ? html.slice(0, k) + script + html.slice(k) : html + script;
  return HtmlService.createHtmlOutput(html)
    .setTitle('إقفال صناديق الكاشير')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- API ---------- */
function api(token, action, p) {
  p = p || {}; BR = ''; SUPER_ = null;
  try {
    if (action === 'status') { const s0 = getSettings_(); return ok_({ hasUsers: users_().length > 0, logoV: s0.logoV || '', wmV: s0.wmV || '', logoWm: s0.logoWm !== false, wmLevel: s0.wmLevel == null ? 1 : s0.wmLevel, wmMode: s0.wmMode || 'cover' }); }
    if (action === 'logo') return ok_(logo_());
    if (action === 'wm') return ok_(wm_());
    if (action === 'setup') return ok_(setup_(p));
    if (action === 'login') return ok_(login_(p));
    const u = session_(token);
    if (!u) return err_('انتهت الجلسة، سجّل الدخول مرة أخرى', 'auth');
    const only = function (roles) { if (roles.indexOf(u.role) < 0) throw new Error('ليست لديك صلاحية لهذا الإجراء'); };
    const admin = function () { if (!isTop_(u)) throw new Error('هذا الإجراء لمدير النظام فقط'); };
    const view = function (tabKey) { if (!(u.role === 'acc' || ((u.role === 'mgr' || u.role === 'own') && perms_(u).tabs.indexOf(tabKey) >= 0))) throw new Error('ليست لديك صلاحية لهذا الإجراء'); };
    BR = pickBranch_(u, p.br);
    switch (action) {
      case 'boot': return ok_(boot_(u, p));
      case 'poll': return ok_(poll_(u, p));
      case 'logout': PropertiesService.getScriptProperties().deleteProperty('s_' + token); CacheService.getScriptCache().remove('u_' + token); return ok_({});
      case 'saveDay': only(['acc', 'sup']); return ok_(saveDay_(u, p));
      case 'deleteDay': only(['acc']); if (!periodOpen_(String(p.date || ''))) throw new Error(CLOSED_MSG); deleteDay_(String(p.date || '')); return ok_({});
      case 'saveSettings': only(['acc']); if (!isTop_(u)) ['store', 'logoPrint', 'logoWm', 'wmLevel', 'wmMode'].forEach(function (k) { delete p[k]; }); return ok_(saveSettings_(p));
      case 'users': admin(); return ok_(users_().map(pub_));
      case 'saveUser': admin(); return ok_(saveUser_(u, p));
      case 'deleteUser': admin(); return ok_(deleteUser_(u, p));
      case 'saveBranch': admin(); return ok_(saveBranch_(u, p));
      case 'deleteBranch': admin(); return ok_(deleteBranch_(u, p));
      case 'brReport': if (!(isAdmin_(u) || u.role === 'own')) throw new Error('ليست لديك صلاحية لهذا الإجراء'); return ok_(brReport_(u, p));
      case 'importDays': admin(); return ok_(importDays_(u, p));
      case 'markSent': only(['lead', 'sup']); return ok_(markSent_(u, p));
      case 'sendMail': only(['lead', 'sup']); return ok_(sendMail_(u, p));
      case 'approve': only(['acc']); return ok_(approve_(u, p));
      case 'saveLogo': admin(); return ok_(saveLogo_(p));
      case 'saveWm': admin(); return ok_(saveWm_(p));
      case 'moveDay': only(['acc']); return ok_(moveDay_(u, p));
      case 'moveBox': only(['acc']); return ok_(moveBox_(u, p));
      case 'cashbox': view('cashbox'); return ok_(cbGet_(String(p.month || '')));
      case 'cbSave': only(['acc']); return ok_(cbSave_(u, p));
      case 'cbDel': only(['acc']); return ok_(cbDel_(u, p));
      case 'cbOpen': only(['acc']); return ok_(cbOpen_(u, p));
      case 'cbCarry': only(['acc']); return ok_(cbCarry_(u, p));
      case 'cbUncarry': only(['acc']); return ok_(cbUncarry_(u, p));
      case 'petty': view('petty'); return ok_(ptGet_());
      case 'ptPerson': only(['acc']); return ok_(ptPerson_(u, p));
      case 'ptTx': only(['acc']); return ok_(ptTx_(u, p));
      case 'ptDel': only(['acc']); return ok_(ptDel_(u, p));
      case 'ptSettle': only(['acc']); return ok_(ptSettle_(u, p));
      case 'ptUnsettle': only(['acc']); return ok_(ptUnsettle_(u, p));
      case 'audit': view('log'); return ok_(audit_(p));
    }
    return err_('إجراء غير معروف');
  } catch (e) {
    return err_(e && e.message ? e.message : String(e));
  }
}
function ok_(d) { return JSON.stringify({ ok: true, data: d }); }
function err_(m, c) { return JSON.stringify({ ok: false, msg: m, code: c || '' }); }

/* ---------- helpers ---------- */
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function addDays_(iso, k) { const d = new Date(iso + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + k); return d.toISOString().slice(0, 10); }
function dayName_(iso) { return DAYS[new Date(iso + 'T12:00:00Z').getUTCDay()]; }
function norm_(s) {
  return String(s == null ? '' : s)
    .replace(/[٠-٩]/g, function (d) { return '٠١٢٣٤٥٦٧٨٩'.indexOf(d); })
    .replace(/[۰-۹]/g, function (d) { return '۰۱۲۳۴۵۶۷۸۹'.indexOf(d); })
    .trim();
}
function hash_(salt, pw) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8));
}
function hiddenSheet_(name, headers) {
  const ss = SS_();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange('A:L').setNumberFormat('@');
    if (headers) s.getRange(1, 1, 1, headers.length).setValues([headers]);
    s.hideSheet();
  }
  return s;
}

/* ---------- users & sessions ---------- */
function users_() {
  const s = hiddenSheet_(USERS, ['username', 'name', 'role', 'salt', 'hash', 'active', 'perms', 'branch']);
  const n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, 8).getValues().map(function (r, i) {
    return { row: i + 2, username: String(r[0]), name: String(r[1]), role: String(r[2]), salt: String(r[3]), hash: String(r[4]), active: r[5] === true || String(r[5]).toUpperCase() === 'TRUE', perms: String(r[6] || ''), branch: String(r[7] || '') };
  }).filter(function (x) { return x.username; });
}
function pub_(x) { return { username: x.username, name: x.name, role: x.role, roleName: ROLE_NAMES[x.role] || x.role, active: x.active, perms: perms_(x), branch: x.branch || '', admin: isTop_(x), allbr: isAdmin_(x), top: isSuper_(x), full: isAdmin_(x) && String(x.perms) === 'FULL' }; }
function checkUser_(n) { if (!/^[A-Za-z0-9_.\-؀-ۿ]{3,30}$/.test(n)) throw new Error('اسم المستخدم من 3 إلى 30 حرفاً بدون مسافات'); }
function checkPw_(pw) { if (pw.length < 4) throw new Error('كلمة المرور 4 أحرف أو أرقام على الأقل'); }

function setup_(p) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    if (users_().length) throw new Error('تم إنشاء حساب المحاسب مسبقاً، سجّل الدخول');
    const un = norm_(p.username).toLowerCase(), pw = norm_(p.password);
    checkUser_(un); checkPw_(pw);
    const salt = Utilities.getUuid();
    hiddenSheet_(USERS).appendRow([un, String(p.name || '').trim().slice(0, 60) || un, 'acc', salt, hash_(salt, pw), true, '', '*']);
    const store = String(p.store || '').trim().slice(0, 120), branch = String(p.branch || '').trim().slice(0, 120);
    if (store || branch) { const st = getSettings_(); if (store) st.store = store; if (branch) st.branch = branch; hiddenSheet_(SETTINGS).getRange(1, 1).setValue(JSON.stringify(st)); }
  } finally { lock.releaseLock(); }
  return login_(p);
}
function login_(p) {
  const un = norm_(p.username).toLowerCase(), pw = norm_(p.password);
  const cache = CacheService.getScriptCache(), key = 'f_' + un;
  const fails = Number(cache.get(key) || 0);
  if (fails >= 8) throw new Error('محاولات كثيرة، حاول بعد 10 دقائق');
  const u = users_().find(function (x) { return x.username.toLowerCase() === un; });
  if (!u || !u.active || hash_(u.salt, pw) !== u.hash) { cache.put(key, String(fails + 1), 600); throw new Error('اسم المستخدم أو كلمة المرور غير صحيحة'); }
  cache.remove(key);
  BR = pickBranch_(u, p.br);
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  const props = PropertiesService.getScriptProperties();
  cleanSessions_(props);
  props.setProperty('s_' + token, JSON.stringify({ u: u.username, h: u.hash.slice(0, 16), exp: Date.now() + SESSION_MS }));
  return { token: token, boot: boot_(u) };
}
/* keep the session store small and clean: drop expired sessions, sessions of deleted users or changed
   passwords, and keep at most 15 devices per user (script properties are limited to 500 KB) */
function cleanSessions_(props) {
  const all = props.getProperties(), now = Date.now(), hs = {}, per = {};
  users_().forEach(function (x) { hs[x.username] = x.hash.slice(0, 16); });
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('s_') !== 0) return;
    let o = null; try { o = JSON.parse(all[k]); } catch (e) { }
    if (!o || o.exp < now || hs[o.u] !== o.h) { props.deleteProperty(k); return; }
    (per[o.u] || (per[o.u] = [])).push({ k: k, exp: o.exp });
  });
  Object.keys(per).forEach(function (u) { per[u].sort(function (a, b) { return b.exp - a.exp; }).slice(14).forEach(function (x) { props.deleteProperty(x.k); }); });
}
function session_(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache(), uv = uver_(), hit = cache.get('u_' + token);
  if (hit) { try { const o = JSON.parse(hit); if (o.uv === uv) return o.u; } catch (e) { } }
  const u = sessionLookup_(token);
  if (u) cache.put('u_' + token, JSON.stringify({ uv: uv, u: u }), 60);
  return u;
}
function uver_() {
  const c = CacheService.getScriptCache();
  let v = c.get('uv');
  if (!v) { v = PropertiesService.getScriptProperties().getProperty('uv') || '1'; c.put('uv', v, 21600); }
  return v;
}
function bumpUv_() {
  const old = Number(PropertiesService.getScriptProperties().getProperty('uv')) || 0;
  const v = String(Math.max(Date.now(), old + 1));
  PropertiesService.getScriptProperties().setProperty('uv', v);
  CacheService.getScriptCache().put('uv', v, 21600);
}
function sessionLookup_(token) {
  const raw = PropertiesService.getScriptProperties().getProperty('s_' + token);
  if (!raw) return null;
  let s; try { s = JSON.parse(raw); } catch (e) { return null; }
  if (s.exp < Date.now()) return null;
  const u = users_().find(function (x) { return x.username === s.u; });
  if (!u || !u.active || u.hash.slice(0, 16) !== s.h) return null;
  if (s.exp - Date.now() < SESSION_MS - 7 * 24 * 3600 * 1000) {
    s.exp = Date.now() + SESSION_MS;
    PropertiesService.getScriptProperties().setProperty('s_' + token, JSON.stringify(s));
  }
  return u;
}
function saveUser_(me, p) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const un = norm_(p.username).toLowerCase();
    const role = ['acc', 'mgr', 'sup', 'lead', 'own'].indexOf(p.role) >= 0 ? p.role : '';
    if (!role) throw new Error('اختر الصلاحية');
    const name = String(p.name || '').trim().slice(0, 60) || un;
    const pw = norm_(p.password || ''), active = p.active !== false;
    const s = hiddenSheet_(USERS);
    const u = users_().find(function (x) { return x.username.toLowerCase() === un; });
    if (p.isNew) {
      checkUser_(un);
      if (u) throw new Error('اسم المستخدم مستخدم من قبل');
      checkPw_(pw);
      const salt = Utilities.getUuid();
      s.appendRow([un, name, role, salt, hash_(salt, pw), active, role === 'acc' ? fullIn_(me, null, p, role) : permsIn_(role, p.perms), branchIn_(role, p.branch)]);
    } else {
      if (!u) throw new Error('المستخدم غير موجود');
      if (isSuper_(u) && !isSuper_(me)) throw new Error('حساب مدير النظام يعدّله صاحبه فقط');
      if (u.username === me.username && (role !== 'acc' || !active || branchIn_(role, p.branch) !== '*')) throw new Error('لا يمكنك تغيير صلاحية حسابك الحالي أو إيقافه');
      s.getRange(u.row, 2, 1, 2).setValues([[name, role]]);
      s.getRange(u.row, 6, 1, 3).setValues([[active, role === 'acc' ? fullIn_(me, u, p, role) : permsIn_(role, p.perms), branchIn_(role, p.branch)]]);
      if (pw) { checkPw_(pw); const salt = Utilities.getUuid(); s.getRange(u.row, 4, 1, 2).setValues([[salt, hash_(salt, pw)]]); }
    }
    bumpUv_();
  } finally { lock.releaseLock(); }
  return users_().map(pub_);
}
/* acc may cover all branches ('*', the system admin) or one; the owner's branches live in his permissions; everyone else works in one branch */
function branchIn_(role, b) {
  const all = branchIds_(); b = String(b || '');
  if (role === 'own') return '*';
  if (role === 'acc' && (b === '*' || !b)) return '*';
  if (all.indexOf(b) < 0) { if (all.length === 1) return 'main'; throw new Error('اختر الفرع'); }
  return b;
}
/* only the system manager grants or removes full permissions; anyone else leaves them as they were */
function fullIn_(me, u, p, role) {
  if (branchIn_(role, p.branch) !== '*') return '';
  const was = !!u && String(u.perms) === 'FULL';
  return (isSuper_(me) ? p.full === true : was) ? 'FULL' : '';
}
function permsIn_(role, p) {
  if (role === 'acc' || !p || typeof p !== 'object') return '';
  const R = ROLE_PERMS[role], o = perms_({ role: role, perms: p });
  if (!R) return '';
  return JSON.stringify(o);
}
function deleteUser_(me, p) {
  const un = norm_(p.username).toLowerCase();
  if (un === me.username) throw new Error('لا يمكنك حذف حسابك الحالي');
  const u = users_().find(function (x) { return x.username.toLowerCase() === un; });
  if (u && isSuper_(u)) throw new Error('لا يمكن حذف حساب مدير النظام');
  if (u) { hiddenSheet_(USERS).deleteRow(u.row); bumpUv_(); }
  return users_().map(pub_);
}

/* ---------- settings ---------- */
function getSettings_() {
  const v = hiddenSheet_(SETTINGS).getRange(1, 1).getValue();
  let o = {}; try { o = v ? JSON.parse(v) : {}; } catch (e) { }
  if (o.signB === 'توقيع مراقب الكاشيرات') o.signB = 'توقيع مسؤول الكاشيرات';
  if (!o.sup12) { o.supHours = 12; }
  const r = Object.assign({}, DEFAULT_SETTINGS, o);
  if (BR) {
    const m = mainSettings_();
    COMPANY_KEYS.forEach(function (k) { r[k] = m[k]; });
    if (!o.branch) r.branch = branchName_(BR);
  }
  r.branches = r.branches || [];
  return r;
}
/* a branch keeps its own settings; the company name, logo, watermark and branch list live with the first branch */
function writeSettings_(st) {
  if (!BR) { hiddenSheet_(SETTINGS).getRange(1, 1).setValue(JSON.stringify(st)); return; }
  const own = Object.assign({}, st); COMPANY_KEYS.forEach(function (k) { delete own[k]; });
  hiddenSheet_(SETTINGS).getRange(1, 1).setValue(JSON.stringify(own));
  const id = BR;
  withBr_('', function () {
    const m = getSettings_();
    COMPANY_KEYS.forEach(function (k) { if (k !== 'branches' && st[k] !== undefined) m[k] = st[k]; });
    m.branches = (m.branches || []).map(function (b) { return b.id === id && st.branch ? { id: b.id, name: String(st.branch).slice(0, 60) } : b; });
    hiddenSheet_(SETTINGS).getRange(1, 1).setValue(JSON.stringify(m));
  });
}
function saveSettings_(p) {
  const s = getSettings_();
  ['store', 'branch', 'signA', 'signB'].forEach(function (k) { if (p[k] != null) s[k] = String(p[k]).slice(0, 120); });
  s.rate = Number(p.rate) || 0;
  s.rateNet = !!p.rateNet;
  s.supPrev = !!p.supPrev;
  s.supHours = Math.max(1, Math.min(72, Number(p.supHours) || 12)); s.sup12 = true;
  s.boxStart = Math.max(1, Math.min(9999, parseInt(p.boxStart, 10) || 1));
  const list = function (a) { return (a || []).map(function (x) { return String(x).trim().slice(0, 60); }).filter(String).slice(0, 200); };
  s.cashiers = list(p.cashiers); s.shifts = list(p.shifts);
  if (p.emails && typeof p.emails === 'object') {
    const em = {};
    Object.keys(p.emails).slice(0, 300).forEach(function (k) { const v = String(p.emails[k] || '').trim().toLowerCase().slice(0, 120); if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) em[String(k).trim().slice(0, 60)] = v; });
    s.emails = em;
  }
  if (p.genders && typeof p.genders === 'object') {
    const gd = {};
    Object.keys(p.genders).slice(0, 300).forEach(function (k) { if (p.genders[k] === 'm') gd[String(k).trim().slice(0, 60)] = 'm'; });
    s.genders = gd;
  }
  if (p.phones && typeof p.phones === 'object') {
    const ph = {};
    Object.keys(p.phones).slice(0, 300).forEach(function (k) { const v = norm_(p.phones[k]).replace(/[^\d+]/g, '').slice(0, 16); if (v) ph[String(k).trim().slice(0, 60)] = v; });
    s.phones = ph;
  }
  if (p.staff && typeof p.staff === 'object') {
    const DAYW = ['السبت', 'الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة'], iso = /^\d{4}-\d{2}-\d{2}$/, mon = /^\d{4}-\d{2}$/;
    const out = {};
    Object.keys(p.staff).slice(0, 300).forEach(function (k) {
      const e = p.staff[k] || {}, name = String(k).trim().slice(0, 60); if (!name) return;
      const o = { start: iso.test(e.start || '') ? e.start : '', end: iso.test(e.end || '') ? e.end : '', shift: String(e.shift || '').trim().slice(0, 30), sched: {} };
      Object.keys(e.sched || {}).slice(0, 240).forEach(function (m) {
        if (!mon.test(m)) return; const v = e.sched[m] || {};
        o.sched[m] = { off: DAYW.indexOf(v.off) >= 0 ? v.off : '', type: v.type === 'بديل' ? 'بديل' : 'أساسي' };
      });
      out[name] = o;
    });
    s.staff = out;
  }
  if (Array.isArray(p.attHide)) s.attHide = p.attHide.map(function (x) { return String(x).trim().slice(0, 60); }).filter(String).slice(0, 300);
  if (Array.isArray(p.comHide)) s.comHide = p.comHide.map(function (x) { return String(x).trim().slice(0, 60); }).filter(String).slice(0, 300);
  if (Array.isArray(p.comBoxes)) s.comBoxes = p.comBoxes.map(Number).filter(function (x) { return x > 0 && x < 100000; }).slice(0, MAX_BOXES);
  if (Array.isArray(p.attBoxes)) s.attBoxes = p.attBoxes.map(Number).filter(function (x) { return x > 0 && x < 100000; }).slice(0, MAX_BOXES);
  if (Array.isArray(p.netsOff)) {
    const off = p.netsOff.filter(function (k) { return NETS.some(function (x) { return x[0] === k; }); });
    if (off.length >= NETS.length) throw new Error('لازم تبقى شبكة واحدة على الأقل');
    s.netsOff = off;
  }
  if (p.boxSlots != null) {
    const nb = Math.max(1, Math.min(MAX_BOXES, parseInt(p.boxSlots, 10) || DEF_BOXES)), was = nBoxes_(s);
    if (nb < was) {
      const used = readAll_().some(function (d) { return (d.boxes || []).slice(nb).some(function (b) { return b && hasData_(b); }); });
      if (used) throw new Error('لا يمكن حذف صندوق عليه بيانات سابقة، استخدم «إخفاء» بدلاً من الحذف');
    }
    s.boxSlots = nb;
  }
  if (Array.isArray(p.boxNos)) s.boxNos = p.boxNos.slice(0, MAX_BOXES).map(function (x) { const v = parseInt(x, 10); return v > 0 && v < 100000 ? v : ''; });
  if (Array.isArray(p.boxLabels)) s.boxLabels = p.boxLabels.slice(0, MAX_BOXES).map(function (x) { return String(x == null ? '' : x).trim().slice(0, 30); });
  if (Array.isArray(p.boxOff)) s.boxOff = p.boxOff.slice(0, MAX_BOXES).map(function (x) { return x === true; });
  if (p.logoPrint != null) s.logoPrint = !!p.logoPrint;
  if (p.logoWm != null) s.logoWm = !!p.logoWm;
  if (p.wmMode != null) s.wmMode = ['cover', 'tile', 'center'].indexOf(p.wmMode) >= 0 ? p.wmMode : 'cover';
  if (p.wmLevel != null) s.wmLevel = Math.max(0, Math.min(2, parseInt(p.wmLevel, 10) || 0));
  if (p.supApprove != null) s.supApprove = !!p.supApprove;
  if (p.allowFuture != null) s.allowFuture = !!p.allowFuture;
  if (p.modCashbox != null) s.modCashbox = !!p.modCashbox;
  if (p.modPetty != null) s.modPetty = !!p.modPetty;
  if (p.fx && typeof p.fx === 'object') {
    const cur = {};
    FX.forEach(function (x) { const c = (p.fx.cur || {})[x[0]] || {}, r = Math.round((Number(c.rate) || 0) * 10000) / 10000; cur[x[0]] = { on: !!c.on && r > 0, rate: r > 0 && r < 1000 ? r : x[2] }; });
    s.fx = { on: !!p.fx.on, mode: p.fx.mode === 'track' ? 'track' : 'convert', cur: cur };
  }
  const seen = {};
  for (let i = 0; i < nBoxes_(s); i++) { const v = bn_(s, i); if (seen[v]) throw new Error('رقم الصندوق ' + v + ' مكرر، كل صندوق لازم يكون له رقم مختلف'); seen[v] = 1; }
  if (p.boxNames) s.boxNames = (p.boxNames || []).slice(0, MAX_BOXES).map(function (x) { return String(x == null ? '' : x).trim().slice(0, 60); });
  writeSettings_(s);
  bump_();
  return s;
}

/* ---------- days ---------- */
function dataSheet_() { return hiddenSheet_(DATA, ['date', 'json', 'savedAt', 'savedBy']); }
function readAll_() {
  const s = dataSheet_(), n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, 4).getDisplayValues().map(function (r, i) {
    try { const d = JSON.parse(dayRaw_(s, i + 2, r[1])); d.date = r[0]; d.savedAt = r[2]; d.savedBy = r[3]; return d; } catch (e) { return null; }
  }).filter(function (d) { return d && d.date; });
}
/* read only the rows that are needed: the big JSON column is fetched just for picked rows */
function readDays_(pick) {
  const s = dataSheet_(), n = s.getLastRow();
  if (n < 2) return { days: [], dates: [] };
  const A = s.getRange(2, 1, n - 1, 1).getDisplayValues(), CD = s.getRange(2, 3, n - 1, 2).getDisplayValues();
  const dates = A.map(function (r) { return r[0]; }), idx = [];
  for (let i = 0; i < dates.length; i++) if (dates[i] && pick(dates[i], CD[i][0], dates)) idx.push(i);
  let B = null;
  if (idx.length > 20) B = s.getRange(2, 2, n - 1, 1).getDisplayValues();
  const days = idx.map(function (i) {
    const raw = dayRaw_(s, i + 2, B ? B[i][0] : null);
    try { const d = JSON.parse(raw); d.date = dates[i]; d.savedAt = CD[i][0]; d.savedBy = CD[i][1]; return d; } catch (e) { return null; }
  }).filter(function (d) { return d; });
  return { days: days, dates: dates.filter(String) };
}
/* a day's JSON can outgrow one cell (50,000 characters): it continues in columns E to L */
const CHUNK = 45000;
function dayRaw_(s, row, first) {
  let raw = first != null ? first : s.getRange(row, 2).getDisplayValue();
  if (raw.length >= CHUNK) raw += s.getRange(row, 5, 1, 8).getDisplayValues()[0].join('');
  return raw;
}
function dayWrite_(s, row, date, json, savedAt, by) {
  const parts = [];
  for (let i = 0; i < json.length; i += CHUNK) parts.push(json.slice(i, i + CHUNK));
  if (!parts.length) parts.push('');
  if (parts.length > 9) throw new Error('بيانات هذا اليوم أكبر من المسموح، قلّل الملاحظات الطويلة');
  ensureRows_(s, row);
  s.getRange(row, 1, 1, 4).setValues([[date, parts[0], savedAt, by]]);
  if (parts.length > 1 || s.getLastColumn() > 4) {
    if (s.getMaxColumns() < 12) s.insertColumnsAfter(s.getMaxColumns(), 12 - s.getMaxColumns());
    const ex = parts.slice(1); while (ex.length < 8) ex.push('');
    s.getRange(row, 5, 1, 8).setNumberFormat('@').setValues([ex]);
  }
}
function findRow_(s, date) {
  const n = s.getLastRow(); if (n < 2) return 0;
  const v = s.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === date) return i + 2;
  return 0;
}
function boot_(u, p) {
  const st = getSettings_(), t = today_(), since = Number(p && p.since) || 0;
  let days, extra = {};
  if (u.role === 'own' && perms_(u).totals) {
    days = [];
  } else if (u.role === 'acc' || u.role === 'own') {
    // an accountant who already has a copy on the phone gets only what changed since then
    const R = readDays_(function (d, sv) { return !since || Date.parse(sv) > since; });
    days = R.days; if (since) { extra.delta = true; extra.dates = R.dates; }
  } else if (u.role === 'sup') {
    const okDates = [t]; if (st.supPrev) okDates.push(addDays_(t, -1));
    let pd = '';
    const R = readDays_(function (d, sv, all) {
      if (!pd) pd = all.filter(function (x) { return x && x < t && okDates.indexOf(x) < 0; }).sort().pop() || '#';
      return okDates.indexOf(d) >= 0 || d === pd;
    });
    const prev = R.days.filter(function (d) { return okDates.indexOf(d.date) < 0; })[0];
    days = R.days.filter(function (d) { return okDates.indexOf(d.date) >= 0; });
    if (prev) days.push({ date: prev.date, namesOnly: true, returns: [], boxes: (prev.boxes || []).map(function (b) { return { cashier: b.cashier, shift: b.shift }; }) });
    const W = supWindow_(u, t, okDates);
    if (W.length) days = days.filter(function (d) { return !(d.namesOnly && W.some(function (w) { return w.date === d.date; })); }).concat(W);
  }
  else {
    const wk = windowDates_(t, perms_(u).days);
    days = readDays_(function (d) { return wk.indexOf(d) >= 0; }).days;
    if (u.role === 'lead') days = leadDays_(days);
  }
  return Object.assign(extra, { rev: rev_(), user: pub_(u), settings: settingsFor_(u, st), today: t, now: new Date().toISOString(), days: days, appUrl: canSend_(u) ? appUrl_() : '', accDays: ACC_EDIT_DAYS, leadDays: perms_(u).days || LEAD_DAYS, sent: sentFor_(u, t), sheetUrl: u.role === 'acc' ? SS_().getUrl() : null,
    br: BR || 'main', branches: userBranches_(u).map(function (id) { return { id: id, name: branchName_(id) }; }), allBranches: isAdmin_(u) ? branchIds_().map(function (id) { return { id: id, name: branchName_(id) }; }) : null });
}
function rev_() {
  const c = CacheService.getScriptCache();
  let r = c.get('rev');
  if (!r) { r = PropertiesService.getScriptProperties().getProperty('rev') || '1'; c.put('rev', r, 21600); }
  return Number(r);
}
function bump_() {
  const r = String(Date.now());
  PropertiesService.getScriptProperties().setProperty('rev', r);
  CacheService.getScriptCache().put('rev', r, 21600);
  return Number(r);
}
function tomb_(date, add) {
  const props = PropertiesService.getScriptProperties();
  let del = {}; try { del = JSON.parse(props.getProperty(delKey_()) || '{}'); } catch (e) { }
  const limit = Date.now() - 40 * 24 * 3600 * 1000;
  Object.keys(del).forEach(function (k) { if (del[k] < limit) delete del[k]; });
  if (add) del[date] = Date.now(); else delete del[date];
  props.setProperty(delKey_(), JSON.stringify(del));
}
function poll_(u, p) {
  const since = Number(p.since) || 0, rev = rev_(), t = today_();
  const out = { rev: rev, user: pub_(u), today: t, now: new Date().toISOString() };
  if (rev <= since) return out;
  const st = getSettings_();
  out.settings = settingsFor_(u, st);
  let days = (u.role === 'own' && perms_(u).totals) ? [] : readDays_(function (d, sv) { return Date.parse(sv) > since; }).days;
  if (u.role === 'sup') {
    const okDates = [t]; if (st.supPrev) okDates.push(addDays_(t, -1));
    const wk = perms_(u).tabs.indexOf('short') >= 0 ? windowDates_(t, perms_(u).days) : [];
    days = days.filter(function (d) { return okDates.indexOf(d.date) >= 0 || wk.indexOf(d.date) >= 0; })
      .map(function (d) { return okDates.indexOf(d.date) >= 0 ? d : supShape_(d); });
  }
  if (u.role === 'lead' || u.role === 'mgr') {
    const wk = windowDates_(t, perms_(u).days);
    days = days.filter(function (d) { return wk.indexOf(d.date) >= 0; });
    if (u.role === 'lead') days = leadDays_(days);
  }
  out.days = days;
  out.sent = sentFor_(u, t);
  let del = {}; try { del = JSON.parse(PropertiesService.getScriptProperties().getProperty(delKey_()) || '{}'); } catch (e) { }
  out.deleted = Object.keys(del).filter(function (k) { return del[k] > since; });
  return out;
}
function hasData_(b) { return !!(b.cashier || b.cash || b.report || NETS.some(function (n) { return b[n[0]]; })); }

function blankBox_(i) {
  const o = { no: i + 1, cashier: '', shift: '', cash: 0, report: 0, notes: '', firstAt: '', by: '' };
  NETS.forEach(function (x) { o[x[0]] = 0; });
  return o;
}
/* accountant approves or rejects the supervisor's requests (boxes and return lines) */
function approve_(u, p) {
  const date = String(p.date || ''), rej = !!p.reject, who = u.name || u.username;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('تاريخ غير صالح');
  if (!periodOpen_(date)) throw new Error(CLOSED_MSG);
  const st = getSettings_(), nowISO = new Date().toISOString();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let doc;
  try {
    const s = dataSheet_(), row = findRow_(s, date);
    if (!row) throw new Error('لا توجد بيانات لهذا اليوم');
    const d = JSON.parse(dayRaw_(s, row)), boxes = d.boxes || [], returns = d.returns || [], audit = [];
    const allB = p.boxes === 'all', allR = p.returns === 'all' || allB;
    const pickB = allB ? null : (p.boxes || []).map(Number), pickR = allR ? null : (p.returns || []).map(String);
    let n = 0;
    boxes.forEach(function (b, i) {
      if (!b || !b.rq || (pickB && pickB.indexOf(i) < 0)) return;
      const rq = b.rq, nb = Object.assign({}, b);
      delete nb.rq;
      if (!rej) {
        const ch = diffBox_(b, rq);
        ['cashier', 'shift', 'cash', 'report', 'notes'].concat(NETS.map(function (x) { return x[0]; })).forEach(function (k) { nb[k] = rq[k]; });
        ['cashSar', 'fxd', 'fxr'].forEach(function (k) { if (rq[k] !== undefined) nb[k] = rq[k]; else delete nb[k]; });
        nb.apAt = nowISO; nb.apBy = who;
        if (hasData_(b)) { nb.edits = (parseInt(b.edits, 10) || 0) + 1; nb.editedBy = rq.by; nb.editedAt = rq.at; }
        audit.push([date, bn_(st, i), rq.cashier || b.cashier || '', 'اعتماد إدخال ' + rq.by + (ch && hasData_(b) ? ': ' + ch : ''), who]);
      } else audit.push([date, bn_(st, i), rq.cashier || b.cashier || '', 'رفض إدخال ' + rq.by, who]);
      boxes[i] = nb; n++;
    });
    for (let k = returns.length - 1; k >= 0; k--) {
      const r = returns[k]; if (!r.rq || (pickR && pickR.indexOf(r.id) < 0)) continue;
      audit.push([date, 'مرتجعات', '', (rej ? 'رفض مرتجع ' : 'اعتماد مرتجع ') + r.amount + (r.note ? ' (' + r.note + ')' : '') + ' — ' + r.rq.by, who]);
      if (rej) returns.splice(k, 1); else { delete r.rq; r.apBy = who; }
      n++;
    }
    if (!n) throw new Error('لا توجد طلبات بانتظار الاعتماد');
    doc = { date: date, boxes: boxes, returns: returns, savedAt: nowISO, savedBy: who };
    dayWrite_(s, row, date, JSON.stringify({ boxes: boxes, returns: returns }), nowISO, who);
    writeDaySheet_(doc, st); writeLog_(doc); appendAudit_(audit); bump_();
    doc.done = n;
  } finally { lock.releaseLock(); }
  return doc;
}
function saveDay_(u, p) {
  const date = String(p.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('تاريخ غير صالح');
  const st = getSettings_(), t = today_();
  const sup = u.role === 'sup', who = u.name || u.username;
  const pa = sup ? perms_(u).approve : undefined, needAp = sup && (typeof pa === 'boolean' ? pa : st.supApprove !== false);
  if (sup && !(date === t || (st.supPrev && date === addDays_(t, -1)))) throw new Error('مسؤول الكاشيرات يحفظ يوم اليوم فقط');
  if (date > t && !st.allowFuture) throw new Error(FUTURE_MSG);
  if (!periodOpen_(date)) throw new Error(CLOSED_MSG);
  const n = function (v) { return Math.round((Number(v) || 0) * 100) / 100; };
  const tx = function (v, l) { return String(v == null ? '' : v).slice(0, l || 120); };
  const now = Date.now(), nowISO = new Date(now).toISOString(), lockMs = (Number(st.supHours) || 12) * 3600 * 1000;
  const expired = function (iso) { return iso && now - Date.parse(iso) > lockMs; };

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let doc, kept = 0;
  try {
    const s = dataSheet_(), row = findRow_(s, date);
    let prev = null;
    if (row) { try { prev = JSON.parse(dayRaw_(s, row)); } catch (e) { prev = null; } }
    const boxes = [], NB = Math.max(nBoxes_(st), (prev && prev.boxes ? prev.boxes.length : 0));
    for (let i = 0; i < NB; i++) boxes.push(Object.assign(blankBox_(i), (prev && prev.boxes && prev.boxes[i]) || {}, { no: i + 1 }));
    const returns = ((prev && prev.returns) || []).slice();

    const inB = p.boxes || {}, audit = [];
    Object.keys(inB).forEach(function (k) {
      const i = parseInt(k, 10); if (!(i >= 0 && i < NB)) return;
      const b = inB[k] || {}, pb = boxes[i];
      if (sup && expired(pb.firstAt)) { kept++; return; }
      const o = { no: i + 1, cashier: tx(b.cashier, 60), shift: tx(b.shift, 30), cash: n(b.cash), report: n(b.report), notes: tx(b.notes, 500) };
      fxIn_(o, b, pb, st);
      NETS.forEach(function (x) { o[x[0]] = n(b[x[0]]); });
      if (needAp) {
        // the supervisor's entry waits for the accountant's approval; reports keep the approved figures
        const same = !diffBox_(pb, o);
        if (same) { delete pb.rq; }
        else { pb.rq = Object.assign({}, o, { at: nowISO, by: who }); delete pb.rq.no; }
        if (!pb.firstAt && (hasData_(o) || pb.rq)) { pb.firstAt = nowISO; pb.by = who; }
        boxes[i] = pb;
        return;
      }
      if (pb.rq && !sup) o.rq = pb.rq;
      if (pb.apAt) { o.apAt = pb.apAt; o.apBy = pb.apBy; }
      const h = hasData_(o);
      o.firstAt = h ? (pb.firstAt || nowISO) : '';
      o.by = h ? (pb.by || who) : '';
      const ch = diffBox_(pb, o), changed = !!(ch && hasData_(pb));
      o.edits = h ? (parseInt(pb.edits, 10) || 0) + (changed ? 1 : 0) : 0;
      o.editedBy = h ? (changed ? who : (pb.editedBy || '')) : '';
      o.editedAt = h ? (changed ? nowISO : (pb.editedAt || '')) : '';
      if (changed) audit.push([date, bn_(st, i), pb.cashier || o.cashier || '', ch, who]);
      boxes[i] = o;
    });
    (p.retDelete || []).forEach(function (id) {
      const k = returns.findIndex(function (r) { return r.id === id; }); if (k < 0) return;
      if (sup && (expired(returns[k].firstAt) || (needAp && !returns[k].rq))) { kept++; return; }
      if (returns[k].rq) { returns.splice(k, 1); return; }
      audit.push([date, 'مرتجعات', '', 'حذف مرتجع ' + returns[k].amount + (returns[k].note ? ' (' + returns[k].note + ')' : ''), who]);
      returns.splice(k, 1);
    });
    (p.retUpsert || []).slice(0, 50).forEach(function (r) {
      const id = tx(r.id, 40); if (!id) return;
      const k = returns.findIndex(function (x) { return x.id === id; }), pr = k >= 0 ? returns[k] : null;
      if (sup && pr && expired(pr.firstAt)) { kept++; return; }
      const o = { id: id, amount: -Math.abs(n(r.amount)), box: Math.max(0, Math.min(NB, parseInt(r.box, 10) || 0)), note: tx(r.note, 120) };
      if (needAp) { if (pr && !pr.rq) { kept++; return; } o.rq = { at: nowISO, by: who }; }
      if (!o.amount && !o.note) { if (k >= 0) returns.splice(k, 1); return; }
      o.firstAt = (pr && pr.firstAt) || nowISO;
      if (!o.rq && pr && !pr.rq && (pr.amount !== o.amount || pr.box !== o.box || (pr.note || '') !== o.note)) {
        const bl = function (x) { return x.box ? bn_(st, x.box - 1) : '—'; };
        audit.push([date, 'مرتجعات', '', 'مرتجع: ' + pr.amount + ' ← ' + o.amount + (pr.box !== o.box ? ' | الصندوق: ' + bl(pr) + ' ← ' + bl(o) : '') + ((pr.note || '') !== o.note ? ' | البيان: ' + (pr.note || '—') + ' ← ' + (o.note || '—') : ''), who]);
      }
      if (k >= 0) returns[k] = o; else returns.push(o);
    });

    doc = { date: date, boxes: boxes, returns: returns, savedAt: nowISO, savedBy: who };
    const vals = [[date, JSON.stringify({ boxes: boxes, returns: returns }), doc.savedAt, doc.savedBy]];
    dayWrite_(s, row || s.getLastRow() + 1, date, vals[0][1], doc.savedAt, doc.savedBy);
    writeDaySheet_(doc, st);
    writeLog_(doc);
    if (audit.length) appendAudit_(audit);
    tomb_(date, false);
    bump_();
  } finally { lock.releaseLock(); }
  try { cleanupDaySheets_(); } catch (e) { }
  doc.kept = kept;
  return doc;
}
function deleteDay_(date) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = dataSheet_(), row = findRow_(s, date);
    if (row) s.deleteRow(row);
    const ss = SS_(), sh = ss.getSheetByName(date);
    if (sh) ss.deleteSheet(sh);
    deleteLogRows_(date);
    tomb_(date, true);
    bump_();
  } finally { lock.releaseLock(); }
}

/* ---------- readable sheets ---------- */
function approvedOnly_(doc) { return Object.assign({}, doc, { returns: (doc.returns || []).filter(function (r) { return !r.rq; }) }); }
function calc_(doc) {
  const rows = [], T = { cash: 0, ret: 0, net: 0, report: 0, diff: 0, nets: {} };
  NETS.forEach(function (x) { T.nets[x[0]] = 0; });
  doc.boxes.forEach(function (b, i) {
    if (!hasData_(b)) return;
    const net = NETS.reduce(function (s, x) { return s + (Number(b[x[0]]) || 0); }, 0);
    const ret = doc.returns.filter(function (r) { return r.box === i + 1; }).reduce(function (s, r) { return s + r.amount; }, 0);
    const total = b.cash + net, diff = Math.round((total - b.report) * 100) / 100;
    rows.push({ b: b, i: i, net: net, ret: ret, total: total, diff: diff });
    T.cash += b.cash; T.net += net; T.report += b.report;
    NETS.forEach(function (x) { T.nets[x[0]] += Number(b[x[0]]) || 0; });
  });
  T.ret = doc.returns.reduce(function (s, r) { return s + r.amount; }, 0);
  return { rows: rows, T: T };
}
function writeDaySheet_(doc, st) {
  doc = approvedOnly_(doc);
  const ss = SS_();
  let sh = ss.getSheetByName(doc.date);
  if (!sh) sh = ss.insertSheet(doc.date, ss.getSheets().length); else sh.clear();
  sh.setRightToLeft(true);
  const c = calc_(doc), rows = c.rows, T = c.T, NV = netsShown_(st, doc.boxes);
  const H = ['الصندوق', 'أمين الصندوق', 'الوردية', 'الكاش', 'مرتجعات مصروفة', 'صافي الكاش']
    .concat(NV.map(function (x) { return x[1]; }))
    .concat(['إجمالي الشبكات', 'الإجمالي الكلي', 'تقرير النظام', 'الزيادة / النقص', 'ملاحظات']);
  const W = H.length, out = [];
  const pad = function (a) { while (a.length < W) a.push(''); return a; };
  out.push(pad([st.store + ' — ' + st.branch]));
  out.push(pad(['تقرير إقفال الصناديق — ' + dayName_(doc.date) + ' ' + doc.date.split('-').reverse().join('/')]));
  out.push(pad([]));
  out.push(H);
  rows.forEach(function (r) {
    const b = r.b;
    out.push([bn_(st, r.i), b.cashier, b.shift, b.cash, r.ret, b.cash + r.ret]
      .concat(NV.map(function (x) { return b[x[0]]; }))
      .concat([r.net, r.total, b.report, r.diff, b.notes]));
  });
  const totRow = out.length + 1;
  out.push(['الإجمالي', '', '', T.cash, T.ret, T.cash + T.ret]
    .concat(NV.map(function (x) { return T.nets[x[0]]; }))
    .concat([T.net, T.cash + T.net, T.report, Math.round((T.cash + T.net - T.report) * 100) / 100, '']));
  out.push(pad([]));
  const cashNet = T.cash + T.ret, repNet = T.report + T.ret, tot = cashNet + T.net, diff = Math.round((tot - repNet) * 100) / 100;
  const sumStart = out.length + 1;
  [['إجمالي الكاش', T.cash], ['المرتجعات', T.ret], ['صافي الكاش', cashNet], ['إجمالي الشبكات', T.net], ['الإجمالي الكلي', tot],
   ['إجمالي تقارير النظام', T.report], ['صافي تقارير النظام', repNet], ['الزيادة / النقص', diff]]
    .forEach(function (r) { out.push(pad([r[0], r[1]])); });
  const sumEnd = out.length;
  if (doc.returns.length) {
    out.push(pad([])); out.push(pad(['المرتجعات', 'من صندوق', 'القيمة', 'ملاحظة']));
    doc.returns.forEach(function (r) { out.push(pad(['', r.box ? bn_(st, r.box - 1) : '', r.amount, r.note])); });
  }
  out.push(pad([]));
  out.push(pad(['آخر حفظ: ' + doc.savedBy + ' — ' + Utilities.formatDate(new Date(doc.savedAt), TZ, 'yyyy/MM/dd HH:mm')]));
  ensureRows_(sh, out.length);
  if (sh.getMaxColumns() < W) sh.insertColumnsAfter(sh.getMaxColumns(), W - sh.getMaxColumns());
  sh.getRange(1, 1, out.length, W).setValues(out);
  sh.getRange(1, 1, 2, 1).setFontWeight('bold');
  sh.getRange(1, 1).setFontSize(13);
  sh.getRange(4, 1, 1, W).setFontWeight('bold').setBackground('#1d3a5f').setFontColor('#ffffff').setWrap(true);
  sh.getRange(totRow, 1, 1, W).setFontWeight('bold').setBackground('#e3eaf3');
  sh.getRange(5, 4, totRow - 4, W - 4).setNumberFormat('#,##0.00');
  sh.getRange(sumStart, 2, sumEnd - sumStart + 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(sumStart, 1, sumEnd - sumStart + 1, 2).setBorder(true, true, true, true, true, true);
  sh.setFrozenRows(4);
  sh.setColumnWidth(2, 140);
  if (sh.getMaxRows() > out.length + 2) sh.deleteRows(out.length + 3, sh.getMaxRows() - out.length - 2);
  if (sh.getMaxColumns() > W) sh.deleteColumns(W + 1, sh.getMaxColumns() - W);
}
function logSheet_() {
  const ss = SS_();
  let s = ss.getSheetByName(LOG);
  if (!s) {
    s = ss.insertSheet(LOG, 0);
    const H = ['التاريخ', 'اليوم', 'الصندوق', 'أمين الصندوق', 'الوردية', 'الكاش', 'مرتجعات مصروفة', 'صافي الكاش']
      .concat(NETS.map(function (x) { return x[1]; }))
      .concat(['إجمالي الشبكات', 'الإجمالي الكلي', 'تقرير النظام', 'الزيادة / النقص', 'ملاحظات', 'أدخله']);
    s.getRange('A:A').setNumberFormat('@');
    s.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#1d3a5f').setFontColor('#ffffff');
    s.setFrozenRows(1); s.setRightToLeft(true);
  }
  return s;
}
function deleteLogRows_(date) {
  const s = logSheet_(), n = s.getLastRow();
  if (n < 2) return;
  const v = s.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] === date) s.deleteRow(i + 2);
}
function logRows_(doc, st) {
  const rows = calc_(doc).rows, dn = dayName_(doc.date);
  const out = rows.map(function (r) {
    const b = r.b;
    return [doc.date, dn, bn_(st, r.i), b.cashier, b.shift, b.cash, r.ret, b.cash + r.ret]
      .concat(NETS.map(function (x) { return b[x[0]]; }))
      .concat([r.net, r.total, b.report, r.diff, b.notes, b.by || doc.savedBy]);
  });
  doc.returns.forEach(function (r) {
    out.push([doc.date, dn, 'مرتجعات', 'من صندوق ' + (r.box ? bn_(st, r.box - 1) : ''), '', 0, r.amount, r.amount]
      .concat(NETS.map(function () { return 0; }))
      .concat([0, r.amount, r.amount, 0, r.note, doc.savedBy]));
  });
  return out;
}
function appendLog_(out) {
  if (!out.length) return;
  const s = logSheet_(), start = s.getLastRow() + 1;
  ensureRows_(s, start + out.length - 1);
  s.getRange(start, 1, out.length, out[0].length).setValues(out);
  s.getRange(start, 6, out.length, 12).setNumberFormat('#,##0.00');
}
function writeLog_(doc) {
  deleteLogRows_(doc.date);
  appendLog_(logRows_(approvedOnly_(doc), getSettings_()));
}
function bn_(st, i) { const v = st.boxNos && parseInt(st.boxNos[i], 10); return v > 0 ? v : (Number(st.boxStart) || 1) + i; }
function ensureRows_(sh, need) {
  const max = sh.getMaxRows();
  if (max < need) sh.insertRowsAfter(max, Math.max(need - max, 200));
}

/* ---------- import ---------- */
function importDays_(u, p) {
  const st = getSettings_(), replace = p.mode === 'replace', makeSheets = !!p.daySheets, who = u.name || u.username;
  const n = function (v) { return Math.round((Number(v) || 0) * 100) / 100; };
  const tx = function (v, l) { return String(v == null ? '' : v).slice(0, l || 120); };
  const list = (p.days || []).slice(0, 62), res = { added: 0, replaced: 0, skipped: 0 };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const s = dataSheet_(), last = s.getLastRow(), idx = {};
    if (last >= 2) s.getRange(2, 1, last - 1, 1).getDisplayValues().forEach(function (r, i) { idx[r[0]] = i + 2; });
    const newRows = [], logOut = [], dates = [], nowISO = new Date().toISOString();
    list.forEach(function (d) {
      const date = String(d.date || '');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
      if (date > today_() && !st.allowFuture) { res.skipped++; res.future = (res.future || 0) + 1; return; }
      const row = idx[date];
      if (row && (!replace || !periodOpen_(date))) { res.skipped++; return; }
      const boxes = [];
      for (let i = 0; i < nBoxes_(st); i++) {
        const b = (d.boxes || [])[i] || {}, o = blankBox_(i);
        o.cashier = tx(b.cashier, 60); o.shift = tx(b.shift, 30); o.cash = n(b.cash); o.report = n(b.report); o.notes = tx(b.notes, 500);
        NETS.forEach(function (x) { o[x[0]] = n(b[x[0]]); });
        if (hasData_(o)) { o.firstAt = nowISO; o.by = who; }
        boxes.push(o);
      }
      const returns = (d.returns || []).slice(0, 50).map(function (r, k) {
        return { id: 'i' + date.replace(/-/g, '') + k, amount: -Math.abs(n(r.amount)), box: Math.max(0, Math.min(nBoxes_(st), parseInt(r.box, 10) || 0)), note: tx(r.note, 120), firstAt: nowISO };
      }).filter(function (r) { return r.amount || r.note; });
      const doc = { date: date, boxes: boxes, returns: returns, savedAt: nowISO, savedBy: who };
      const vals = [date, JSON.stringify({ boxes: boxes, returns: returns }), nowISO, who];
      if (row) { s.getRange(row, 1, 1, 4).setValues([vals]); deleteLogRows_(date); res.replaced++; }
      else { newRows.push(vals); idx[date] = -1; res.added++; }
      Array.prototype.push.apply(logOut, logRows_(doc, st));
      if (makeSheets) writeDaySheet_(doc, st);
      dates.push(date);
    });
    if (newRows.length) {
      const start = s.getLastRow() + 1;
      ensureRows_(s, start + newRows.length - 1);
      s.getRange(start, 1, newRows.length, 4).setValues(newRows);
    }
    appendLog_(logOut);
    if (dates.length) {
      const props = PropertiesService.getScriptProperties();
      let del = {}; try { del = JSON.parse(props.getProperty(delKey_()) || '{}'); } catch (e) { }
      dates.forEach(function (d) { delete del[d]; });
      props.setProperty(delKey_(), JSON.stringify(del));
      bump_();
    }
  } finally { lock.releaseLock(); }
  return res;
}

/* ---------- shortage notices (cashier lead) ---------- */
function sentSheet_() {
  const ss = SS_();
  let s = ss.getSheetByName(SENT);
  if (!s) {
    s = ss.insertSheet(SENT);
    s.getRange('A:A').setNumberFormat('@');
    s.getRange(1, 1, 1, 10).setValues([['التاريخ', 'اليوم', 'الصندوق', 'الموظفة', 'الجوال', 'العجز / الزيادة', 'أُرسل بواسطة', 'وقت الإرسال', 'رقم داخلي', 'الطابع الزمني']])
      .setFontWeight('bold').setBackground('#1d3a5f').setFontColor('#ffffff');
    s.setFrozenRows(1); s.setRightToLeft(true);
  }
  return s;
}
function sentFor_(u, t) {
  if (['lead', 'acc', 'mgr', 'own'].indexOf(u.role) < 0 && !(u.role === 'sup' && perms_(u).tabs.indexOf('short') >= 0)) return [];
  const s = sentSheet_(), n = s.getLastRow(), from = (u.role === 'acc' || u.role === 'own') ? addDays_(t, -60) : addDays_(t, -(perms_(u).days - 1));
  if (n < 2) return [];
  const w = Math.min(14, s.getMaxColumns());
  return s.getRange(2, 1, n - 1, w).getDisplayValues()
    .filter(function (r) { return r[0] >= from && r[0] <= t; })
    .map(function (r) { return { date: r[0], box: Number(r[8]), name: r[3], by: r[6], at: r[9], ch: r[11] || 'واتساب', ack: r[12] || '', seen: r[13] || '' }; });
}
function markSent_(u, p) {
  const t = today_(), date = String(p.date || ''), i = parseInt(p.box, 10);
  const pm = perms_(u);
  if (!pm.wa) throw new Error('ليست لديك صلاحية الإرسال');
  if (windowDates_(t, pm.days).indexOf(date) < 0) throw new Error('التقرير متاح لآخر ' + pm.days + ' أيام فقط');
  if (!(i >= 0 && i < MAX_BOXES)) throw new Error('رقم صندوق غير صالح');
  const ds = dataSheet_(), row = findRow_(ds, date);
  if (!row) throw new Error('لا توجد بيانات لهذا اليوم');
  const doc = JSON.parse(dayRaw_(ds, row)), b = doc.boxes[i] || {};
  const net = NETS.reduce(function (a, x) { return a + (Number(b[x[0]]) || 0); }, 0);
  const diff = Math.round(((Number(b.cash) || 0) + net - (Number(b.report) || 0)) * 100) / 100;
  const st = getSettings_(), phone = ((st.phones || {})[String(b.cashier || '').trim()]) || '';
  const tk = /^[A-Za-z0-9]{16,40}$/.test(String(p.token || '')) ? String(p.token) : Utilities.getUuid().replace(/-/g, '');
  logSent_(st, date, i, b, phone, diff, u.name || u.username, tk, 'واتساب');
  return sentFor_(u, t);
}
function logSent_(st, date, i, b, contact, diff, by, tk, ch) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
  const now = new Date(), s = sentSheet_(), nr = s.getLastRow() + 1;
  ensureRows_(s, nr);
  if (s.getMaxColumns() < 14) s.insertColumnsAfter(s.getMaxColumns(), 14 - s.getMaxColumns());
  if (!s.getRange(1, 14).getValue()) s.getRange(1, 11, 1, 4).setValues([['رمز الإشعار', 'طريقة الإرسال', 'أكدت الاطلاع', 'فتحت الإشعار']]).setFontWeight('bold').setBackground('#1d3a5f').setFontColor('#ffffff');
  s.getRange(nr, 1, 1, 14).setValues([[date, dayName_(date), bn_(st, i), b.cashier || '', contact, diff, by,
    Utilities.formatDate(now, TZ, 'yyyy/MM/dd HH:mm:ss'), i, now.toISOString(), tk, ch, '', '']]);
  bump_();
  } finally { lock.releaseLock(); }
}
/* ---------- the employee's notice: a branded page she opens from WhatsApp or e-mail ---------- */
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (ch) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]; }); }
function money_(v) { const n = Math.abs(Number(v) || 0); return n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function noticeFind_(tk) {
  if (!/^[A-Za-z0-9]{16,40}$/.test(tk)) return null;
  const ids = branchIds_();
  for (let k = 0; k < ids.length; k++) { BR = ids[k] === 'main' ? '' : ids[k]; const f = noticeFind1_(tk); if (f) return f; }
  BR = ''; return null;
}
function noticeFind1_(tk) {
  const s = sentSheet_(), n = s.getLastRow(); if (n < 2 || s.getMaxColumns() < 13) return null;
  const v = s.getRange(2, 1, n - 1, Math.min(14, s.getMaxColumns())).getDisplayValues();
  for (let k = v.length - 1; k >= 0; k--) if (v[k][10] === tk) return { row: k + 2, r: v[k] };
  return null;
}
function noticeData_(f) {
  const r = f.r, st = getSettings_(), ds = dataSheet_(), row = findRow_(ds, r[0]);
  let b = {}; if (row) { try { b = JSON.parse(dayRaw_(ds, row)).boxes[Number(r[8])] || {}; } catch (e) { } }
  const diff = Number(String(r[5]).replace(/,/g, '')) || 0;
  return { date: r[0], day: r[1], box: r[2], name: r[3], diff: diff, by: r[6], at: r[7], ch: r[11], ack: r[12], cash: Number(b.cash) || 0, report: Number(b.report) || 0, store: st.store, branch: st.branch };
}
/* staff are female by default; the accountant marks male staff in settings */
function male_(st, name) { return ((st || getSettings_()).genders || {})[String(name || '').trim()] === 'm'; }
function noticeHTML_(d, logo, forMail, link) {
  const M = male_(null, d.name), T = d.title || senderTitle_(d.by);
  const short = d.diff < 0, col = short ? '#c0392b' : '#b9770e', bg = short ? '#fdecea' : '#fef5e7';
  const title = short ? 'إشعار عجز في الصندوق' : 'إشعار زيادة في الصندوق';
  const date = d.date.split('-').reverse().join('/');
  const row = function (k, v) { return '<tr><td style="padding:9px 0;color:#5b6b7f;border-bottom:1px solid #edf1f6">' + k + '</td><td style="padding:9px 0;text-align:left;font-weight:600;color:#1d2b3a;border-bottom:1px solid #edf1f6;direction:ltr">' + v + '</td></tr>'; };
  return '<div style="max-width:440px;margin:0 auto;background:#ffffff;border-radius:18px;overflow:hidden;box-shadow:0 6px 28px rgba(29,58,95,.12);font-family:Tahoma,Arial,sans-serif;direction:rtl;text-align:right">' +
    '<div style="background:#1d3a5f;padding:22px 20px 18px;text-align:center">' +
    (logo ? '<img src="' + logo + '" alt="" style="max-height:64px;max-width:70%;background:#fff;border-radius:10px;padding:6px 10px;margin-bottom:10px">' : '') +
    '<div style="color:#fff;font-size:18px;font-weight:700">' + esc_(d.store) + '</div><div style="color:#c9d6e6;font-size:13px;margin-top:2px">' + esc_(d.branch) + '</div></div>' +
    '<div style="padding:22px 22px 8px;position:relative">' + (!forMail && wmSrc_() ? '<div style="position:absolute;inset:90px 0 40px;background:url(' + wmSrc_() + ') center/62% auto no-repeat;opacity:' + [.05, .08, .12][getSettings_().wmLevel || 0] + ';pointer-events:none"></div>' : '') + '<div style="position:relative"><div style="font-size:13px;color:#5b6b7f">' + (M ? 'الموظف الكريم' : 'الموظفة الكريمة') + '</div><div style="font-size:19px;font-weight:700;color:#1d2b3a;margin:2px 0 14px">' + esc_(d.name) + '</div>' +
    '<div style="background:' + bg + ';border:1px solid ' + col + '33;border-radius:14px;padding:16px;text-align:center"><div style="color:' + col + ';font-size:14px;font-weight:700">' + title + '</div>' +
    '<div style="color:' + col + ';font-size:32px;font-weight:800;margin-top:6px;direction:ltr">' + money_(d.diff) + ' <span style="font-size:15px">ريال</span></div></div>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:14px;font-size:14px">' + row('التاريخ', esc_(d.day) + ' ' + date) + row('رقم الصندوق', esc_(d.box)) +
    (d.cash ? row('الكاش', money_(d.cash)) : '') + '</table>' +
    '<p style="font-size:13.5px;line-height:1.9;color:#33475b;margin:16px 0 4px">نرجو الاطلاع على تفاصيل صندوقك أعلاه، ومراجعة ' + T + ' في حال وجود أي ملاحظة.</p>' +
    '</div>' + (forMail && link ? '<div style="text-align:center;margin:18px 0 6px"><a href="' + link + '" style="display:inline-block;background:#1d3a5f;color:#fff;text-decoration:none;padding:12px 26px;border-radius:12px;font-weight:700;font-size:14px">عرض الإشعار وتأكيد الاطلاع</a></div>' : '') +
    '</div><div style="padding:12px 22px 18px;border-top:1px solid #edf1f6;margin-top:10px;font-size:11px;color:#8a97a8;text-align:center">تم إرسال هذا الإشعار بواسطة ' + T + ': ' + esc_(d.by) + '<br>' + esc_(String(d.at).slice(0, 16)) + '</div></div>';
}
/* the notice opens only for its recipient: she proves it with the last 4 digits of her phone (or her e-mail).
   Opening or confirming from a device logged into the system (the supervisor, the accountant) never counts. */
function noticeKey_(f) {
  const r = f.r, st = getSettings_(), name = String(r[3] || '').trim();
  const ph = String((st.phones || {})[name] || (/\d{4}/.test(r[4]) ? r[4] : '')).replace(/\D/g, '');
  if (ph.length >= 4) return { kind: 'phone', val: ph.slice(-4) };
  const em = String((st.emails || {})[name] || (String(r[4]).indexOf('@') > 0 ? r[4] : '')).trim().toLowerCase();
  if (em) return { kind: 'email', val: em };
  return { kind: 'name', val: name.replace(/\s+/g, ' ') };
}
function noticeCheck_(f, code) {
  const k = noticeKey_(f), c0 = String(code || '').trim(), cache = CacheService.getScriptCache(), key = 'nv_' + f.r[10];
  const fails = Number(cache.get(key) || 0);
  if (fails >= 6) throw new Error(male_(null, f.r[3]) ? 'محاولات كثيرة، حاول بعد 15 دقيقة' : 'محاولات كثيرة، حاولي بعد 15 دقيقة');
  const ok = k.kind === 'phone' ? c0.replace(/[٠-٩]/g, function (d) { return '٠١٢٣٤٥٦٧٨٩'.indexOf(d); }).replace(/\D/g, '') === k.val
    : k.kind === 'email' ? c0.toLowerCase() === k.val : c0.replace(/\s+/g, ' ') === k.val;
  if (!ok) { cache.put(key, String(fails + 1), 900); throw new Error(k.kind === 'phone' ? 'آخر 4 أرقام غير صحيحة' : k.kind === 'email' ? 'البريد غير صحيح' : 'الاسم غير صحيح'); }
  cache.remove(key);
}
function noticeStaff_(sess) { try { return sess ? session_(String(sess)) : null; } catch (e) { return null; } }
function openNotice(tk, code, sess) {
  const f = noticeFind_(String(tk || '')); if (!f) throw new Error('الرابط غير صالح');
  const d = noticeData_(f), html = noticeHTML_(d, logo_().data, false), staff = noticeStaff_(sess);
  if (staff) return { preview: true, by: staff.name || staff.username, html: html, ack: d.ack, seen: f.r[13] || '' };
  noticeCheck_(f, code);
  let seen = f.r[13] || '';
  if (!seen) {
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      const s = sentSheet_(); if (s.getMaxColumns() < 14) s.insertColumnsAfter(s.getMaxColumns(), 14 - s.getMaxColumns());
      seen = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm'); s.getRange(f.row, 14).setValue(seen); bump_();
    } finally { lock.releaseLock(); }
  }
  return { html: html, ack: d.ack, seen: seen };
}
function ackNotice(tk, code, sess) {
  const f = noticeFind_(String(tk || '')); if (!f) throw new Error('الرابط غير صالح');
  if (noticeStaff_(sess)) throw new Error(male_(null, f.r[3]) ? 'تأكيد الاطلاع يكون من الموظف نفسه فقط' : 'تأكيد الاطلاع يكون من الموظفة نفسها فقط');
  noticeCheck_(f, code);
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const s = sentSheet_(); let v = f.r[12];
    if (!v) { v = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm'); s.getRange(f.row, 13).setValue(v); if (s.getMaxColumns() >= 14 && !f.r[13]) s.getRange(f.row, 14).setValue(v); bump_(); }
    return v;
  } finally { lock.releaseLock(); }
}
function noticePage_(tk) {
  const f = noticeFind_(tk), st = getSettings_(), logo = logo_().data;
  const wrap = function (body) {
    return HtmlService.createHtmlOutput('<!doctype html><html dir="rtl" lang="ar"><head><meta charset="utf-8"></head><body style="margin:0;background:#eef2f7;padding:22px 14px;font-family:Tahoma,Arial,sans-serif">' + body + '</body></html>')
      .setTitle('إشعار الصندوق').addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  };
  if (!f) return wrap('<div style="max-width:420px;margin:40px auto;background:#fff;border-radius:16px;padding:28px;text-align:center;color:#5b6b7f">الرابط غير صالح أو انتهى.</div>');
  const k = noticeKey_(f), name = esc_(f.r[3]), M = male_(st, f.r[3]), w = M ? 'اكتب' : 'اكتبي';
  const ask = k.kind === 'phone' ? 'للتحقق، ' + w + ' آخر 4 أرقام من رقم جوالك' : k.kind === 'email' ? 'للتحقق، ' + w + ' بريدك الإلكتروني' : 'للتحقق، ' + w + ' اسمك كما هو مسجّل';
  const inp = k.kind === 'phone' ? 'inputmode="numeric" maxlength="4" placeholder="••••" style="letter-spacing:10px;text-align:center;direction:ltr;' : k.kind === 'email' ? 'type="email" inputmode="email" dir="ltr" placeholder="name@example.com" style="' : 'placeholder="الاسم" style="';
  const S = 'width:100%;box-sizing:border-box;border:1px solid #d5dde8;border-radius:12px;padding:13px;font:600 17px Tahoma,Arial;margin:10px 0 12px;background:#f8fafc"';
  const B = 'width:100%;border:0;background:#1d3a5f;color:#fff;font:700 15px Tahoma,Arial;padding:14px;border-radius:12px;cursor:pointer';
  const gate = '<div id="gate" style="max-width:440px;margin:0 auto;background:#fff;border-radius:18px;overflow:hidden;box-shadow:0 6px 28px rgba(29,58,95,.12)">' +
    '<div style="background:#1d3a5f;padding:20px;text-align:center">' + (logo ? '<img src="' + logo + '" style="max-height:56px;max-width:70%;background:#fff;border-radius:10px;padding:6px 10px;margin-bottom:8px">' : '') +
    '<div style="color:#fff;font-size:17px;font-weight:700">' + esc_(st.store) + '</div><div style="color:#c9d6e6;font-size:13px">' + esc_(st.branch) + '</div></div>' +
    '<div style="padding:22px"><div style="font-size:13px;color:#5b6b7f">' + (M ? 'إشعار خاص بالموظف' : 'إشعار خاص بالموظفة') + '</div><div style="font-size:19px;font-weight:700;color:#1d2b3a;margin:2px 0 14px">' + name + '</div>' +
    '<div style="font-size:14px;color:#33475b">' + ask + '</div><input id="code" ' + inp + S + '>' +
    '<button id="go" style="' + B + '">عرض الإشعار</button><div id="err" style="color:#c0392b;font-size:13px;margin-top:10px;text-align:center"></div></div></div>';
  const js = '<script>(function(){var tk=' + JSON.stringify(tk) + ',code="",sess="";try{sess=JSON.parse(localStorage.getItem("ccd-token")||"\\"\\"")||""}catch(e){}' +
    'var $=function(i){return document.getElementById(i)},run=function(){return google.script.run};' +
    'function show(r,pre){$("gate").style.display="none";var w=$("view");w.innerHTML=r.html;var a=$("act");' +
    'if(pre){a.innerHTML=\'<div style="background:#fff7e6;border:1px solid #f0c36d;color:#8a5a00;border-radius:12px;padding:12px;font-size:13.5px;line-height:1.8">معاينة فقط: هذا الجهاز مسجّل في النظام باسم <b>\'+r.by+\'</b>. فتح الإشعار وتأكيد الاطلاع يُحسبان من ' + (M ? 'الموظف نفسه' : 'الموظفة نفسها') + ' فقط.</div>\';return}' +
    'if(r.ack){a.innerHTML=\'<div style="color:#1e8449;font-weight:700">✓ تم تأكيد اطلاعك بتاريخ \'+r.ack+\'</div>\';return}' +
    'a.innerHTML=\'<button id="ack" style="' + B + '">أقر بأنني اطلعت على هذا الإشعار</button><div id="ok" style="color:#1e8449;font-weight:700;margin-top:10px"></div>\';' +
    '$("ack").onclick=function(){var b=this;b.disabled=true;b.textContent="جارٍ التسجيل…";run().withSuccessHandler(function(t){b.style.display="none";$("ok").textContent="✓ شكراً لك، تم تسجيل اطلاعك "+t}).withFailureHandler(function(e){b.disabled=false;b.textContent="أقر بأنني اطلعت على هذا الإشعار";$("ok").style.color="#c0392b";$("ok").textContent=e&&e.message||"تعذّر التسجيل"}).ackNotice(tk,code,"")}}' +
    'if(sess)run().withSuccessHandler(function(r){if(r&&r.preview)show(r,true)}).withFailureHandler(function(){}).openNotice(tk,"",sess);' +
    '$("go").onclick=function(){var b=this;code=$("code").value;$("err").textContent="";b.disabled=true;b.textContent="جارٍ التحقق…";' +
    'run().withSuccessHandler(function(r){show(r,false)}).withFailureHandler(function(e){b.disabled=false;b.textContent="عرض الإشعار";$("err").textContent=e&&e.message||"تعذّر التحقق"}).openNotice(tk,code,"")};' +
    '$("code").onkeydown=function(e){if(e.key==="Enter")$("go").click()}})()</script>';
  return wrap(gate + '<div id="view"></div><div id="act" style="max-width:440px;margin:14px auto 0;text-align:center"></div>' + js);
}
function appUrl_() { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }
function sendMail_(u, p) {
  const t = today_(), date = String(p.date || ''), i = parseInt(p.box, 10), pm = perms_(u);
  if (!pm.wa) throw new Error('ليست لديك صلاحية الإرسال');
  if (windowDates_(t, pm.days).indexOf(date) < 0) throw new Error('التقرير متاح لآخر ' + pm.days + ' أيام فقط');
  const ds = dataSheet_(), row = findRow_(ds, date); if (!row) throw new Error('لا توجد بيانات لهذا اليوم');
  const b = (JSON.parse(dayRaw_(ds, row)).boxes || [])[i] || {}, st = getSettings_();
  const mail = ((st.emails || {})[String(b.cashier || '').trim()]) || '';
  if (!mail) throw new Error((male_(st, b.cashier) ? 'لا يوجد إيميل لهذا الموظف' : 'لا يوجد إيميل لهذه الموظفة') + '، يضيفه المحاسب من الإعدادات');
  const sent = sentFor_(u, t).filter(function (e) { return e.date === date && e.box === i && e.ch === 'إيميل'; }).length;
  if (sent >= 3) throw new Error('أُرسل هذا الإشعار بالإيميل 3 مرات، ما يمكن إرساله أكثر');
  const net = NETS.reduce(function (a, x) { return a + (Number(b[x[0]]) || 0); }, 0);
  const diff = Math.round(((Number(b.cash) || 0) + net - (Number(b.report) || 0)) * 100) / 100;
  const tk = Utilities.getUuid().replace(/-/g, ''), by = u.name || u.username, link = appUrl_() ? appUrl_() + '?n=' + tk : '';
  const d = { date: date, day: dayName_(date), box: bn_(st, i), name: b.cashier || '', diff: diff, by: by, at: Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm'), cash: Number(b.cash) || 0, report: Number(b.report) || 0, store: st.store, branch: st.branch, title: ROLE_NAMES[u.role] };
  const lg = logo_().data, opt = { name: st.store + ' — ' + st.branch };
  let logoRef = '';
  const m = /^data:(image\/[a-z]+);base64,(.+)$/.exec(lg || '');
  if (m) { opt.inlineImages = { logo: Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'logo') }; logoRef = 'cid:logo'; }
  const im = /^data:image\/(png|jpeg);base64,([A-Za-z0-9+\/=]+)$/.exec(String(p.img || ''));
  if (im && im[2].length < 2000000) {
    const card = Utilities.newBlob(Utilities.base64Decode(im[2]), 'image/' + im[1], 'اشعار-' + d.box + '.' + (im[1] === 'png' ? 'png' : 'jpg'));
    opt.inlineImages = Object.assign(opt.inlineImages || {}, { card: card }); opt.attachments = [card];
    opt.htmlBody = '<div style="background:#eef2f7;padding:22px 10px;text-align:center;font-family:Tahoma,Arial,sans-serif;direction:rtl">' +
      '<img src="cid:card" alt="إشعار" style="max-width:440px;width:100%;border-radius:16px">' +
      (link ? '<div style="margin:16px 0 4px"><a href="' + link + '" style="display:inline-block;background:#1d3a5f;color:#fff;text-decoration:none;padding:12px 26px;border-radius:12px;font-weight:700;font-size:14px">عرض الإشعار وتأكيد الاطلاع</a><div style="font-size:12px;color:#5b6b7f;margin-top:6px">للتحقق يُطلب منك آخر 4 أرقام من جوالك</div></div>' : '') +
      '<div style="font-size:11px;color:#8a97a8;margin-top:10px">تم إرسال هذا الإشعار بواسطة ' + ROLE_NAMES[u.role] + ': ' + esc_(by) + '</div></div>';
  } else opt.htmlBody = '<div style="background:#eef2f7;padding:22px 10px">' + noticeHTML_(d, logoRef, true, link) + '</div>';
  MailApp.sendEmail(mail, (diff < 0 ? 'إشعار عجز' : 'إشعار زيادة') + ' — صندوق ' + d.box + ' — ' + date.split('-').reverse().join('/'),
    d.name + '، ' + (diff < 0 ? 'عجز' : 'زيادة') + ' بقيمة ' + money_(diff) + ' ريال في صندوق ' + d.box + ' بتاريخ ' + date + (link ? '\n' + link : '') + '\nتم الإرسال بواسطة ' + ROLE_NAMES[u.role] + ': ' + by, opt);
  logSent_(st, date, i, b, mail, diff, by, tk, 'إيميل');
  return sentFor_(u, t);
}
/* run this once from the editor after updating, so Google asks for the e-mail permission */
function authorizeEmail() { return MailApp.getRemainingDailyQuota(); }

/* a supervisor with the shortages screen sees the earlier days of the window as cash and difference only */
function supShape_(d) {
  const L = leadDays_([d])[0];
  L.boxes = L.boxes.map(function (x, i) { const b = (d.boxes || [])[i] || {}; return Object.assign({ cashier: b.cashier || '', shift: b.shift || '' }, x); });
  return L;
}
function supWindow_(u, t, okDates) {
  const pm = perms_(u); if (pm.tabs.indexOf('short') < 0) return [];
  const wk = windowDates_(t, pm.days).filter(function (d) { return okDates.indexOf(d) < 0; });
  if (!wk.length) return [];
  return readDays_(function (d) { return wk.indexOf(d) >= 0; }).days.map(supShape_);
}
function canSend_(u) { return (u.role === 'lead' || u.role === 'sup') && perms_(u).wa; }
function senderTitle_(name) {
  const u = users_().filter(function (x) { return (x.name || x.username) === name; })[0];
  return u && u.role === 'sup' ? ROLE_NAMES.sup : ROLE_NAMES.lead;
}
function leadDays_(days) {
  return days.map(function (d) {
    return {
      date: d.date, savedAt: d.savedAt, returns: [],
      boxes: (d.boxes || []).map(function (b) {
        if (!hasData_(b)) return {};
        const net = NETS.reduce(function (a, x) { return a + (Number(b[x[0]]) || 0); }, 0);
        return { cashier: b.cashier || '', cash: Number(b.cash) || 0, diff: Math.round(((Number(b.cash) || 0) + net - (Number(b.report) || 0)) * 100) / 100 };
      })
    };
  });
}

/* ---------- review additions ---------- */
function settingsFor_(u, st) {
  if (u.role === 'acc') return st;
  const o = Object.assign({}, st);
  o.staffShift = {}; Object.keys(st.staff || {}).forEach(function (k) { if (st.staff[k] && st.staff[k].shift) o.staffShift[k] = st.staff[k].shift; });
  if (!(u.role === 'mgr' && perms_(u).tabs.indexOf('att') >= 0)) delete o.staff;
  if (!canSend_(u)) { delete o.phones; delete o.emails; }
  return o;
}
function diffBox_(a, b) {
  const L = [['cashier', 'الكاشير'], ['shift', 'الوردية'], ['cash', 'الكاش']].concat(NETS).concat([['report', 'التقرير'], ['notes', 'ملاحظات']]), out = [];
  L.forEach(function (x) {
    const k = x[0], va = a[k] == null ? '' : a[k], vb = b[k] == null ? '' : b[k];
    if (String(va) !== String(vb)) out.push(x[1] + ': ' + (va === '' ? '—' : va) + ' ← ' + (vb === '' ? '—' : vb));
  });
  return out.join(' | ');
}
function audit_(p) {
  const s = SS_().getSheetByName('سجل التعديلات');
  if (!s || s.getLastRow() < 2) return [];
  const f = String(p.from || ''), t = String(p.to || '9999'), bx = String(p.box || '');
  const v = s.getRange(2, 1, s.getLastRow() - 1, 6).getDisplayValues(), out = [];
  for (let i = v.length - 1; i >= 0 && out.length < 3000; i--) {
    const r = v[i], d = String(r[0]).replace(/\//g, '-');
    if (d < f || d > t || (bx && String(r[1]) !== bx)) continue;
    out.push({ date: d, box: r[1], cashier: r[2], change: r[3], by: r[4], at: r[5] });
  }
  return out;
}
function appendAudit_(rows) {
  const ss = SS_(), name = 'سجل التعديلات';
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange('A:A').setNumberFormat('@');
    s.getRange(1, 1, 1, 6).setValues([['التاريخ', 'الصندوق', 'الكاشير', 'التعديل (قبل ← بعد)', 'عدّله', 'وقت التعديل']])
      .setFontWeight('bold').setBackground('#1d3a5f').setFontColor('#ffffff');
    s.setFrozenRows(1); s.setRightToLeft(true); s.setColumnWidth(4, 420);
  }
  const at = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss'), start = s.getLastRow() + 1;
  ensureRows_(s, start + rows.length - 1);
  s.getRange(start, 1, rows.length, 6).setValues(rows.map(function (r) { return r.concat([at]); }));
}
function cleanupDaySheets_() {
  const c = CacheService.getScriptCache(), t = today_(), key = 'clean_' + t;
  if (c.get(key)) return;
  c.put(key, '1', 21600);
  const d0 = new Date(t + 'T12:00:00Z'); d0.setUTCMonth(d0.getUTCMonth() - KEEP_MONTHS);
  const limit = Utilities.formatDate(d0, 'UTC', 'yyyy-MM') + '-01', ss = SS_();
  ss.getSheets().forEach(function (sh) {
    const n = sh.getName();
    if (/^\d{4}-\d{2}-\d{2}( - ف\d+)?$/.test(n) && n.slice(0, 10) < limit && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
}


/* ---------- moving a whole day to another date ---------- */
function moveDay_(u, p) {
  const from = String(p.from || ''), to = String(p.to || ''), iso = /^\d{4}-\d{2}-\d{2}$/;
  if (!iso.test(from) || !iso.test(to)) throw new Error('تاريخ غير صالح');
  if (from === to) throw new Error('اختر تاريخاً مختلفاً');
  const st = getSettings_(), t = today_(), who = u.name || u.username;
  if (!periodOpen_(from) || !periodOpen_(to)) throw new Error(CLOSED_MSG);
  if (to > t && !st.allowFuture) throw new Error(FUTURE_MSG);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let doc;
  try {
    const s = dataSheet_(), row = findRow_(s, from);
    if (!row) throw new Error('لا توجد بيانات محفوظة على ' + from);
    const trow = findRow_(s, to);
    if (trow) {
      let td = null; try { td = JSON.parse(dayRaw_(s, trow)); } catch (e) { }
      const busy = td && ((td.boxes || []).some(hasData_) || (td.returns || []).length);
      if (busy) throw new Error('يوجد بيانات محفوظة على تاريخ ' + to.split('-').reverse().join('/') + '، احذفها أولاً أو اختر تاريخاً آخر');
    }
    const body = JSON.parse(dayRaw_(s, row)), nowISO = new Date().toISOString();
    doc = { date: to, boxes: body.boxes || [], returns: body.returns || [], savedAt: nowISO, savedBy: who };
    s.getRange(row, 1, 1, 4).setValues([[to, s.getRange(row, 2).getDisplayValue(), nowISO, who]]);
    if (trow) s.deleteRow(trow);
    const ss = SS_(), old = ss.getSheetByName(from);
    if (old) ss.deleteSheet(old);
    writeDaySheet_(doc, st);
    deleteLogRows_(from); writeLog_(doc);
    // shortage notices sent for that day follow it, so their links keep working
    const ns = ss.getSheetByName(SENT);
    if (ns && ns.getLastRow() > 1) {
      const v = ns.getRange(2, 1, ns.getLastRow() - 1, 2).getDisplayValues();
      v.forEach(function (r, i) { if (r[0] === from) ns.getRange(i + 2, 1, 1, 2).setValues([[to, dayName_(to)]]); });
    }
    appendAudit_([[to, '—', '', 'نقل بيانات اليوم كاملة من ' + from.split('-').reverse().join('/') + ' إلى ' + to.split('-').reverse().join('/'), who]]);
    tomb_(from, true); tomb_(to, false);
    bump_();
  } finally { lock.releaseLock(); }
  return doc;
}

/* one box (with its returns) moves to the same box number on another date */
function moveBox_(u, p) {
  const from = String(p.from || ''), to = String(p.to || ''), i = parseInt(p.box, 10), iso = /^\d{4}-\d{2}-\d{2}$/;
  if (!iso.test(from) || !iso.test(to)) throw new Error('تاريخ غير صالح');
  if (from === to) throw new Error('اختر تاريخاً مختلفاً');
  const st = getSettings_(), t = today_(), who = u.name || u.username, fmtD = function (d) { return d.split('-').reverse().join('/'); };
  if (!(i >= 0 && i < MAX_BOXES)) throw new Error('صندوق غير صالح');
  if (!periodOpen_(from) || !periodOpen_(to)) throw new Error(CLOSED_MSG);
  if (to > t && !st.allowFuture) throw new Error(FUTURE_MSG);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let a, z;
  try {
    const s = dataSheet_(), row = findRow_(s, from);
    if (!row) throw new Error('لا توجد بيانات محفوظة على ' + fmtD(from));
    const src = JSON.parse(dayRaw_(s, row)); src.boxes = src.boxes || []; src.returns = src.returns || [];
    const b = src.boxes[i];
    if (!b || !hasData_(b)) throw new Error('هذا الصندوق فاضي');
    if (b.rq) throw new Error('الصندوق بانتظار الاعتماد، اعتمده أو ارفضه أولاً');
    const trow = findRow_(s, to);
    let dst = { boxes: [], returns: [] };
    if (trow) { try { dst = JSON.parse(dayRaw_(s, trow)); } catch (e) { } dst.boxes = dst.boxes || []; dst.returns = dst.returns || []; }
    const NB = Math.max(nBoxes_(st), dst.boxes.length, i + 1);
    for (let k = 0; k < NB; k++) dst.boxes[k] = Object.assign(blankBox_(k), dst.boxes[k] || {}, { no: k + 1 });
    if (hasData_(dst.boxes[i]) || dst.boxes[i].rq) throw new Error('الصندوق ' + bn_(st, i) + ' على تاريخ ' + fmtD(to) + ' فيه بيانات، احذفها أولاً أو انقل اليوم كامل');
    dst.boxes[i] = Object.assign({}, b, { no: i + 1 });
    src.boxes[i] = blankBox_(i);
    const mine = function (r) { return +r.box === i + 1; };
    dst.returns = dst.returns.concat(src.returns.filter(mine)); src.returns = src.returns.filter(function (r) { return !mine(r); });
    const nowISO = new Date().toISOString();
    z = { date: to, boxes: dst.boxes, returns: dst.returns, savedAt: nowISO, savedBy: who };
    a = { date: from, boxes: src.boxes, returns: src.returns, savedAt: nowISO, savedBy: who };
    dayWrite_(s, trow || s.getLastRow() + 1, to, JSON.stringify({ boxes: z.boxes, returns: z.returns }), nowISO, who);
    dayWrite_(s, findRow_(s, from), from, JSON.stringify({ boxes: a.boxes, returns: a.returns }), nowISO, who);
    writeDaySheet_(z, st); writeDaySheet_(a, st); writeLog_(z); writeLog_(a);
    const ns = SS_().getSheetByName(SENT);
    if (ns && ns.getLastRow() > 1) {
      const v = ns.getRange(2, 1, ns.getLastRow() - 1, 9).getDisplayValues();
      v.forEach(function (r, k) { if (r[0] === from && Number(r[8]) === i) ns.getRange(k + 2, 1, 1, 2).setValues([[to, dayName_(to)]]); });
    }
    appendAudit_([[to, bn_(st, i), b.cashier || '', 'نقل الصندوق من ' + fmtD(from) + ' إلى ' + fmtD(to), who]]);
    tomb_(to, false); tomb_(from, false);
    bump_();
  } finally { lock.releaseLock(); }
  return { from: a, to: z };
}

/* ---------- main cash box: monthly ledger ---------- */
const CB = '_cashbox', CBM = '_cashbox_months';
function cbSheet_() { return hiddenSheet_(CB, ['id', 'month', 'kind', 'date', 'desc', 'amount', 'by', 'at']); }
function cbmSheet_() { return hiddenSheet_(CBM, ['month', 'opening', 'src', 'carried', 'carriedAt', 'by']); }
function cbMonthOk_(m) { if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(m)) throw new Error('شهر غير صالح'); return m; }
function nextMonth_(m) { const y = +m.slice(0, 4), k = +m.slice(5, 7); return k === 12 ? (y + 1) + '-01' : y + '-' + ('0' + (k + 1)).slice(-2); }
function prevMonth_(m) { const y = +m.slice(0, 4), k = +m.slice(5, 7); return k === 1 ? (y - 1) + '-12' : y + '-' + ('0' + (k - 1)).slice(-2); }
function cbMeta_(m) {
  const s = cbmSheet_(), n = s.getLastRow();
  if (n >= 2) { const v = s.getRange(2, 1, n - 1, 6).getDisplayValues(); for (let i = 0; i < v.length; i++) if (v[i][0] === m) return { row: i + 2, month: m, opening: Number(v[i][1]) || 0, src: v[i][2] || '', carried: v[i][3] === '1', carriedAt: v[i][4] || '', by: v[i][5] || '' }; }
  return { row: 0, month: m, opening: 0, src: '', carried: false, carriedAt: '', by: '' };
}
function cbMetaPut_(o) {
  const s = cbmSheet_(), vals = [[o.month, String(Math.round(o.opening * 100) / 100), o.src || '', o.carried ? '1' : '', o.carriedAt || '', o.by || '']];
  if (o.row) s.getRange(o.row, 1, 1, 6).setValues(vals); else { const r = s.getLastRow() + 1; ensureRows_(s, r); s.getRange(r, 1, 1, 6).setValues(vals); o.row = r; }
}
function cbEntries_(m) {
  const s = cbSheet_(), n = s.getLastRow(); if (n < 2) return [];
  return s.getRange(2, 1, n - 1, 8).getDisplayValues().map(function (r, i) { return { row: i + 2, id: r[0], month: r[1], kind: r[2], date: r[3], desc: r[4], amount: Number(r[5]) || 0, by: r[6], at: r[7] }; })
    .filter(function (e) { return e.month === m && e.id; });
}
/* cash sales of each day of the month: the approved cash after cash returns */
function cbSales_(m) {
  return readDays_(function (d) { return d.slice(0, 7) === m; }).days.map(function (d) {
    const T = calc_(approvedOnly_(d)).T, net = Math.round((T.cash + T.ret) * 100) / 100;
    const fx = Math.round((d.boxes || []).reduce(function (a, b) { return a + (b && b.fxd ? (Number(b.cash) || 0) - (Number(b.cashSar) || 0) : 0); }, 0) * 100) / 100;
    return { date: d.date, cash: Math.round(T.cash * 100) / 100, ret: Math.round(T.ret * 100) / 100, net: net, fx: fx };
  }).filter(function (x) { return x.cash || x.ret; }).sort(function (a, b) { return a.date < b.date ? -1 : 1; });
}
function cbTotals_(meta, entries, sales) {
  const r2 = function (v) { return Math.round(v * 100) / 100; };
  const salesIn = r2(sales.reduce(function (a, x) { return a + x.net; }, 0));
  const otherIn = r2(entries.filter(function (e) { return e.kind === 'in'; }).reduce(function (a, e) { return a + e.amount; }, 0));
  const out = r2(entries.filter(function (e) { return e.kind === 'out'; }).reduce(function (a, e) { return a + e.amount; }, 0));
  const inT = r2(salesIn + otherIn);
  return { opening: meta.opening, salesIn: salesIn, otherIn: otherIn, inTotal: inT, out: out, net: r2(inT - out), closing: r2(meta.opening + inT - out) };
}
function cbGet_(m) {
  cbMonthOk_(m);
  const meta = cbMeta_(m), entries = cbEntries_(m), sales = cbSales_(m), prev = cbMeta_(prevMonth_(m));
  const pub = function (e) { return { id: e.id, kind: e.kind, date: e.date, desc: e.desc, amount: e.amount, by: e.by, at: e.at }; };
  return { month: m, opening: meta.opening, src: meta.src, carried: meta.carried, carriedAt: meta.carriedAt, carriedBy: meta.by, prevCarried: prev.carried,
    entries: entries.map(pub).sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.at < b.at ? -1 : 1); }), sales: sales, totals: cbTotals_(meta, entries, sales) };
}
function cbOpenMonth_(m) { const meta = cbMeta_(m); if (meta.carried) throw new Error('رصيد هذا الشهر مُرحَّل، ألغِ الترحيل أولاً للتعديل'); return meta; }
function cbSave_(u, p) {
  const m = cbMonthOk_(String(p.month || '')), kind = p.kind === 'out' ? 'out' : p.kind === 'in' ? 'in' : '';
  if (!kind) throw new Error('نوع الحركة غير صالح');
  const date = String(p.date || ''); if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date.slice(0, 7) !== m) throw new Error('التاريخ لازم يكون داخل الشهر');
  if (date > today_() && !getSettings_().allowFuture) throw new Error(FUTURE_MSG);
  const amount = Math.round((Number(p.amount) || 0) * 100) / 100; if (!(amount > 0) || amount > 100000000) throw new Error('اكتب مبلغاً صحيحاً أكبر من صفر');
  const desc = String(p.desc || '').trim().slice(0, 120); if (!desc) throw new Error('اكتب البيان');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    cbOpenMonth_(m);
    const s = cbSheet_(), who = u.name || u.username, at = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm');
    const id = String(p.id || '');
    if (id) {
      const e = cbEntries_(m).filter(function (x) { return x.id === id; })[0]; if (!e) throw new Error('الحركة غير موجودة');
      s.getRange(e.row, 3, 1, 6).setValues([[kind, date, desc, String(amount), who, at]]);
      appendAudit_([[date, 'الصندوق', '', 'تعديل ' + (e.kind === 'out' ? 'مخرج' : 'مدخل') + ': ' + e.desc + ' ' + e.amount + ' ← ' + desc + ' ' + amount, who]]);
    } else {
      const r = s.getLastRow() + 1; ensureRows_(s, r);
      s.getRange(r, 1, 1, 8).setValues([[Utilities.getUuid().slice(0, 12), m, kind, date, desc, String(amount), who, at]]);
    }
    bump_();
  } finally { lock.releaseLock(); }
  return cbGet_(m);
}
function cbDel_(u, p) {
  const m = cbMonthOk_(String(p.month || '')), id = String(p.id || '');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    cbOpenMonth_(m);
    const e = cbEntries_(m).filter(function (x) { return x.id === id; })[0]; if (!e) throw new Error('الحركة غير موجودة');
    cbSheet_().deleteRow(e.row);
    appendAudit_([[e.date, 'الصندوق', '', 'حذف ' + (e.kind === 'out' ? 'مخرج' : 'مدخل') + ': ' + e.desc + ' ' + e.amount, u.name || u.username]]);
    bump_();
  } finally { lock.releaseLock(); }
  return cbGet_(m);
}
function cbOpen_(u, p) {
  const m = cbMonthOk_(String(p.month || '')), v = Math.round((Number(p.amount) || 0) * 100) / 100;
  if (!isFinite(v) || Math.abs(v) > 100000000) throw new Error('مبلغ غير صالح');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const meta = cbOpenMonth_(m), was = meta.opening;
    meta.opening = v; meta.src = 'manual'; meta.by = u.name || u.username; cbMetaPut_(meta);
    appendAudit_([[m + '-01', 'الصندوق', '', 'الرصيد الافتتاحي: ' + was + ' ← ' + v, meta.by]]);
    bump_();
  } finally { lock.releaseLock(); }
  return cbGet_(m);
}
function cbCarry_(u, p) {
  const m = cbMonthOk_(String(p.month || ''));
  if (m >= today_().slice(0, 7) && !p.force) throw new Error('الشهر ما انتهى بعد، تأكيد الترحيل مطلوب');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const meta = cbOpenMonth_(m), T = cbTotals_(meta, cbEntries_(m), cbSales_(m)), who = u.name || u.username;
    const nx = cbMeta_(nextMonth_(m));
    if (nx.carried) throw new Error('الشهر التالي مُرحَّل، ما يمكن تغيير رصيده الافتتاحي');
    meta.carried = true; meta.carriedAt = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm'); meta.by = who; cbMetaPut_(meta);
    nx.opening = T.closing; nx.src = 'carry'; nx.by = who; cbMetaPut_(nx);
    appendAudit_([[m + '-01', 'الصندوق', '', 'ترحيل رصيد الإقفال ' + T.closing + ' إلى ' + nextMonth_(m), who]]);
    bump_();
  } finally { lock.releaseLock(); }
  return cbGet_(m);
}
function cbUncarry_(u, p) {
  const m = cbMonthOk_(String(p.month || ''));
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const meta = cbMeta_(m); if (!meta.carried) throw new Error('هذا الشهر غير مُرحَّل');
    const nx = cbMeta_(nextMonth_(m)); if (nx.carried) throw new Error('الشهر التالي مُرحَّل أيضاً، ألغِ ترحيله أولاً');
    meta.carried = false; meta.carriedAt = ''; meta.by = u.name || u.username; cbMetaPut_(meta);
    if (nx.src === 'carry') { nx.opening = 0; nx.src = ''; cbMetaPut_(nx); }
    appendAudit_([[m + '-01', 'الصندوق', '', 'إلغاء ترحيل رصيد ' + m, meta.by]]);
    bump_();
  } finally { lock.releaseLock(); }
  return cbGet_(m);
}


/* ---------- branch management ---------- */
function saveBranch_(u, p) {
  const name = String(p.name || '').trim().replace(/\s+/g, ' ').slice(0, 60);
  if (!name) throw new Error('اكتب اسم الفرع');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    withBr_('', function () {
      const m = getSettings_(), list = (m.branches || []).slice(), id = String(p.id || '');
      if (id === 'main') m.branch = name;
      else if (id) { const b = list.filter(function (x) { return x.id === id; })[0]; if (!b) throw new Error('الفرع غير موجود'); b.name = name; }
      else {
        if (list.length >= 30) throw new Error('الحد الأعلى 30 فرعاً');
        let n = 2; list.forEach(function (b) { n = Math.max(n, (parseInt(String(b.id).slice(1), 10) || 1) + 1); });
        list.push({ id: 'f' + n, name: name });
      }
      if ([m.branch].concat(list.map(function (b) { return b.name; })).filter(function (x) { return x === name; }).length > 1) throw new Error('يوجد فرع بنفس الاسم');
      m.branches = list; hiddenSheet_(SETTINGS).getRange(1, 1).setValue(JSON.stringify(m));
    });
    if (p.id && p.id !== 'main') withBr_(p.id, function () { const st = getSettings_(); st.branch = name; writeSettings_(st); });
    bumpUv_(); bump_();
  } finally { lock.releaseLock(); }
  return branchIds_().map(function (id) { return { id: id, name: branchName_(id) }; });
}
function deleteBranch_(u, p) {
  const id = String(p.id || '');
  if (!id || id === 'main') throw new Error('ما يمكن حذف الفرع الرئيسي');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    if (branchIds_().indexOf(id) < 0) throw new Error('الفرع غير موجود');
    const used = withBr_(id, function () { const s = SpreadsheetApp.getActive().getSheetByName(bname_(DATA)); return s && s.getLastRow() > 1; });
    if (used) throw new Error('الفرع عليه بيانات محفوظة، ما يمكن حذفه');
    if (users_().some(function (x) { return x.branch === id; })) throw new Error('فيه مستخدمين مرتبطين بهذا الفرع، انقلهم أو احذفهم أولاً');
    withBr_('', function () { const m = getSettings_(); m.branches = (m.branches || []).filter(function (b) { return b.id !== id; }); hiddenSheet_(SETTINGS).getRange(1, 1).setValue(JSON.stringify(m)); });
    bumpUv_(); bump_();
  } finally { lock.releaseLock(); }
  return branchIds_().map(function (x) { return { id: x, name: branchName_(x) }; });
}
/* sales of several branches for a period: one line per branch, the caller adds them up when it wants a combined report */
function brReport_(u, p) {
  const iso = /^\d{4}-\d{2}-\d{2}$/, from = String(p.from || ''), to = String(p.to || '');
  if (!iso.test(from) || !iso.test(to) || from > to) throw new Error('فترة غير صالحة');
  const al = userBranches_(u), want = Array.isArray(p.ids) && p.ids.length ? p.ids.map(String).filter(function (id) { return al.indexOf(id) >= 0; }) : al;
  const r2 = function (v) { return Math.round(v * 100) / 100; }, tot = u.role === 'own' && perms_(u).totals;
  const out = want.map(function (id) {
    return withBr_(id === 'main' ? '' : id, function () {
      const o = { id: id, name: branchName_(id), days: 0, cash: 0, ret: 0, nets: {}, net: 0, total: 0, report: 0, diff: 0, short: 0, over: 0, boxes: 0, perDay: [] };
      NETS.forEach(function (x) { o.nets[x[0]] = 0; });
      readDays_(function (d) { return d >= from && d <= to; }).days.forEach(function (d) {
        const C = calc_(approvedOnly_(d)); if (!C.rows.length && !C.T.ret) return;
        o.days++; o.cash += C.T.cash; o.ret += C.T.ret; o.net += C.T.net; o.report += C.T.report; o.boxes += C.rows.length;
        NETS.forEach(function (x) { o.nets[x[0]] += C.T.nets[x[0]]; });
        let dd = 0; C.rows.forEach(function (r) { dd += r.diff; if (r.diff < 0) o.short += r.diff; else o.over += r.diff; });
        o.diff += dd;
        o.perDay.push({ date: d.date, cash: r2(C.T.cash), net: r2(C.T.net), ret: r2(C.T.ret), total: r2(C.T.cash + C.T.net), report: r2(C.T.report), diff: r2(dd) });
      });
      ['cash', 'ret', 'net', 'report', 'diff', 'short', 'over'].forEach(function (k) { o[k] = r2(o[k]); });
      NETS.forEach(function (x) { o.nets[x[0]] = r2(o.nets[x[0]]); });
      o.total = r2(o.cash + o.net);
      o.perDay.sort(function (a, b) { return a.date < b.date ? -1 : 1; });
      return o;
    });
  });
  if (!tot) return out;
  return out.map(function (o) {
    return { id: o.id, name: o.name, days: o.days, cash: o.cash, nets: o.nets, net: o.net, total: o.total,
      perDay: o.perDay.map(function (d) { return { date: d.date, cash: d.cash, net: d.net, total: d.total }; }) };
  });
}


/* ---------- petty cash custody: each custodian has an advance, spends it on invoices, and settles ---------- */
const PTP = '_petty_people', PTX = '_petty', PTS = '_petty_settle';
function ptpSheet_() { return hiddenSheet_(PTP, ['id', 'name', 'amount', 'active', 'by', 'at']); }
function ptxSheet_() { return hiddenSheet_(PTX, ['id', 'pid', 'kind', 'date', 'desc', 'ref', 'amount', 'settle', 'by', 'at']); }
function ptsSheet_() { return hiddenSheet_(PTS, ['id', 'pid', 'no', 'date', 'from', 'to', 'exp', 'refill', 'balance', 'by', 'at']); }
function ptRows_(sh, w) { const n = sh.getLastRow(); return n < 2 ? [] : sh.getRange(2, 1, n - 1, w).getDisplayValues().map(function (r, i) { r.row = i + 2; return r; }).filter(function (r) { return r[0]; }); }
const r2s_ = function (v) { return Math.round((Number(v) || 0) * 100) / 100; };
function ptGet_() {
  const P = ptRows_(ptpSheet_(), 6), X = ptRows_(ptxSheet_(), 10), S = ptRows_(ptsSheet_(), 11);
  const tx = X.map(function (r) { return { id: r[0], pid: r[1], kind: r[2], date: r[3], desc: r[4], ref: r[5], amount: Number(r[6]) || 0, settle: r[7], by: r[8], at: r[9] }; });
  const st = S.map(function (r) { return { id: r[0], pid: r[1], no: Number(r[2]) || 0, date: r[3], from: r[4], to: r[5], exp: Number(r[6]) || 0, refill: Number(r[7]) || 0, balance: Number(r[8]) || 0, by: r[9], at: r[10] }; });
  const people = P.map(function (r) {
    const id = r[0], my = tx.filter(function (t) { return t.pid === id; });
    const sum = function (k, open) { return r2s_(my.filter(function (t) { return t.kind === k && (!open || !t.settle); }).reduce(function (a, t) { return a + t.amount; }, 0)); };
    const funded = sum('fund'), spent = sum('exp'), returned = sum('ret');
    return { id: id, name: r[1], amount: Number(r[2]) || 0, active: r[3] !== '0', funded: funded, spent: spent, returned: returned, balance: r2s_(funded - spent - returned), openExp: sum('exp', true), openCount: my.filter(function (t) { return t.kind === 'exp' && !t.settle; }).length };
  });
  return { people: people, tx: tx.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.at < b.at ? -1 : 1); }), settles: st.sort(function (a, b) { return b.no - a.no; }) };
}
function ptAt_() { return Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm'); }
function ptPerson_(u, p) {
  const name = String(p.name || '').trim().replace(/\s+/g, ' ').slice(0, 60); if (!name) throw new Error('اكتب اسم صاحب العهدة');
  const amount = r2s_(p.amount); if (!(amount >= 0) || amount > 10000000) throw new Error('مبلغ العهدة غير صالح');
  const who = u.name || u.username, lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = ptpSheet_(), all = ptRows_(sh, 6), id = String(p.id || '');
    if (all.some(function (r) { return r[1] === name && r[0] !== id; })) throw new Error('يوجد صاحب عهدة بنفس الاسم');
    if (id) {
      const r = all.filter(function (x) { return x[0] === id; })[0]; if (!r) throw new Error('صاحب العهدة غير موجود');
      sh.getRange(r.row, 2, 1, 5).setValues([[name, String(amount), p.active === false ? '0' : '1', who, ptAt_()]]);
    } else {
      const nid = 'p' + Utilities.getUuid().slice(0, 8), row = sh.getLastRow() + 1; ensureRows_(sh, row);
      sh.getRange(row, 1, 1, 6).setValues([[nid, name, String(amount), '1', who, ptAt_()]]);
      if (amount > 0) ptAdd_(nid, 'fund', today_(), 'صرف العهدة', '', amount, '', who);
      appendAudit_([[today_(), 'العهدة', name, 'صاحب عهدة جديد بمبلغ ' + amount, who]]);
    }
    bump_();
  } finally { lock.releaseLock(); }
  return ptGet_();
}
function ptAdd_(pid, kind, date, desc, ref, amount, settle, who) {
  const sh = ptxSheet_(), row = sh.getLastRow() + 1; ensureRows_(sh, row);
  const id = 't' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  sh.getRange(row, 1, 1, 10).setValues([[id, pid, kind, date, desc, ref, String(r2s_(amount)), settle || '', who, ptAt_()]]);
  return id;
}
function ptTx_(u, p) {
  const kind = ['fund', 'exp', 'ret'].indexOf(p.kind) >= 0 ? p.kind : ''; if (!kind) throw new Error('نوع الحركة غير صالح');
  const date = String(p.date || ''); if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('تاريخ غير صالح');
  if (date > today_() && !getSettings_().allowFuture) throw new Error(FUTURE_MSG);
  const amount = r2s_(p.amount); if (!(amount > 0) || amount > 10000000) throw new Error('اكتب مبلغاً صحيحاً أكبر من صفر');
  const desc = String(p.desc || '').trim().slice(0, 120); if (!desc) throw new Error('اكتب البيان');
  const ref = String(p.ref || '').trim().slice(0, 40), who = u.name || u.username;
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const pr = ptRows_(ptpSheet_(), 6).filter(function (r) { return r[0] === String(p.pid || ''); })[0]; if (!pr) throw new Error('صاحب العهدة غير موجود');
    const id = String(p.id || '');
    if (id) {
      const sh = ptxSheet_(), r = ptRows_(sh, 10).filter(function (x) { return x[0] === id; })[0]; if (!r) throw new Error('الحركة غير موجودة');
      if (r[7]) throw new Error('هذه الحركة ضمن تصفية سابقة ومقفلة');
      sh.getRange(r.row, 3, 1, 5).setValues([[kind, date, desc, ref, String(amount)]]); sh.getRange(r.row, 9, 1, 2).setValues([[who, ptAt_()]]);
      appendAudit_([[date, 'العهدة', pr[1], 'تعديل حركة: ' + r[4] + ' ' + r[6] + ' ← ' + desc + ' ' + amount, who]]);
    } else ptAdd_(pr[0], kind, date, desc, ref, amount, '', who);
    bump_();
  } finally { lock.releaseLock(); }
  return ptGet_();
}
function ptDel_(u, p) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = ptxSheet_(), r = ptRows_(sh, 10).filter(function (x) { return x[0] === String(p.id || ''); })[0]; if (!r) throw new Error('الحركة غير موجودة');
    if (r[7]) throw new Error('هذه الحركة ضمن تصفية سابقة ومقفلة');
    sh.deleteRow(r.row);
    appendAudit_([[r[3], 'العهدة', '', 'حذف حركة: ' + r[4] + ' ' + r[6], u.name || u.username]]);
    bump_();
  } finally { lock.releaseLock(); }
  return ptGet_();
}
/* settling closes the open invoices of one custodian; replenishing pays him back what he spent */
function ptSettle_(u, p) {
  const pid = String(p.pid || ''), who = u.name || u.username;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let sid = '';
  try {
    const pr = ptRows_(ptpSheet_(), 6).filter(function (r) { return r[0] === pid; })[0]; if (!pr) throw new Error('صاحب العهدة غير موجود');
    const sh = ptxSheet_(), open = ptRows_(sh, 10).filter(function (r) { return r[1] === pid && !r[7] && r[2] === 'exp'; });
    if (!open.length) throw new Error('ما فيه مصروفات مفتوحة للتصفية');
    const exp = r2s_(open.reduce(function (a, r) { return a + (Number(r[6]) || 0); }, 0)), dates = open.map(function (r) { return r[3]; }).sort();
    const S = ptsSheet_(), prev = ptRows_(S, 11).filter(function (r) { return r[1] === pid; }), no = prev.reduce(function (a, r) { return Math.max(a, Number(r[2]) || 0); }, 0) + 1;
    sid = 's' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    open.forEach(function (r) { sh.getRange(r.row, 8).setValue(sid); });
    const refill = p.refill ? exp : 0;
    if (refill) ptAdd_(pid, 'fund', today_(), 'استعاضة تصفية رقم ' + no, '', refill, sid, who);
    const bal = ptGet_().people.filter(function (x) { return x.id === pid; })[0].balance;
    const row = S.getLastRow() + 1; ensureRows_(S, row);
    S.getRange(row, 1, 1, 11).setValues([[sid, pid, String(no), today_(), dates[0], dates[dates.length - 1], String(exp), String(refill), String(bal), who, ptAt_()]]);
    appendAudit_([[today_(), 'العهدة', pr[1], 'تصفية رقم ' + no + ' بمبلغ ' + exp + (refill ? ' مع الاستعاضة' : ''), who]]);
    bump_();
  } finally { lock.releaseLock(); }
  const g = ptGet_(); g.settled = sid; return g;
}
function ptUnsettle_(u, p) {
  const pid = String(p.pid || ''), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const S = ptsSheet_(), mine = ptRows_(S, 11).filter(function (r) { return r[1] === pid; }).sort(function (a, b) { return (Number(b[2]) || 0) - (Number(a[2]) || 0); });
    if (!mine.length) throw new Error('ما فيه تصفية سابقة');
    const last = mine[0], sid = last[0], sh = ptxSheet_(), rows = ptRows_(sh, 10).filter(function (r) { return r[7] === sid; });
    rows.slice().sort(function (a, b) { return b.row - a.row; }).forEach(function (r) { if (r[2] === 'fund' && /^استعاضة تصفية/.test(r[4])) sh.deleteRow(r.row); else sh.getRange(r.row, 8).setValue(''); });
    S.deleteRow(last.row);
    appendAudit_([[today_(), 'العهدة', '', 'إلغاء التصفية رقم ' + last[2], u.name || u.username]]);
    bump_();
  } finally { lock.releaseLock(); }
  return ptGet_();
}
