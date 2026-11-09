/**
 * CEO Group — Operations backend (Google Apps Script)          v4 · Phase 1
 * --------------------------------------------------------------------------
 * Bound to the CEO Operations sheet (Extensions → Apps Script).
 * Deploy → Manage deployments → Edit → New version   (keeps the same Web App URL)
 *
 * Phase 1:
 *  • Students live in a SEPARATE sheet (CEO Students DB). Its link goes in Settings!B8.
 *    syncFromDb() copies it into the "Students" tab every 10 minutes (installed by setup).
 *  • The landing page looks the phone up (action=lookup) → shows the student's name,
 *    the module running now in their group and the doctor.
 *  • Unregistered numbers: the student types a name; only attendance is recorded.
 *
 * Form Responses 1:
 *  A Timestamp | B Phone | C Accreditation | D Module | E Page Opened At | F Submission ID | G Source
 *  H..O formulas (date, day, time, entry, quality, phone key, student name, group) | P Typed Name (unregistered)
 */

const CODE_VERSION = 'v8.9';

const SETTINGS = {
  RESPONSES_SHEET: 'Form Responses 1',
  CATALOG_SHEET: 'Module Catalog',
  TIMEZONE: 'Africa/Cairo',
  RAW_COLS: 7,          // A..G
  FORMULA_COL: 8,       // H
  FORMULA_COUNT: 8,     // H..O
  TYPED_NAME_COL: 16,   // P
  ACC_ORDER: [
    'EIU MBA', 'EIU MBA Healthcare', 'EIU BBA 1 Year', 'EIU DBA',
    'AITU MBA', 'AITU DBA', 'QUALIFI level 7 leadership', 'QUALIFI level 7 Innovation'
  ]
};

const T = {
  STUDENTS: 'Students', TRACKER: 'Tracker', SCHEDULE: 'Group Schedule', DASH: 'Dashboard', SETTINGS: 'Settings',
  DB_TAB: 'Students',                 // tab name inside the Students DB sheet
  DASH_FIRST_ROW: 9,
  FR: "'Form Responses 1'", GS: "'Group Schedule'"
};
// Operations › Students columns
const SC = { KEYMATCH: 1, KEY: 2, PHONE: 3, NAME: 4, ACC: 5, GROUP: 6, STATUS: 7, START_MOD: 8, START_DATE: 9, F_FIRST: 10, F_COUNT: 7, NOTE: 17, SOURCE: 18, DONE: 19, PLAN_NOTE: 20, WIDTH: 20 };
// Students DB columns
const DBC = { PHONE: 1, NAME: 2, ACC: 3, GROUP: 4, STATUS: 5, START_MOD: 6, START_DATE: 7, NOTES: 8, SOURCE: 9, DONE: 10, WIDTH: 10 };
const ACTIVE_STATUS = 'Still studying';
const AUTO_SOURCE = /^(EIU|AITU) · /;   // rows written by importSources_ (manual rows have another / empty source)

/* =========================== Web endpoints =========================== */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'catalog';
  if (action === 'catalog') return json_({ ok: true, catalog: getCatalog_() });
  if (action === 'lookup') return json_(lookup_(p.phone));
  if (action === 'lecture') return json_(lectureInfo_(p.phone));
  if (action === 'evalcheck') return json_(evalCheck_(p.phone));
  if (action === 'pathcheck') return json_(pathCheck_(p.phone));
  if (action === 'ping') return json_({ ok: true, time: new Date().toISOString() });
  if (action === 'version') return json_({ ok: true, version: CODE_VERSION, timezone: SpreadsheetApp.getActive().getSpreadsheetTimeZone() });
  if (action === 'check') return json_(checkSubmission_(String(p.id || '')));
  return json_({ ok: false, error: 'Unknown action' });
}

function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Invalid request' }); }
  if (data.website) return json_({ ok: true });   // honeypot
  if (data.form === 'evaluation') return json_(submitEvaluation_(data));   // evaluation page (evaluation.html)
  if (data.form === 'path') return json_(submitPath_(data));               // evaluation page › academic path tab

  const phone = normalizePhone_(data.phone);
  if (!/^\+?\d{8,15}$/.test(phone)) return json_({ ok: false, error: 'Invalid phone number' });

  // Who is this? (server decides – the page cannot fake module/accreditation)
  const found = findStudent_(phone);
  let accreditation = '', module = '', typedName = '', masked = '';
  if (found.length) {
    const wanted = String(data.accreditation || '').trim();
    const en = found.filter(function (x) { return x.acc === wanted; })[0] || found.filter(function (x) { return x.active; })[0] || found[0];
    accreditation = en.acc; module = en.module || ''; masked = en.name;
  } else {
    typedName = String(data.name || '').replace(/\s+/g, ' ').trim().slice(0, 80);
    if (typedName.length < 3) {
      // old page (v1.x) still sends accreditation + module: accept them
      const catalog = getCatalog_(), acc = String(data.accreditation || '').trim(), mod = String(data.module || '').trim();
      if (catalog[acc] && catalog[acc].indexOf(mod) !== -1) { accreditation = acc; module = mod; }
      else return json_({ ok: false, error: 'Name is required' });
    }
  }

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return json_({ ok: false, error: 'Server busy, please try again' }); }
  try {
    const sh = sheet_(SETTINGS.RESPONSES_SHEET);
    const now = new Date();
    const submissionId = String(data.submissionId || Utilities.getUuid()).slice(0, 64);
    const row = firstEmptyRow_(sh);
    if (row > 2) {
      const ids = sh.getRange(2, 6, row - 2, 1).getValues();
      for (let i = ids.length - 1; i >= Math.max(0, ids.length - 300); i--) {
        if (ids[i][0] === submissionId) return json_({ ok: true, timestamp: now.toISOString(), resent: true, known: found.length > 0 });
      }
    }
    const duplicate = isDuplicateToday_(sh, row, phone, module, now);
    sh.getRange(row, 2).setNumberFormat('@');
    sh.getRange(row, 1, 1, SETTINGS.RAW_COLS).setValues([[
      now, phone, accreditation, module, parseDate_(data.pageOpenedAt) || '', submissionId, String(data.source || 'Landing Page')
    ]]);
    sh.getRange(row, SETTINGS.TYPED_NAME_COL).setValue(typedName);
    writeResponseFormulas_(sh, row, 1);
    return json_({ ok: true, timestamp: now.toISOString(), duplicate: duplicate, row: row,
                   known: found.length > 0, name: masked || typedName, module: module, accreditation: accreditation });
  } finally {
    lock.releaseLock();
  }
}

/* =========================== Phone lookup =========================== */
/** Public lookup used by the page. Returns a masked name only (privacy). */
function lookup_(rawPhone) {
  const phone = normalizePhone_(rawPhone);
  if (!/^\+?\d{8,15}$/.test(phone)) return { ok: false, error: 'Invalid phone number' };
  let found = findStudent_(phone);
  if (!found.length) return { ok: true, found: false };
  if (found.some(function (x) { return x.active; })) found = found.filter(function (x) { return x.active; });   // hide old / inactive programmes
  return { ok: true, found: true, name: found[0].name,
           enrollments: found.map(function (x) { return { accreditation: x.acc, group: x.group, module: x.module, doctor: x.doctor, moduleEnds: x.moduleEnds, active: x.active }; }) };
}

/* =========================== Lecture / doctor evaluation (evaluation.html) =========================== */
// Tab "Evaluations" (created automatically): one evaluation per phone + programme + module run.
const EV = { SHEET: 'Evaluations', HEAD: ['Timestamp', 'Phone', 'Student', 'Accreditation', 'Group', 'Module', 'Doctor',
             'Doctor rating (1-5)', 'Lecture rating (1-5)', 'Suggestions', 'Module ends', 'Key', 'Submission ID', 'Material rating (1-5)', 'Interaction rating (1-5)'], KEY: 12, SUB: 13 };

function evalSheet_() {
  let sh = sheet_(EV.SHEET);
  if (!sh) {
    sh = SpreadsheetApp.getActive().insertSheet(EV.SHEET);
    sh.getRange(1, 1, 1, EV.HEAD.length).setValues([EV.HEAD]).setFontWeight('bold').setBackground('#0c2536').setFontColor('#ffffff').setWrap(true);
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm');
    [150, 120, 160, 130, 170, 260, 190, 90, 90, 380, 100, 160, 120].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.setTabColor('#7b5cd6');
  }
  if (String(sh.getRange(1, 14).getValue()).trim() !== EV.HEAD[13]) {   // v8.9: material + interaction questions (older rows stay empty there)
    sh.getRange(1, 10).setValue(EV.HEAD[9]);
    sh.getRange(1, 14, 1, 2).setValues([[EV.HEAD[13], EV.HEAD[14]]]).setFontWeight('bold').setBackground('#0c2536').setFontColor('#ffffff').setWrap(true);
    sh.setColumnWidth(14, 110); sh.setColumnWidth(15, 110);
  }
  return sh;
}
function evalKey_(phone, en) { return phoneKey_(phone) + '|' + en.acc + '|' + en.module + '|' + en.moduleEnds; }
function evalDone_(keys) {
  const sh = sheet_(EV.SHEET), out = {};
  if (!sh) return out;
  const last = lastDataRow_(sh, 1);
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, EV.SUB).getValues().forEach(function (r) {
    const k = String(r[EV.KEY - 1]);
    if (keys.indexOf(k) !== -1) out[k] = { at: r[0], sub: String(r[EV.SUB - 1]) };
  });
  return out;
}

/** Phone → registered? which module/doctor can be evaluated now, and was it already evaluated? */
function evalCheck_(rawPhone) {
  const phone = normalizePhone_(rawPhone);
  if (!/^\+?\d{8,15}$/.test(phone)) return { ok: false, error: 'Invalid phone number' };
  let found = findStudent_(phone);
  if (!found.length) return { ok: true, found: false, contacts: readContacts_().map(function (c) { return { label: c.label, whatsapp: c.whatsapp }; }) };
  if (found.some(function (x) { return x.active; })) found = found.filter(function (x) { return x.active; });
  const keys = found.map(function (en) { return evalKey_(phone, en); });
  const done = evalDone_(keys);
  return { ok: true, found: true, name: found[0].name,
           enrollments: found.map(function (en, i) {
             const d = done[keys[i]];
             return { accreditation: en.acc, group: en.group, module: en.module, doctor: en.doctor, moduleEnds: en.moduleEnds,
                      evaluated: !!d, evaluatedAt: d ? Utilities.formatDate(new Date(d.at), SETTINGS.TIMEZONE, 'yyyy-MM-dd') : '' };
           }) };
}

/** Saves one evaluation. The server decides the module + doctor; a second evaluation of the same module run is refused. */
function submitEvaluation_(data) {
  const phone = normalizePhone_(data.phone);
  if (!/^\+?\d{8,15}$/.test(phone)) return { ok: false, error: 'Invalid phone number' };
  const dr = Number(data.doctorRating), lr = Number(data.lectureRating), mr = Number(data.materialRating), ir = Number(data.interactionRating);
  if ([dr, lr, mr, ir].some(function (x) { return !(x >= 1 && x <= 5 && Math.round(x) === x); })) return { ok: false, error: 'Ratings must be 1-5' };
  const comment = String(data.comment || '').replace(/\s+\n/g, '\n').trim().slice(0, 1500);
  let found = findStudent_(phone);
  if (!found.length) return { ok: false, error: 'not registered', code: 'NOT_FOUND' };
  if (found.some(function (x) { return x.active; })) found = found.filter(function (x) { return x.active; });
  const en = found.filter(function (x) { return x.acc === String(data.accreditation || '').trim(); })[0] || found[0];
  if (!en.module) return { ok: false, error: 'No module running now', code: 'NO_MODULE' };
  const key = evalKey_(phone, en), sub = String(data.submissionId || Utilities.getUuid()).slice(0, 64);
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return { ok: false, error: 'Server busy, please try again' }; }
  try {
    const prev = evalDone_([key])[key];
    if (prev) return prev.sub === sub ? { ok: true, repeat: true, module: en.module, doctor: en.doctor }
                                      : { ok: false, error: 'already evaluated', code: 'DUP', evaluatedAt: Utilities.formatDate(new Date(prev.at), SETTINGS.TIMEZONE, 'yyyy-MM-dd') };
    const sh = evalSheet_();
    const studentName = studentFullName_(phone, en.acc) || en.name;
    const row = [new Date(), phone, studentName, en.acc, en.group, en.module, en.doctor || '', dr, lr, comment, en.moduleEnds, key, sub, mr, ir];
    const r = lastDataRow_(sh, 1) + 1;
    sh.getRange(r, 2).setNumberFormat('@');
    sh.getRange(r, 1, 1, row.length).setValues([row]);
    return { ok: true, module: en.module, doctor: en.doctor, accreditation: en.acc };
  } finally { lock.releaseLock(); }
}
function studentFullName_(phone, acc) {
  const st = sheet_(T.STUDENTS), key = phoneKey_(phone), last = lastDataRow_(st, SC.PHONE);
  if (last < 2) return '';
  const r = st.getRange(2, 1, last - 1, SC.ACC).getValues().filter(function (x) { return String(x[SC.KEY - 1]) === key && String(x[SC.ACC - 1]) === acc; })[0];
  return r ? String(r[SC.NAME - 1]) : '';
}

/** Dashboard › Evaluations block (details: Doctors Report tab). */
function evalDashboard_() {
  evalSheet_();   // makes sure the Evaluations tab exists
  const db = sheet_(T.DASH);
  if (!db) return;
  const b = 72, mk = 'TEXT(TODAY(),"yyyy-mm")';
  db.getRange(b, 15).setValue('Evaluations').setFontSize(13).setFontWeight('bold').setFontColor('#0c2536');
  const mets = [
    ['Evaluations this month', '=SUMPRODUCT((TEXT(Evaluations!A2:A,"yyyy-mm")=' + mk + ')*(Evaluations!A2:A<>""))'],
    ['Avg doctor rating (this month)', '=IFERROR(ROUND(AVERAGEIFS(Evaluations!H:H,Evaluations!A:A,">="&EOMONTH(TODAY(),-1)+1),2),"—")'],
    ['Avg lecture rating (this month)', '=IFERROR(ROUND(AVERAGEIFS(Evaluations!I:I,Evaluations!A:A,">="&EOMONTH(TODAY(),-1)+1),2),"—")'],
    ['Low ratings (≤2) this month', '=COUNTIFS(Evaluations!A:A,">="&EOMONTH(TODAY(),-1)+1,Evaluations!H:H,"<=2")+COUNTIFS(Evaluations!A:A,">="&EOMONTH(TODAY(),-1)+1,Evaluations!H:H,">2",Evaluations!I:I,"<=2")']
  ];
  mets.forEach(function (m, i) { db.getRange(b + 1 + i, 15).setValue(m[0]); db.getRange(b + 1 + i, 16).setFormula(m[1]).setFontWeight('bold'); });
  db.getRange(b + 4, 16).setFontColor('#a52f2f');
  db.getRange(b + 5, 15, 1, 2).clearContent();   // was "Lowest-rated doctor" (old Evaluation Report tab)
  db.getRange(b + 6, 15).setValue('Details: Doctors Report tab').setFontStyle('italic').setFontSize(9).setFontColor('#6b7f8e');
}

/* =========================== Academic path (evaluation.html › second tab) =========================== */
// "Academic Path Services": prices + who can see each service (editable by the team).
// "Academic Path Requests": one row per student request, Status followed by the team.
const AP = {
  SVC: 'Academic Path Services', REQ: 'Academic Path Requests', REP: 'Academic Path Report',
  SVC_HEAD: ['ID', 'Active (Yes/No)', 'Order', 'Title (Arabic)', 'Title (English)', 'Description (Arabic)', 'Description (English)',
             'Price (USD)', 'Old price (USD)', 'Badge (Arabic)', 'Badge (English)', 'Who can see it (accreditations, comma separated)',
             'Options – one per line:  Title :: details :: price   ({MODULES} = the student\'s modules)', 'Opportunity (Sale-up sheet)'],
  REQ_HEAD: ['Timestamp', 'Phone', 'Student', 'Accreditation', 'Group', 'Service ID', 'Service', 'Option', 'Price (USD)', 'Student note',
             'Status', 'Status updated', 'Follow-up WhatsApp', 'Team notes', 'Key', 'Submission ID'],
  STATUSES: ['New', 'Contacted', 'Enrolled', 'Not interested'],
  R: { STATUS: 11, STATUS_AT: 12, KEY: 15, SUB: 16 }
};
const EIU_SPECS = 'General Management :: International Business Environment + Corporate Communication\n' +
  'Finance Management :: Financial Institutions & Managerial Economics + Research Methodology or Entrepreneurship\n' +
  'Entrepreneurship & Innovation Management :: Research for Strategic Development + Entrepreneurship\n' +
  'Marketing Management :: Advanced Marketing & International Marketing + Corporate Communication\n' +
  'Human Resource Management :: Organizational Behavior & Business Ethics + International Business Environment\n' +
  'Supply Chain Management :: Supply Chain Economics & Operations Strategy + Project Quality Management';
const KHDA_PROGRAMS =
  'Contemporary Strategic Management Program :: HR · Finance · Leadership · Marketing · Strategic Management – 120 hours\n' +
  'Executive Director Program :: HR · Finance · Leadership · Marketing · Strategic Management – 120 hours\n' +
  'Leadership Development Program :: HR · Finance · Marketing · Leadership – 100 hours\n' +
  'Project Management for the 21st Century :: HR · Finance · Leadership · Marketing · Project Management – 120 hours\n' +
  'Regional Director Program :: HR · Finance · Leadership · Marketing · International Business – 120 hours\n' +
  'Entrepreneurship Program :: HR · Finance · Marketing · Leadership · Entrepreneurship – 120 hours\n' +
  'Chairman Program :: HR · Finance · Leadership · Marketing · Strategic Management – 120 hours\n' +
  'Area Manager Program :: HR · Finance · Leadership · Marketing · International Business – 120 hours\n' +
  'International Business Environment Program :: HR · Finance · Leadership · Marketing · International Business – 120 hours\n' +
  'Training Program for Top Managers :: HR · Finance · Leadership · Marketing · Strategic Management · Project Management – 120 hours\n' +
  'Finance for Managers :: Financial Management · Accounting Management · Economics · Strategic Management – 100 hours :: 550';
const AP_DEFAULT = [
  ['DM-EIU', 'Yes', 1, 'تخصص إضافي (Double Major)', 'Double Major', 'شهادة تانية من نفس جامعتك EIU بتخصص مختلف عن تخصصك الأساسي، من غير ما تدرس مواد إضافية: بتقدّم رسالة في التخصص الجديد بس.', 'A second certificate from your own university (EIU) in a different specialisation – no extra modules, you only submit a dissertation in the new specialisation.', 1000, '', '', '', 'EIU MBA', '', ''],
  ['DM-AITU', 'Yes', 2, 'تخصص إضافي (Double Major)', 'Double Major', 'شهادة تانية من نفس جامعتك AITU بتخصص مختلف عن تخصصك الأساسي، من غير ما تدرس مواد إضافية: بتقدّم رسالة في التخصص الجديد بس.', 'A second certificate from your own university (AITU) in a different specialisation – no extra modules, you only submit a dissertation in the new specialisation.', 1000, '', '', '', 'AITU MBA', '', ''],
  ['DD-AITU-MBA', 'Yes', 3, 'شهادة مزدوجة (Double Degree) – ماجستير من AITU', 'Double Degree – Master\'s from AITU', 'إضافة لشهادتك من EIU، تقدر تحصل على شهادة ماجستير في نفس التخصص من جامعة AITU.', 'Alongside your EIU degree, receive a Master\'s in the same specialisation from AITU.', 1000, '', '', '', 'EIU MBA, EIU MBA Healthcare', '', 'Dual Degree'],
  ['DD-AITU-DBA', 'Yes', 4, 'شهادة مزدوجة (Double Degree) – دكتوراه من AITU', 'Double Degree – Doctorate from AITU', 'إضافة لشهادتك من EIU، تقدر تحصل على شهادة الدكتوراه في نفس التخصص من جامعة AITU.', 'Alongside your EIU degree, receive a doctorate in the same field from AITU.', 1600, '', '', '', 'EIU DBA', '', 'Dual Degree'],
  ['DD-EIU-MBA', 'Yes', 5, 'شهادة مزدوجة (Double Degree) – ماجستير من EIU', 'Double Degree – Master\'s from EIU', 'إضافة لشهادتك من AITU، تقدر تحصل على شهادة ماجستير في نفس التخصص من جامعة EIU.', 'Alongside your AITU degree, receive a Master\'s in the same specialisation from EIU.', 1000, '', '', '', 'AITU MBA', '', 'Dual Degree'],
  ['DD-EIU-DBA', 'Yes', 6, 'شهادة مزدوجة (Double Degree) – دكتوراه من EIU', 'Double Degree – Doctorate from EIU', 'إضافة لشهادتك من AITU، تقدر تحصل على شهادة الدكتوراه في نفس التخصص من جامعة EIU.', 'Alongside your AITU degree, receive a doctorate in the same field from EIU.', 2100, '', '', '', 'AITU DBA', '', 'Dual Degree'],
  ['DD-BBA', 'Yes', 7, 'شهادة مزدوجة (Double Degree) – بكالوريوس', 'Double Degree – Bachelor\'s', 'شهادة بكالوريوس تانية في نفس التخصص من جامعة شريكة.', 'A second bachelor\'s degree in the same field from a partner university.', 1000, '', '', '', 'EIU BBA 1 Year, EIU BBA 3 Years', '', 'Dual Degree'],
  ['ALT-THESIS', 'Yes', 8, 'مادتين بدل الرسالة', 'Two modules instead of the thesis', 'تقدر تستبدل رسالة الماجستير بمادتين دراسيتين حسب تخصصك.', 'Replace the MBA thesis with two taught modules in your specialisation.', 200, 340, 'منحة حالية', 'Current scholarship', 'EIU MBA', EIU_SPECS, ''],
  ['PRO-CERT', 'Yes', 9, 'شهادة احترافية في مادة (Professional Certificate)', 'Professional Certificate in a module', 'شهادة من EIU باسم المادة اللي درستها.', 'An EIU certificate carrying the name of a module you studied.', 300, '', '', '', 'EIU MBA, EIU MBA Healthcare, EIU BBA 1 Year, EIU BBA 3 Years, EIU DBA', '{MODULES}', ''],
  ['LEGAL', 'Yes', 10, 'توثيق الشهادة', 'Certificate attestation', 'توثيق شهادتك من دولة الجامعة ومن دولتك.', 'Attestation of your certificate in the university\'s country and in your own country.', 750, 1000, 'منحة لفترة محدودة', 'Limited-time scholarship', 'EIU MBA, EIU MBA Healthcare, EIU BBA 1 Year, EIU BBA 3 Years, EIU DBA, AITU MBA, AITU DBA', '', 'Legalization'],
  ['KHDA', 'Yes', 11, 'برامج إدارية معتمدة من KHDA', 'KHDA management programmes', 'شهادة لكل مادة + شهادة باسم البرنامج من هيئة المعرفة KHDA (حكومة دبي)، من غير دراسة إضافية لأنك درست المواد في برنامجك.', 'A certificate for each module plus a programme certificate from KHDA (Dubai Government) – no extra study, you already took the modules.', 650, '', '', '', 'EIU MBA, EIU MBA Healthcare, EIU DBA, AITU MBA, AITU DBA', KHDA_PROGRAMS, 'KHDA Certificate']
];

function apEnsure_() {
  const ss = SpreadsheetApp.getActive();
  let sv = sheet_(AP.SVC);
  if (!sv) {
    sv = ss.insertSheet(AP.SVC);
    sv.getRange(1, 1, 1, AP.SVC_HEAD.length).setValues([AP.SVC_HEAD]).setFontWeight('bold').setBackground('#0c2536').setFontColor('#ffffff').setWrap(true);
    sv.getRange(2, 1, AP_DEFAULT.length, AP.SVC_HEAD.length).setValues(AP_DEFAULT).setWrap(true).setVerticalAlignment('top');
    sv.setFrozenRows(1);
    [110, 70, 55, 230, 210, 300, 280, 80, 80, 130, 130, 260, 420].forEach(function (w, i) { sv.setColumnWidth(i + 1, w); });
    sv.getRange('B2:B200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No']).setAllowInvalid(false).build());
    sv.setTabColor('#2e7d5b');
  }
  let rq = sheet_(AP.REQ);
  if (!rq) {
    rq = ss.insertSheet(AP.REQ);
    rq.getRange(1, 1, 1, AP.REQ_HEAD.length).setValues([AP.REQ_HEAD]).setFontWeight('bold').setBackground('#0c2536').setFontColor('#ffffff').setWrap(true);
    rq.setFrozenRows(1);
    rq.getRange('B:B').setNumberFormat('@'); rq.getRange('M:M').setNumberFormat('@');
    rq.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm'); rq.getRange('L:L').setNumberFormat('dd/mm/yyyy hh:mm');
    rq.getRange('K2:K5000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(AP.STATUSES).setAllowInvalid(false).build());
    [140, 120, 170, 130, 170, 100, 230, 260, 80, 260, 110, 130, 130, 220, 120, 110].forEach(function (w, i) { rq.setColumnWidth(i + 1, w); });
    rq.setTabColor('#2e7d5b');
  }
}
/** Updates rows that still hold an older default text (rows the team edited are left alone). */
function apMigrate_() {
  const sv = sheet_(AP.SVC);
  if (!sv || lastDataRow_(sv, 1) < 2) return;
  const n = lastDataRow_(sv, 1) - 1, v = sv.getRange(2, 1, n, AP.SVC_HEAD.length).getValues();
  const OLD = { 'DM-EIU': 'شهادة بتخصص تاني من نفس الجامعة EIU، برسالة في التخصص الجديد.', 'DM-AITU': 'رسالة إضافية في تخصص مختلف عن تخصصك الأساسي في AITU، وتحصل على شهادة بالتخصص الجديد.' };
  v.forEach(function (r, i) {
    const id = String(r[0]).trim();
    if (!OLD[id] || String(r[5]).trim() !== OLD[id]) return;
    const d = AP_DEFAULT.filter(function (x) { return x[0] === id; })[0];
    sv.getRange(i + 2, 4, 1, 4).setValues([[d[3], d[4], d[5], d[6]]]);
    if (String(r[12]).trim() === EIU_SPECS.trim()) sv.getRange(i + 2, 13).setValue('');
  });
}
function apServices_() {
  const sv = sheet_(AP.SVC);
  const rows = sv && lastDataRow_(sv, 1) >= 2 ? sv.getRange(2, 1, lastDataRow_(sv, 1) - 1, AP.SVC_HEAD.length).getValues() : AP_DEFAULT;
  return rows.filter(function (r) { return String(r[0]).trim() && String(r[1]).trim().toLowerCase() !== 'no'; })
    .map(function (r) {
      return { id: String(r[0]).trim(), order: Number(r[2]) || 99, title: { ar: String(r[3]), en: String(r[4] || r[3]) }, desc: { ar: String(r[5]), en: String(r[6] || r[5]) },
               price: Number(r[7]) || 0, oldPrice: Number(r[8]) || 0, badge: { ar: String(r[9] || ''), en: String(r[10] || r[9] || '') },
               accs: String(r[11] || '').split(',').map(function (x) { return x.trim(); }).filter(String), options: String(r[12] || '').trim(), opp: String(r[13] || '').trim() };
    }).sort(function (a, b) { return a.order - b.order; });
}
function apOptions_(svc, acc) {
  if (!svc.options) return [];
  if (svc.options === '{MODULES}') return (getCatalog_()[acc] || []).map(function (m) { return { title: m, details: '', price: 0 }; });
  return svc.options.split('\n').map(function (l) {
    const p = l.split('::').map(function (x) { return x.trim(); });
    return { title: p[0], details: p[1] || '', price: Number(p[2]) || 0 };
  }).filter(function (o) { return o.title; });
}
function apRequests_(key) {
  const rq = sheet_(AP.REQ), out = {};
  if (!rq) return out;
  const last = lastDataRow_(rq, 1);
  if (last < 2) return out;
  rq.getRange(2, 1, last - 1, AP.R.SUB).getValues().forEach(function (r) {
    const k = String(r[AP.R.KEY - 1]);
    if (k.indexOf(key + '|') === 0) out[k] = { at: r[0], sub: String(r[AP.R.SUB - 1]), option: String(r[7]), status: String(r[AP.R.STATUS - 1]) };
  });
  return out;
}

/* =========================== Personal development path – clients from the Sale-up sheet =========================== */
// Settings: B39 Sale-up sheet link · B40 tabs to take (comma separated) · B41 WhatsApp for these clients · B42 last sync
const DEV = { SHEET: 'Development Clients', WA_DEFAULT: '201041091902', URL_DEFAULT: 'https://docs.google.com/spreadsheets/d/1dMaV0kuY4OI3I1WQ9iuvKE3LmQH5ZozDWnk0szBiQWw/edit', TABS_DEFAULT: 'سلمي احمد, صابرين مصطفى',
  HEAD: ['Phone', 'Phone Key', 'Name', 'Email', 'Programmes', 'Accreditations', 'Available opportunities', 'Opportunity details',
         'Studying status', 'Certificate status', 'Employee', 'Programme keys', 'Source tab'] };
/** Opportunity names used in the Sale-up sheet → the extra rows added to Academic Path Services (the team sets prices and texts there). */
const DEV_OPPS = [
  ['GRAD', 'Graduation Ceremony', 'حفل التخرج', 'Graduation ceremony', 'احجز مكانك في حفل التخرج واستلم شهادتك في احتفال رسمي.', 'Book your place at the graduation ceremony and receive your certificate at an official event.'],
  ['WEBINAR', 'Webinar', 'ويبينار', 'Webinar', 'ويبينارات متخصصة مع خبراء في مجالك.', 'Specialised webinars with experts in your field.'],
  ['SINGLE-MBA', 'Single MBA', 'ماجستير إدارة الأعمال (MBA)', 'MBA', 'ابدأ ماجستير إدارة الأعمال من جامعة معتمدة.', 'Start an accredited MBA.'],
  ['SINGLE-DBA', 'Single DBA', 'دكتوراه إدارة الأعمال (DBA)', 'DBA', 'الخطوة الجاية بعد الماجستير: دكتوراه إدارة الأعمال من جامعة معتمدة.', 'The next step after your master\'s: an accredited DBA.'],
  ['PHD', 'PHD', 'دكتوراه الفلسفة (PhD)', 'PhD', 'برنامج دكتوراه الفلسفة في مجالك.', 'A PhD programme in your field.'],
  ['SINGLE-COURSE', 'Single Course', 'كورس احترافي', 'Professional course', 'كورسات قصيرة معتمدة تطوّر مهاراتك في شغلك.', 'Short accredited courses that build skills for your work.']
];
const DEV_TAGS = { 'LEGAL': 'Legalization', 'KHDA': 'KHDA Certificate', 'DD-AITU-MBA': 'Dual Degree', 'DD-AITU-DBA': 'Dual Degree', 'DD-EIU-MBA': 'Dual Degree', 'DD-EIU-DBA': 'Dual Degree', 'DD-BBA': 'Dual Degree' };

function devSettings_() {
  const set = sheet_(T.SETTINGS);
  if (!String(set.getRange('A39').getValue()).trim()) {
    set.getRange('A39:A42').setValues([['Sale-up sheet link (development path)'], ['Sale-up tabs to take'], ['Development path WhatsApp'], ['Last sync of development clients']]).setFontWeight('bold');
    set.getRange('B40').setValue(DEV.TABS_DEFAULT); set.getRange('B41').setNumberFormat('@').setValue(DEV.WA_DEFAULT);
    if (!String(set.getRange('B39').getValue()).trim()) set.getRange('B39').setValue(DEV.URL_DEFAULT);
    set.getRange('C39').setValue('Clients of these tabs see "مسار التطوير الشخصي" on the evaluation page. Copied into the Development Clients tab every hour.');
    set.getRange('C41').setValue('After a request the client is sent straight to this WhatsApp number.');
    set.getRange('B39:B41').setBackground('#fff4d6');
  }
  return { url: String(set.getRange('B39').getValue() || '').trim() || DEV.URL_DEFAULT, tabs: String(set.getRange('B40').getValue() || DEV.TABS_DEFAULT).split(',').map(function (x) { return x.trim(); }).filter(String),
           wa: String(set.getRange('B41').getValue() || DEV.WA_DEFAULT).replace(/\D/g, '') || DEV.WA_DEFAULT, set: set };
}
/** "Single MBA EIU | Double MBA EIU AITU | MBA in Healthcare management EIU" → ['EIU MBA', 'AITU MBA', 'EIU MBA Healthcare'] (used by "Who can see it"). */
function devProgrammeKeys_(programmes) {
  const out = [];
  String(programmes || '').split('|').forEach(function (p) {
    const s = p.toUpperCase(), lvl = /DBA/.test(s) ? 'DBA' : /MBA/.test(s) ? 'MBA' : /BBA/.test(s) ? 'BBA' : '';
    if (!lvl) return;
    ['EIU', 'AITU'].forEach(function (u) { if (s.indexOf(u) !== -1) out.push(u + ' ' + lvl + (lvl === 'MBA' && /HEALTH/.test(s) ? ' Healthcare' : lvl === 'BBA' ? ' 1 Year' : '')); });
  });
  return out.filter(function (x, i, a) { return a.indexOf(x) === i; });
}
/** One cell can hold two numbers glued together ("249912977246201102547889") or separated – returns each number. */
function splitPhones_(raw) {
  const out = [];
  String(raw == null ? '' : raw).split(/[\/|,;\n]|\s{2,}|\s-\s/).forEach(function (part) {
    const d = part.replace(/\D/g, '');
    if (d.length < 8) return;
    if (d.length <= 15) { out.push(d); return; }
    const CC = ['20', '966', '971', '965', '974', '968', '973', '962', '964', '249', '218', '216', '212', '44', '1'];
    for (let i = 10; i <= d.length - 10; i++) {
      if (i <= 13 && d.length - i <= 13 && CC.some(function (c) { return d.substr(i, c.length) === c; })) { out.push(d.slice(0, i), d.slice(i)); return; }
    }
    out.push(d.slice(0, 15));
  });
  return out.filter(function (x, i, a) { return a.indexOf(x) === i; });
}
/** Copies the chosen Sale-up tabs into "Development Clients" (a mirror – edit the Sale-up sheet, not this tab). */
function syncDevClients_() {
  const c = devSettings_();
  if (!c.url) return 'Sale-up link empty (Settings!B39)';
  const book = SpreadsheetApp.openByUrl(c.url), out = [], seen = {};
  const norm = function (x) { return String(x).replace(/\s+/g, ' ').trim().replace(/ى/g, 'ي').replace(/ة/g, 'ه').replace(/[أإآ]/g, 'ا'); };
  const all = book.getSheets(), missing = [];
  c.tabs.forEach(function (name) {
    const sh = all.filter(function (x) { return norm(x.getName()) === norm(name); })[0];
    if (!sh) { missing.push(name); return; }
    const v = sh.getDataRange().getValues(); if (v.length < 2) return;
    const h = v[0].map(function (x) { return String(x).toLowerCase(); });
    const col = function (re) { return h.findIndex(function (x) { return re.test(x); }); };
    const C = { opp: col(/available opportunities/), mob: col(/^mob/), ar: col(/name in arabic/), en: col(/name in english/), mail: col(/e-?mail/), prog: col(/^programs/),
                det: col(/opportunity details/), stat: col(/studying status/), cert: col(/certificate status/), acc: col(/accreditations/), emp: col(/assigned employee/) };
    const g = function (r, i) { return i >= 0 ? String(r[i] === null || r[i] === undefined ? '' : r[i]).trim() : ''; };
    for (let r = 1; r < v.length; r++) {
      const nums = splitPhones_(g(v[r], C.mob)), keys = nums.map(function (n) { return n.slice(-9); });
      if (!keys.length || keys.some(function (k) { return seen[k]; })) continue;
      keys.forEach(function (k) { seen[k] = true; });
      const phone = nums.join(' / '), key = keys.join(', ');
      out.push([phone, key, g(v[r], C.en) || g(v[r], C.ar), g(v[r], C.mail), g(v[r], C.prog), g(v[r], C.acc), g(v[r], C.opp), g(v[r], C.det),
                g(v[r], C.stat), g(v[r], C.cert), g(v[r], C.emp), devProgrammeKeys_(g(v[r], C.prog)).join(', '), name]);
    }
  });
  let dc = sheet_(DEV.SHEET);
  if (!dc) { dc = SpreadsheetApp.getActive().insertSheet(DEV.SHEET); dc.setTabColor('#d99a1e'); }
  dc.clear();
  dc.getRange(1, 1, 1, DEV.HEAD.length).setValues([DEV.HEAD]).setFontWeight('bold').setBackground('#0c2536').setFontColor('#ffffff');
  dc.getRange('A:B').setNumberFormat('@');
  if (out.length) dc.getRange(2, 1, out.length, DEV.HEAD.length).setValues(out);
  dc.setFrozenRows(1);
  [130, 100, 230, 210, 260, 160, 360, 360, 130, 260, 120, 180, 120].forEach(function (w, i) { dc.setColumnWidth(i + 1, w); });
  const msg = 'clients ' + out.length + ' from ' + c.tabs.join(' + ') + (missing.length ? ' · tab not found: ' + missing.join(', ') : '');
  c.set.getRange('B42').setValue(Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, 'dd/MM/yyyy HH:mm') + ' · ' + msg);
  return msg;
}
function devClient_(phone) {
  const dc = sheet_(DEV.SHEET), key = phoneKey_(phone);
  if (!dc || !key) return null;
  const last = lastDataRow_(dc, 1); if (last < 2) return null;
  const r = dc.getRange(2, 1, last - 1, DEV.HEAD.length).getValues().filter(function (x) { return String(x[1]).split(',').map(function (k) { return k.trim(); }).indexOf(key) !== -1; })[0];
  if (!r) return null;
  return { name: maskName_(r[2]), fullName: String(r[2]), programmes: String(r[4]), accreditations: String(r[5]),
           opps: String(r[6]).split('|').map(function (x) { return x.trim().toLowerCase(); }).filter(String),
           keys: String(r[11]).split(',').map(function (x) { return x.trim(); }).filter(String) };
}
/** Services the client may see: the row's opportunity is one of the client's, and "Who can see it" is empty or matches one of the client's programmes. */
function devServices_(cl) {
  return apServices_().filter(function (s) {
    return s.opp && cl.opps.indexOf(s.opp.toLowerCase()) !== -1 && (!s.accs.length || s.accs.some(function (a) { return cl.keys.indexOf(a) !== -1; }));
  });
}
/** Services tab: column N "Opportunity (Sale-up sheet)" + one row for every Sale-up opportunity that has none yet. */
function devEnsureServices_() {
  const sv = sheet_(AP.SVC); if (!sv) return;
  if (String(sv.getRange(1, 14).getValue()).trim() !== AP.SVC_HEAD[13]) {
    sv.getRange(1, 14).setValue(AP.SVC_HEAD[13]).setFontWeight('bold').setBackground('#0c2536').setFontColor('#ffffff').setWrap(true);
    sv.setColumnWidth(14, 190);
    const last = lastDataRow_(sv, 1);
    if (last >= 2) { const v = sv.getRange(2, 1, last - 1, 14).getValues(); v.forEach(function (r) { if (!String(r[13]).trim() && DEV_TAGS[String(r[0]).trim()]) r[13] = DEV_TAGS[String(r[0]).trim()]; }); sv.getRange(2, 1, v.length, 14).setValues(v); }
  }
  const last = lastDataRow_(sv, 1), v = last >= 2 ? sv.getRange(2, 1, last - 1, 14).getValues() : [];
  const ids = v.map(function (r) { return String(r[0]).trim(); }), tags = v.map(function (r) { return String(r[13]).trim().toLowerCase(); });
  let order = v.reduce(function (m, r) { return Math.max(m, Number(r[2]) || 0); }, 0);
  const add = DEV_OPPS.filter(function (o) { return ids.indexOf(o[0]) === -1 && tags.indexOf(o[1].toLowerCase()) === -1; })
    .map(function (o) { return [o[0], 'Yes', ++order, o[2], o[3], o[4], o[5], '', '', '', '', '', '', o[1]]; });
  if (add.length) sv.getRange(last + 1, 1, add.length, 14).setValues(add).setWrap(true).setVerticalAlignment('top');
}

/** Phone → registered? + the services each of the student's programmes can see. */
/** Works even before setup was run again: adds the opportunity column / rows and copies the clients on first use. */
function devReady_() {
  const sv = sheet_(AP.SVC);
  if (!sv || String(sv.getRange(1, 14).getValue()).trim() !== AP.SVC_HEAD[13]) { apEnsure_(); devEnsureServices_(); }
  const dc = sheet_(DEV.SHEET);
  if (!dc || lastDataRow_(dc, 1) < 2) { try { syncDevClients_(); } catch (e) { Logger.log('development clients: ' + e); } }
}
function pathCheck_(rawPhone) {
  const phone = normalizePhone_(rawPhone);
  if (!/^\+?\d{8,15}$/.test(phone)) return { ok: false, error: 'Invalid phone number' };
  devReady_();
  const c = devSettings_(), cl = devClient_(phone);
  if (!cl) return { ok: true, found: false, whatsapp: c.wa };
  const pk = phoneKey_(phone), done = apRequests_(pk), tz = SETTINGS.TIMEZONE, acc0 = cl.keys[0] || '';
  return { ok: true, found: true, name: cl.name,
    enrollments: [{ accreditation: cl.accreditations, group: cl.programmes, whatsapp: c.wa,
      services: devServices_(cl).map(function (s) {
        const req = Object.keys(done).filter(function (k) { return k.indexOf(pk + '|DEV|' + s.id + '|') === 0; })
                      .map(function (k) { return { option: done[k].option, at: Utilities.formatDate(new Date(done[k].at), tz, 'yyyy-MM-dd') }; });
        // prices stay in the sheet only – the page never receives them
        return { id: s.id, title: s.title, desc: s.desc, price: 0, oldPrice: 0, badge: s.badge,
                 options: apOptions_(s, acc0).map(function (o) { return { title: o.title, details: o.details, price: 0 }; }), requested: req };
      }) }] };
}

/** Saves "I'm interested" for one service (+ option). The same service/option is saved once per student. */
function submitPath_(data) {
  const phone = normalizePhone_(data.phone);
  if (!/^\+?\d{8,15}$/.test(phone)) return { ok: false, error: 'Invalid phone number' };
  devReady_();
  const c = devSettings_(), cl = devClient_(phone);
  if (!cl) return { ok: false, error: 'not registered', code: 'NOT_FOUND', whatsapp: c.wa };
  const svc = devServices_(cl).filter(function (s) { return s.id === String(data.serviceId || ''); })[0];
  if (!svc) return { ok: false, error: 'Service not available', code: 'NOT_ELIGIBLE' };
  const opts = apOptions_(svc, cl.keys[0] || '');
  let opt = null;
  if (opts.length) {
    opt = opts.filter(function (o) { return o.title === String(data.option || ''); })[0];
    if (!opt) return { ok: false, error: 'Please choose an option', code: 'OPTION' };
  }
  const price = opt && opt.price ? opt.price : svc.price;
  const note = String(data.note || '').replace(/\s+\n/g, '\n').trim().slice(0, 1000);
  const key = phoneKey_(phone) + '|DEV|' + svc.id + '|' + (opt ? opt.title : '');
  const sub = String(data.submissionId || Utilities.getUuid()).slice(0, 64);
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return { ok: false, error: 'Server busy, please try again' }; }
  try {
    apEnsure_();
    const prev = apRequests_(phoneKey_(phone))[key];
    if (prev) return { ok: true, already: true, whatsapp: c.wa, at: Utilities.formatDate(new Date(prev.at), SETTINGS.TIMEZONE, 'yyyy-MM-dd') };
    const rq = sheet_(AP.REQ), r = lastDataRow_(rq, 1) + 1;
    rq.getRange(r, 2).setNumberFormat('@'); rq.getRange(r, 13).setNumberFormat('@');
    rq.getRange(r, 1, 1, AP.REQ_HEAD.length).setValues([[new Date(), phone, cl.fullName, cl.accreditations, cl.programmes, svc.id, svc.title.en,
      opt ? opt.title : '', price || '', note, 'New', '', c.wa, '', key, sub]]);
    return { ok: true, whatsapp: c.wa, service: svc.title };
  } finally { lock.releaseLock(); }
}

/** "Academic Path Report" tab (rebuilt by setup). */
/** "Academic Path Report" (refreshed every 10 minutes): the month in numbers, requests waiting for the team, results per service / group, and the monthly history. */
function pathReportTab_() {
  const tz = SETTINGS.TIMEZONE, W = 9, rows = [], sh = sheet_(AP.REQ);
  const row = function (v, k, cc, color) { rows.push([v, k, color ? cc : 0, color]); };
  const last = sh ? lastDataRow_(sh, 1) : 1;
  const req = (last >= 2 ? sh.getRange(2, 1, last - 1, 14).getValues() : []).map(function (r) {
    const d = r[0] instanceof Date ? r[0] : toDate_(r[0]);
    return d ? { d: d, day: dayStart_(d), mon: Utilities.formatDate(d, tz, 'yyyy-MM'), phone: String(r[1]), student: String(r[2]), group: String(r[4] || ''),
                 service: String(r[6] || r[5] || ''), option: String(r[7] || ''), price: Number(r[8]) || 0, status: String(r[10] || 'New').trim() || 'New',
                 at: r[11] instanceof Date ? r[11] : toDate_(r[11]), wa: String(r[12] || ''), note: String(r[13] || '') } : null;
  }).filter(Boolean);
  const today = dayStart_(new Date()), thisM = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  const pm = new Date(); pm.setDate(1); pm.setMonth(pm.getMonth() - 1); const prevM = Utilities.formatDate(pm, tz, 'yyyy-MM');
  const money = function (n) { return '$' + Math.round(n).toLocaleString('en-US'); };
  const sum = function (l) { return l.reduce(function (a, x) { return a + x.price; }, 0); };
  const conv = function (l) { const c = l.filter(function (x) { return x.status === 'Enrolled' || x.status === 'Not interested'; }).length; return c ? Math.round(l.filter(function (x) { return x.status === 'Enrolled'; }).length / c * 100) + '%' : '—'; };
  const by = function (l, st) { return l.filter(function (x) { return x.status === st; }); };
  row(['Development Path Report (مسار التطوير الشخصي)'], 'title');
  row(['Updated ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm') + ' · from the Academic Path Requests tab (Status column) · conversion = Enrolled ÷ (Enrolled + Not interested)'], 'note');
  // 1 · this month
  const tm = req.filter(function (x) { return x.mon === thisM; }), lm = req.filter(function (x) { return x.mon === prevM; });
  row([], 'gap'); row(['1 · This month in numbers'], 'section'); const cP1 = rows.length;
  row(['Metric', 'This month', 'Last month', 'All time'], 'head');
  [['Requests', tm.length, lm.length, req.length],
   ['Still New (not contacted)', by(tm, 'New').length, by(lm, 'New').length, by(req, 'New').length],
   ['Contacted', by(tm, 'Contacted').length, by(lm, 'Contacted').length, by(req, 'Contacted').length],
   ['Enrolled', by(tm, 'Enrolled').length, by(lm, 'Enrolled').length, by(req, 'Enrolled').length],
   ['Not interested', by(tm, 'Not interested').length, by(lm, 'Not interested').length, by(req, 'Not interested').length],
   ['Conversion', conv(tm), conv(lm), conv(req)],
   ['Value enrolled (USD)', money(sum(by(tm, 'Enrolled'))), money(sum(by(lm, 'Enrolled'))), money(sum(by(req, 'Enrolled')))],
   ['Value still open – New + Contacted (USD)', money(sum(by(tm, 'New').concat(by(tm, 'Contacted')))), money(sum(by(lm, 'New').concat(by(lm, 'Contacted')))), money(sum(by(req, 'New').concat(by(req, 'Contacted'))))]
  ].forEach(function (m) { row(m, 'Held'); });
  // 2 · waiting for the team
  const age = function (x) { return Math.floor((new Date() - (x.status === 'Contacted' && x.at ? x.at : x.d)) / 864e5); };
  const wait = req.filter(function (x) { return x.status === 'New' || (x.status === 'Contacted' && age(x) >= 7); })
    .sort(function (a, b) { return (a.status === 'New' ? 0 : 1) - (b.status === 'New' ? 0 : 1) || age(b) - age(a); });
  row([], 'gap'); row(['2 · Waiting for the team – New requests, and Contacted with no update for 7+ days (' + wait.length + ')'], 'section');
  row(['Request date', 'Student', 'Phone', 'Group', 'Service', 'Option', 'Price (USD)', 'Status', 'Waiting (days)'], 'head');
  if (!wait.length) row(['Nothing waiting.'], 'note');
  wait.forEach(function (x) { const a = age(x); row([shortDate_(x.d), x.student, x.phone, x.group, x.service, x.option || '—', x.price, x.status, a], x.status === 'New' ? 'Today' : 'Held', 9, a >= 3 ? '#a52f2f' : null); });
  // 3 · per service
  const svc = {}; req.forEach(function (x) { (svc[x.service] = svc[x.service] || []).push(x); });
  row([], 'gap'); row(['3 · Results per service (all time)'], 'section'); const cP3 = rows.length;
  row(['Service', 'Requests', 'New', 'Contacted', 'Enrolled', 'Not interested', 'Conversion', 'Value enrolled (USD)', 'Value open (USD)'], 'head');
  const sk = Object.keys(svc).sort(function (a, b) { return svc[b].length - svc[a].length; });
  if (!sk.length) row(['No requests yet.'], 'note');
  sk.forEach(function (k) { const l = svc[k]; row([k, l.length, by(l, 'New').length, by(l, 'Contacted').length, by(l, 'Enrolled').length, by(l, 'Not interested').length, conv(l), money(sum(by(l, 'Enrolled'))), money(sum(by(l, 'New').concat(by(l, 'Contacted'))))], 'Held'); });
  // 4 · per group (share of active students who asked)
  const act = {}; const dcs = sheet_(DEV.SHEET), dl = dcs ? lastDataRow_(dcs, 1) : 1;   // clients in the development list, per programme
  if (dl >= 2) dcs.getRange(2, 5, dl - 1, 1).getValues().forEach(function (r) { act[r[0]] = (act[r[0]] || 0) + 1; });
  const grp = {}; req.forEach(function (x) { (grp[x.group] = grp[x.group] || []).push(x); });
  row([], 'gap'); row(['4 · Results per programme (all time)'], 'section');
  row(['Programme', 'Clients in the list', 'Clients who asked', '% who asked', 'Requests', 'Enrolled', 'Conversion', 'Value enrolled (USD)'], 'head');
  const gk = Object.keys(grp).sort(function (a, b) { return grp[b].length - grp[a].length; });
  if (!gk.length) row(['No requests yet.'], 'note');
  gk.forEach(function (g) { const l = grp[g], who = {}; l.forEach(function (x) { who[phoneKey_(x.phone)] = true; }); const n = Object.keys(who).length, a = act[g] || 0;
    row([g || '(no group)', a, n, a ? Math.round(n / a * 100) + '%' : '—', l.length, by(l, 'Enrolled').length, conv(l), money(sum(by(l, 'Enrolled')))], 'Held'); });
  // history · month by month
  const mon = {}; req.forEach(function (x) { (mon[x.mon] = mon[x.mon] || []).push(x); });
  row([], 'gap'); row(['History – month by month (newest first)'], 'histsec'); const cPH = rows.length;
  row(['Month', 'Requests', 'Enrolled', 'Not interested', 'Still open', 'Conversion', 'Value enrolled (USD)'], 'head');
  const mk = Object.keys(mon).sort().reverse();
  if (!mk.length) row(['No requests yet.'], 'note');
  mk.forEach(function (m) { const l = mon[m]; row([m, l.length, by(l, 'Enrolled').length, by(l, 'Not interested').length, by(l, 'New').length + by(l, 'Contacted').length, conv(l), money(sum(by(l, 'Enrolled')))], 'Finished'); });
  const STS = AP.STATUSES;
  writeReport_(AP.REP, '#2e7d5b', W, rows, [300, 200, 140, 190, 260, 200, 110, 140, 120], [
    { row: cP1, type: 'PIE', title: 'Requests by status (all time)', header: ['Status', 'Requests'], data: STS.map(function (st) { return [st, by(req, st).length]; }).filter(function (x) { return x[1]; }),
      colors: ['#d99a1e', '#1f8fb5', '#2e7d5b', '#c0392b'], h: 260 },
    { row: cP3, type: 'BAR', title: 'Requests per service – by status', header: ['Service'].concat(STS), stacked: true, colors: ['#d99a1e', '#1f8fb5', '#2e7d5b', '#c0392b'],
      data: sk.map(function (k) { return [k].concat(STS.map(function (st) { return by(svc[k], st).length; })); }), h: Math.max(240, 60 + sk.length * 34) },
    { row: cPH, type: 'COLUMN', title: 'Requests and enrolled – per month', header: ['Month', 'Requests', 'Enrolled'], colors: ['#174560', '#2e7d5b'],
      data: mk.slice().reverse().map(function (m) { return [m, mon[m].length, by(mon[m], 'Enrolled').length]; }), h: 280 }
  ]);
  return req.length;
}

function buildPathReport_() {
  apEnsure_();
  apMigrate_();
  pathReportTab_();
  const S = "'" + AP.REQ + "'";
  const db = sheet_(T.DASH);
  if (db) {
    const b = 80;
    db.getRange(b, 15).setValue('Academic path').setFontSize(13).setFontWeight('bold').setFontColor('#0c2536');
    const R = S + '!';
    [['Requests this month', '=COUNTIFS(' + R + 'A:A,">="&EOMONTH(TODAY(),-1)+1)'],
     ['New – not contacted yet', '=COUNTIF(' + R + 'K:K,"New")'],
     ['Enrolled (all time)', '=COUNTIF(' + R + 'K:K,"Enrolled")'],
     ['Value of requests this month (USD)', '=SUMIFS(' + R + 'I:I,' + R + 'A:A,">="&EOMONTH(TODAY(),-1)+1)']]
      .forEach(function (m, i) { db.getRange(b + 1 + i, 15).setValue(m[0]); db.getRange(b + 1 + i, 16).setFormula(m[1]).setFontWeight('bold'); });
    db.getRange(b + 2, 16).setFontColor('#a52f2f');
    db.getRange(b + 5, 15).setValue('Details: Academic Path Report tab').setFontStyle('italic').setFontSize(9).setFontColor('#6b7f8e');
  }
}

const DOC_OFFICIAL = ["DR. Hesham Shaarawy", "DR. Eslam Abd Elhameed", "DR. Ghassan Mohamed", "DR. Hany Eid", "DR. Hisham Abd Elhamid", "DR. Nahla Mohamed Saadeldeen", "DR. Nihal Fathi Shehab", "DR. Osama Hanafi", "DR. Pakinam Seif Eldeen", "DR. Aya Mohamed ElShahawy", "DR. Mary George", "DR. Olfat Mustafa", "DR. Ghada Elkady", "DR. Mona Hashim", "DR. Mohamed Samir", "DR. Nour Elakkad", "DR. Moustafa Alshenawy", "DR. Sara Faried", "DR. Nermin Salah El-Din", "DR. Bassel Shalaby", "DR. Mohamed Nageh Hassan", "DR. Wael Aly", "DR. Kareem Mousa", "DR. Hoda Hassan Aly", "DR. Samia Samy", "DR. Aisha Hamdy", "DR. Amira Elsokary", "DR. Youmna Eid", "DR. Hams Mohamed Elmahdy", "DR. Janit Maher", "DR. Mohammed Adel", "DR. Hesham Safwat"];
const DOC_EXTRA = [];   // only the official list is selectable; anything else stays as written (red mark) until a name is added to the Doctors tab
const DOC_ALIASES = [["hesham shaarawy", "DR. Hesham Shaarawy"], ["hesham sharawy", "DR. Hesham Shaarawy"], ["hisham sharawy", "DR. Hesham Shaarawy"], ["h sharawy", "DR. Hesham Shaarawy"], ["sharawy", "DR. Hesham Shaarawy"], ["shaarawy", "DR. Hesham Shaarawy"], ["eslam", "DR. Eslam Abd Elhameed"], ["eslam abd elhameed", "DR. Eslam Abd Elhameed"], ["eslam abdelhameed", "DR. Eslam Abd Elhameed"], ["ghassan", "DR. Ghassan Mohamed"], ["hany", "DR. Hany Eid"], ["hani", "DR. Hany Eid"], ["hisham abd", "DR. Hisham Abd Elhamid"], ["hesham abd", "DR. Hisham Abd Elhamid"], ["abd elhameed", "DR. Hisham Abd Elhamid"], ["abd elhamid", "DR. Hisham Abd Elhamid"], ["abdelhameed", "DR. Hisham Abd Elhamid"], ["abd-elhameed", "DR. Hisham Abd Elhamid"], ["abdelhamed", "DR. Hisham Abd Elhamid"], ["abd elhamed", "DR. Hisham Abd Elhamid"], ["nahla", "DR. Nahla Mohamed Saadeldeen"], ["nihal", "DR. Nihal Fathi Shehab"], ["osama", "DR. Osama Hanafi"], ["pakinam", "DR. Pakinam Seif Eldeen"], ["bakinam", "DR. Pakinam Seif Eldeen"], ["aya", "DR. Aya Mohamed ElShahawy"], ["mary", "DR. Mary George"], ["olfat", "DR. Olfat Mustafa"], ["olfaat", "DR. Olfat Mustafa"], ["olft", "DR. Olfat Mustafa"], ["ghada", "DR. Ghada Elkady"], ["mona", "DR. Mona Hashim"], ["mohamed samir", "DR. Mohamed Samir"], ["nour", "DR. Nour Elakkad"], ["shenawy", "DR. Moustafa Alshenawy"], ["alshenawy", "DR. Moustafa Alshenawy"], ["sara", "DR. Sara Faried"], ["nermin", "DR. Nermin Salah El-Din"], ["bassel", "DR. Bassel Shalaby"], ["nageh", "DR. Mohamed Nageh Hassan"], ["naggeh", "DR. Mohamed Nageh Hassan"], ["wael", "DR. Wael Aly"], ["kareem", "DR. Kareem Mousa"], ["hoda", "DR. Hoda Hassan Aly"], ["samia", "DR. Samia Samy"], ["aisha", "DR. Aisha Hamdy"], ["amira", "DR. Amira Elsokary"], ["youmna", "DR. Youmna Eid"], ["yomna", "DR. Youmna Eid"], ["hams", "DR. Hams Mohamed Elmahdy"], ["janit", "DR. Janit Maher"], ["mohammed adel", "DR. Mohammed Adel"], ["mohamed adel", "DR. Mohammed Adel"], ["hesham safwat", "DR. Hesham Safwat"], ["h safwat", "DR. Hesham Safwat"], ["safwat", "DR. Hesham Safwat"]];

/* =========================== Doctors: one official list everywhere =========================== */
let DOC_MANUAL_ = null;
/** Names the team typed into the Doctors tab themselves (not one of the official names) – kept and selectable. */
function docManual_() {
  if (DOC_MANUAL_) return DOC_MANUAL_;
  DOC_MANUAL_ = [];
  const dc = sheet_('Doctors'); if (!dc) return DOC_MANUAL_;
  const last = lastDataRow_(dc, 1); if (last < 2) return DOC_MANUAL_;
  const prev = PropertiesService.getScriptProperties().getProperty('DOC_LIST_V') === 'v6.4';
  dc.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
    const n = String(r[0] || '').trim();
    if (!n || DOC_OFFICIAL.indexOf(n) !== -1) return;
    DOC_MANUAL_.push(n);
  });
  if (!prev) DOC_MANUAL_ = [];   // first v6.4 run: the tab still holds the old list (Hesham Safwat, AI Grading) – start from the official list
  return DOC_MANUAL_;
}
function officialDoctor_(raw) {
  const s = String(raw || '').trim();
  if (!s) return '';
  if (DOC_OFFICIAL.indexOf(s) !== -1 || DOC_EXTRA.indexOf(s) !== -1 || docManual_().indexOf(s) !== -1) return s;
  return aliasDoctor_(s);
}
function aliasDoctor_(s) {
  const n = ' ' + s.toLowerCase().replace(/[.\/:\-]/g, ' ').replace(/\b(dr|doctor)\b/g, ' ').replace(/د\s/g, ' ').replace(/\s+/g, ' ').trim() + ' ';
  let best = null;
  DOC_ALIASES.forEach(function (a) {
    const al = a[0].replace(/-/g, ' ');
    if (n.indexOf(' ' + al + ' ') !== -1 || (al.length > 5 && n.indexOf(al) !== -1)) { if (!best || al.length > best[1]) best = [a[1], al.length]; }
  });
  return best ? best[0] : '';
}
/** Doctors tab = official list (+ data kept), Lists doctor dropdown, alias table, Group Schedule / Evaluations names. */
function unifyDoctors_() {
  DOC_MANUAL_ = null;
  const ss = SpreadsheetApp.getActive(), dc = sheet_('Doctors');
  const manual = docManual_().filter(function (n) { return !aliasDoctor_(n); });
  const all = DOC_OFFICIAL.concat(DOC_EXTRA).concat(manual);
  if (dc) {
    const last = lastDataRow_(dc, 1);
    const old = last >= 2 ? dc.getRange(2, 1, last - 1, 6).getValues() : [];
    const info = {}, spell = {};
    old.forEach(function (r) {
      const o = officialDoctor_(r[0]) || '';
      if (!o) return;
      spell[o] = (spell[o] || []).concat([String(r[0]), String(r[1] || '')].filter(String));
      if (!info[o]) info[o] = [r[3], r[4], r[5]];
    });
    const rows = all.map(function (n, i) {
      const sp = (spell[n] || []).join(', ').split(',').map(function (x) { return x.trim(); }).filter(function (x, j, a) { return x && x !== n && a.indexOf(x) === j; }).join(', ');
      return [n, sp, 'Yes'].concat(info[n] || ['', '', '']);
    });
    dc.getRange(2, 1, Math.max(last - 1, rows.length) + 5, 8).clearContent();
    dc.getRange(2, 1, rows.length, 6).setValues(rows);
    dc.getRange(2, 7, rows.length, 2).setFormulas(rows.map(function (r, i) {
      const k = i + 2;
      return ["=COUNTIFS('Group Schedule'!$I:$I,$A" + k + ",'Group Schedule'!$H:$H,\"Running\")", "=COUNTIF('Group Schedule'!$I:$I,$A" + k + ")"];
    }));
  }
  const ls = sheet_('Lists');
  if (ls) {
    ls.getRange(2, 6, 200, 1).clearContent();
    ls.getRange(2, 6, all.length + 1, 1).setValues([['All']].concat(all.map(function (n) { return [n]; })));
    // alias table I:J – every spelling → official (unknown old names keep their own name for history)
    const last = lastDataRow_(ls, 9);
    const cur = last >= 2 ? ls.getRange(2, 9, last - 1, 2).getValues() : [];
    const map = {};
    cur.forEach(function (r) { const k = String(r[0]).trim(); if (k) map[k] = officialDoctor_(k) || String(r[1] || k); });
    all.forEach(function (n) { map[n] = n; });
    const out = Object.keys(map).map(function (k) { return [k, map[k]]; });
    ls.getRange(2, 9, Math.max(cur.length, out.length) + 5, 2).clearContent();
    ls.getRange(2, 9, out.length, 2).setValues(out);
  }
  // Reports: doctor rows (lectures / assignments by doctor) use the official names
  const rp = sheet_('Reports');
  if (rp) {
    const n = Math.min(Math.max(rp.getLastRow(), 1), 600), a = rp.getRange(1, 1, n, 1).getValues();
    let sec = '', seen = {}, ch = false;
    a.forEach(function (r) {
      const v = String(r[0] || '');
      if (/^\d+ · /.test(v)) { sec = v; seen = {}; return; }
      if (!/by doctor/i.test(sec) || !v) return;
      const o = officialDoctor_(v);
      if (o && o !== v && !seen[o]) { r[0] = o; ch = true; }
      seen[r[0]] = true;
    });
    if (ch) rp.getRange(1, 1, n, 1).setValues(a);
  }
  [['Group Schedule', 9], [EV.SHEET, 7], [LEC.SHEET, 7], [AS.SHEET, AS.DOCTOR]].forEach(function (x) {
    const sh = sheet_(x[0]); if (!sh) return;
    const last = lastDataRow_(sh, 1); if (last < 2) return;
    const rg = sh.getRange(2, x[1], last - 1, 1), v = rg.getValues();
    let ch = false;
    v.forEach(function (r) { const o = officialDoctor_(r[0]); if (o && o !== r[0]) { r[0] = o; ch = true; } });
    if (ch) rg.setValues(v);
  });
  PropertiesService.getScriptProperties().setProperty('DOC_LIST_V', 'v6.4');
  DOC_MANUAL_ = null;
}

/* =========================== Lectures: the master schedule =========================== */
// Lectures tab: A Date | B Day* | C Time | D Group | E Module | F Lecture # | G Doctor | H Admin | I Zoom link | J Recording link
//               K Doctor attended? | L Reason | M New date | N Notes | O Recorded at | P Status* | Q Month key* | R Schedule ID | S Doctor key*
//               T Operation Action | U Ops admin | V Zoom account | W Material | X Assignment        (* = formula)
const LX = { WIDTH: 24, VALUE_COLS: [1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 18, 20, 21, 22, 23, 24],
             HEAD_EXTRA: ['Operation Action', 'Ops admin', 'Zoom account', 'Material', 'Assignment'] };
const GROUP_ACC = { 'EIU MBA · Group 1': 'EIU MBA', 'EIU MBA · Group 2': 'EIU MBA', 'EIU MBA · Group 3 New': 'EIU MBA', 'EIU MBA · Golden': 'EIU MBA',
                    'EIU MBA Healthcare': 'EIU MBA Healthcare', 'EIU BBA · Top Up': 'EIU BBA 1 Year', 'AITU MBA': 'AITU MBA', 'DBA (EIU / AITU)': 'EIU DBA' };
function lectureGroup_(label) {
  const raw = String(label || '').replace(/\s+/g, ' ').trim(), l = raw.toLowerCase();
  if (!l) return '';
  if (/^(group ?1|g\.? ?1)$/.test(l)) return 'EIU MBA · Group 1';
  if (/^(group ?2|g\.? ?2)$/.test(l)) return 'EIU MBA · Group 2';
  if (/^(group ?3( new)?|g\.? ?3)$/.test(l)) return 'EIU MBA · Group 3 New';
  if (/^golden/.test(l)) return 'EIU MBA · Golden';
  if (/^health ?care/.test(l)) return 'EIU MBA Healthcare';
  if (/^bba/.test(l)) return 'EIU BBA · Top Up';
  if (/^(aitu|atiu)/.test(l)) return 'AITU MBA';
  if (/^dba/.test(l)) return 'DBA (EIU / AITU)';
  if (/^(l7|qualifi)/.test(l)) return 'Qualifi L7';
  if (/^priv/.test(l)) { const rest = raw.replace(/^priv(ate)?[\s\-]*/i, '').trim(); return 'Private' + (rest ? ' · ' + rest : ''); }
  if (l === 'saged' || l === 'sagged') return 'Private · ' + raw;
  return raw;   // unknown label – kept as written, flagged by the import
}
function lectureModule_(raw, group, catalog) {
  const s = String(raw || '').replace(/\*/g, '').replace(/\s+/g, ' ').trim();
  if (!s) return '';
  const acc = GROUP_ACC[group];
  const tries = acc === 'EIU DBA' ? [catalog['EIU DBA'] || [], catalog['AITU DBA'] || []] : [catalog[acc] || []];
  const clean = s.replace(/^(DBA|Qualifi)\s*-\s*/i, '');
  for (let i = 0; i < tries.length; i++) { const m = catalogMatch_(tries[i], clean) || catalogMatch_(tries[i], s); if (m) return m; }
  return clean;
}
function cairoKey_(d) { return Utilities.formatDate(d, SETTINGS.TIMEZONE, 'yyyyMMdd'); }

/** One-time import from the old live lectures sheet (Settings!B33): fills the blank rows of each block, keeps dates ≥ from, merges attendance already recorded. */
function importOldLectures_(url, fromDate) {
  const res = readOldLectures_(url, fromDate), out = res.out, unknown = res.unknown;
  writeLectures_(out.map(function (o) { return o.row; }), true);
  return 'lectures ' + out.length + (Object.keys(unknown).length ? ' · rows without a group: ' + Object.keys(unknown).length : '');
}
function readOldLectures_(url, fromDate) {
  const book = SpreadsheetApp.openByUrl(url), src = book.getSheetByName('Lectures');
  if (!src) throw new Error('Tab "Lectures" not found in the old lectures sheet');
  const v = src.getDataRange().getValues(), catalog = getCatalog_();
  const from = cairoKey_(fromDate || new Date(2026, 9, 1));
  const out = [], unknown = {};
  let blk = null;
  for (let r = 1; r < v.length; r++) {
    const x = v[r], d = toDate_(x[0]);
    if (!d) continue;
    const label = String(x[3] || '').trim(), mod = String(x[5] || '').trim(), num = String(x[4] || '').trim();
    const startsBlock = !blk || (label && lectureGroup_(label) !== blk.group) || mod || num === '1' || (label && !blk.group);
    if (startsBlock) blk = { group: lectureGroup_(label) || (blk && !label && !mod ? blk.group : ''), module: '', doctor: '', admin: '', zoom: '', time: '', degree: '' };
    if (label) blk.group = lectureGroup_(label);
    if (mod) blk.module = lectureModule_(mod, blk.group, catalog);
    if (String(x[8] || '').trim()) blk.doctor = officialDoctor_(x[8]) || String(x[8]).trim();
    if (String(x[9] || '').trim()) blk.admin = String(x[9]).trim();
    if (/^https?:/.test(String(x[11] || '').trim())) blk.zoom = String(x[11]).trim();
    if (String(x[2] || '').trim()) blk.time = String(x[2]).trim();
    if (cairoKey_(d) < from) continue;
    const action = String(x[13] === true ? 'TRUE' : x[13] === false ? '' : x[13] || '').trim();
    let att = '', note = '';
    if (/^cancel/i.test(action)) { att = 'لم يحضر'; note = 'Cancel (من الشيت القديم)'; }
    else if (/^(done|move to drive|true)/i.test(action)) att = 'حضر';
    if (d.getTime() > Date.now() + 200 * 864e5) note = (note ? note + ' · ' : '') + '⚠ راجع التاريخ';
    const known = !!GROUP_ACC[blk.group] || /^(Qualifi L7|Private)/.test(blk.group);
    if (!known) { unknown[r + 1] = true; note = (note ? note + ' · ' : '') + '⚠ راجع الجروب (في الشيت القديم: ' + (label || 'فاضي') + ')'; }
    const drive = String(x[12] || '').trim();
    out.push({ d: d, row: [d, '', String(x[2] || '').trim() || blk.time, blk.group || '(check group)', blk.module, num, blk.doctor, '', blk.zoom,
               /^https?:/.test(drive) ? drive : '', att, '', '', note, '', '', '', '', '', action, String(x[10] || '').trim(), String(x[15] || '').trim(),
               String(x[16] || '').trim(), String(x[17] || '').trim()] });
  }
  out.sort(function (a, b) { return a.d - b.d || String(a.row[3]).localeCompare(String(b.row[3])); });
  return { out: out, unknown: unknown };
}

/** Live lectures sheet (Settings!B33) → Lectures tab, from today to the last row, exactly as written there.
 *  Merged cells count for every row they cover (that is how the sheet shows them). Nothing is filled in or guessed;
 *  only the names are written the way this sheet writes them (group, module, doctor). Lectures before today are left as they are. */
/** Rows of the live lectures sheet from today on, in Lectures-tab shape (merged cells counted for every row they cover). */
function readLiveRows_(url) {
  const book = SpreadsheetApp.openByUrl(url), src = book.getSheetByName('Lectures');
  if (!src) throw new Error('Tab "Lectures" not found in the live lectures sheet');
  const rg = src.getDataRange(), v = rg.getValues();
  rg.getMergedRanges().forEach(function (m) {
    const r0 = m.getRow() - 1, c0 = m.getColumn() - 1, val = (v[r0] || [])[c0];
    for (let r = r0; r < r0 + m.getNumRows() && r < v.length; r++) for (let c = c0; c < c0 + m.getNumColumns(); c++) v[r][c] = val;
  });
  const catalog = getCatalog_(), today = cairoKey_(new Date()), t = function (x) { return String(x === null || x === undefined ? '' : x).trim(); };
  const srcTz = (book.getSpreadsheetTimeZone && book.getSpreadsheetTimeZone()) || SETTINGS.TIMEZONE;
  // the live sheet keeps its dates in its own time zone: take the calendar day it shows there (otherwise the dates after
  // Egypt's winter-time change on 29/10 land one day early)
  const day = function (d) { return Utilities.parseDate(Utilities.formatDate(d, srcTz, 'yyyy-MM-dd'), SETTINGS.TIMEZONE, 'yyyy-MM-dd'); };
  const live = [];
  for (let r = 1; r < v.length; r++) {
    const x = v[r], d0 = toDate_(x[0]), d = d0 ? day(d0) : null;
    if (!d || cairoKey_(d) < today) continue;
    const group = lectureGroup_(x[3]), raw = t(x[5]).replace(/\s+/g, ' ');
    const num = typeof x[4] === 'number' ? x[4] : t(x[4]);
    const row = new Array(LX.WIDTH).fill('');
    row[0] = d; row[2] = t(x[2]).replace(/(AM|PM)\s*(?::|to)\s*/i, '$1 - '); row[3] = group; row[4] = raw ? lectureModule_(raw, group, catalog) : ''; row[5] = num;
    row[6] = officialDoctor_(x[8]) || t(x[8]);
    row[8] = /^https?:/.test(t(x[11])) ? t(x[11]) : ''; row[9] = /^https?:/.test(t(x[12])) ? t(x[12]) : '';
    row[19] = x[13] === true ? 'TRUE' : x[13] === false ? '' : t(x[13]);
    row[21] = t(x[15]); row[22] = t(x[16]); row[23] = t(x[17]);
    live.push(row);
  }
  return live;
}
function importLiveFromToday_(url) {
  const book = SpreadsheetApp.openByUrl(url), src = book.getSheetByName('Lectures');
  if (!src) throw new Error('Tab "Lectures" not found in the live lectures sheet');
  const rg = src.getDataRange(), v = rg.getValues();
  rg.getMergedRanges().forEach(function (m) {
    const r0 = m.getRow() - 1, c0 = m.getColumn() - 1, val = (v[r0] || [])[c0];
    for (let r = r0; r < r0 + m.getNumRows() && r < v.length; r++) for (let c = c0; c < c0 + m.getNumColumns(); c++) v[r][c] = val;
  });
  const catalog = getCatalog_(), today = cairoKey_(new Date()), t = function (x) { return String(x === null || x === undefined ? '' : x).trim(); };
  const srcTz = (book.getSpreadsheetTimeZone && book.getSpreadsheetTimeZone()) || SETTINGS.TIMEZONE;
  // the live sheet keeps its dates in its own time zone: take the calendar day it shows there (otherwise the dates after
  // Egypt's winter-time change on 29/10 land one day early)
  const day = function (d) { return Utilities.parseDate(Utilities.formatDate(d, srcTz, 'yyyy-MM-dd'), SETTINGS.TIMEZONE, 'yyyy-MM-dd'); };
  const live = [];
  for (let r = 1; r < v.length; r++) {
    const x = v[r], d0 = toDate_(x[0]), d = d0 ? day(d0) : null;
    if (!d || cairoKey_(d) < today) continue;
    const group = lectureGroup_(x[3]), raw = t(x[5]).replace(/\s+/g, ' ');
    const num = typeof x[4] === 'number' ? x[4] : t(x[4]);
    const row = new Array(LX.WIDTH).fill('');
    row[0] = d; row[2] = t(x[2]); row[3] = group; row[4] = raw ? lectureModule_(raw, group, catalog) : ''; row[5] = num;
    row[6] = officialDoctor_(x[8]) || t(x[8]);
    row[8] = /^https?:/.test(t(x[11])) ? t(x[11]) : ''; row[9] = /^https?:/.test(t(x[12])) ? t(x[12]) : '';
    row[19] = x[13] === true ? 'TRUE' : x[13] === false ? '' : t(x[13]);
    row[21] = t(x[15]); row[22] = t(x[16]); row[23] = t(x[17]);
    live.push(row);
  }
  const lw = sheet_(LEC.SHEET), last = lastDataRow_(lw, 1);
  const cur = last >= 2 ? lw.getRange(2, 1, last - 1, LX.WIDTH).getValues() : [];
  const past = cur.filter(function (r) { return r[0] instanceof Date && cairoKey_(r[0]) < today; })
                  .map(function (r) { const c = r.slice(); c[1] = ''; c[15] = ''; c[16] = ''; c[18] = ''; return c; });
  const rows = past.concat(live);
  rows.sort(function (a, b) { return a[0] - b[0] || String(a[3]).localeCompare(String(b[3])); });
  writeLectures_(rows, false);
  return 'from ' + Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, 'dd/MM/yyyy') + ': ' + live.length + ' lectures copied · lectures before today kept: ' + past.length;
}

/** One-time (v8.8): lectures the team added to the live sheet after the 28/10 import. Only rows that are not in the Lectures tab
 *  (same group + date + lecture #) are added – nothing already in the tab is changed. From now on the Lectures tab is the only schedule. */
function addMissingLiveRows_(url) {
  const live = readLiveRows_(url), lw = sheet_(LEC.SHEET), last = lastDataRow_(lw, 1);
  const cur = last >= 2 ? lw.getRange(2, 1, last - 1, LX.WIDTH).getValues().filter(function (r) { return r[0] instanceof Date; }) : [];
  const have = {}; cur.forEach(function (r) { have[r[3] + '|' + cairoKey_(r[0]) + '|' + String(r[5]).trim()] = true; });
  const add = live.filter(function (r) { const k = r[3] + '|' + cairoKey_(r[0]) + '|' + String(r[5]).trim(); if (have[k]) return false; have[k] = true; return true; });
  if (!add.length) return 'nothing new in the live sheet';
  const rows = cur.map(function (r) { const c = r.slice(); c[1] = ''; c[15] = ''; c[16] = ''; c[18] = ''; return c; }).concat(add);
  rows.sort(function (a, b) { return a[0] - b[0] || String(a[3]).localeCompare(String(b[3])); });
  writeLectures_(rows, false);
  add.sort(function (a, b) { return a[0] - b[0]; });
  return 'added ' + add.length + ' lectures from the live sheet (' + Utilities.formatDate(add[0][0], SETTINGS.TIMEZONE, 'dd/MM') + ' → ' + Utilities.formatDate(add[add.length - 1][0], SETTINGS.TIMEZONE, 'dd/MM/yyyy') + ')';
}

/** Replaces the Lectures values (formula columns untouched). merge=true keeps attendance already recorded in the tab (same day + group). */
function writeLectures_(rows, merge) {
  const lw = sheet_(LEC.SHEET);
  lw.getRange(1, 20, 1, LX.HEAD_EXTRA.length).setValues([LX.HEAD_EXTRA]).setFontWeight('bold').setBackground('#0c2536').setFontColor('#ffffff');
  const last = Math.max(lastDataRow_(lw, 1), 1);
  if (merge && last >= 2) {
    const cur = lw.getRange(2, 1, last - 1, 15).getValues(), rec = {};
    cur.forEach(function (r) {
      if (!(r[0] instanceof Date) || !(r[10] || r[11] || r[12] || r[13])) return;
      const k = cairoKey_(r[0]) + '|' + r[3];
      if (!rec[k]) rec[k] = r.slice(10, 15);
    });
    rows.forEach(function (r) { const k = cairoKey_(r[0]) + '|' + r[3]; if (rec[k]) { r[10] = rec[k][0]; r[11] = rec[k][1]; r[12] = rec[k][2]; r[13] = rec[k][3] || r[13]; r[14] = rec[k][4]; delete rec[k]; } });
  }
  const n = Math.max(last - 1, rows.length);
  // clear value columns only (B, P, Q, S keep their formulas)
  [[1, 1], [3, 13], [18, 1], [20, 5]].forEach(function (c) { lw.getRange(2, c[0], n, c[1]).clearContent(); });
  if (!rows.length) return;
  const col = function (a, b) { return rows.map(function (r) { return r.slice(a - 1, b); }); };
  lw.getRange(2, 1, rows.length, 1).setValues(col(1, 1));
  lw.getRange(2, 3, rows.length, 13).setValues(col(3, 15));
  lw.getRange(2, 18, rows.length, 1).setValues(col(18, 18));
  lw.getRange(2, 20, rows.length, 5).setValues(col(20, 24));
  ensureValidations_();
}

/** Lectures tab columns: H (Admin) is no longer used – emptied and hidden; J holds the Drive link of the recording. */
function lectureColumns_() {
  const lw = sheet_(LEC.SHEET); if (!lw) return;
  const last = lastDataRow_(lw, 1);
  if (last >= 2) lw.getRange(2, 8, last - 1, 1).clearContent();
  lw.getRange(1, 8).setValue('(not used)');
  lw.getRange(1, 9, 1, 2).setValues([['Zoom link', 'Drive link (recording)']]);
  lw.getRange(1, 10).setNote('بعد ما المحاضرة تتسجل وتترفع على الدرايف: حط لينك الدرايف هنا – بيظهر للطالب في صفحة "محاضرتي القادمة" تحت "تسجيلات المحاضرات".');
  lw.hideColumns(8);
}

/** Dropdowns set by the script (Google does not always keep the ones imported from Excel). Lists tab feeds them, so new reasons / agents can be added there. */
function ensureValidations_() {
  const ls = sheet_('Lists'), dc = sheet_('Doctors');
  const dv = function (rng, list, strict) {
    const b = SpreadsheetApp.newDataValidation();
    (Array.isArray(list) ? b.requireValueInList(list, true) : b.requireValueInRange(list, true)).setAllowInvalid(!strict);
    rng.setDataValidation(b.build());
  };
  const lw = sheet_(LEC.SHEET);
  if (lw) {
    const n = Math.max(lastDataRow_(lw, 1), 2) + 1000;
    dv(lw.getRange(2, 4, n, 1), Object.keys(GROUP_ACC).concat(['Qualifi L7 · Innovation', 'Qualifi L7 · Leadership', 'Qualifi L4 & L5', 'BBA 1st Year', 'BBA 2nd Year', 'Qualifi L7', 'Private']), false);
    if (dc) dv(lw.getRange(2, 7, n, 1), dc.getRange('A2:A200'), false);
    dv(lw.getRange(2, LEC.STATUS, n, 1), ['حضر', 'لم يحضر'], true);
    if (ls) dv(lw.getRange(2, LEC.REASON, n, 1), ls.getRange('M2:M40'), false);
    lw.getRange(2, LEC.NEWDATE, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());
  }
  const as = sheet_(AS.SHEET);
  if (as && ls) {
    const n = Math.max(lastDataRow_(as, 1), 2) + 3000;
    dv(as.getRange(2, AS.AGENT, n, 1), ls.getRange('A2:A100'), false);
    dv(as.getRange(2, AS.CASE, n, 1), ls.getRange('B2:B30'), false);
    if (dc) dv(as.getRange(2, AS.DOCTOR, n, 1), dc.getRange('A2:A200'), false);
    dv(as.getRange(2, 6, n, 1), ls.getRange('H2:H40'), false);
    dv(as.getRange(2, 7, n, 1), ls.getRange('G3:G40'), false);
    dv(as.getRange(2, AS.ACTION, n, 1), ['Pending', 'Done', 'canceled'], true);
    dv(as.getRange(2, AS.CS, n, 1), ['Done'], true);
    dv(as.getRange(2, AS.FEES, n, 1), ['Pending', 'Done'], true);
  }
}

/** Keeps the Lectures tab in date order (values only – formulas stay). */
function sortLectures_() {
  const lw = sheet_(LEC.SHEET), last = lastDataRow_(lw, 1);
  if (last < 3) return 0;
  const v = lw.getRange(2, 1, last - 1, LX.WIDTH).getValues().filter(function (r) { return r[0] instanceof Date; });
  const sorted = v.slice().sort(function (a, b) { return a[0] - b[0] || String(a[3]).localeCompare(String(b[3])) || String(a[2]).localeCompare(String(b[2])); });
  if (sorted.every(function (r, i) { return r === v[i]; })) return 0;
  writeLectures_(sorted, false);
  return sorted.length;
}

/** Group Schedule follows the Lectures tab: each run of lectures (group + module) sets the module start, the doctor, and the next module's start ends it. */
function syncScheduleFromLectures_() {
  const lw = sheet_(LEC.SHEET), gs = sheet_(T.SCHEDULE);
  if (!lw || !gs) return 'no Lectures / Group Schedule tab';
  const last = lastDataRow_(lw, 1);
  if (last < 2) return 'no lectures';
  const groups = {};
  gs.getRange(2, 2, Math.max(lastDataRow_(gs, 1) - 1, 1), 1).getValues().forEach(function (r) { if (r[0]) groups[r[0]] = true; });
  const now = Date.now(), lo = now - 150 * 864e5, hi = now + 150 * 864e5;   // ignore far-away rows (old data / mistyped years)
  const L = lw.getRange(2, 1, last - 1, 11).getValues()
    .filter(function (r) { return r[0] instanceof Date && r[0] >= lo && r[0] <= hi && groups[r[3]] && String(r[4]).trim(); })
    .sort(function (a, b) { return a[0] - b[0]; });
  const blocks = {};   // group → [{module, start, last, docs}]
  L.forEach(function (r) {
    const g = r[3], arr = blocks[g] = blocks[g] || [], b = arr[arr.length - 1], m = String(r[4]).trim();
    if (b && b.module === m) { b.last = r[0]; if (r[6]) b.docs[r[6]] = (b.docs[r[6]] || 0) + 1; }
    else { const nb = { module: m, start: r[0], last: r[0], docs: {}, first: String(r[5]).trim() }; if (r[6]) nb.docs[r[6]] = 1; arr.push(nb); }
  });
  const gl = lastDataRow_(gs, 1);
  const rows = gl >= 2 ? gs.getRange(2, 1, gl - 1, 9).getValues() : [];
  let updated = 0, added = 0;
  const day = 864e5;
  Object.keys(blocks).forEach(function (g) {
    blocks[g].forEach(function (b) {
      const doc = Object.keys(b.docs).sort(function (x, y) { return b.docs[y] - b.docs[x]; })[0] || '';
      let best = -1, bd = 1e15;
      rows.forEach(function (r, i) {
        if (r[1] !== g || String(r[3]).trim() !== b.module || !(r[4] instanceof Date)) return;
        const dist = Math.abs(r[4] - b.start); if (dist < bd) { bd = dist; best = i; }
      });
      if (best !== -1 && bd <= 75 * day) {
        const r = rows[best];
        if (b.first === '1' && cairoKey_(r[4]) !== cairoKey_(b.start)) { r[4] = b.start; r._ch = true; }   // the start moves only when lecture 1 is in the tab
        if (doc && officialDoctor_(r[8]) !== doc) { r[8] = doc; r._ch = true; }
        if (r[5] instanceof Date && r[5] < b.last) { r[5] = b.last; r._ch = true; }
      } else {
        const ids = rows.filter(function (x) { return x[1] === g; }).map(function (x) { return String(x[0]); });
        const pre = ids.length ? ids[0].replace(/-\d+$/, '') : g.replace(/[^A-Za-z0-9]/g, '').slice(0, 4).toUpperCase();
        const num = ids.reduce(function (m, id) { const k = Number((id.match(/-(\d+)$/) || [])[1] || 0); return Math.max(m, k); }, 0) + 1;
        const nr = [pre + '-' + ('0' + num).slice(-2), g, GROUP_ACC[g] === 'EIU DBA' ? 'EIU DBA / AITU DBA' : (GROUP_ACC[g] || ''), b.module, b.start, new Date(b.last.getTime() + 14 * day), 'Lectures', '', doc];
        nr._ch = true; nr._new = true; rows.push(nr); added++;
      }
    });
    // each module ends the day before the next module of the same group starts
    const gr = rows.filter(function (x) { return x[1] === g && x[4] instanceof Date; }).sort(function (a, b) { return a[4] - b[4]; });
    gr.forEach(function (r, i) {
      if (i + 1 < gr.length) {
        let e = new Date(gr[i + 1][4].getTime() - day);
        if (e < r[4]) e = r[4];
        if (!(r[5] instanceof Date) || cairoKey_(r[5]) !== cairoKey_(e)) { r[5] = e; r._ch = true; }
      }
    });
  });
  rows.forEach(function (r, i) {
    if (!r._ch) return;
    const k = i + 2;
    if (r._new) {
      gs.getRange(k, 1, 1, 7).setValues([r.slice(0, 7)]);
      gs.getRange(k, 8).setFormula('=IF(TODAY()<E' + k + ',"Upcoming",IF(TODAY()<=F' + k + ',"Running","Finished"))');
      gs.getRange(k, 9).setValue(r[8]);
      gs.getRange(k, 10).setFormula('=B' + k + '&"|"&H' + k);
      gs.getRange(k, 5, 1, 2).setNumberFormat('dd/mm/yyyy');
    } else { gs.getRange(k, 5, 1, 2).setValues([[r[4], r[5]]]); gs.getRange(k, 9).setValue(r[8]); updated++; }
  });
  return 'Group Schedule ← Lectures: updated ' + updated + ' · new modules ' + added;
}
function nightly() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(280000); } catch (e) { return; }
  try { sortLectures_(); Logger.log(syncScheduleFromLectures_()); } finally { lock.releaseLock(); }
}
function menuLectures() { const n = sortLectures_(); toast_('Lectures sorted (' + n + ' rows) · ' + syncScheduleFromLectures_()); }

/* =========================== Attendance: per lecture (report) and per student (Students tab) =========================== */
// Students tab, added columns (values written by the script every 10 minutes)
const SA = { LAST: 21, LAST_ATT: 22, ABS: 23 };
const AR_SHEET = 'Attendance Report';

/** Lecture start / end today in minutes, from "06:00 PM - 08:00 PM". */
function lecMinutes_(time) {
  const m = String(time || '').replace(/(AM|PM)\s*(?::|to)\s*/i, '$1 - ').match(/(\d{1,2}):(\d{2})\s*(AM|PM)\s*[-–]\s*(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!m) return null;
  const f = function (h, mi, ap) { h = +h % 12 + (/pm/i.test(ap) ? 12 : 0); return h * 60 + +mi; };
  return { start: f(m[1], m[2], m[3]), end: f(m[4], m[5], m[6]) };
}

/** Everything both outputs need: lectures by group, check-ins by phone + day, students by group. */
function attendanceData_() {
  const tz = SETTINGS.TIMEZONE, now = new Date(), today = dayStart_(now);
  const nowMin = +Utilities.formatDate(now, tz, 'HH') * 60 + +Utilities.formatDate(now, tz, 'mm');
  // check-ins: phone key → { day: true }
  const fr = sheet_(SETTINGS.RESPONSES_SHEET), frLast = fr ? lastDataRow_(fr, 1) : 1;
  const ins = {}; let first = null;
  if (frLast >= 2) fr.getRange(2, 1, frLast - 1, 13).getValues().forEach(function (r) {
    const d = toDate_(r[7]) || toDate_(r[0]); const k = String(r[12] || '').trim() || phoneKey_(r[1]);
    if (!d || !k) return;
    const day = dayStart_(d); (ins[k] = ins[k] || {})[day] = true; if (!first || day < first) first = day;
  });
  // tracking starts on the first check-in day (or Settings!B37 when filled)
  const set = sheet_(T.SETTINGS);
  if (!String(set.getRange('A37').getValue()).trim()) {
    set.getRange('A37').setValue('Attendance counted from').setFontWeight('bold');
    set.getRange('C37').setValue('Lectures before this day are not counted as absences (the check-in page started then). Change it if needed.');
  }
  let from = toDate_(set.getRange('B37').getValue());
  if (!from && first) { from = new Date(+String(first).slice(0, 4), +String(first).slice(4, 6) - 1, +String(first).slice(6, 8)); set.getRange('B37').setValue(from).setNumberFormat('dd/mm/yyyy'); }
  const fromKey = from ? dayStart_(from) : 99999999;
  // lectures
  const lw = sheet_(LEC.SHEET), last = lastDataRow_(lw, 1);
  const lv = last >= 2 ? lw.getRange(2, 1, last - 1, LX.WIDTH).getValues() : [];
  const byGroup = {};
  lv.forEach(function (r) {
    if (!(r[0] instanceof Date) || !String(r[3]).trim()) return;
    const cancelled = String(r[10]).trim() === 'لم يحضر' || /^cancel/i.test(String(r[19]).trim());
    const moved = cancelled && r[12] instanceof Date ? r[12] : null;
    const d = moved || r[0], day = dayStart_(d), mm = lecMinutes_(r[2]);
    const n = Number(r[5]), isLecture = String(r[5]).trim() !== '' && !isNaN(n);
    let status;
    if (cancelled && !moved) status = 'Cancelled';
    else if (day < today || (day === today && mm && nowMin >= mm.end)) status = 'Held';
    else if (day === today) status = 'Today';
    else status = 'Upcoming';
    (byGroup[r[3]] = byGroup[r[3]] || []).push({ d: d, day: day, time: String(r[2] || ''), group: String(r[3]), module: String(r[4] || ''), n: isLecture ? n : String(r[5] || '').trim(),
      isLecture: isLecture, doctor: String(r[6] || ''), status: status, reason: String(r[11] || '').trim(), moved: !!moved, tracked: day >= fromKey,
      zoom: /^https?:/.test(String(r[8] || '').trim()), wasCancelled: cancelled, origDay: dayStart_(r[0]) });
  });
  Object.keys(byGroup).forEach(function (g) {
    const seen = {};   // the same lecture typed twice in the Lectures tab is counted once
    byGroup[g] = byGroup[g].filter(function (x) { const k = x.day + '|' + x.module + '|' + x.n; if (seen[k]) return false; seen[k] = true; return true; });
    const L = byGroup[g].sort(function (a, b) { return a.d - b.d; }), tot = {};
    L.forEach(function (x) { if (x.isLecture) tot[x.module] = Math.max(tot[x.module] || 0, x.n); });
    L.forEach(function (x) { x.of = tot[x.module] || 0; });
  });
  // students
  const st = sheet_(T.STUDENTS), sl = lastDataRow_(st, SC.PHONE);
  const sv = sl >= 2 ? st.getRange(2, 1, sl - 1, SC.WIDTH).getValues() : [];
  return { today: today, from: from, fromKey: fromKey, ins: ins, byGroup: byGroup, students: sv, st: st };
}
function attended_(D, key, day) { return !!(D.ins[key] && D.ins[key][day]); }
function startedBy_(s, day) { const d = toDate_(s[SC.START_DATE - 1]); return !d || dayStart_(d) <= day; }
function shortDate_(d) { return Utilities.formatDate(d, SETTINGS.TIMEZONE, 'dd/MM/yyyy'); }

/** Students tab: last lecture of the student's group, Attend / ما حضرش, and absences per module since tracking began. */
function refreshStudentAttendance_(D) {
  D = D || attendanceData_();
  const st = D.st, n = D.students.length;
  st.getRange(1, SA.LAST, 1, 3).setValues([['Last lecture', 'Last lecture attendance', 'Absences per module']])
    .setFontWeight('bold').setBackground('#0c2536').setFontColor('#ffffff');
  if (!n) return 0;
  const out = D.students.map(function (s) {
    const key = String(s[SC.KEY - 1] || '').trim(), L = D.byGroup[s[SC.GROUP - 1]];
    if (!key || !L || String(s[SC.STATUS - 1]) !== ACTIVE_STATUS) return ['', '', ''];
    const held = L.filter(function (x) { return x.status === 'Held' && x.isLecture && x.tracked && startedBy_(s, x.day); });
    if (!held.length) return ['', '', ''];
    const lastL = held[held.length - 1];
    const mods = [], per = {};
    held.forEach(function (x) {
      if (!per[x.module]) { per[x.module] = { held: 0, missed: [] }; mods.push(x.module); }
      per[x.module].held++;
      if (!attended_(D, key, x.day)) per[x.module].missed.push('#' + x.n);
    });
    const abs = mods.map(function (m) { const p = per[m]; return m + ': غاب ' + p.missed.length + ' من ' + p.held + (p.missed.length ? ' (' + p.missed.join('، ') + ')' : ''); }).join('\n');
    return [shortDate_(lastL.d) + ' · ' + lastL.module + ' · محاضرة ' + lastL.n + ' من ' + lastL.of, attended_(D, key, lastL.day) ? 'Attend' : 'ما حضرش', abs];
  });
  st.getRange(2, SA.LAST, n, 3).setValues(out);
  st.getRange(2, SA.ABS, n, 1).setWrap(true);
  return n;
}

/** "Attendance Report" tab: one table per running group – current module lecture by lecture, then the upcoming modules. */
function buildAttendanceReport_(D, runs) {
  D = D || attendanceData_(); runs = runs || moduleRuns_(D);
  let rp = sheet_(AR_SHEET);
  if (!rp) { rp = SpreadsheetApp.getActive().insertSheet(AR_SHEET); rp.setTabColor('#0c2536'); }
  rp.clear();
  const tz = SETTINGS.TIMEZONE, W = 9, aCharts = [];
  const rows = [], fmt = [];   // fmt: [rowIndex, kind]
  const push = function (r, kind) { while (r.length < W) r.push(''); rows.push(r); fmt.push(kind); };
  push(['Attendance Report'], 'title');
  push(['Updated ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm') + ' · attendance counted from ' + (D.from ? shortDate_(D.from) : '—') +
        ' · "Should attend" = students "' + ACTIVE_STATUS + '" in the group'], 'note');
  const active = {}; D.students.forEach(function (s) { if (String(s[SC.STATUS - 1]) === ACTIVE_STATUS) (active[s[SC.GROUP - 1]] = active[s[SC.GROUP - 1]] || []).push(s); });
  const running = Object.keys(D.byGroup).filter(function (g) {
    return D.byGroup[g].some(function (x) { return Math.abs(x.d - new Date()) < 35 * 864e5; });
  }).sort();
  const groups = running.filter(function (g) { return (active[g] || []).length; });
  const noStudents = running.filter(function (g) { return groups.indexOf(g) === -1; });
  groups.forEach(function (g) {
    const L = D.byGroup[g], studs = active[g] || [];
    const past = L.filter(function (x) { return x.day <= D.today; });
    const cur = (past.length ? past[past.length - 1] : L[0]).module;
    const curL = L.filter(function (x) { return x.module === cur; });
    push([], 'gap');
    push([g + '  —  ' + studs.length + ' active students  —  current module: ' + cur], 'group');
    const chart = { row: rows.length, type: 'COLUMN', title: g + ' – attendance % per lecture', header: ['Lecture', 'Attendance %'], data: [], max: 100, colors: ['#2e7d5b'] };
    aCharts.push(chart);
    push(['Module', 'Lecture', 'Date', 'Day', 'Doctor', 'Status', 'Should attend', 'Attended', 'Attendance %'], 'head');
    curL.forEach(function (x) {
      const lab = x.isLecture ? x.n + ' of ' + x.of : x.n;
      let should = '', came = '', pct = '', status = x.status === 'Cancelled' ? 'Cancelled' + (x.reason ? ' – ' + x.reason : '') : x.status;
      if (x.moved) status += ' (new date)';
      if (x.status === 'Held' && x.isLecture) {
        if (!x.tracked) { should = '—'; came = '—'; pct = 'before tracking'; }
        else {
          const exp = studs.filter(function (s) { return startedBy_(s, x.day); });
          const att = exp.filter(function (s) { return attended_(D, String(s[SC.KEY - 1]).trim(), x.day); }).length;
          should = exp.length; came = att; pct = exp.length ? Math.round(att / exp.length * 100) + '%' : '';
        }
      }
      push([x.module, lab, shortDate_(x.d), Utilities.formatDate(x.d, tz, 'EEEE'), x.doctor, status, should, came, pct], x.status);
      if (typeof should === 'number' && should) chart.data.push(['#' + x.n + ' · ' + Utilities.formatDate(x.d, tz, 'dd/MM'), Math.round(came / should * 100)]);
    });
    // upcoming modules (after the current one)
    const firstDay = {}, up = [];   // modules whose first lecture is still ahead, in date order
    L.forEach(function (x) { if (!(x.module in firstDay)) { firstDay[x.module] = x.day; if (x.module !== cur && x.day > D.today) up.push(x.module); } });
    up.forEach(function (m) {
      const ml = L.filter(function (x) { return x.module === m; }), lec = ml.filter(function (x) { return x.isLecture; });
      const docs = ml.map(function (x) { return x.doctor; }).filter(function (d, i, a) { return d && a.indexOf(d) === i; }).join(' / ');
      push([m, (lec.length ? Math.max.apply(null, lec.map(function (x) { return x.n; })) : ml.length) + ' lectures', shortDate_(ml[0].d) + ' → ' + shortDate_(ml[ml.length - 1].d),
            Utilities.formatDate(ml[0].d, tz, 'EEEE'), docs, 'Upcoming module', '', '', ''], 'upmod');
    });
    // history: modules this group already finished (newest first) – averages per lecture
    const curRun = runs.filter(function (r) { return r.group === g && r.module === cur; }).pop();
    runs.filter(function (r) { return r.group === g && r.status === 'Finished' && r !== curRun; }).sort(function (a, b) { return b.end - a.end; }).forEach(function (r) {
      const at = runAttendance_(D, r, studs);
      push([r.module, r.held + ' of ' + r.total + ' held', shortDate_(r.start) + ' → ' + shortDate_(r.end), '', r.doctor, 'Finished' + (r.cancelled ? ' · ' + r.cancelled + ' cancelled' : ''),
            at ? at.should : '—', at ? at.came : '—', at ? at.pct : noAttLabel_(r)], 'hist');
    });
  });
  if (!groups.length) push(['No running groups in the Lectures tab.'], 'note');
  if (noStudents.length) { push([], 'gap'); push(['Not in this report – lectures exist but no "' + ACTIVE_STATUS + '" students in the Students tab: ' + noStudents.join(' · ')], 'note'); }
  rp.getRange(1, 1, rows.length, W).setValues(rows);
  // formatting
  const COL = { Held: '#ffffff', Today: '#fff4d6', Upcoming: '#f3f6f9', Cancelled: '#fde8e8', upmod: '#e8f1fb', hist: '#f2f2f2' };
  rp.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor('#0c2536');
  rp.getRange(2, 1).setFontStyle('italic').setFontColor('#6b7f8e');
  fmt.forEach(function (k, i) {
    const r = i + 1;
    if (k === 'group') rp.getRange(r, 1, 1, W).setBackground('#0c2536').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);
    else if (k === 'head') rp.getRange(r, 1, 1, W).setBackground('#174560').setFontColor('#ffffff').setFontWeight('bold');
    else if (COL[k]) {
      rp.getRange(r, 1, 1, W).setBackground(COL[k]);
      if (k === 'upmod' || k === 'hist') rp.getRange(r, 1, 1, W).setFontStyle('italic');
      const p = parseInt(rows[i][8], 10);
      if (!isNaN(p)) rp.getRange(r, 9).setFontWeight('bold').setFontColor(p >= 75 ? '#1e7a46' : p >= 50 ? '#9a6a00' : '#a52f2f');
    }
  });
  rp.getRange(1, 2, rows.length, 8).setHorizontalAlignment('center');
  [300, 90, 160, 95, 210, 200, 105, 90, 115].forEach(function (w, i) { rp.setColumnWidth(i + 1, w); });
  rp.setFrozenRows(2);
  aCharts.forEach(function (c, i) { const next = i + 1 < aCharts.length ? aCharts[i + 1].row : rows.length + 2; c.h = Math.min(320, Math.max(200, (next - c.row - 1) * 21)); });
  try { drawCharts_(rp, W + 2, aCharts); } catch (e) { Logger.log('charts attendance: ' + e); }
  return groups.length;
}
/* =========================== Module runs (from the Lectures tab) – shared by the reports =========================== */
/** One run = one module taught to one group (consecutive lectures). Status: Running / Upcoming / Finished. */
function moduleRuns_(D) {
  const runs = [];
  Object.keys(D.byGroup).forEach(function (g) {
    const open = {};   // module → its current run (a group can alternate two modules, e.g. two Qualifi cohorts)
    D.byGroup[g].forEach(function (x) {
      if (!x.module) return;
      let run = open[x.module];
      if (!run || x.d - run.lastD > 60 * 864e5) { run = open[x.module] = { group: g, module: x.module, rows: [], lastD: x.d }; runs.push(run); }
      run.rows.push(x); run.lastD = x.d;
    });
  });
  runs.forEach(function (r) {
    const lec = r.rows.filter(function (x) { return x.isLecture; }), base = lec.length ? lec : r.rows;
    r.start = base[0].d; r.end = base[base.length - 1].d;
    r.startKey = dayStart_(r.start); r.endKey = dayStart_(r.end);
    r.total = lec.length ? Math.max.apply(null, lec.map(function (x) { return x.n; })) : 0;
    r.held = lec.filter(function (x) { return x.status === 'Held'; }).length;
    r.cancelled = lec.filter(function (x) { return x.status === 'Cancelled'; }).length;
    const dc = {}; r.rows.forEach(function (x) { if (x.doctor) dc[x.doctor] = (dc[x.doctor] || 0) + 1; });
    r.doctors = Object.keys(dc).sort(function (a, b) { return dc[b] - dc[a]; });
    r.doctor = r.doctors.join(' / ');
    const allDone = lec.length && lec.every(function (x) { return x.status === 'Held' || x.status === 'Cancelled'; });
    r.status = r.startKey > D.today ? 'Upcoming' : (allDone || r.endKey < D.today ? 'Finished' : 'Running');
    r.evals = [];
  });
  return runs;
}
/** Evaluations tab rows, each attached to the run of the same group + module it was given in. */
function attachEvaluations_(runs) {
  const sh = sheet_(EV.SHEET), out = [];
  if (!sh) return out;
  const last = lastDataRow_(sh, 1);
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 15).getValues().forEach(function (r) {
    const d = toDate_(r[0]) || (r[0] instanceof Date ? r[0] : null), dr = Number(r[7]), lr = Number(r[8]);
    if (!d || !(dr || lr)) return;
    const e = { d: d, group: String(r[4]), module: String(r[5]), doctor: officialDoctor_(r[6]) || String(r[6] || '').trim(), dr: dr || null, lr: lr || null,
                mr: Number(r[13]) || null, ir: Number(r[14]) || null };
    const cand = runs.filter(function (x) { return x.group === e.group && x.module === e.module; });
    let best = null;
    cand.forEach(function (x) { const dist = d < x.start ? x.start - d : (d > x.end ? d - x.end : 0); if (!best || dist < best.dist) best = { run: x, dist: dist }; });
    if (best && best.dist < 60 * 864e5) best.run.evals.push(e);
    out.push(e);
  });
  return out;
}
function avg_(list, f) { const v = list.map(f).filter(function (x) { return x; }); return v.length ? Math.round(v.reduce(function (a, b) { return a + b; }, 0) / v.length * 10) / 10 : ''; }
/** Average attendance of a run's tracked, held lectures (same counting as the Attendance Report). */
function runAttendance_(D, run, studs) {
  let s = 0, a = 0, n = 0;
  run.rows.forEach(function (x) {
    if (x.status !== 'Held' || !x.isLecture || !x.tracked) return;
    const exp = studs.filter(function (st) { return startedBy_(st, x.day); });
    if (!exp.length) return;
    s += exp.length; a += exp.filter(function (st) { return attended_(D, String(st[SC.KEY - 1]).trim(), x.day); }).length; n++;
  });
  return n ? { should: Math.round(s / n), came: Math.round(a / n), pct: Math.round(a / s * 100) + '%', n: n } : null;
}
function noAttLabel_(run) { return run.rows.some(function (x) { return x.status === 'Held' && x.isLecture && x.tracked; }) ? 'no students in Students tab' : 'before tracking'; }
function activeByGroup_(D) { const o = {}; D.students.forEach(function (s) { if (String(s[SC.STATUS - 1]) === ACTIVE_STATUS) (o[s[SC.GROUP - 1]] = o[s[SC.GROUP - 1]] || []).push(s); }); return o; }

/** Writes a list of [row values, kind] blocks into a fresh report tab with the house formatting. */
/* =========================== Charts beside the reports (data in the hidden "Chart Data" tab) =========================== */
const CHART_SHEET = 'Chart Data';
const CHART_BAND = { 'Manager Summary': 1, 'Attendance Report': 41, 'Doctors Report': 81, 'Academic Path Report': 121, 'Students Follow-up': 161 };
const CHART_COLORS = ['#174560', '#2e7d5b', '#c0392b', '#d99a1e', '#7a4bb3', '#6b7f8e', '#1f8fb5', '#9a6a00'];
/** specs: { row, title, type: 'COLUMN' | 'BAR' | 'LINE' | 'PIE', header: [...], data: [[label, n, ...]], stacked, max, h, w } */
function drawCharts_(rp, col, specs) {
  let cd = sheet_(CHART_SHEET);
  if (!cd) cd = SpreadsheetApp.getActive().insertSheet(CHART_SHEET);
  if (cd.getMaxColumns() < 240) cd.insertColumnsAfter(cd.getMaxColumns(), 240 - cd.getMaxColumns());
  cd.hideSheet();
  const band = CHART_BAND[rp.getName()] || 201;
  // lay the data blocks out first; if the layout is the one already drawn, keep the charts and only rewrite the numbers
  let r0 = 1;
  const plan = (specs || []).map(function (sp) {
    const data = (sp.data || []).filter(function (d) { return d.slice(1).some(function (x) { return typeof x === 'number'; }); });
    if (!data.length) return null;
    const block = [sp.header].concat(data); if (r0 + block.length > 600) return null;
    const p = { sp: sp, block: block, r: r0 }; r0 += block.length + 1; return p;
  }).filter(Boolean);
  const sig = JSON.stringify(plan.map(function (p) { return [p.sp.row, p.r, p.block.length, p.block[0].length, p.sp.type, p.sp.title, p.sp.h]; }));
  const P = PropertiesService.getDocumentProperties ? PropertiesService.getDocumentProperties() : PropertiesService.getScriptProperties();
  const key = 'CHARTS_' + rp.getName(), same = P.getProperty(key) === sig && rp.getCharts().length === plan.length;
  cd.getRange(1, band, 600, 40).clearContent();
  plan.forEach(function (p) { cd.getRange(p.r, band, p.block.length, p.block[0].length).setValues(p.block); });
  if (same) return;
  rp.getCharts().forEach(function (c) { rp.removeChart(c); });
  P.setProperty(key, sig);
  plan.forEach(function (p) {
    const sp = p.sp, block = p.block, w = sp.header.length;
    const rg = cd.getRange(p.r, band, block.length, w);
    const axis = { viewWindow: { min: 0 } };
    if (sp.max) axis.viewWindow.max = sp.max;
    const b = rp.newChart().setChartType(Charts.ChartType[sp.type]).addRange(rg).setNumHeaders(1)
      .setPosition(Math.max(sp.row, 1), col, 12, 4)
      .setOption('title', sp.title).setOption('titleTextStyle', { color: '#0c2536', fontSize: 13, bold: true })
      .setOption('width', sp.w || 560).setOption('height', sp.h || 300)
      .setOption('colors', sp.colors || CHART_COLORS)
      .setOption('legend', { position: sp.type === 'PIE' || w > 2 ? 'right' : 'none' })
      .setOption('backgroundColor', '#ffffff');
    if (sp.type === 'BAR') b.setOption('hAxis', axis); else if (sp.type !== 'PIE') b.setOption('vAxis', axis);
    if (sp.stacked) b.setOption('isStacked', true);
    if (sp.type === 'PIE') b.setOption('pieHole', 0.45).setOption('pieSliceText', 'value');
    if (sp.type === 'COLUMN' || sp.type === 'BAR') b.setOption('annotations', { alwaysOutside: true });
    rp.insertChart(b.build());
  });
}

function writeReport_(name, color, W, rows, widths, charts) {
  let rp = sheet_(name);
  if (!rp) { rp = SpreadsheetApp.getActive().insertSheet(name); rp.setTabColor(color); }
  rp.clear();
  const vals = rows.map(function (r) { const v = r[0].slice(); while (v.length < W) v.push(''); return v; });
  rp.getRange(1, 1, vals.length, W).setValues(vals);
  const BG = { Held: '#ffffff', Today: '#fff4d6', Upcoming: '#f3f6f9', Cancelled: '#fde8e8', upmod: '#e8f1fb', Running: '#eaf6ee', Finished: '#ffffff', hist: '#f7f7f7' };
  rp.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor('#0c2536');
  rp.getRange(2, 1).setFontStyle('italic').setFontColor('#6b7f8e');
  rows.forEach(function (r, i) {
    const k = r[1], row = i + 1;
    if (k === 'group') rp.getRange(row, 1, 1, W).setBackground('#0c2536').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);
    else if (k === 'section') rp.getRange(row, 1, 1, W).setBackground('#2e7d5b').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);
    else if (k === 'histsec') rp.getRange(row, 1, 1, W).setBackground('#6b7f8e').setFontColor('#ffffff').setFontWeight('bold');
    else if (k === 'head') rp.getRange(row, 1, 1, W).setBackground('#174560').setFontColor('#ffffff').setFontWeight('bold');
    else if (BG[k]) { rp.getRange(row, 1, 1, W).setBackground(BG[k]); if (k === 'upmod' || k === 'hist') rp.getRange(row, 1, 1, W).setFontStyle('italic'); }
    if (r[2]) rp.getRange(row, r[2], 1, 1).setFontWeight('bold').setFontColor(r[3]);
  });
  rp.getRange(3, 2, Math.max(vals.length - 2, 1), W - 1).setHorizontalAlignment('center');
  widths.forEach(function (w, i) { rp.setColumnWidth(i + 1, w); });
  rp.setFrozenRows(2);
  try { drawCharts_(rp, W + 2, charts); } catch (e) { Logger.log('charts ' + name + ': ' + e); }
  return rp;
}
function pctColor_(p) { p = parseInt(p, 10); return isNaN(p) ? null : (p >= 75 ? '#1e7a46' : p >= 50 ? '#9a6a00' : '#a52f2f'); }
function rateColor_(r) { return r === '' || typeof r !== 'number' ? null : (r >= 4 ? '#1e7a46' : r >= 3 ? '#9a6a00' : '#a52f2f'); }

/** "Doctors Report": who teaches what now / next, ratings per doctor and per module, and the history of finished modules. */
function buildDoctorsReport_(D, runs, evals) {
  const tz = SETTINGS.TIMEZONE, W = 13, rows = [], act = activeByGroup_(D);
  const row = function (v, k, cc, color) { rows.push([v, k, color ? cc : 0, color]); };
  const sd = shortDate_;
  row(['Doctors Report'], 'title');
  row(['Updated ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm') + ' · from the Lectures tab and the student evaluations (1–5) · finished modules move down to History'], 'note');
  // 1 · doctors now & next
  const dc = sheet_('Doctors'), dl = dc ? lastDataRow_(dc, 1) : 1;
  const names = [];
  (dl >= 2 ? dc.getRange(2, 1, dl - 1, 1).getValues() : []).forEach(function (r) { const n = officialDoctor_(r[0]) || String(r[0]).trim(); if (n && names.indexOf(n) === -1) names.push(n); });
  runs.forEach(function (r) { r.doctors.forEach(function (d) { if (names.indexOf(d) === -1) names.push(d); }); });
  const evBy = {}; evals.forEach(function (e) { (evBy[e.doctor] = evBy[e.doctor] || []).push(e); if (names.indexOf(e.doctor) === -1) names.push(e.doctor); });
  const docRows = names.map(function (n) {
    const mine = runs.filter(function (r) { return r.doctors.indexOf(n) !== -1; });
    const now = mine.filter(function (r) { return r.status === 'Running'; }), next = mine.filter(function (r) { return r.status === 'Upcoming'; }).sort(function (a, b) { return a.start - b.start; });
    return { n: n, now: now, next: next, done: mine.filter(function (r) { return r.status === 'Finished'; }).length, ev: evBy[n] || [], rank: now.length ? 0 : next.length ? 1 : 2 };
  }).sort(function (a, b) { return a.rank - b.rank || a.n.localeCompare(b.n); });
  row([], 'gap');
  row(['1 · Doctors – teaching now and next'], 'section'); const cD1 = rows.length;
  row(['Doctor', 'Status', 'Module', 'Group', 'From', 'To', 'Lectures', 'Avg doctor rating', 'Avg lecture rating', 'Avg material', 'Avg interaction', 'Evaluations', 'Modules finished'], 'head');
  docRows.forEach(function (d) {
    const dr = avg_(d.ev, function (e) { return e.dr; }), lr = avg_(d.ev, function (e) { return e.lr; }), mr = avg_(d.ev, function (e) { return e.mr; }), ir = avg_(d.ev, function (e) { return e.ir; });
    const list = d.now.concat(d.next);
    const R = function (v) { return v === '' ? 'No ratings yet' : v; };   // every cell says something – no empty cells
    if (!list.length) { row([d.n, 'No lectures scheduled', '—', '—', '—', '—', '—', R(dr), R(lr), R(mr), R(ir), d.ev.length, d.done], 'hist', 8, rateColor_(dr)); return; }
    list.forEach(function (r, i) {
      const st = r.status === 'Running' ? 'Teaching now' : 'Upcoming – starts ' + sd(r.start);
      row([d.n, st, r.module || '—', r.group, sd(r.start), sd(r.end), r.status === 'Running' ? r.held + ' of ' + r.total : r.total + ' lectures',
           R(dr), R(lr), R(mr), R(ir), d.ev.length, d.done], r.status === 'Running' ? 'Running' : 'Upcoming', 8, rateColor_(dr));
    });
  });
  // 2 · modules running / upcoming with their ratings so far
  const live = runs.filter(function (r) { return r.status !== 'Finished'; }).sort(function (a, b) { return (a.status === 'Running' ? 0 : 1) - (b.status === 'Running' ? 0 : 1) || a.start - b.start; });
  row([], 'gap');
  row(['2 · Modules – running and upcoming (ratings so far)'], 'section'); const cD2 = rows.length;
  row(['Group', 'Module', 'Doctor', 'Start', 'End', 'Status', 'Lectures', 'Avg doctor rating', 'Avg lecture rating', 'Avg material', 'Avg interaction', 'Evaluations', 'Avg attendance'], 'head');
  live.forEach(function (r) {
    const dr = avg_(r.evals, function (e) { return e.dr; }), lr = avg_(r.evals, function (e) { return e.lr; }), mr = avg_(r.evals, function (e) { return e.mr; }), ir = avg_(r.evals, function (e) { return e.ir; }), at = runAttendance_(D, r, act[r.group] || []);
    row([r.group, r.module || '—', r.doctor || 'No doctor yet', sd(r.start), sd(r.end), r.status, r.status === 'Running' ? r.held + ' of ' + r.total : r.total + ' lectures',
         dr === '' ? 'No ratings yet' : dr, lr === '' ? 'No ratings yet' : lr, mr === '' ? 'No ratings yet' : mr, ir === '' ? 'No ratings yet' : ir, r.evals.length, at ? at.pct : (r.status === 'Upcoming' ? 'Not started' : noAttLabel_(r))],
        r.status, 8, rateColor_(dr));
  });
  // 3 · history
  const hist = runs.filter(function (r) { return r.status === 'Finished'; }).sort(function (a, b) { return b.end - a.end; });
  row([], 'gap');
  row(['History – finished modules (newest first)'], 'histsec');
  row(['Group', 'Module', 'Doctor', 'Start', 'End', 'Lectures held', 'Cancelled', 'Avg doctor rating', 'Avg lecture rating', 'Avg material', 'Avg interaction', 'Evaluations', 'Avg attendance'], 'head');
  if (!hist.length) row(['No finished modules yet.'], 'note');
  hist.forEach(function (r) {
    const dr = avg_(r.evals, function (e) { return e.dr; }), lr = avg_(r.evals, function (e) { return e.lr; }), mr = avg_(r.evals, function (e) { return e.mr; }), ir = avg_(r.evals, function (e) { return e.ir; }), at = runAttendance_(D, r, act[r.group] || []);
    row([r.group, r.module || '—', r.doctor || 'No doctor', sd(r.start), sd(r.end), r.held + ' of ' + r.total, r.cancelled, dr === '' ? 'No ratings' : dr, lr === '' ? 'No ratings' : lr, mr === '' ? 'No ratings' : mr, ir === '' ? 'No ratings' : ir, r.evals.length, at ? at.pct : noAttLabel_(r)], 'Finished', 8, rateColor_(dr));
  });
  const rated = docRows.filter(function (d) { return d.ev.length; }).map(function (d) { return [d.n, avg_(d.ev, function (e) { return e.dr; }) || 0, avg_(d.ev, function (e) { return e.lr; }) || 0, avg_(d.ev, function (e) { return e.mr; }) || 0, avg_(d.ev, function (e) { return e.ir; }) || 0]; })
    .sort(function (a, b) { return b[1] - a[1]; });
  const load = docRows.filter(function (d) { return d.now.length || d.next.length; }).map(function (d) { return [d.n, d.now.length, d.next.length]; });
  writeReport_('Doctors Report', '#7a4bb3', W, rows, [230, 170, 280, 190, 95, 95, 95, 120, 120, 110, 120, 95, 115], [
    { row: cD1, type: 'BAR', title: 'Average rating per doctor (1–5)', header: ['Doctor', 'Doctor', 'Lecture', 'Material', 'Interaction'], data: rated, max: 5,
      colors: ['#7a4bb3', '#1f8fb5', '#2e7d5b', '#d99a1e'], h: Math.max(240, 60 + rated.length * 34) },
    { row: cD2, type: 'BAR', title: 'Modules per doctor – teaching now / upcoming', header: ['Doctor', 'Teaching now', 'Upcoming'], data: load, stacked: true,
      colors: ['#2e7d5b', '#9fb3c8'], h: Math.max(240, 60 + load.length * 28) }
  ]);
  return names.length;
}

/* =========================== Manager Summary (one tab, refreshed every 10 minutes) =========================== */
function dayKeyAdd_(key, n) { const d = new Date(+String(key).slice(0, 4), +String(key).slice(4, 6) - 1, +String(key).slice(6, 8)); d.setDate(d.getDate() + n); return dayStart_(d); }
function buildManagerSummary_(D, runs, evals) {
  const tz = SETTINGS.TIMEZONE, W = 6, rows = [], act = activeByGroup_(D), today = D.today;
  const row = function (v, k, cc, color) { rows.push([v, k, color ? cc : 0, color]); };
  const w0 = dayKeyAdd_(today, -6), p0 = dayKeyAdd_(today, -13), p1 = dayKeyAdd_(today, -7);   // last 7 days incl. today, the 7 before
  const inW = function (k) { return k >= w0 && k <= today; }, inP = function (k) { return k >= p0 && k <= p1; };
  const allL = []; Object.keys(D.byGroup).forEach(function (g) { D.byGroup[g].forEach(function (x) { if (x.isLecture) allL.push(x); }); });
  const attOver = function (inside, group) {
    let s = 0, a = 0;
    allL.forEach(function (x) {
      if (x.status !== 'Held' || !x.tracked || !inside(x.day) || (group && x.group !== group)) return;
      const exp = (act[x.group] || []).filter(function (st) { return startedBy_(st, x.day); });
      s += exp.length; a += exp.filter(function (st) { return attended_(D, String(st[SC.KEY - 1]).trim(), x.day); }).length;
    });
    return s ? Math.round(a / s * 100) : '';
  };
  const cnt = function (f) { return allL.filter(f).length; };
  const evIn = function (inside) { return evals.filter(function (e) { return inside(dayStart_(e.d)); }); };
  const pct = function (v) { return v === '' ? '—' : v + '%'; };
  const arrow = function (a, b, unit) { if (a === '' || b === '' || a === '—' || b === '—') return ''; const d = Math.round((a - b) * 10) / 10; return d > 0 ? '▲ ' + d + (unit || '') : d < 0 ? '▼ ' + Math.abs(d) + (unit || '') : 'زي ما هو'; };
  row(['ملخص المدير'], 'title');
  row(['آخر تحديث ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm') + ' · "الأسبوع ده" = آخر 7 أيام · الطلبة اللي محتاجين متابعة في تاب Students Follow-up'], 'note');
  // 1 · the week in numbers
  const evW = evIn(inW), evP = evIn(inP);
  row([], 'gap'); row(['1 · الأسبوع بالأرقام'], 'section'); const cR1 = rows.length;
  row(['', 'الأسبوع ده', 'الأسبوع اللي فات', 'الفرق'], 'head');
  const aw = attOver(inW), ap0 = attOver(inP), dw = avg_(evW, function (e) { return e.dr; }), dp = avg_(evP, function (e) { return e.dr; });
  const hw = cnt(function (x) { return x.status === 'Held' && inW(x.day); }), hp = cnt(function (x) { return x.status === 'Held' && inP(x.day); });
  const cw = cnt(function (x) { return x.wasCancelled && inW(x.origDay); }), cp = cnt(function (x) { return x.wasCancelled && inP(x.origDay); });
  [['محاضرات اتعملت', hw, hp, arrow(hw, hp)],
   ['محاضرات اتلغت', cw, cp, arrow(cw, cp)],
   ['نسبة الحضور', pct(aw), pct(ap0), arrow(aw, ap0, '%')],
   ['تقييم الدكاترة (من 5)', dw === '' ? '—' : dw, dp === '' ? '—' : dp, arrow(dw, dp)],
   ['عدد التقييمات', evW.length, evP.length, arrow(evW.length, evP.length)]
  ].forEach(function (x) { row(x, 'Held'); });
  row(['طلبة بيدرسوا دلوقتي', D.students.filter(function (s) { return String(s[SC.STATUS - 1]) === ACTIVE_STATUS; }).length, '', ''], 'Held');
  // 2 · attendance per group this week
  const groups = Object.keys(act).filter(function (g) { return (D.byGroup[g] || []).some(function (x) { return Math.abs(x.d - new Date()) < 35 * 864e5; }); }).sort();   // running groups
  row([], 'gap'); row(['2 · الحضور لكل جروب'], 'section'); const cR2 = rows.length;
  row(['الجروب', 'محاضرات الأسبوع ده', 'الحضور الأسبوع ده', 'الحضور الأسبوع اللي فات', 'طلبة بيدرسوا'], 'head');
  const gChart = [];
  groups.forEach(function (g) {
    const h = cnt(function (x) { return x.group === g && x.status === 'Held' && inW(x.day); });
    const a1 = attOver(inW, g), a2 = attOver(inP, g);
    row([g, h, h ? pct(a1) : 'مفيش محاضرات', pct(a2), (act[g] || []).length], 'Held', 3, a1 === '' ? null : a1 >= 75 ? '#1e7a46' : a1 >= 50 ? '#9a6a00' : '#a52f2f');
    if (a1 !== '') gChart.push([g, a1]);
  });
  // 3 · next 7 days – what is missing
  const n7 = dayKeyAdd_(today, 7), gaps = [];
  allL.forEach(function (x) {
    if (x.day < today || x.day > n7 || x.status === 'Cancelled' || x.status === 'Held') return;
    const miss = []; if (!x.zoom) miss.push('لينك الزووم'); if (!x.doctor) miss.push('الدكتور'); if (!lecMinutes_(x.time)) miss.push('الميعاد');
    if (miss.length) gaps.push([shortDate_(x.d) + ' · ' + Utilities.formatDate(x.d, tz, 'EEEE'), x.group, x.module, 'محاضرة ' + x.n, 'ناقص: ' + miss.join('، ')]);
  });
  gaps.sort(function (a, b) { return a[0].slice(0, 10).split('/').reverse().join('').localeCompare(b[0].slice(0, 10).split('/').reverse().join('')); });
  row([], 'gap'); row(['3 · محاضرات الأسبوع الجاي ناقصها حاجة (' + gaps.length + ')'], 'section');
  row(['اليوم', 'الجروب', 'المادة', 'المحاضرة', 'المطلوب'], 'head');
  if (!gaps.length) row(['كل محاضرات الأسبوع الجاي فيها لينك ودكتور وميعاد.'], 'note');
  gaps.forEach(function (g) { row(g, 'Today', 5, '#a52f2f'); });
  // 4 · cancellations this month
  const mStart = Number(Utilities.formatDate(new Date(), tz, 'yyyyMM') + '01');
  const canc = allL.filter(function (x) { return x.wasCancelled && x.origDay >= mStart && x.origDay <= today; });
  const byR = {}, byG = {};
  canc.forEach(function (x) { const r = x.reason || 'من غير سبب مكتوب'; byR[r] = (byR[r] || 0) + 1; byG[x.group] = (byG[x.group] || 0) + 1; });
  const rk = Object.keys(byR).sort(function (a, b) { return byR[b] - byR[a]; }), gk = Object.keys(byG).sort(function (a, b) { return byG[b] - byG[a]; });
  row([], 'gap'); row(['4 · المحاضرات اللي اتلغت الشهر ده (' + canc.length + ')'], 'section'); const cR4 = rows.length;
  if (!canc.length) row(['مفيش محاضرات اتلغت الشهر ده.'], 'note');
  else {
    row(['السبب', 'عدد المحاضرات'], 'head');
    rk.forEach(function (r) { row([r, byR[r]], 'Held'); });
    row([], 'gap');
    row(['الجروب', 'اتلغت', 'اتعملت', 'نسبة الإلغاء'], 'head');
    gk.forEach(function (g) { const held = allL.filter(function (x) { return x.group === g && x.status === 'Held' && x.day >= mStart; }).length;
      row([g, byG[g], held, Math.round(byG[g] / (held + byG[g]) * 100) + '%'], 'Held'); });
  }
  const weeks = [];
  for (let i = 5; i >= 0; i--) {
    const e = dayKeyAdd_(today, -7 * i), b0 = dayKeyAdd_(e, -6), p = attOver(function (k) { return k >= b0 && k <= e; });
    weeks.push([dayKeyToIso_(b0).slice(5).split('-').reverse().join('/') + '–' + dayKeyToIso_(e).slice(5).split('-').reverse().join('/'), p === '' ? null : p]);
  }
  writeReport_('Manager Summary', '#c0392b', W, rows, [280, 170, 170, 190, 220, 120], [
    { row: cR1, type: 'LINE', title: 'نسبة الحضور – آخر 6 أسابيع', header: ['الأسبوع', 'الحضور %'], data: weeks, max: 100, h: 280 },
    { row: cR2, type: 'BAR', title: 'نسبة الحضور الأسبوع ده – لكل جروب', header: ['الجروب', 'الحضور %'], data: gChart, max: 100, colors: ['#2e7d5b'], h: Math.max(220, 60 + gChart.length * 32) },
    { row: cR4, type: 'PIE', title: 'أسباب الإلغاء الشهر ده', header: ['السبب', 'محاضرات'], data: rk.map(function (k) { return [k, byR[k]]; }), h: 260 }
  ]);
  return buildFollowUp_(D);
}

/** "Students Follow-up" tab: active students who stopped attending (2+ lectures in a row) or never checked in. */
function buildFollowUp_(D) {
  const W = 8, rows = [], act = activeByGroup_(D);
  const row = function (v, k, cc, color) { rows.push([v, k, color ? cc : 0, color]); };
  const follow = [];
  D.students.forEach(function (s) {
    if (String(s[SC.STATUS - 1]) !== ACTIVE_STATUS) return;
    const key = String(s[SC.KEY - 1] || '').trim(), L = D.byGroup[s[SC.GROUP - 1]]; if (!key || !L) return;
    const held = L.filter(function (x) { return x.status === 'Held' && x.isLecture && x.tracked && startedBy_(s, x.day); });
    if (held.length < 2) return;
    let streak = 0; for (let i = held.length - 1; i >= 0 && !attended_(D, key, held[i].day); i--) streak++;
    if (streak < 2) return;
    const lastAt = held.filter(function (x) { return attended_(D, key, x.day); }).pop();
    const cur = held[held.length - 1].module, inMod = held.filter(function (x) { return x.module === cur; });
    const miss = inMod.filter(function (x) { return !attended_(D, key, x.day); }).length;
    follow.push([String(s[SC.NAME - 1]), String(s[SC.PHONE - 1]), String(s[SC.GROUP - 1]), cur, lastAt ? 'بطّل يحضر' : 'ما سجلش حضور خالص',
                 streak, lastAt ? shortDate_(lastAt.d) : '—', 'غاب ' + miss + ' من ' + inMod.length]);
  });
  const STOP = 'بطّل يحضر';
  follow.sort(function (a, b) { return (a[4] === b[4] ? 0 : a[4] === STOP ? -1 : 1) || b[5] - a[5] || a[2].localeCompare(b[2]) || a[0].localeCompare(b[0]); });
  const fg = {}; follow.forEach(function (f) { const x = fg[f[2]] = fg[f[2]] || [0, 0]; x[f[4] === STOP ? 0 : 1]++; });
  row(['متابعة الطلبة'], 'title');
  row(['آخر تحديث ' + Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, 'dd/MM/yyyy HH:mm') + ' · طلبة "' + ACTIVE_STATUS + '" غابوا آخر محاضرتين أو أكتر ورا بعض · الحضور محسوب من ' + (D.from ? shortDate_(D.from) : '—')], 'note');
  row([], 'gap'); row(['1 · لكل جروب'], 'section'); const cR1 = rows.length;
  row(['الجروب', 'بطّل يحضر', 'ما سجلش حضور خالص', 'طلبة بيدرسوا', 'نسبة اللي محتاجين متابعة'], 'head');
  const gk = Object.keys(fg).sort(function (a, b) { return (fg[b][0] + fg[b][1]) - (fg[a][0] + fg[a][1]); });
  if (!gk.length) row(['مفيش طلبة محتاجين متابعة.'], 'note');
  gk.forEach(function (g) { const t = (act[g] || []).length; row([g, fg[g][0], fg[g][1], t, t ? Math.round((fg[g][0] + fg[g][1]) / t * 100) + '%' : ''], 'Held', 2, fg[g][0] ? '#a52f2f' : null); });
  const stopped = follow.filter(function (f) { return f[4] === STOP; }).length;
  row([], 'gap'); row(['2 · الطلبة (' + follow.length + ') – بطّل يحضر: ' + stopped + ' · ما سجلش حضور خالص: ' + (follow.length - stopped)], 'section');
  row(['الطالب', 'الرقم', 'الجروب', 'المادة الحالية', 'الحالة', 'غاب ورا بعض', 'آخر حضور', 'في المادة دي'], 'head');
  follow.forEach(function (f) { row(f, f[4] === STOP ? 'Today' : 'Held', 6, f[5] >= 4 ? '#a52f2f' : '#9a6a00'); });
  writeReport_('Students Follow-up', '#a52f2f', W, rows, [260, 140, 190, 230, 150, 110, 110, 120], [
    { row: cR1, type: 'COLUMN', title: 'الطلبة المحتاجين متابعة – لكل جروب', header: ['الجروب', 'بطّل يحضر', 'ما سجلش حضور خالص'], stacked: true,
      colors: ['#c0392b', '#d99a1e'], data: gk.map(function (g) { return [g, fg[g][0], fg[g][1]]; }), h: 300 }
  ]);
  return follow.length;
}

/** v7.0: the old report tabs are replaced by Manager Summary, Attendance Report and Doctors Report. */
function removeOldReports_() {
  const ss = SpreadsheetApp.getActive(); let n = 0;
  ['Reports', 'Lecture Report', 'Evaluation Report'].forEach(function (name) { const sh = ss.getSheetByName(name); if (sh) { ss.deleteSheet(sh); n++; } });
  return n;
}

function refreshAttendance_(onlyIfDue) {
  const D = attendanceData_();
  refreshStudentAttendance_(D);
  const P = PropertiesService.getScriptProperties(), last = Number(P.getProperty('REPORTS_AT') || 0);
  if (onlyIfDue && Date.now() - last < 30 * 60000) return 0;   // the report tabs + charts are heavy: rebuilt every 30 minutes
  P.setProperty('REPORTS_AT', String(Date.now()));
  const runs = moduleRuns_(D), evals = attachEvaluations_(runs);
  buildDoctorsReport_(D, runs, evals);
  buildManagerSummary_(D, runs, evals);
  try { pathReportTab_(); } catch (e) { Logger.log('path report: ' + e); }
  return buildAttendanceReport_(D, runs);
}
function menuLiveAdd() { try { const u = String(sheet_(T.SETTINGS).getRange('B33').getValue() || '').trim(); const m = addMissingLiveRows_(u); PropertiesService.getScriptProperties().setProperty('LIVE_ADD', 'v8.8'); toast_('Lectures: ' + m); } catch (e) { toast_('Failed: ' + e.message); } }
function menuDevClients() { try { apEnsure_(); devEnsureServices_(); toast_('Development clients: ' + syncDevClients_()); } catch (e) { toast_('Development clients failed: ' + e.message); try { sheet_(T.SETTINGS).getRange('B42').setValue('Sync failed: ' + e.message); } catch (x) {} } }
function menuAttendance() { toast_('Attendance updated · groups ' + refreshAttendance_()); }

/* =========================== CS Schedule → Lectures (one-time import, Settings!B35) =========================== */
const CS_TABS = {   // CS Schedule tab (trimmed) → group in the Lectures tab
  'MBA-EIU G1': 'EIU MBA · Group 1', 'MBA- EIU G2': 'EIU MBA · Group 2', 'MBA-EIU G3': 'EIU MBA · Group 3 New', 'Golden': 'EIU MBA · Golden',
  'Health care': 'EIU MBA Healthcare', 'BBA TOP UP': 'EIU BBA · Top Up', 'MBA-AITU': 'AITU MBA', 'DBA': 'DBA (EIU / AITU)',
  'innovation of L7 Qualifi': 'Qualifi L7 · Innovation', 'leadership of L7 Qualifi': 'Qualifi L7 · Leadership', 'l4&5': 'Qualifi L4 & L5',
  'BBA 1Y': 'BBA 1st Year', 'BBA 2Y': 'BBA 2nd Year'
};
/** Possible dates of a cell: real dates may have day and month swapped (typed 2/8 → read 8 Feb or 2 Aug). */
function dateCands_(v) {
  const out = [];
  const add = function (y, m, d) { const x = new Date(y, m - 1, d); if (x.getFullYear() === y && x.getMonth() === m - 1 && x.getDate() === d && y > 2019 && y < 2031) out.push(x); };
  if (v instanceof Date) {
    const y = v.getFullYear(), m = v.getMonth() + 1, d = v.getDate();
    add(y, m, d); if (d <= 12 && d !== m) add(y, d, m);
  } else if (typeof v === 'string') {
    let m = v.trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) { add(+m[1], +m[2], +m[3]); return out; }
    m = v.trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (m) { const a = +m[1], b = +m[2], y = +m[3]; add(y, b, a); if (a <= 12 && a !== b) add(y, a, b); }
  }
  return out;
}
function pickDate_(cands, prev) {
  if (!cands.length) return null;
  if (!prev) return cands[0];
  const slack = 7 * 864e5;
  return cands.slice().sort(function (a, b) {
    const pa = a < prev - slack ? 1 : 0, pb = b < prev - slack ? 1 : 0;
    return pa - pb || Math.abs(a - prev) - Math.abs(b - prev);
  })[0];
}
/** "6:00 PM - 9:00 PM (Cairo Time)", "Mon & Thr 5-8", "Sat & Tue 6 PM" (+ "2 hours") → "06:00 PM - 09:00 PM". '' when the text has no time (e.g. "الاحد"). */
function csTime_(text, hours) {
  const t = String(text || '');
  const f = function (h, m, ap) { h = +h; m = +(m || 0); if (!ap) ap = h >= 1 && h <= 11 ? 'PM' : 'AM'; ap = ap.toUpperCase(); return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m + ' ' + ap; };
  let m = t.match(/(\d{1,2})(?::(\d{2}))?\s*(am|pm)?\s*[-–]\s*(\d{1,2})(?::(\d{2}))?\s*(am|pm)?/i);
  if (m && +m[1] <= 12 && +m[4] <= 12) return f(m[1], m[2], m[3] || m[6]) + ' - ' + f(m[4], m[5], m[6] || m[3]);
  m = t.match(/(\d{1,2})(?::(\d{2}))?\s*(am|pm)\b/i);
  if (!m || +m[1] > 12) return '';
  const start = f(m[1], m[2], m[3]);
  if (!hours) return start;
  let h = (+m[1] % 12) + (/pm/i.test(m[3]) ? 12 : 0) + hours;
  return start + ' - ' + f(((h + 11) % 12) + 1, m[2], h % 24 >= 12 ? 'PM' : 'AM');
}
function readCsTab_(sh, group, timeline, catalog) {
  const v = sh.getDataRange().getValues();
  const title = v.slice(0, 2).map(function (r) { return r.join(' '); }).join(' ');
  let hi = -1;
  for (let i = 0; i < Math.min(v.length, 5); i++) if (v[i].some(function (c) { return /محاضرة\s*\d+/.test(String(c)) && !/^نهاية/.test(String(c).trim()); })) { hi = i; break; }
  if (hi === -1) return [];
  const h = v[hi].map(function (c) { return String(c).replace(/\s+/g, ' ').trim(); });
  const lcols = [];
  h.forEach(function (c, i) { const m = c.match(/^محاضرة\s*(\d+)/); if (m) lcols.push([i, +m[1]]); });
  const mc = h.findIndex(function (c) { return /اسم الماد/.test(c); }), sc = h.findIndex(function (c) { return /تاريخ البداي/.test(c); });
  const out = [];
  let prev = null;
  for (let r = hi + 1; r < v.length; r++) {
    const row = v[r];
    if (row.some(function (c) { return /Program Summary|نظام البريك/.test(String(c)); })) break;
    let modRaw = mc >= 0 ? String(row[mc] || '').trim() : '';
    if (!modRaw) continue;
    const module = GROUP_ACC[group] ? lectureModule_(modRaw, group, catalog) : modRaw.replace(/\s+/g, ' ');
    const start = sc >= 0 ? pickDate_(dateCands_(row[sc]), prev) : null;
    if (start) prev = prev && start < prev - 30 * 864e5 ? prev : start;
    let doc = '';
    for (let c = lcols.length ? lcols[lcols.length - 1][0] + 1 : 0; c < row.length; c++) { const o = officialDoctor_(row[c]); if (o) { doc = o; break; } }
    const tl = timeline[group + '|' + module] || timeline[group + '|' + modRaw.toLowerCase()] || {};
    if (!doc && tl.doctor) doc = tl.doctor;
    const time = tl.time || csTime_(v[0].join(' '), tl.hours);
    const seen = {};
    lcols.forEach(function (lc) {
      const d = pickDate_(dateCands_(row[lc[0]]), prev);
      if (!d) return;
      const k = cairoKey_(d); if (seen[k]) return; seen[k] = true;
      prev = d;
      out.push({ d: d, group: group, module: module, n: lc[1], doctor: doc, time: time });
    });
  }
  return out;
}
function readTimeline_(book, catalog) {
  const sh = book.getSheets().filter(function (s) { return s.getName().trim() === 'Timeline'; })[0];
  const out = {};
  if (!sh) return out;
  const v = sh.getDataRange().getValues();
  const hi = v.findIndex(function (r) { return String(r[0]).trim() === 'Group' && /Subject/.test(String(r[1])); });
  if (hi === -1) return out;
  for (let r = hi + 1; r < v.length; r++) {
    const g = CS_TABS[String(v[r][0]).trim()]; if (!g) continue;
    const raw = String(v[r][1] || '').trim();
    const mod = GROUP_ACC[g] ? lectureModule_(raw, g, catalog) : raw;
    const hrs = Number((String(v[r][6] || '').match(/^(\d+)\s*hour/) || [])[1]) || 0;
    const rec = { time: csTime_(v[r][5], hrs), hours: hrs, doctor: officialDoctor_(v[r][8]) || '' };
    out[g + '|' + mod] = rec; out[g + '|' + raw.toLowerCase()] = rec;
  }
  return out;
}
/** Builds every lecture from the CS Schedule tables (from fromDate on) into the Lectures tab.
 *  CS Schedule decides dates + doctor; links, attendance and Operation columns of the same group + module + lecture # are kept. */
function importCsSchedule_(url, fromDate) {
  const book = SpreadsheetApp.openByUrl(url), catalog = getCatalog_(), from = cairoKey_(fromDate || new Date(2026, 9, 1));
  const timeline = readTimeline_(book, catalog);
  let cs = [];
  book.getSheets().forEach(function (sh) { const g = CS_TABS[sh.getName().trim()]; if (g) cs = cs.concat(readCsTab_(sh, g, timeline, catalog)); });
  cs = cs.filter(function (x) { return cairoKey_(x.d) >= from; });
  const lw = sheet_(LEC.SHEET), last = lastDataRow_(lw, 1);
  const cur = last >= 2 ? lw.getRange(2, 1, last - 1, LX.WIDTH).getValues().filter(function (r) { return r[0] instanceof Date; }) : [];
  const byNum = {}, byDay = {};
  cur.forEach(function (r, i) {
    byNum[r[3] + '|' + r[4] + '|' + String(r[5]).trim()] = i;
    byDay[r[3] + '|' + cairoKey_(r[0])] = i;
  });
  // the team's own sheets only: CS Schedule first, then the old live lectures sheet (Settings!B33) for the same group + day
  const oldDay = {}, modDoc = {};
  const oldUrl = String(sheet_(T.SETTINGS).getRange('B33').getValue() || '').trim();
  if (oldUrl) try { readOldLectures_(oldUrl, fromDate).out.forEach(function (o) { const r = o.row; oldDay[r[3] + '|' + cairoKey_(r[0])] = { doctor: officialDoctor_(r[6]) || r[6], time: /\d/.test(String(r[2])) ? r[2] : '' }; }); } catch (e) { Logger.log('old lectures sheet: ' + e); }
  cs.forEach(function (x) { if (x.doctor) modDoc[x.group + '|' + x.module] = x.doctor; });
  cs.forEach(function (x) { const o = oldDay[x.group + '|' + cairoKey_(x.d)]; if (o && o.doctor && !modDoc[x.group + '|' + x.module]) modDoc[x.group + '|' + x.module] = o.doctor; });
  const used = {}, rows = [];
  let carried = 0;
  cs.forEach(function (x) {
    let i = byNum[x.group + '|' + x.module + '|' + x.n];
    if (i !== undefined && Math.abs(cur[i][0] - x.d) > 21 * 864e5) i = undefined;   // same module again a year later = a different lecture
    if (i === undefined || used[i]) i = byDay[x.group + '|' + cairoKey_(x.d)];
    const old = i !== undefined && !used[i] ? cur[i] : null;
    if (old) { used[i] = true; carried++; }
    const r = old ? old.slice() : new Array(LX.WIDTH).fill('');
    r[0] = x.d; r[1] = ''; r[3] = x.group; r[4] = x.module; r[5] = x.n;
    const od = oldDay[x.group + '|' + cairoKey_(x.d)] || {};
    r[6] = x.doctor || od.doctor || modDoc[x.group + '|' + x.module] || '';   // never a guess: blank when neither sheet names the doctor
    r[2] = x.time || od.time || ''; r[7] = '';
    r[15] = ''; r[16] = ''; r[18] = '';
    rows.push(r);
  });
  let kept = 0;
  cur.forEach(function (r, i) {
    if (used[i]) return;
    const c = r.slice(); c[1] = ''; c[15] = ''; c[16] = ''; c[18] = '';
    if (CS_GROUPS_[c[3]] && cairoKey_(c[0]) >= from && !/ليست في CS Schedule/.test(String(c[13]))) c[13] = (c[13] ? c[13] + ' · ' : '') + 'ليست في CS Schedule – راجعها';
    rows.push(c); kept++;
  });
  rows.sort(function (a, b) { return a[0] - b[0] || String(a[3]).localeCompare(String(b[3])) || String(a[2]).localeCompare(String(b[2])); });
  writeLectures_(rows, false);
  return 'CS Schedule lectures ' + cs.length + ' · links/attendance carried over ' + carried + ' · other lectures kept ' + kept;
}

/** One-time repair of the first v6.2 import (Settings!D36): a Zoom link had been copied to every later lecture of the same module
 *  (also the same module a year later), and a few lectures were matched to the same module # a year later.
 *  Future lectures cannot have attendance, and a link is only kept for 60 days from the first lecture that used it in the group. */
function repairCsImport_() {
  const lw = sheet_(LEC.SHEET), last = lastDataRow_(lw, 1);
  if (last < 2) return 'nothing to repair';
  const v = lw.getRange(2, 1, last - 1, LX.WIDTH).getValues();
  const today = cairoKey_(new Date()), DAY = 864e5;
  const idx = v.map(function (r, i) { return i; }).filter(function (i) { return v[i][0] instanceof Date; })
    .sort(function (a, b) { return v[a][0] - v[b][0]; });
  const first = {};
  idx.forEach(function (i) { const z = String(v[i][8] || '').trim(); if (/^https?:/.test(z) && !first[v[i][3] + '|' + z]) first[v[i][3] + '|' + z] = v[i][0]; });
  const MOVE = [8, 9, 10, 11, 12, 14, 19];   // I J K L M O T
  let moved = 0, cleared = 0;
  idx.forEach(function (i) {
    const r = v[i]; if (cairoKey_(r[0]) <= today || /⚠ راجع التاريخ/.test(String(r[13]))) return;   // ⚠ = typo date from the old sheet, left for review
    const z = String(r[8] || '').trim();
    const lateLink = /^https?:/.test(z) && r[0] - first[r[3] + '|' + z] > 60 * DAY;
    const att = String(r[10] || '').trim() || String(r[11] || '').trim() || r[14];
    if (!lateLink && !att) return;
    if (att) {   // a past lecture's record that landed on a future date → give it back to the same group / module / lecture # before today
      const t = idx.filter(function (j) { const x = v[j]; return x[3] === r[3] && x[4] === r[4] && String(x[5]) === String(r[5]) && cairoKey_(x[0]) <= today && !String(x[10] || '').trim() && !String(x[8] || '').trim(); }).pop();
      if (t !== undefined) { MOVE.forEach(function (c) { if (!String(v[t][c] || '').trim()) v[t][c] = r[c]; }); moved++; }
    }
    MOVE.forEach(function (c) { r[c] = ''; });
    cleared++;
  });
  const put = function (c, n) { lw.getRange(2, c, v.length, n).setValues(v.map(function (r) { return r.slice(c - 1, c - 1 + n); })); };
  put(9, 5); put(15, 1); put(20, 1);
  return 'future lectures cleaned ' + cleared + ' · records moved back ' + moved;
}
const CS_GROUPS_ = (function () { const o = {}; Object.keys(CS_TABS).forEach(function (k) { o[CS_TABS[k]] = true; }); return o; })();

/* ---------- cancellation reasons ---------- */
const CANCEL_REASONS = ['اعتذر الدكتور', 'تأخر الدكتور عن الدخول', 'ظرف مرضي للدكتور', 'ظرف طارئ / شخصي للدكتور', 'الدكتور لم يحضر بدون إبلاغ', 'سفر الدكتور',
  'لينك الزوم لا يعمل / لينك خطأ', 'اللينك لم يُرسل للطلاب أو اتبعت متأخر', 'الدكتور خرج من اللينك ولم يرجع', 'مشكلة في حساب الزوم (اشتراك / تعارض اجتماعين)',
  'مشكلة في الصوت أو الصورة أو مشاركة الشاشة', 'انقطاع الإنترنت', 'انقطاع الكهرباء', 'عدم حضور الطلاب / عدد قليل جداً', 'تأجيل بطلب من الطلاب',
  'تأجيل بقرار من الإدارة', 'تعارض مع محاضرة أو امتحان آخر', 'إجازة / عطلة رسمية', 'مناسبة دينية / أعياد', 'لم يتم تعيين دكتور / تغيير الدكتور',
  'المادة العلمية غير جاهزة', 'خطأ في الجدول أو الميعاد', 'أخرى'];
function ensureReasons_() {
  const ls = sheet_('Lists'); if (!ls) return;
  const last = lastDataRow_(ls, 13);
  const cur = last >= 2 ? ls.getRange(2, 13, last - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }).filter(String) : [];
  const old = ['تأخر عن الدخول', 'ظرف مرضي', 'ظرف طارئ / شخصي', 'مشكلة تقنية (زوم / إنترنت)', 'تأجيل بطلب من الإدارة أو الطلاب'];   // first list – kept for old records
  const list = CANCEL_REASONS.slice(0, -1).concat(cur.filter(function (x) { return CANCEL_REASONS.indexOf(x) === -1 && old.indexOf(x) === -1 && x !== 'أخرى'; })).concat(['أخرى']);
  ls.getRange(2, 13, Math.max(cur.length, list.length) + 2, 1).clearContent();
  ls.getRange(2, 13, list.length, 1).setValues(list.map(function (x) { return [x]; }));
  ls.getRange(1, 13).setValue('Reason (not attended)');
}

/* ---------- WhatsApp follow-up numbers (Settings!A21:C40, created by setup – editable) ---------- */
const WA_HEAD_ROW = 20;
const WA_DEFAULT = [
  ['EIU MBA', '201044015001', 'EIU MBA · Group 1, EIU MBA · Group 2, EIU MBA · Group 3 New, EIU MBA'],
  ['EIU MBA – Golden', '201158882585', 'EIU MBA · Golden'],
  ['EIU MBA Healthcare', '201070988446', 'EIU MBA Healthcare'],
  ['EIU BBA', '201070988446', 'EIU BBA · Top Up, EIU BBA 1 Year, EIU BBA 3 Years'],
  ['AITU MBA', '201070988446', 'AITU MBA'],
  ['DBA (EIU / AITU)', '201158882585', 'DBA (EIU / AITU), EIU DBA, AITU DBA'],
  ['Pre DBA', '201158882585', ''],
  ['UK programmes', '201158882585', ''],
  ['Qualifi', '201158882585', 'QUALIFI level 7 leadership, QUALIFI level 7 Innovation']
];
function readContacts_() {
  const set = sheet_(T.SETTINGS);
  let rows = WA_DEFAULT;
  if (set && String(set.getRange(WA_HEAD_ROW, 1).getValue()).trim() === 'WhatsApp follow-up') {
    const v = [];   // the table ends at its first empty row (links further down in Settings are not phone numbers)
    const all = set.getRange(WA_HEAD_ROW + 1, 1, 12, 3).getValues();
    for (let i = 0; i < all.length && String(all[i][0]).trim(); i++) {
      const num = String(all[i][1]).replace(/\D/g, '');
      if (!(all[i][1] instanceof Date) && !/https?:/.test(String(all[i][1])) && num.length >= 8 && num.length <= 15) v.push(all[i]);
    }
    if (v.length) rows = v;
  }
  return rows.map(function (r) {
    return { label: String(r[0]).trim(), whatsapp: String(r[1]).replace(/\D/g, ''),
             keys: String(r[2] || '').split(',').map(function (k) { return k.trim(); }).filter(String) };
  });
}
function contactFor_(contacts, group, acc) {
  const byGroup = contacts.filter(function (c) { return c.keys.indexOf(group) !== -1; })[0];
  const byAcc = contacts.filter(function (c) { return c.keys.indexOf(acc) !== -1; })[0];
  return (byGroup || byAcc || contacts[0] || { whatsapp: '' }).whatsapp;
}
function writeContactsTable_() {
  const set = sheet_(T.SETTINGS);
  if (!set || String(set.getRange(WA_HEAD_ROW, 1).getValue()).trim() === 'WhatsApp follow-up') return;
  set.getRange(WA_HEAD_ROW, 1, 1, 3).setValues([['WhatsApp follow-up', 'WhatsApp number', 'Groups / accreditations it follows (comma separated)']])
     .setFontWeight('bold').setBackground('#0c2536').setFontColor('#ffffff');
  set.getRange(WA_HEAD_ROW + 1, 2, WA_DEFAULT.length, 1).setNumberFormat('@');
  set.getRange(WA_HEAD_ROW + 1, 1, WA_DEFAULT.length, 3).setValues(WA_DEFAULT);
  set.getRange(WA_HEAD_ROW + 1 + WA_DEFAULT.length, 1).setValue('Column A is what an unregistered student chooses on the landing page. Add rows here for new programmes.');
}

/** "My next lecture": for a registered phone → next lectures of each group (date, time, doctor, Zoom link)
 *  + the WhatsApp number that follows that group. Unknown phone → found:false + the list of follow-up numbers. */
/** Lecture times are written in Egypt time; the student page shows them in Saudi time ("06:00 PM - 08:00 PM" → "07:00 PM - 09:00 PM" in winter). */
function ksaTime_(day, text) {
  const raw = String(text || '').trim();
  const m = raw.replace(/\(.*?\)/g, '').replace(/(AM|PM)\s*(?::|to|الى|إلى)\s*/i, '$1 - ').match(/(\d{1,2})(?::(\d{2}))?\s*(AM|PM)?\s*[-–]\s*(\d{1,2})(?::(\d{2}))?\s*(AM|PM)/i);
  if (!m || !(day instanceof Date)) return raw;
  const ymd = Utilities.formatDate(day, SETTINGS.TIMEZONE, 'yyyy-MM-dd');
  const conv = function (h, mi, ap) {
    h = +h % 12 + (/pm/i.test(ap) ? 12 : 0);
    const d = Utilities.parseDate(ymd + ' ' + (h < 10 ? '0' : '') + h + ':' + (mi || '00'), SETTINGS.TIMEZONE, 'yyyy-MM-dd HH:mm');
    return Utilities.formatDate(d, 'Asia/Riyadh', 'hh:mm a');
  };
  return conv(m[1], m[2], m[3] || m[6]) + ' - ' + conv(m[4], m[5], m[6]);
}

/** The next lecture and the rest of the same module run (a gap over 60 days = the same module again much later, not shown). */
function moduleRun_(list) {
  // start from the next real lecture (an "Assign" row of the module that is ending is not "the next lecture")
  const first = list.findIndex(function (x) { return /^\d+$/.test(String(x.n).trim()); });
  if (first > 0) list = list.slice(first);
  if (!list.length) return [];
  const out = [list[0]], DAY = 864e5;
  const t = function (x) { return new Date(x.date + 'T00:00:00Z').getTime(); };
  for (let i = 1; i < list.length; i++) {
    if (list[i].rescheduled) { out.push(list[i]); continue; }   // a moved lecture is always shown
    if (list[i].module !== list[0].module) continue;
    if (t(list[i]) - t(out[out.length - 1]) > 60 * DAY) break;
    out.push(list[i]);
  }
  return out.sort(function (a, b) { return a._k - b._k; });
}
function lectureInfo_(rawPhone) {
  const contacts = readContacts_();
  const base = lookup_(rawPhone);
  if (!base.ok) return base;
  if (!base.found) return { ok: true, found: false, contacts: contacts.map(function (c) { return { label: c.label, whatsapp: c.whatsapp }; }) };
  const lw = sheet_(LEC.SHEET);
  const last = lw ? lastDataRow_(lw, 1) : 1;
  const rows = last >= 2 ? lw.getRange(2, 1, last - 1, 20).getValues() : [];
  const tz = SETTINGS.TIMEZONE, dk = function (d) { return Number(Utilities.formatDate(d, tz, 'yyyyMMdd')); };
  const today = dk(new Date());
  const enrollments = base.enrollments.map(function (e) {
    const list = [], recs = [];
    rows.forEach(function (r) {
      if (!(r[0] instanceof Date) || String(r[3]) !== e.group) return;
      const drive = String(r[9] || '').trim();
      if (/^https?:/.test(drive) && dk(r[0]) <= today) recs.push({ date: Utilities.formatDate(r[0], tz, 'yyyy-MM-dd'), module: String(r[4] || ''), n: r[5] || '', drive: drive, _k: dk(r[0]) });
      const status = /^cancel/i.test(String(r[19] || '').trim()) ? 'لم يحضر' : String(r[10] || '').trim(), moved = r[12] instanceof Date ? r[12] : null;
      let d = r[0], rescheduled = false;
      if (status === 'لم يحضر') { if (!moved) return; d = moved; rescheduled = true; }
      if (dk(d) < today) return;
      const z = String(r[8] || '').trim();
      list.push({ date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'), time: ksaTime_(d, r[2]), module: String(r[4] || e.module || ''),
                  n: r[5] || '', doctor: String(r[6] || '') || e.doctor || '', zoom: /^https?:/.test(z) ? z : '',
                  today: dk(d) === today, rescheduled: rescheduled, _k: dk(d) });
    });
    list.sort(function (a, b) { return a._k - b._k; });
    return { accreditation: e.accreditation, group: e.group, module: e.module, doctor: e.doctor, whatsapp: contactFor_(contacts, e.group, e.accreditation),
             // the next lecture and every lecture after it in the same module (to the module's last lecture)
             lectures: moduleRun_(list).map(function (x) { delete x._k; return x; }),
             recordings: recs.sort(function (a, b) { return b._k - a._k; }).slice(0, 6).map(function (x) { delete x._k; return x; }) };
  });
  return { ok: true, found: true, name: base.name, enrollments: enrollments };
}

/** All enrollments for a phone, with the module running now + doctor. */
function findStudent_(phone) {
  const key = phoneKey_(phone);
  if (!key) return [];
  const st = sheet_(T.STUDENTS);
  const last = lastDataRow_(st, SC.PHONE);
  if (last < 2) return [];
  const rows = st.getRange(2, 1, last - 1, SC.STATUS).getValues().filter(function (r) { return String(r[SC.KEY - 1]) === key || phoneKey_(r[SC.PHONE - 1]) === key; });
  if (!rows.length) return [];
  const today = dayStart_(new Date());
  const sched = readSchedule_();
  const byId = {}; sched.forEach(function (s) { byId[s.id] = s; });
  // running plan rows from Tracker
  const tr = sheet_(T.TRACKER), trLast = lastDataRow_(tr, 1);
  const trVals = trLast >= 2 ? tr.getRange(2, 1, trLast - 1, 8).getValues() : [];
  return rows.map(function (r) {
    const acc = String(r[SC.ACC - 1]), group = String(r[SC.GROUP - 1]), mk = key + '|' + acc;
    let s = null;
    for (let i = 0; i < trVals.length; i++) {
      if (trVals[i][0] !== mk) continue;
      const x = byId[trVals[i][6]];
      if (x && x.start <= today && today <= x.end) { s = x; break; }
    }
    const active = String(r[SC.STATUS - 1]) === ACTIVE_STATUS;
    if (!s && active) s = sched.filter(function (x) { return x.group === group && x.accs.indexOf(acc) !== -1 && x.start <= today && today <= x.end; })[0] || null;
    return { acc: acc, group: group, name: maskName_(r[SC.NAME - 1]), active: String(r[SC.STATUS - 1]) === ACTIVE_STATUS,
             module: s ? s.module : '', doctor: s ? s.doctor : '', moduleEnds: s ? dayKeyToIso_(s.end) : '' };
  });
}

/** "Alghamdi, Faten Sami S" → "Faten A." · "Ahmed Mohsen Hassouna" → "Ahmed M." */
function maskName_(name) {
  let s = String(name || '').replace(/\s+/g, ' ').trim();
  if (s.indexOf(',') > 0) { const p = s.split(','); s = (p[1] + ' ' + p[0]).trim(); }
  const t = s.split(' ').filter(String);
  if (!t.length) return '';
  return t[0].charAt(0).toUpperCase() + t[0].slice(1) + (t[1] ? ' ' + t[1].charAt(0).toUpperCase() + '.' : '');
}

/* =========================== Students DB sync =========================== */
/** Trigger target (every 10 min) + menu. Copies the Students DB into the Students tab. */
function syncFromDb() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return; }
  try { return syncFromDb_(); } finally { lock.releaseLock(); }
}

function syncFromDb_() {
  const set = sheet_(T.SETTINGS);
  const url = String(set.getRange('B8').getValue() || '').trim();
  if (!url) throw new Error('Paste the Students DB link in Settings!B8 first.');
  const dbSheet = SpreadsheetApp.openByUrl(url).getSheetByName(T.DB_TAB);
  if (!dbSheet) throw new Error('Tab "' + T.DB_TAB + '" not found in the Students DB.');

  const dbLast = lastDataRow_(dbSheet, DBC.PHONE);
  const dbVals = dbLast >= 2 ? dbSheet.getRange(2, 1, dbLast - 1, DBC.WIDTH).getValues() : [];
  const st = sheet_(T.STUDENTS);
  const stLast = Math.max(lastDataRow_(st, SC.PHONE), lastDataRow_(st, SC.KEYMATCH));
  const stVals = stLast >= 2 ? st.getRange(2, 1, stLast - 1, SC.WIDTH).getValues() : [];
  const opsRow = {}; stVals.forEach(function (v, i) { if (v[0]) opsRow[v[0]] = i + 2; });

  const seen = {}, appended = [];
  let updated = 0, duplicates = 0;
  dbVals.forEach(function (d) {
    const key = phoneKey_(d[DBC.PHONE - 1]), acc = String(d[DBC.ACC - 1] || '').trim();
    if (!key || !acc) return;
    const mk = key + '|' + acc;
    if (seen[mk]) { duplicates++; return; }
    seen[mk] = true;
    const want = [mk, key, String(d[DBC.PHONE - 1]).replace(/\.0+$/, ''), String(d[DBC.NAME - 1] || '').trim(), acc,
                  String(d[DBC.GROUP - 1] || '').trim(), String(d[DBC.STATUS - 1] || '').trim(),
                  String(d[DBC.START_MOD - 1] || '').trim(), d[DBC.START_DATE - 1] instanceof Date ? d[DBC.START_DATE - 1] : ''];
    const note = String(d[DBC.NOTES - 1] || ''), source = String(d[DBC.SOURCE - 1] || '') || 'Students DB', done = String(d[DBC.DONE - 1] || '');
    const r = opsRow[mk];
    if (r) {
      const cur = stVals[r - 2];
      const same = want.every(function (v, i) { return sameVal_(v, cur[i]); }) && sameVal_(note, cur[SC.NOTE - 1]) && sameVal_(source, cur[SC.SOURCE - 1]) && sameVal_(done, cur[SC.DONE - 1]);
      if (!same) {
        st.getRange(r, SC.KEY, 1, 2).setNumberFormat('@');
        st.getRange(r, 1, 1, 9).setValues([want]);
        st.getRange(r, SC.NOTE, 1, 3).setValues([[note, source, done]]);
        updated++;
      }
    } else {
      appended.push({ want: want, note: note, source: source, done: done });
    }
  });
  // rows that disappeared from the DB
  let removed = 0;
  stVals.forEach(function (v, i) {
    if (v[0] && !seen[v[0]] && v[SC.STATUS - 1] !== 'Removed from DB') { st.getRange(i + 2, SC.STATUS).setValue('Removed from DB'); removed++; }
  });
  // append new students
  if (appended.length) {
    const start = stLast + 1;
    st.getRange(start, SC.KEY, appended.length, 2).setNumberFormat('@');
    st.getRange(start, 1, appended.length, 9).setValues(appended.map(function (a) { return a.want; }));
    st.getRange(start, SC.NOTE, appended.length, 3).setValues(appended.map(function (a) { return [a.note, a.source, a.done]; }));
    // live formulas only for active students (keeps the sheet fast); syncStudents_ adds them if a student becomes active later
    st.getRange(start, SC.F_FIRST, appended.length, SC.F_COUNT).setFormulas(appended.map(function (a, i) {
      return a.want[SC.STATUS - 1] === ACTIVE_STATUS ? studentFormulas_(start + i) : ['', '', '', '', '', '', ''];
    }));
    if (start > 2) st.getRange(2, 1, 1, SC.WIDTH).copyTo(st.getRange(start, 1, appended.length, SC.WIDTH), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
  const planned = syncStudents_();
  if (planned || appended.length) refreshDashboard_();
  const summary = 'added ' + appended.length + ' · updated ' + updated + ' · removed ' + removed + ' · plans ' + planned + (duplicates ? ' · duplicates skipped ' + duplicates : '');
  set.getRange('B9').setValue(new Date()).setNumberFormat('dd/mm/yyyy hh:mm');
  set.getRange('C9').setValue(summary);
  try { refreshAttendance_(true); } catch (e) { Logger.log('attendance: ' + e); }
  return summary;
}

function sameVal_(a, b) {
  // dates are compared by calendar day (Cairo) – the same date can come back from Sheets a few hours/seconds apart
  if (a instanceof Date || b instanceof Date) {
    const ka = a instanceof Date ? Utilities.formatDate(a, SETTINGS.TIMEZONE, 'yyyyMMdd') : String(a == null ? '' : a).trim();
    const kb = b instanceof Date ? Utilities.formatDate(b, SETTINGS.TIMEZONE, 'yyyyMMdd') : String(b == null ? '' : b).trim();
    return ka === kb;
  }
  return String(a == null ? '' : a).trim() === String(b == null ? '' : b).trim();
}

/** Creates Tracker plans for active students that are not imported and have no plan yet. */
function syncStudents_() {
  const st = sheet_(T.STUDENTS), tr = sheet_(T.TRACKER);
  const last = lastDataRow_(st, SC.PHONE);
  if (last < 2) return 0;
  const vals = st.getRange(2, 1, last - 1, SC.WIDTH).getValues();
  const fJ = st.getRange(2, SC.F_FIRST, last - 1, 1).getFormulas();
  const trLast = lastDataRow_(tr, 1);
  const tracked = {};
  if (trLast >= 2) tr.getRange(2, 1, trLast - 1, 1).getValues().forEach(function (x) { tracked[x[0]] = true; });
  const sched = readSchedule_(), catalog = getCatalog_();
  const newTracker = [];
  let added = 0;
  vals.forEach(function (v, i) {
    const r = i + 2, mk = v[0], key = String(v[SC.KEY - 1]), acc = String(v[SC.ACC - 1]), group = String(v[SC.GROUP - 1]);
    if (!mk || !key || !acc) return;
    if (v[SC.STATUS - 1] !== ACTIVE_STATUS) return;
    if (!fJ[i][0]) st.getRange(r, SC.F_FIRST, 1, SC.F_COUNT).setFormulas([studentFormulas_(r)]);
    if (tracked[mk]) return;
    const startMod = String(v[SC.START_MOD - 1] || '').trim(), startDate = v[SC.START_DATE - 1];
    let plan, note = '';
    if (!sched.some(function (s) { return s.group === group; })) { plan = []; note = 'No schedule for group "' + group + '" – attendance only'; }
    else if (startMod && startDate instanceof Date) {
      const done = String(v[SC.DONE - 1] || '').split(';').map(function (x) { return x.trim(); }).filter(String);
      plan = buildPlanFromHistory_(sched, group, startMod, startDate, done, goLiveDay_());
      if (!plan.length) note = 'No remaining modules (all modules already dated)';
    } else {
      plan = buildPlan_(sched, group, acc, startMod, new Date(), catalog[acc]);
      if (!plan.length) note = 'No current / upcoming module in the schedule';
    }
    if (!plan.length) { if (String(v[SC.PLAN_NOTE - 1]) !== note) st.getRange(r, SC.PLAN_NOTE).setValue(note); return; }
    if (groupsFor_(sched, acc).indexOf(group) === -1) st.getRange(r, SC.PLAN_NOTE).setValue('⚠ Group does not match accreditation');
    else if (v[SC.PLAN_NOTE - 1]) st.getRange(r, SC.PLAN_NOTE).setValue('');
    plan.forEach(function (p, k) { newTracker.push([mk, key, v[SC.NAME - 1], acc, group, k + 1, p.id, p.module]); });
    tracked[mk] = true; added++;
  });
  if (newTracker.length) appendTracker_(newTracker);
  return added;
}

/** Same rule as the initial import: last dated module (if it matches the schedule) + every module not yet taken. */
function buildPlanFromHistory_(sched, group, lastMod, lastDate, doneMods, goLive) {
  const rows = sched.filter(function (s) { return s.group === group; });
  const ld = dayStart_(lastDate), plan = [], seen = {};
  const done = {}; doneMods.forEach(function (m) { done[m] = true; }); done[lastMod] = true;
  const ldMs = new Date(lastDate).getTime();
  const match = rows.filter(function (s) { return s.module === lastMod && Math.abs(dayKeyToDate_(s.start).getTime() - ldMs) <= 21 * 864e5; })[0];
  let curStart = ld;
  if (match) { plan.push(match); seen[match.module] = true; curStart = match.start; }
  rows.forEach(function (s) { if (s.start > curStart && !done[s.module] && !seen[s.module]) { plan.push(s); seen[s.module] = true; } });
  return plan.filter(function (p) { return p.end >= goLive; });
}
function dayKeyToDate_(n) { const s = String(n); return new Date(Number(s.slice(0, 4)), Number(s.slice(4, 6)) - 1, Number(s.slice(6, 8))); }
function goLiveDay_() { const v = sheet_(T.SETTINGS).getRange('B2').getValue(); return v instanceof Date ? dayStart_(v) : dayStart_(new Date()); }

/* =========================== Import from EIU / AITU Data Input sheets =========================== */
const SRC = [
  { key: 'EIU',  url: 'B11', tab: 'MBA Groups',       acc: 'EIU MBA',            status: 'Studying Status', group: { col: 'Study Group' }, modules: true },
  { key: 'EIU',  url: 'B11', tab: 'Healthcare Group', acc: 'EIU MBA Healthcare', status: 'Status',          group: 'EIU MBA Healthcare', modules: true },
  { key: 'EIU',  url: 'B11', tab: 'Top Up Group',     acc: 'EIU BBA 1 Year',     status: 'Studying Status', group: 'EIU BBA · Top Up', modules: true },
  { key: 'EIU',  url: 'B11', tab: 'DBA Group',        acc: 'EIU DBA',            status: 'Studying Status', group: 'DBA (EIU / AITU)', modules: false },
  { key: 'AITU', url: 'B12', tab: 'MBA Groups',       acc: 'AITU MBA',           status: 'Studying Status', group: 'AITU MBA', modules: true },
  { key: 'AITU', url: 'B12', tab: 'DBA Group',        acc: 'AITU DBA',           status: 'Studying Status', group: 'DBA (EIU / AITU)', modules: false }
];
const EIU_GROUPS = { 'Group 1': 'EIU MBA · Group 1', 'Group 2': 'EIU MBA · Group 2', 'Group3 New': 'EIU MBA · Group 3 New', 'Golden Package': 'EIU MBA · Golden' };

function importSources() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(280000); }            // the 10-minute sync may be running – wait for it instead of skipping the hour
  catch (e) {
    try { sheet_(T.SETTINGS).getRange('D13').setValue('Hourly import skipped (sync busy) at ' + Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, 'dd/MM HH:mm')); } catch (x) {}
    return;
  }
  try { const r = importSources_(); sheet_(T.SETTINGS).getRange('D13').setValue('');
    try { syncDevClients_(); } catch (e) { sheet_(T.SETTINGS).getRange('B42').setValue('Sync failed: ' + e.message); }
    return r; } finally { lock.releaseLock(); }
}

/** Reads every student (any status) from the source sheets and upserts them into the Students DB. */
function importSources_() {
  const set = sheet_(T.SETTINGS), catalog = getCatalog_();
  const books = {};
  const students = [], seen = {};
  SRC.forEach(function (src) {
    const url = String(set.getRange(src.url).getValue() || '').trim();
    if (!url) return;
    if (!books[src.url]) books[src.url] = SpreadsheetApp.openByUrl(url);
    const sh = books[src.url].getSheetByName(src.tab);
    if (!sh) throw new Error('Tab "' + src.tab + '" not found in ' + src.key + ' sheet');
    const vals = sh.getDataRange().getValues();
    const head = vals[0].map(function (h) { return String(h).replace(/\s+/g, ' ').trim(); });
    const pc = head.indexOf('Phone'), nc = head.indexOf('Name'), sc = head.indexOf(src.status);
    if (pc < 0 || nc < 0 || sc < 0) throw new Error('Columns Phone / Name / ' + src.status + ' not found in ' + src.key + ' › ' + src.tab);
    const gc = src.group.col ? head.indexOf(src.group.col) : -1;
    const mcols = {};
    if (src.modules) {
      const used = {};
      head.forEach(function (h, i) { const m = catalogMatch_(catalog[src.acc] || [], h); if (m && !used[m]) { mcols[i] = m; used[m] = true; } });
    }
    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      const key = phoneKey_(row[pc]), name = String(row[nc] || '').replace(/\s+/g, ' ').trim();
      if (!key || !name || name.charAt(0) === '#') continue;   // skip broken rows (#N/A, #REF! …)
      const mk = key + '|' + src.acc;
      const status = String(row[sc] || '').replace(/\s+/g, ' ').trim();
      if (seen[mk] !== undefined) {
        // same phone twice in the source: keep the active row (e.g. an old "suspended" row + a new "Still studying" row)
        const prev = students[seen[mk]];
        if (prev[4] !== ACTIVE_STATUS && status === ACTIVE_STATUS) students.splice(seen[mk], 1, null);
        else continue;
      }
      seen[mk] = students.length;
      let group = src.group;
      if (gc >= 0) { const g = String(row[gc] || '').replace(/\s+/g, ' ').trim(); group = EIU_GROUPS[g] || ('EIU MBA · ' + (g || '(no group)')); }
      const dated = {};
      Object.keys(mcols).forEach(function (c) { const d = toDate_(row[c]); if (d) dated[mcols[c]] = d; });
      let lastMod = '', lastDate = '';
      Object.keys(dated).forEach(function (m) { if (!lastDate || dated[m] > lastDate) { lastDate = dated[m]; lastMod = m; } });
      students.push([phoneDigits_(row[pc]), name, src.acc, group, status,
                     lastMod, lastDate, src.key + ' · ' + src.tab, Object.keys(dated).sort().join('; ')]);
    }
  });
  for (let i = students.length - 1; i >= 0; i--) if (!students[i]) students.splice(i, 1);
  if (!students.length) return 'no source links in Settings!B11 / B12';

  // upsert into the Students DB
  const db = SpreadsheetApp.openByUrl(String(set.getRange('B8').getValue()).trim()).getSheetByName(T.DB_TAB);
  const last = lastDataRow_(db, DBC.PHONE);
  const cur = last >= 2 ? db.getRange(2, 1, last - 1, DBC.WIDTH).getValues() : [];
  const at = {}; cur.forEach(function (v, i) { at[phoneKey_(v[0]) + '|' + String(v[DBC.ACC - 1]).trim()] = i; });
  let added = 0, updated = 0, gone = 0;
  const inSource = {};
  students.forEach(function (s) {
    const mk = phoneKey_(s[0]) + '|' + s[2]; inSource[mk] = true;
    const i = at[mk];
    if (i === undefined) {
      cur.push([s[0], s[1], s[2], s[3], s[4], s[5], s[6], '', s[7], s[8]]); added++;
      return;
    }
    const row = cur[i];
    if (!AUTO_SOURCE.test(String(row[DBC.SOURCE - 1] || '')) && String(row[DBC.SOURCE - 1] || '') !== '') return;   // manual row: CS owns it
    const next = [s[0], s[1], s[2], s[3], s[4], s[5], s[6], row[DBC.NOTES - 1], s[7], s[8]];
    const diff = next.findIndex(function (v, k) { return !sameVal_(v, row[k]); });
    if (diff !== -1) {
      if (updated < 3) Logger.log('import update ' + mk + ' · column ' + (diff + 1) + ': "' + row[diff] + '" → "' + next[diff] + '"');
      cur[i] = next; updated++;
    }
  });
  cur.forEach(function (row, i) {
    const mk = phoneKey_(row[0]) + '|' + String(row[DBC.ACC - 1]).trim();
    if (AUTO_SOURCE.test(String(row[DBC.SOURCE - 1] || '')) && !inSource[mk] && row[DBC.STATUS - 1] !== 'Not in source') { cur[i][DBC.STATUS - 1] = 'Not in source'; gone++; }
  });
  if (added || updated || gone) {
    db.getRange(2, DBC.PHONE, cur.length, 1).setNumberFormat('@');
    db.getRange(2, 1, cur.length, DBC.WIDTH).setValues(cur);
  }
  const summary = 'sources ' + students.length + ' · new ' + added + ' · updated ' + updated + ' · not in source ' + gone;
  set.getRange('B13').setValue(new Date()).setNumberFormat('dd/mm/yyyy hh:mm');
  set.getRange('C13').setValue(summary);
  return summary + ' | Students DB → ' + syncFromDb_();
}

function catalogMatch_(mods, header) {
  const s = String(header || '').replace(/\s+/g, ' ').trim(), low = s.toLowerCase();
  if (!s) return null;
  const code = s.match(/([A-Za-z]{2,5})\s*(\d{3})/);
  if (code) {
    const c = (code[1] + code[2]).toUpperCase();
    for (let i = 0; i < mods.length; i++) if (mods[i].toUpperCase().replace(/\s/g, '').indexOf(c) === 0) return mods[i];
  }
  for (let i = 0; i < mods.length; i++) {
    const ml = mods[i].toLowerCase();
    if (low === ml || low.indexOf(ml) === 0 || low.indexOf(ml) !== -1) return mods[i];
  }
  for (let i = 0; i < mods.length; i++) {
    const name = mods[i].split(' - ').slice(1).join(' - ').toLowerCase();
    if (name && (low.indexOf(name.slice(0, 25)) !== -1 || name.indexOf(low.slice(0, 25)) !== -1)) return mods[i];
  }
  return null;
}
function toDate_(v) {
  let d = null;
  if (v instanceof Date) d = v;
  else if (typeof v === 'string') { const m = v.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])); }
  if (!d || isNaN(d.getTime())) return null;
  const y = d.getFullYear();
  return (y >= 2015 && y <= 2029) ? d : null;
}
function phoneDigits_(v) { return firstNumber_(v); }

function installTriggers_(withImport) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['syncFromDb', 'importSources', 'nightly'].indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncFromDb').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('nightly').timeBased().everyDays(1).atHour(3).create();
  if (withImport) ScriptApp.newTrigger('importSources').timeBased().everyHours(1).create();
}

/* =========================== Setup + menu =========================== */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(SETTINGS.TIMEZONE);
  if (!sheet_(SETTINGS.RESPONSES_SHEET)) throw new Error('Sheet "' + SETTINGS.RESPONSES_SHEET + '" not found');
  sheet_(SETTINGS.RESPONSES_SHEET).getRange('B:B').setNumberFormat('@');
  const set = sheet_(T.SETTINGS);
  if (!String(set.getRange('A33').getValue()).trim()) {
    set.getRange('A33:A34').setValues([['Old lectures sheet link (one-time import)'], ['Lectures import status']]).setFontWeight('bold');
    set.getRange('C33').setValue('Paste the link of the old live lectures sheet, then run setup: lectures from 1 Oct 2026 are copied into the Lectures tab (once).');
    set.getRange('B33').setBackground('#fff4d6');
  }
  const backup = String(set.getRange('B15').getValue() || '').trim();
  if (backup && !String(set.getRange('B16').getValue() || '').trim()) Logger.log(restoreFromBackup_(backup));
  repairResponses_();
  unifyDoctors_();
  ensureValidations_();
  const oldLec = String(set.getRange('B33').getValue() || '').trim();
  if (oldLec && !String(set.getRange('B34').getValue() || '').trim()) {
    const msg = importOldLectures_(oldLec, new Date(2026, 9, 1));
    set.getRange('B34').setValue('Imported ' + Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, 'dd/MM/yyyy HH:mm') + ' · ' + msg);
    Logger.log(msg);
  }
  // v6.6: the live lectures sheet is the only source – from today to its last row, as written (CS Schedule no longer imported)
  if (oldLec && PropertiesService.getScriptProperties().getProperty('LIVE_V') !== 'v6.7') {   // v6.7: copied again with the dates fixed
    PropertiesService.getScriptProperties().setProperty('LIVE_V', 'v6.7');
    const msg = importLiveFromToday_(oldLec);
    set.getRange('B34').setValue('Imported ' + Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, 'dd/MM/yyyy HH:mm') + ' · ' + msg);
    set.getRange('A33').setValue('Live lectures sheet link');
    set.getRange('C33').setValue('Lectures from the import day to the last row of this sheet were copied once, exactly as written (merged cells included). After that the Lectures tab is edited here.');
    if (String(set.getRange('A35').getValue()).trim()) set.getRange('C35').setValue('Not used any more (v6.6) – the live lectures sheet is the only source.');
    Logger.log(msg);
  }
  if (oldLec && PropertiesService.getScriptProperties().getProperty('LIVE_ADD') !== 'v8.8') {
    PropertiesService.getScriptProperties().setProperty('LIVE_ADD', 'v8.8');
    try { const m2 = addMissingLiveRows_(oldLec); set.getRange('B34').setValue(String(set.getRange('B34').getValue()) + ' · ' + Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, 'dd/MM HH:mm') + ' ' + m2); Logger.log(m2); }
    catch (e) { Logger.log('live sheet (new lectures): ' + e); PropertiesService.getScriptProperties().deleteProperty('LIVE_ADD'); }
  }
  ensureReasons_();
  lectureColumns_();
  ensureValidations_();
  Logger.log('Old report tabs removed: ' + removeOldReports_());
  apEnsure_(); devEnsureServices_();
  try { Logger.log('Development clients: ' + syncDevClients_()); } catch (e) { Logger.log('Development clients (run the menu item again later): ' + e); try { sheet_(T.SETTINGS).getRange('B42').setValue('Sync failed: ' + e.message); } catch (x) {} }
  Logger.log(syncScheduleFromLectures_());
  Logger.log('Attendance Report groups: ' + refreshAttendance_());
  writeContactsTable_();
  evalDashboard_();
  buildPathReport_();
  const url = String(set.getRange('B8').getValue() || '').trim();
  if (url) {
    const hasSources = !!String(set.getRange('B11').getValue() || set.getRange('B12').getValue() || '').trim();
    Logger.log(hasSources ? 'Import from EIU / AITU: ' + importSources_() : 'Students DB sync: ' + syncFromDb_());
    installTriggers_(hasSources);
    Logger.log('Auto-sync every 10 minutes: ON' + (hasSources ? ' · import from EIU / AITU every hour: ON' : ''));
  } else {
    Logger.log('⚠ Settings!B8 is empty – paste the Students DB link and run setup again.');
  }
  refreshDashboard_();
  Logger.log('Setup OK · ' + CODE_VERSION + ' · timezone ' + ss.getSpreadsheetTimeZone());
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('CEO Tools')
    .addItem('Import students from EIU / AITU sheets now', 'menuImport')
    .addItem('Sync students from Students DB now', 'menuSync')
    .addItem('Rebuild dashboard module list', 'menuDash')
    .addItem('Repair response formulas', 'menuRepair')
    .addItem('Lectures: sort by date + update Group Schedule', 'menuLectures')
    .addItem('Attendance: update report + Students columns now', 'menuAttendance')
    .addItem('Development path: copy Salma + Sabreen clients now', 'menuDevClients')
    .addItem('Lectures: add new lectures from the live sheet (once)', 'menuLiveAdd')
    .addToUi();
}
function menuImport() { toast_(importSources_()); }
function menuSync() { toast_('Students DB → ' + syncFromDb_()); }
function menuDash() { refreshDashboard_(); toast_('Dashboard rebuilt'); }
function menuRepair() { toast_('Formulas restored on ' + repairResponses_() + ' rows'); }

/* =========================== Assignments + Lectures automation =========================== */
// Assignments: A Account Manager | B Session Date | C Module | D Case | E Dr. Name | F Program | G Acc. | H Link | I Students
//              J Action | K Action time | L CS Action | M CS time | N Dr. Fees | O Fees time | P Note | Q Sent to doctor | R..V formulas
const AS = { SHEET: 'Assignments', AGENT: 1, DATE: 2, MODULE: 3, CASE: 4, DOCTOR: 5, LINK: 8, ACTION: 10, ACTION_AT: 11,
             CS: 12, CS_AT: 13, FEES: 14, FEES_AT: 15, NOTE: 16, SENT: 17, WIDTH: 17 };
// Lectures: K Doctor attended? | L Reason | M New date | O Recorded at
const LEC = { SHEET: 'Lectures', STATUS: 11, REASON: 12, NEWDATE: 13, AT: 15 };

/** Simple trigger: runs on every manual edit. */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet(), name = sh.getName();
    const r1 = Math.max(2, e.range.getRow()), r2 = e.range.getRow() + e.range.getNumRows() - 1;
    if (r2 < 2) return;
    const c1 = e.range.getColumn(), c2 = c1 + e.range.getNumColumns() - 1;
    const hit = function (c) { return c1 <= c && c2 >= c; };
    if (name === AS.SHEET) processAssignments_(r1, r2, { action: hit(AS.ACTION), cs: hit(AS.CS), fees: hit(AS.FEES) });
    else if (name === LEC.SHEET && (hit(LEC.STATUS) || hit(LEC.REASON) || hit(LEC.NEWDATE))) stampLectures_(sh, r1, r2);
    else if (name === AP.REQ && hit(AP.R.STATUS)) for (let r = r1; r <= r2; r++) sh.getRange(r, AP.R.STATUS_AT).setValue(sh.getRange(r, AP.R.STATUS).getValue() ? new Date() : '');
  } catch (err) { console.error(err); toast_('Error: ' + err.message); }
}

/** Stamps the date & time of every Action / CS Action / Dr. Fees change; fills Session date for new rows. */
function processAssignments_(r1, r2, edited) {
  const sh = sheet_(AS.SHEET);
  if (!sh || r2 < r1) return 0;
  edited = edited || {};
  const vals = sh.getRange(r1, 1, r2 - r1 + 1, AS.WIDTH).getValues();
  const now = new Date(), today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  let changed = 0;
  vals.forEach(function (v, i) {
    const r = r1 + i, upd = {};
    const started = [AS.AGENT, AS.MODULE, AS.CASE, AS.DOCTOR, AS.LINK, AS.ACTION].some(function (c) { return v[c - 1] !== '' && v[c - 1] !== null; });
    if (!started) return;
    if (!v[AS.DATE - 1]) upd[AS.DATE] = today;
    const action = String(v[AS.ACTION - 1] || '').trim();
    if (edited.action) {
      upd[AS.ACTION_AT] = action ? now : '';
      if (/^pending$/i.test(action) && !v[AS.SENT - 1]) upd[AS.SENT] = now;
    }
    if (edited.cs) upd[AS.CS_AT] = String(v[AS.CS - 1] || '').trim() ? now : '';
    if (edited.fees) upd[AS.FEES_AT] = String(v[AS.FEES - 1] || '').trim() ? now : '';
    const keys = Object.keys(upd);
    if (!keys.length) return;
    keys.forEach(function (c) { sh.getRange(r, Number(c)).setValue(upd[c]); });
    changed++;
  });
  return changed;
}

function stampLectures_(sh, r1, r2) {
  const v = sh.getRange(r1, 1, r2 - r1 + 1, LEC.AT).getValues(), now = new Date();
  v.forEach(function (row, i) {
    if (!row[0]) return;
    const filled = row[LEC.STATUS - 1] !== '' || row[LEC.REASON - 1] !== '' || row[LEC.NEWDATE - 1] !== '';
    sh.getRange(r1 + i, LEC.AT).setValue(filled ? now : '');
  });
}

/* =========================== One-time restore from a backup copy =========================== */
/** Copies Settings links, check-ins and open assignments from the backup copy (Settings!B15) into this sheet. */
function restoreFromBackup_(url) {
  const bk = SpreadsheetApp.openByUrl(url), set = sheet_(T.SETTINGS), out = [];
  // 1) Settings links
  const bset = bk.getSheetByName(T.SETTINGS);
  if (bset) ['B8', 'B10', 'B11', 'B12'].forEach(function (a) {
    const v = bset.getRange(a).getValue();
    if (v !== '' && v !== null) set.getRange(a).setValue(v);
  });
  // 2) check-ins (A..G raw + P typed name)
  const bfr = bk.getSheetByName(SETTINGS.RESPONSES_SHEET), fr = sheet_(SETTINGS.RESPONSES_SHEET);
  if (bfr) {
    const n = lastDataRow_(bfr, 1) - 1;
    if (n > 0) {
      const raw = bfr.getRange(2, 1, n, SETTINGS.RAW_COLS).getValues();
      const typed = bfr.getRange(2, SETTINGS.TYPED_NAME_COL, n, 1).getValues();
      const cur = lastDataRow_(fr, 1) - 1;
      if (cur > 0) { fr.getRange(2, 1, cur, SETTINGS.RAW_COLS).clearContent(); fr.getRange(2, SETTINGS.TYPED_NAME_COL, cur, 1).clearContent(); }
      fr.getRange(2, 2, n, 1).setNumberFormat('@');
      fr.getRange(2, 1, n, SETTINGS.RAW_COLS).setValues(raw);
      fr.getRange(2, SETTINGS.TYPED_NAME_COL, n, 1).setValues(typed);
      repairResponses_();
    }
    out.push('check-ins ' + Math.max(n, 0));
  }
  // 3) assignments: keep only open ones (Pending / no action yet); Done and canceled are dropped
  const bas = bk.getSheetByName(AS.SHEET), as = sheet_(AS.SHEET);
  if (bas) {
    const head = bas.getRange(1, 1, 1, bas.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    const col = function (name) { return head.indexOf(name); };
    const n = lastDataRow_(bas, 1) - 1;
    const rows = n > 0 ? bas.getRange(2, 1, n, head.length).getValues() : [];
    const isNew = col('CS Action') !== -1;
    const keep = [];
    rows.forEach(function (v) {
      if (!v.slice(0, 9).some(function (x) { return x !== '' && x !== null; })) return;
      const g = function (name) { const i = col(name); return i === -1 ? '' : v[i]; };
      const action = String(g('Action') || '').trim();
      if (/^(done|canceled|cancelled)$/i.test(action)) return;
      const act = /^pending$/i.test(action) ? 'Pending' : '';
      const id = String(g('Request ID') || '').match(/^AS-(\d+)$/);
      if (!act && id && Number(id[1]) <= 2364) return;      // old CS Tasks history that is not Pending
      let note = String(g('Note') || '').trim();
      let fees = String(g('Dr. Fees') || '').trim();
      if (fees && ['Pending', 'Done'].indexOf(fees) === -1) { note = [note, 'Dr. Fees: ' + fees].filter(String).join(' · '); fees = ''; }
      const sent = isNew ? g('Sent to doctor') : (g('Sent date') || (act ? g('Session Date Request') : ''));
      keep.push([g('Account Manager'), g('Session Date Request'), g('Moudule'), g('Case'), g('Dr. Name'), g('Program'), g('Acc.'),
                 g('Zoom link // Assignment link'), g('عدد الطلاب تصحيح'), act,
                 isNew ? g('Action date & time') : (act ? sent : ''), isNew ? g('CS Action') : '', isNew ? g('CS date & time') : '',
                 fees, isNew ? g('Fees date & time') : '', note, act ? sent : '']);
    });
    const cur = Math.max(lastDataRow_(as, 1), lastDataRow_(as, AS.ACTION)) - 1;
    if (cur > 0) as.getRange(2, 1, cur, AS.WIDTH).clearContent();
    if (keep.length) as.getRange(2, 1, keep.length, AS.WIDTH).setValues(keep);
    out.push('open assignments ' + keep.length + ' of ' + rows.length);
  }
  // 4) lecture attendance already recorded in the backup (same date + group + module + lecture #)
  const blec = bk.getSheetByName(LEC.SHEET), lec = sheet_(LEC.SHEET);
  if (blec && lec) {
    const key = function (r) { return Utilities.formatDate(new Date(r[0]), SETTINGS.TIMEZONE, 'yyyyMMdd') + '|' + r[3] + '|' + r[4] + '|' + r[5]; };
    const bn = lastDataRow_(blec, 1) - 1, n = lastDataRow_(lec, 1) - 1;
    if (bn > 0 && n > 0) {
      const map = {};
      blec.getRange(2, 1, bn, LEC.AT).getValues().forEach(function (r) { if (r[0] && (r[LEC.STATUS - 1] || r[LEC.REASON - 1])) map[key(r)] = r; });
      const cur = lec.getRange(2, 1, n, LEC.AT).getValues();
      let k = 0;
      cur.forEach(function (r, i) {
        const b = r[0] && map[key(r)];
        if (b) { lec.getRange(i + 2, LEC.STATUS, 1, 5).setValues([b.slice(LEC.STATUS - 1, LEC.AT)]); k++; }
      });
      out.push('lecture records ' + k);
    }
  }
  const msg = 'Restored ' + Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, 'dd/MM/yyyy HH:mm') + ' · ' + out.join(' · ');
  set.getRange('B16').setValue(msg);
  return msg;
}

function toast_(msg) { try { SpreadsheetApp.getActive().toast(msg, 'CEO Tools', 6); } catch (e) {} }

/* Editor tests */
function testLookup() { Logger.log(JSON.stringify(lookup_('966531146380'))); }
function testSubmit() {
  Logger.log(doPost({ postData: { contents: JSON.stringify({ phone: '01000000000', name: 'Test Student', submissionId: 'TEST-' + Date.now(), source: 'Editor test' }) } }).getContent());
}

/* =========================== Formulas =========================== */
function responseFormulas_(r) {
  return [
    '=IF($A' + r + '="","",INT($A' + r + '))',
    '=IF($A' + r + '="","",TEXT($A' + r + ',"dddd"))',
    '=IF($A' + r + '="","",MOD($A' + r + ',1))',
    '=IF($A' + r + '="","",IF(COUNTIFS($M$2:$M' + r + ',$M' + r + ',$D$2:$D' + r + ',$D' + r + ',$H$2:$H' + r + ',$H' + r + ')>1,"Duplicate","First"))',
    '=IF($A' + r + '="","",IF($B' + r + '="","Missing phone",IF($C' + r + '="",IF($P' + r + '<>"","Unregistered","Missing accreditation"),IF($D' + r + '="","No module running",IF($K' + r + '="Duplicate","Duplicate","Complete")))))',
    '=IF($A' + r + '="","",RIGHT(SUBSTITUTE($B' + r + ',"+",""),9))',
    '=IF($A' + r + '="","",IFERROR(INDEX(Students!$D:$D,MATCH($M' + r + '&"|"&$C' + r + ',Students!$A:$A,0)),IF($P' + r + '<>"",$P' + r + '&" (unregistered)","Not in student list")))',
    '=IF($A' + r + '="","",IFERROR(INDEX(Students!$F:$F,MATCH($M' + r + '&"|"&$C' + r + ',Students!$A:$A,0)),""))'
  ];
}
function writeResponseFormulas_(sh, fromRow, n) {
  const f = [];
  for (let i = 0; i < n; i++) f.push(responseFormulas_(fromRow + i));
  sh.getRange(fromRow, SETTINGS.FORMULA_COL, n, SETTINGS.FORMULA_COUNT).setFormulas(f);
  sh.getRange(fromRow, 8, n, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(fromRow, 10, n, 1).setNumberFormat('hh:mm AM/PM');
}
function repairResponses_() {
  const sh = sheet_(SETTINGS.RESPONSES_SHEET);
  const n = Math.max(firstEmptyRow_(sh) - 2 + 200, 999);
  writeResponseFormulas_(sh, 2, n);
  return n;
}
function studentFormulas_(r) {
  const FR = T.FR, look = 'MATCH($A' + r + '&"|1",Tracker!$N:$N,0)';
  return [
    '=IFERROR(INDEX(Tracker!$H:$H,' + look + '),"—")',
    '=IFERROR(INDEX(Tracker!$M:$M,' + look + '),"")',
    '=IFERROR(INDEX(Tracker!$K:$K,' + look + '),"")',
    '=IF(COUNTIFS(' + FR + '!$M:$M,$B' + r + ',' + FR + '!$C:$C,$E' + r + ',' + FR + '!$H:$H,TODAY())>0,"✔ Today","")',
    '=COUNTIFS(' + FR + '!$M:$M,$B' + r + ',' + FR + '!$C:$C,$E' + r + ',' + FR + '!$K:$K,"First")',
    '=IF($N' + r + '=0,"",MAXIFS(' + FR + '!$A:$A,' + FR + '!$M:$M,$B' + r + ',' + FR + '!$C:$C,$E' + r + '))',
    '=COUNTIFS(Tracker!$A:$A,$A' + r + ',Tracker!$M:$M,"Missing")'
  ];
}
function trackerFormulas_(r) {
  const FR = T.FR, GS = T.GS;
  const crit = FR + '!$M:$M,$B' + r + ',' + FR + '!$C:$C,$D' + r + ',' + FR + '!$D:$D,$H' + r;
  const win = FR + '!$H:$H,">="&$I' + r + ',' + FR + '!$H:$H,"<="&$J' + r;
  return [
    '=INDEX(' + GS + '!$E:$E,MATCH($G' + r + ',' + GS + '!$A:$A,0))',
    '=INDEX(' + GS + '!$F:$F,MATCH($G' + r + ',' + GS + '!$A:$A,0))',
    '=COUNTIFS(' + crit + ',' + win + ',' + FR + '!$K:$K,"First")',
    '=IF($K' + r + '=0,"",MAXIFS(' + FR + '!$A:$A,' + crit + ',' + win + '))',
    '=IF($K' + r + '>0,"Attended",IF(IFERROR(INDEX(Students!$G:$G,MATCH($A' + r + ',Students!$A:$A,0)),"")<>"Still studying","Inactive",IF($J' + r + '<Settings!$B$2,"Before tracking",IF(TODAY()<$I' + r + ',"Upcoming",IF(TODAY()<=$J' + r + ',"In progress","Missing")))))',
    '=$A' + r + '&"|"&IF(AND($I' + r + '<=TODAY(),$J' + r + '>=TODAY()),1,0)'
  ];
}
function dashFormulas_(r) {
  const GS = T.GS, m = 'MATCH($B' + r + ',' + GS + '!$A:$A,0)', tc = 'Tracker!$G:$G,$B' + r;
  return [
    '=INDEX(' + GS + '!$B:$B,' + m + ')', '=INDEX(' + GS + '!$D:$D,' + m + ')',
    '=INDEX(' + GS + '!$E:$E,' + m + ')', '=INDEX(' + GS + '!$F:$F,' + m + ')',
    '=IF(TODAY()<$E' + r + ',"Upcoming",IF(TODAY()<=$F' + r + ',"Running","Finished"))',
    '=COUNTIF(Tracker!$G:$G,$B' + r + ')', '=COUNTIFS(' + tc + ',Tracker!$M:$M,"Attended")',
    '=IF($H' + r + '=0,"",$I' + r + '/$H' + r + ')', '=SUMIFS(Tracker!$K:$K,' + tc + ')',
    '=COUNTIFS(' + tc + ',Tracker!$M:$M,"Missing")', '=COUNTIFS(' + tc + ',Tracker!$M:$M,"In progress")'
  ];
}

/* =========================== Tracker + dashboard =========================== */
function appendTracker_(rows) {
  const tr = sheet_(T.TRACKER);
  const start = lastDataRow_(tr, 1) + 1;
  tr.getRange(start, 2, rows.length, 1).setNumberFormat('@');
  tr.getRange(start, 1, rows.length, 8).setValues(rows);
  tr.getRange(start, 3, rows.length, 1).setFormulas(rows.map(function (_, i) {
    return ['=IFERROR(INDEX(Students!$D:$D,MATCH($A' + (start + i) + ',Students!$A:$A,0)),"")'];
  }));
  tr.getRange(start, 9, rows.length, 6).setFormulas(rows.map(function (_, i) { return trackerFormulas_(start + i); }));
  if (start > 2) tr.getRange(2, 1, 1, 14).copyTo(tr.getRange(start, 1, rows.length, 14), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

function refreshDashboard_() {
  const db = sheet_(T.DASH), tr = sheet_(T.TRACKER);
  if (!db || !tr) return;
  const trLast = lastDataRow_(tr, 1);
  const used = {};
  if (trLast >= 2) tr.getRange(2, 7, trLast - 1, 1).getValues().forEach(function (x) { if (x[0]) used[x[0]] = true; });
  const ids = readSchedule_().filter(function (s) { return used[s.id]; }).map(function (s) { return s.id; });
  const first = T.DASH_FIRST_ROW;
  const oldLast = Math.max(lastDataRow_(db, 2), first - 1);
  if (ids.length) {
    db.getRange(first, 2, ids.length, 1).setValues(ids.map(function (id) { return [id]; }));
    db.getRange(first, 3, ids.length, 11).setFormulas(ids.map(function (_, i) { return dashFormulas_(first + i); }));
    const newLast = first + ids.length - 1;
    if (newLast > oldLast) db.getRange(first, 2, 1, 12).copyTo(db.getRange(oldLast + 1, 2, newLast - oldLast, 12), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
  const clearFrom = first + ids.length;
  if (oldLast >= clearFrom) db.getRange(clearFrom, 2, oldLast - clearFrom + 1, 12).clearContent();
}

/* =========================== Shared helpers =========================== */
function readSchedule_() {
  const sh = sheet_(T.SCHEDULE), last = lastDataRow_(sh, 1);
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 9).getValues()
    .filter(function (r) { return r[0] && r[1] && r[3] && r[4] instanceof Date; })
    .map(function (r) {
      return { id: String(r[0]), group: String(r[1]), accs: String(r[2]).split('/').map(function (a) { return a.trim(); }),
               module: String(r[3]).trim(), start: dayStart_(r[4]), end: r[5] instanceof Date ? dayStart_(r[5]) : dayStart_(r[4]),
               doctor: String(r[8] || '').trim() };
    })
    .sort(function (a, b) { return a.start - b.start; });
}
function groupsFor_(sched, acc) {
  const out = [];
  sched.forEach(function (s) { if (s.accs.indexOf(acc) !== -1 && out.indexOf(s.group) === -1) out.push(s.group); });
  return out;
}
function buildPlan_(sched, group, acc, startModule, refDate, catalogMods) {
  const rows = sched.filter(function (s) { return s.group === group; });
  if (!rows.length) return [];
  const ref = dayStart_(refDate || new Date());
  let occ = null;
  if (startModule) {
    occ = rows.filter(function (s) { return s.module === startModule && s.start <= ref && ref <= s.end; })[0] ||
          rows.filter(function (s) { return s.module === startModule && s.start >= ref; })[0] || null;
  }
  if (!occ) occ = rows.filter(function (s) { return s.start <= ref && ref <= s.end; })[0] ||
                  rows.filter(function (s) { return s.start > ref; })[0] || null;
  if (!occ) return [];
  const plan = [occ], seen = {}; seen[occ.module] = true;
  const need = (catalogMods && catalogMods.length) ? catalogMods.length : 99;
  rows.forEach(function (s) {
    if (plan.length >= need) return;
    if (s.start > occ.start && !seen[s.module]) { plan.push(s); seen[s.module] = true; }
  });
  return plan;
}
function checkSubmission_(id) {
  if (!id) return { ok: false, error: 'Missing id' };
  const sh = sheet_(SETTINGS.RESPONSES_SHEET);
  const row = firstEmptyRow_(sh);
  if (row <= 2) return { ok: true, found: false };
  const start = Math.max(2, row - 500);
  const vals = sh.getRange(start, 1, row - start, 11).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][5]) === id) {
      return { ok: true, found: true, timestamp: vals[i][0] instanceof Date ? vals[i][0].toISOString() : '', duplicate: vals[i][10] === 'Duplicate' };
    }
  }
  return { ok: true, found: false };
}
function getCatalog_() {
  const sh = sheet_(SETTINGS.CATALOG_SHEET);
  const values = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 4).getValues();
  const map = {};
  values.forEach(function (r) {
    const acc = String(r[0]).trim(), mod = String(r[1]).replace(/\s+/g, ' ').trim(), active = String(r[3]).trim().toLowerCase();
    if (!acc || !mod || active === 'no') return;
    if (!map[acc]) map[acc] = [];
    if (map[acc].indexOf(mod) === -1) map[acc].push(mod);
  });
  const ordered = {};
  SETTINGS.ACC_ORDER.forEach(function (a) { if (map[a]) ordered[a] = map[a]; });
  Object.keys(map).forEach(function (a) { if (!ordered[a]) ordered[a] = map[a]; });
  return ordered;
}
function firstEmptyRow_(sh) {
  const last = Math.max(sh.getLastRow(), 2);
  const colA = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = colA.length - 1; i >= 0; i--) if (colA[i][0] !== '' && colA[i][0] !== null) return i + 3;
  return 2;
}
function isDuplicateToday_(sh, row, phone, module, now) {
  if (row <= 2) return false;
  const start = Math.max(2, row - 3000);
  const vals = sh.getRange(start, 1, row - start, 4).getValues();
  const today = Utilities.formatDate(now, SETTINGS.TIMEZONE, 'yyyy-MM-dd'), key = phoneKey_(phone);
  for (let i = vals.length - 1; i >= 0; i--) {
    const ts = vals[i][0];
    if (!(ts instanceof Date)) continue;
    if (Utilities.formatDate(ts, SETTINGS.TIMEZONE, 'yyyy-MM-dd') !== today) continue;
    if (phoneKey_(vals[i][1]) === key && String(vals[i][3]) === module) return true;
  }
  return false;
}
/** First phone number in a cell. "966569989030-998917989906" → 966569989030 · "+20 100 123 4567" → 201001234567 */
function firstNumber_(p) {
  const s = String(p == null ? '' : p).replace(/\.0+$/, '')
    .replace(/[٠-٩]/g, function (d) { return d.charCodeAt(0) - 0x0660; }).replace(/[۰-۹]/g, function (d) { return d.charCodeAt(0) - 0x06F0; });
  const groups = s.match(/\d+/g) || [];
  let acc = '';
  for (let i = 0; i < groups.length; i++) {
    if (acc.length >= 8 && groups[i].length >= 8) break;   // a second number starts here
    acc += groups[i];
  }
  return acc.length >= 8 ? acc : '';
}
function phoneKey_(p) { const d = firstNumber_(p); return d ? d.slice(-9) : ''; }
function normalizePhone_(p) {
  let s = String(p || '')
    .replace(/[٠-٩]/g, function (d) { return d.charCodeAt(0) - 0x0660; })
    .replace(/[۰-۹]/g, function (d) { return d.charCodeAt(0) - 0x06F0; })
    .replace(/[\s\-().]/g, '');
  if (s.indexOf('00') === 0) s = '+' + s.slice(2);
  return s;
}
function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function lastDataRow_(sh, col) {
  const n = sh.getLastRow();
  if (n < 2) return 1;
  const v = sh.getRange(2, col, n - 1, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] !== '' && v[i][0] !== null) return i + 2;
  return 1;
}
/** Calendar day in Cairo as a sortable number (20261019) – immune to the script's own time zone. */
function dayStart_(d) { return Number(Utilities.formatDate(new Date(d), SETTINGS.TIMEZONE, 'yyyyMMdd')); }
function dayKeyToIso_(n) { const s = String(n); return s.slice(0, 4) + '-' + s.slice(4, 6) + '-' + s.slice(6, 8); }
function parseDate_(iso) { if (!iso) return null; const d = new Date(iso); return isNaN(d.getTime()) ? null : d; }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
