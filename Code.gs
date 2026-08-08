/**
 * EHFL — Tax & Flexi Declaration Portal, backend
 * FY 2026-27
 *
 * Deploy: Extensions ▸ Apps Script ▸ Deploy ▸ New deployment
 *   Type    : Web app
 *   Execute as    : Me
 *   Who has access: Anyone
 * Then paste the /exec URL into API in index.html.
 *
 * Security model
 * --------------
 *  · Login returns ONE employee's record. The master sheet never leaves the server.
 *  · Passcode = first 4 characters of PAN + DDMMYYYY. Never stored, only compared.
 *  · A token is issued per session, held 45 minutes in CacheService, and required
 *    on submit. Tokens are not written to the sheet.
 *  · Five failed attempts on an employee code locks it for 15 minutes.
 *  · Every login attempt is written to an Audit sheet.
 */

var SS = SpreadsheetApp.getActive();
var SH_MASTER  = 'Master';       // employee master
var SH_DECL    = 'Declarations'; // investment declarations on record
var SH_RENT    = 'HRA';          // rent rows on record
var SH_OUT_D   = 'OUT_Declaration';
var SH_OUT_H   = 'OUT_HRA';
var SH_OUT_R   = 'OUT_Regime';
var SH_OUT_F   = 'OUT_Flexi';
var SH_AUDIT   = 'Audit';

var LOCK_TRIES = 5;
var LOCK_MINS  = 15;
var TOKEN_MINS = 45;

/* ───────────── helpers ───────────── */
function sheet_(name){
  var s = SS.getSheetByName(name);
  if(!s) s = SS.insertSheet(name);
  return s;
}
function rows_(name){
  var s = SS.getSheetByName(name);
  if(!s) return [];
  var v = s.getDataRange().getValues();
  if(v.length < 2) return [];
  var head = v[0].map(function(h){return String(h).trim()});
  return v.slice(1).map(function(r){
    var o = {};
    head.forEach(function(h,i){ o[h] = r[i]; });
    return o;
  });
}
function norm_(s){ return String(s == null ? '' : s).replace(/[\s_\-\.]/g,'').toLowerCase(); }
function pick_(obj, names){
  var keys = Object.keys(obj);
  for(var i=0;i<names.length;i++){
    for(var j=0;j<keys.length;j++){
      if(norm_(keys[j]) === norm_(names[i])) return obj[keys[j]];
    }
  }
  return '';
}
function n_(v){
  var x = parseFloat(String(v==null?'':v).replace(/[^0-9.\-]/g,''));
  return isNaN(x) ? 0 : x;
}
function ddmmyyyy_(v){
  if(v instanceof Date && !isNaN(v)){
    var d = ('0'+v.getDate()).slice(-2), m = ('0'+(v.getMonth()+1)).slice(-2);
    return d + m + v.getFullYear();
  }
  var s = String(v||'').trim();
  var m1 = s.match(/^(\d{1,2})[\-\/](\d{1,2})[\-\/](\d{4})$/);
  if(m1) return ('0'+m1[1]).slice(-2) + ('0'+m1[2]).slice(-2) + m1[3];
  var m2 = s.match(/^(\d{4})[\-\/](\d{1,2})[\-\/](\d{1,2})$/);
  if(m2) return ('0'+m2[3]).slice(-2) + ('0'+m2[2]).slice(-2) + m2[1];
  return s.replace(/\D/g,'');
}
function out_(obj, cb){
  var body = JSON.stringify(obj);
  if(cb) return ContentService.createTextOutput(cb + '(' + body + ')')
           .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(body)
           .setMimeType(ContentService.MimeType.JSON);
}
function audit_(code, event, detail){
  try{
    var s = sheet_(SH_AUDIT);
    if(s.getLastRow() === 0) s.appendRow(['Timestamp','Employee Code','Event','Detail']);
    s.appendRow([new Date(), code, event, detail || '']);
  }catch(e){}
}

/* ───────────── entry points ───────────── */
function doGet(e){
  var p  = e.parameter || {};
  var cb = p.callback;
  try{
    if(p.action === 'login') return out_(login_(p.code, p.p), cb);
    return out_({status:'error', message:'Unknown request.'}, cb);
  }catch(err){
    return out_({status:'error', message:String(err && err.message || err)}, cb);
  }
}
function doPost(e){
  try{
    var body = JSON.parse(e.postData.contents);
    if(body.action === 'submit') return out_(submit_(body));
    return out_({status:'error', message:'Unknown request.'});
  }catch(err){
    return out_({status:'error', message:String(err && err.message || err)});
  }
}

/* ───────────── login ───────────── */
function login_(code, pass){
  code = String(code||'').trim();
  pass = String(pass||'').trim().toUpperCase();
  if(!code || !pass) return {status:'error', message:'Employee code and passcode are both required.'};

  var cache = CacheService.getScriptCache();
  var fkey  = 'fail_' + code;
  var fails = Number(cache.get(fkey) || 0);
  if(fails >= LOCK_TRIES){
    audit_(code, 'LOCKED', '');
    return {status:'error', message:'Too many failed attempts. Try again in ' + LOCK_MINS + ' minutes.'};
  }

  var master = rows_(SH_MASTER);
  var rec = null;
  for(var i=0;i<master.length;i++){
    if(String(pick_(master[i], ['Employee Code','emp_code','code'])).trim() === code){ rec = master[i]; break; }
  }
  if(!rec){
    cache.put(fkey, String(fails+1), LOCK_MINS*60);
    audit_(code, 'FAIL', 'no such code');
    return {status:'error', message:'Employee code or passcode is incorrect.'};
  }

  var pan = String(pick_(rec, ['Employee Pan Card','PAN','pan'])).trim().toUpperCase();
  var dob = ddmmyyyy_(pick_(rec, ['Date of Birth','DOB','dob']));
  var expect = pan.substring(0,4) + dob;

  if(!pan || !dob || expect !== pass){
    cache.put(fkey, String(fails+1), LOCK_MINS*60);
    audit_(code, 'FAIL', 'bad passcode');
    return {status:'error', message:'Employee code or passcode is incorrect.'};
  }
  cache.remove(fkey);

  // already submitted?
  var done = rows_(SH_OUT_R).some(function(r){
    return String(pick_(r,['employee_code','Employee Code'])).trim() === code;
  });

  var token = Utilities.getUuid();
  cache.put('tok_' + token, code, TOKEN_MINS*60);
  audit_(code, 'LOGIN', done ? 'resubmission' : 'first');

  return {status:'ok', token: token, emp: buildEmp_(rec, code)};
}

/* Assemble exactly one employee's payload. Nothing else is exposed. */
function buildEmp_(rec, code){
  var declRow = null, decl = rows_(SH_DECL);
  for(var i=0;i<decl.length;i++){
    if(String(pick_(decl[i], ['Employee Code','emp_code'])).trim() === code){ declRow = decl[i]; break; }
  }
  var rent = new Array(12).fill(0);
  var MN = [4,5,6,7,8,9,10,11,12,1,2,3], MY = [2026,2026,2026,2026,2026,2026,2026,2026,2026,2027,2027,2027];
  var rrows = rows_(SH_RENT), land = {name:'', pan:'', addr:'', city:'', pin:''};
  rrows.forEach(function(r){
    if(String(pick_(r,['Employee Code','emp_code'])).trim() !== code) return;
    var amt = n_(pick_(r,['Amount','Rent']));
    var from = pick_(r,['Rent From','From Month (M)']);
    var mo, yr;
    if(from instanceof Date){ mo = from.getMonth()+1; yr = from.getFullYear(); }
    else { mo = n_(from); yr = n_(pick_(r,['From Year (yyyy)'])); }
    for(var i=0;i<12;i++) if(MN[i]===mo && MY[i]===yr) rent[i] = amt;
    if(!land.name) land.name = String(pick_(r,['Landlord Name'])||'');
    if(!land.pan)  land.pan  = String(pick_(r,['Landlord Pan Card','Landlord PAN'])||'');
    if(!land.city) land.city = String(pick_(r,['City'])||'');
    if(!land.pin)  land.pin  = String(pick_(r,['Pincode'])||'');
    if(!land.addr) land.addr = String(pick_(r,['Address'])||'');
  });

  // monthly actuals: columns named  Apr_Basic, Apr_HRA, Apr_Flex, Apr_NPS, Apr_Gross, Apr_TDS ...
  var MON = ['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar'];
  var paid = n_(pick_(rec, ['Paid Months','paid_months']));
  var actuals = [];
  for(var i=0;i<paid;i++){
    actuals.push({
      basic: n_(pick_(rec,[MON[i]+'_Basic'])), hra:  n_(pick_(rec,[MON[i]+'_HRA'])),
      flex:  n_(pick_(rec,[MON[i]+'_Flex'])),  nps:  n_(pick_(rec,[MON[i]+'_NPS'])),
      other: n_(pick_(rec,[MON[i]+'_Other'])), gross:n_(pick_(rec,[MON[i]+'_Gross'])),
      tds:   n_(pick_(rec,[MON[i]+'_TDS'])),   pt:   n_(pick_(rec,[MON[i]+'_PT'])) || 200
    });
  }

  var regimeNow = '';
  rows_(SH_OUT_R).forEach(function(r){
    if(String(pick_(r,['employee_code','Employee Code'])).trim() === code)
      regimeNow = /new/i.test(String(pick_(r,['Tax Regime Name']))) ? 'new' : 'old';
  });

  return {
    code: code,
    name: String(pick_(rec,['Employee Name','emp_name'])||''),
    location: String(pick_(rec,['Location'])||''),
    doj: String(pick_(rec,['Date of Joining','DOJ'])||''),
    monthlyBasic: n_(pick_(rec,['Monthly Basic','basic_monthly'])),
    monthlyGross: n_(pick_(rec,['Monthly Gross','gross_monthly'])),
    annualGross:  n_(pick_(rec,['Annual Fixed Pay','AFP','Annual Gross'])),
    paidMonths: paid,
    actuals: actuals,
    metro: String(pick_(rec,['Metro'])||'Y').toUpperCase().charAt(0) === 'N' ? 'N' : 'Y',
    regime: regimeNow || (/new/i.test(String(pick_(rec,['Tax Regime']))) ? 'new' : 'old'),
    npsPct: n_(pick_(rec,['NPS Percent','nps_pct'])),
    flexi: {
      meal:   n_(pick_(rec,['Meal'])),   tele:  n_(pick_(rec,['Telecom'])),
      health: n_(pick_(rec,['Health'])), fuel:  n_(pick_(rec,['Fuel'])),
      driver: n_(pick_(rec,['Driver'])), lta:   n_(pick_(rec,['LTA'])),
      wear:   n_(pick_(rec,['Wear'])),   books: n_(pick_(rec,['Books'])),
      lnd:    n_(pick_(rec,['LND']))
    },
    rentExisting: rent,
    landlordName: land.name, landlordPan: land.pan,
    rentAddr: land.addr, rentCity: land.city, rentPin: land.pin,
    decl: declRow ? mapDecl_(declRow) : {}
  };
}

/* declaration sheet column → portal key */
var DECL_MAP = {
  lip:'Life Insurance Premium [LIP]', ppf:'Public Provident Fund [PPF]',
  elss:'ELSS Investment in the Unit of Government approved plan framed under equity Linked Saving Scheme (For Self Only)',
  nsc:'Subscription to National Saving Certificate [NSC] VII Issue', ulip:'Unit Linked Plan [ULIP]',
  fd5:'Tax Saving Fixed Deposit [5 Years and above] or 5 Year Post Office Time Deposit [POTD] Scheme',
  hlp:'Housing. Loan [Principal Repayment]', sip:'SIP - more than 3 years lock in period',
  ssa:'Sukanya Samriddhi Account', stamp:'Stamp Duty & Registration Charges',
  tuit:'Children\u2019s Tuition Fee', pens:'Pension Plan from Insurance Companies/Mutual Funds',
  nps80c:'New Pension Scheme [NPS]', scss:'Senior Citizen\u2019s Saving Scheme [SCSS]',
  kvp:'Kisan Vikas Patra (KVP)', infra:'Long term Infrastructure Bonds',
  pmsby:'Pradhan Mantri Suraksha Bima Yojana',
  nps1b:'(80CCD(1B)) Sec 124(3) - Contribution under National Pension Scheme (Max \u20b950,000)',
  d80self:'(80D) Sec 126 - Medical Insurance Premium, Preventive Health Check-up for Individual, Spouse & children',
  d80par:'(80D) Sec 126 - Medical Insurance Premium, Preventive Health Check-up for Parents not being Sr. Citizens',
  d80sr:'(80D) Sec 126 - Medical Insurance Premium, Preventive Health Check-up for Parents are Sr. Citizens',
  int24:'Interest paid on self-occupied property', letInc:'Income from let-out property',
  letInt:'Interest paid for let-out property',
  letTax:'House government/municipal tax paid (let-out property)',
  d80e:'(80E) Sec 129 - Payment of interest on the loan for higher education. (Deduction allowed in Initial Assessment Year and 7 preceding years)',
  d80tta:'(80TTA) Sec 153 - Deduction in respect of interest on deposits in savings account',
  d80eeb:'(80EEB) Sec 132 - Payment of interest on loan for purchase of Electric Vehicle',
  d80ddb:'(80DDB) Sec 128 - Medical treatment for self/dependents of specified diseases (Below 60 years)',
  d80dd:'(80DD) Sec 127 - Medical treatment for handicapped dependent - Disability between 40% to 80%',
  d80u:'(80U) Sec 154 - Permanent physical disability for tax assessee only (Person with General Disability below 80%)',
  don100:'Donation (Upto 100%)',
  don50:'Donation (Upto 50%) - consider 50% of amount filled for calculations'
};
function mapDecl_(row){
  var o = {};
  Object.keys(DECL_MAP).forEach(function(k){ o[k] = n_(pick_(row, [DECL_MAP[k]])); });
  return o;
}

/* ───────────── submit ───────────── */
function submit_(b){
  var cache = CacheService.getScriptCache();
  var code  = cache.get('tok_' + String(b.token||''));
  if(!code) return {status:'error', message:'Your session has expired. Sign in again to save.'};
  if(String(b.code).trim() !== code)
    return {status:'error', message:'Session does not match this employee.'};

  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try{
    upsert_(SH_OUT_D, b.decl,   'emp_code');
    upsert_(SH_OUT_R, b.regime, 'employee_code');
    upsert_(SH_OUT_F, b.flexi,  'emp_code');
    replaceRows_(SH_OUT_H, b.hra, 'emp_code', code);

    var c = b.computed || {};
    var s = sheet_('OUT_Summary');
    if(s.getLastRow() === 0)
      s.appendRow(['Timestamp','Employee Code','Regime','Taxable Income','Total Tax',
                   'TDS Paid','Balance','Monthly TDS','HRA Exemption','Perquisite over 7.5L']);
    s.appendRow([new Date(), code, c.regime, c.taxable, c.total_tax, c.tds_paid,
                 c.balance, c.monthly_tds, c.hra_exemption, c.perquisite_over_75L]);

    audit_(code, 'SUBMIT', c.regime + ' / tax ' + c.total_tax);
    return {status:'ok'};
  } finally { lock.releaseLock(); }
}

/* write one row per employee, replacing any earlier row */
function upsert_(name, obj, keyCol){
  if(!obj) return;
  var s = sheet_(name);
  var head = Object.keys(obj);
  if(s.getLastRow() === 0) s.appendRow(head);
  var existing = s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(String);
  if(existing.join('|') !== head.join('|') && s.getLastRow() <= 1){
    s.clear(); s.appendRow(head); existing = head;
  }
  var key = String(obj[keyCol]);
  var col = existing.indexOf(keyCol) + 1;
  var vals = existing.map(function(h){ return obj[h] === undefined ? '' : obj[h]; });
  if(col > 0 && s.getLastRow() > 1){
    var colVals = s.getRange(2, col, s.getLastRow()-1, 1).getValues();
    for(var i=0;i<colVals.length;i++){
      if(String(colVals[i][0]).trim() === key){
        s.getRange(i+2, 1, 1, vals.length).setValues([vals]); return;
      }
    }
  }
  s.appendRow(vals);
}
/* replace all rows for one employee (HRA is many rows per person) */
function replaceRows_(name, list, keyCol, code){
  var s = sheet_(name);
  if(!list || !list.length) return;
  var head = Object.keys(list[0]);
  if(s.getLastRow() === 0) s.appendRow(head);
  var existing = s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(String);
  var col = existing.indexOf(keyCol) + 1;
  if(col > 0 && s.getLastRow() > 1){
    var v = s.getRange(2, col, s.getLastRow()-1, 1).getValues();
    for(var i=v.length-1;i>=0;i--) if(String(v[i][0]).trim() === String(code)) s.deleteRow(i+2);
  }
  var out = list.map(function(o){ return existing.map(function(h){ return o[h]===undefined?'':o[h]; }); });
  s.getRange(s.getLastRow()+1, 1, out.length, existing.length).setValues(out);
}
