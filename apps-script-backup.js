// ================================================================
// Apps Script - ניהול מלאי בית מלאכה + מעקב מכונות קרח בהשכרה
// גיבוי - ספטמבר 2026
// Google Sheets ID: 125xRue7T-5WNsinT2fA0zlJCwyz6zyeWOeTo8trdNrM
// ================================================================

function doGet(e) {
  // פעולות מכונות מטופלות בנפרד (בתחתית הקובץ)
  if (MACHINE_ACTIONS.indexOf(e.parameter.action) >= 0) return out(machineAction(e.parameter.action, e.parameter));
  const ss = SpreadsheetApp.openById('125xRue7T-5WNsinT2fA0zlJCwyz6zyeWOeTo8trdNrM');
  const items = ss.getSheetByName('פריטים');
  const log = ss.getSheetByName('יומן');
  const action = e.parameter.action;

  if (action === 'getItems') {
    return out(items.getDataRange().getValues());
  }
  if (action === 'getLog') {
    try { return out(log.getDataRange().getValues()); }
    catch(err) { return out([]); }
  }
  if (action === 'take') {
    const id = e.parameter.id;
    const customer = e.parameter.customer || '';
    const role = e.parameter.role || 'טכנאי';
    const takeQty = parseInt(e.parameter.qty) || 1;
    const data = items.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        const qty = parseInt(data[i][5]) || 0;
        if (qty <= 0) return out({success:false, msg:'אזל מהמלאי'});
        if (qty < takeQty) return out({success:false, msg:'מלאי לא מספיק, יש רק ' + qty + ' יחידות'});
        const newQty = qty - takeQty;
        const min = parseInt(data[i][6]) || 3;
        items.getRange(i+1,6).setValue(newQty);
        const now = new Date();
        log.appendRow([
          Utilities.formatDate(now,'Asia/Jerusalem','dd/MM/yyyy'),
          Utilities.formatDate(now,'Asia/Jerusalem','HH:mm'),
          data[i][1], data[i][2], customer, role, takeQty
        ]);
        if (newQty < min) {
          sendLowStockAlert(data[i][1], newQty, min, data[i][9], data[i][8]);
        }
        return out({success:true, name:data[1], qty:newQty});
      }
    }
    return out({success:false, msg:'פריט לא נמצא'});
  }
  if (action === 'techView') {
    const id = e.parameter.id;
    const data = items.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        return out({
          id: data[i][0], name: data[i][1], type: data[i][2],
          desc: data[i][3], category: data[i][4],
          qty: parseInt(data[i][5]) || 0, min: parseInt(data[i][6]) || 3,
          emoji: data[i][7] || '📦', image: data[i][8] || '',
          sku: data[i][9] || ''
        });
      }
    }
    return out({error:'לא נמצא'});
  }
  if (action === 'addItem') {
    items.appendRow([
      e.parameter.id, e.parameter.name, e.parameter.type,
      e.parameter.desc, e.parameter.category,
      parseInt(e.parameter.qty)||0, parseInt(e.parameter.min)||3,
      e.parameter.emoji||'📦', e.parameter.image||'', e.parameter.sku||''
    ]);
    return out({success:true});
  }
  if (action === 'updateItem') {
    const id = e.parameter.id;
    const data = items.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        if (e.parameter.qty !== undefined) items.getRange(i+1,6).setValue(parseInt(e.parameter.qty)||0);
        if (e.parameter.name) items.getRange(i+1,2).setValue(e.parameter.name);
        if (e.parameter.type) items.getRange(i+1,3).setValue(e.parameter.type);
        if (e.parameter.desc) items.getRange(i+1,4).setValue(e.parameter.desc);
        if (e.parameter.category) items.getRange(i+1,5).setValue(e.parameter.category);
        if (e.parameter.min) items.getRange(i+1,7).setValue(parseInt(e.parameter.min)||3);
        if (e.parameter.emoji) items.getRange(i+1,8).setValue(e.parameter.emoji);
        if (e.parameter.image !== undefined) items.getRange(i+1,9).setValue(e.parameter.image);
        if (e.parameter.sku !== undefined) items.getRange(i+1,10).setValue(e.parameter.sku);
        return out({success:true});
      }
    }
    return out({success:false});
  }
  if (action === 'deleteItem') {
    const id = e.parameter.id;
    const data = items.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        items.deleteRow(i+1);
        return out({success:true});
      }
    }
    return out({success:false});
  }
  return out({error:'פעולה לא מוכרת'});
}

function sendLowStockAlert(itemName, qty, min, sku, imageUrl) {
  itemName = itemName || 'פריט לא ידוע';
  qty = (qty === undefined || qty === null || qty === '') ? 0 : qty;
  min = (min === undefined || min === null || min === '') ? 0 : min;
  sku = sku || '';
  imageUrl = imageUrl || '';

  const subject = 'התראת מלאי נמוך: ' + itemName;

  const body = 'התראת מלאי נמוך\n\n' +
    'פריט: ' + itemName + '\n' +
    (sku ? 'מק"ט: ' + sku + '\n' : '') +
    'כמות נוכחית: ' + qty + '\n' +
    'מינימום נדרש: ' + min + '\n\n' +
    'נשלח אוטומטית ממערכת ניהול מלאי בית מלאכה';

  const htmlBody =
    '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:15px;color:#222;line-height:1.6">' +
      '<h2 style="color:#c0392b;margin-bottom:12px">⚠️ התראת מלאי נמוך</h2>' +
      (imageUrl ? '<img src="' + imageUrl + '" alt="' + itemName + '" style="max-width:180px;border-radius:8px;margin-bottom:12px;display:block" />' : '') +
      '<table style="border-collapse:collapse">' +
        '<tr><td style="padding:4px 12px 4px 0;color:#555">פריט:</td><td style="font-weight:bold">' + itemName + '</td></tr>' +
        (sku ? '<tr><td style="padding:4px 12px 4px 0;color:#555">מק"ט:</td><td>' + sku + '</td></tr>' : '') +
        '<tr><td style="padding:4px 12px 4px 0;color:#555">כמות נוכחית:</td><td style="color:#c0392b;font-weight:bold">' + qty + '</td></tr>' +
        '<tr><td style="padding:4px 12px 4px 0;color:#555">מינימום נדרש:</td><td>' + min + '</td></tr>' +
      '</table>' +
      '<p style="margin-top:16px;color:#888;font-size:13px">נשלח אוטומטית ממערכת ניהול מלאי בית מלאכה</p>' +
    '</div>';

  MailApp.sendEmail({
    to: 'yossi.vogel2@gmail.com',
    subject: subject,
    body: body,
    htmlBody: htmlBody
  });
}

function testEmailAlert() {
  MailApp.sendEmail({
    to: 'yossi.vogel2@gmail.com',
    subject: 'בדיקת מערכת מלאי',
    body: 'המערכת עובדת'
  });
}

function out(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
// מעקב מכונות קרח בהשכרה
// כל מה שמתחת לשורה הזאת שייך למכונות בלבד ולא נוגע במלאי
// ================================================================

const SS_ID = '125xRue7T-5WNsinT2fA0zlJCwyz6zyeWOeTo8trdNrM';
const ADMIN_KEY = 'Vogel';
const PHOTO_FOLDER = 'מכונות-תמונות';

const ST_W = 'בבית המלאכה';
const ST_R = 'בשכירות';
const ST_M = 'בתחזוקה';
const ST_F = 'תקולה בבדיקה';

const M_SHEETS = {
  machines:  ['מכונות', ['מזהה','מספר מכונה','דגם','מצב','הערות']],
  rentals:   ['שכירויות', ['מזהה','מזהה מכונה','מספר מכונה','לקוח','כתובת','איש קשר','תאריך התחלה','תאריך סיום','ימים','ציוד','תמונה','פתח','סגר','אופן סגירה','הערות']],
  mlog:      ['יומן מכונות', ['תאריך','שעה','מספר מכונה','פעולה','לקוח','מי','פרטים']],
  customers: ['לקוחות', ['שם','כתובת','איש קשר','חודשי']],
  people:    ['אנשים', ['שם','תפקיד','קוד']],
  equipment: ['ציוד נלווה', ['שם']]
};

const MACHINE_ACTIONS = [
  'mCheckCode','mView','mFind','mInstall','mEnd','mTransfer','mSwap','mSetStatus','mEditRental',
  'mAll','mAddMachine','mUpdateMachine','mDeleteMachine',
  'mAddPerson','mResetCode','mDeletePerson','mSetRole',
  'mUpdateCustomer','mDeleteCustomer','mAddEquipment','mDeleteEquipment'
];

// פעולות שמשנות נתונים: רצות עם נעילה כדי ששתי סריקות באותו רגע לא יתנגשו
const WRITE_ACTIONS = [
  'mInstall','mEnd','mTransfer','mSwap','mSetStatus','mEditRental',
  'mAddMachine','mUpdateMachine','mDeleteMachine',
  'mAddPerson','mResetCode','mDeletePerson','mSetRole',
  'mUpdateCustomer','mDeleteCustomer','mAddEquipment','mDeleteEquipment'
];

// קבלת נתונים בשליחה (משמש את דף הסריקה, כולל תמונות)
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { return out({success:false, msg:'בקשה לא תקינה'}); }
  if (MACHINE_ACTIONS.indexOf(p.action) >= 0) return out(machineAction(p.action, p));
  return out({error:'פעולה לא מוכרת'});
}

// להריץ פעם אחת מתוך העורך כדי לאשר הרשאות ולהקים את הגיליונות
function authorizeOnce() {
  Object.keys(M_SHEETS).forEach(function (k) { mSheet(k); });
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  if (!it.hasNext()) DriveApp.createFolder(PHOTO_FOLDER);
  MailApp.getRemainingDailyQuota();
  Logger.log('הכל מוכן');
}

function machineAction(action, p) {
  const needLock = WRITE_ACTIONS.indexOf(action) >= 0;
  const lock = LockService.getScriptLock();
  try {
    if (needLock) lock.waitLock(20000);
    return runMachineAction(action, p);
  } catch (err) {
    return {success:false, msg:'שגיאה: ' + err.message};
  } finally {
    if (needLock) { try { lock.releaseLock(); } catch (e2) {} }
  }
}

function runMachineAction(action, p) {
  const who = whoIs(p);
  const isManager = who && who.role === 'manager';
  const isAdmin = p.key === ADMIN_KEY;

  // ---------- בדיקת קוד אישי ----------
  if (action === 'mCheckCode') {
    if (!who || isAdmin) return {success:false, msg:'קוד שגוי'};
    return {success:true, name:who.name, role:who.role};
  }

  // ---------- תצוגת מכונה בסריקה ----------
  if (action === 'mView') {
    const m = findMachine(p.id);
    if (!m) return {success:false, msg:'מכונה לא נמצאה'};
    const r = openRental(m.id);
    const customers = customersList();
    const available = machinesList().filter(function (x) { return x.status === ST_W && String(x.id) !== String(m.id); })
      .map(function (x) { return {id:x.id, number:x.number, model:x.model}; });
    return {
      success:true, machine:m,
      rental: r ? rentalOut(r, customers) : null,
      customers: customers, equipment: equipmentList(), available: available
    };
  }

  // ---------- חיפוש מכונה לפי מספר (הקלדה ידנית) ----------
  if (action === 'mFind') {
    const key = normNum(p.number);
    if (!key) return {success:false, msg:'לא הוקלד מספר'};
    const list = machinesList();
    let hit = list.filter(function (x) { return normNum(x.number) === key; });
    if (!hit.length) hit = list.filter(function (x) { return normNum(x.number).indexOf(key) >= 0; });
    if (hit.length === 1) return {success:true, id:hit[0].id};
    if (hit.length > 1) return {success:false, msg:'נמצאו כמה מכונות, להקליד מספר מלא', options:hit.map(function (x) { return x.number; })};
    return {success:false, msg:'לא נמצאה מכונה עם המספר הזה'};
  }

  // ---------- כל הנתונים לאפליקציית הניהול ----------
  if (action === 'mAll') {
    if (!isAdmin) return {success:false, msg:'אין הרשאה'};
    const customers = customersList();
    return {
      success:true,
      machines: machinesList(),
      rentals: rentalsList().map(function (r) { return rentalOut(r, customers); }),
      log: logList(),
      customers: customers,
      people: rowsOf('people').map(function (r) { return {name:String(r[0]), role:roleCode(r[1]), code:String(r[2])}; }),
      equipment: equipmentList(),
      today: todayIso()
    };
  }

  // מכאן והלאה כל פעולה דורשת זיהוי
  if (!who) return {success:false, msg:'קוד אישי לא תקין. יש להזין קוד מחדש', badCode:true};

  // ---------- התקנה אצל לקוח ----------
  if (action === 'mInstall') {
    const m = findMachine(p.id);
    if (!m) return {success:false, msg:'מכונה לא נמצאה'};
    if (openRental(m.id)) return {success:false, msg:'המכונה כבר בשכירות. יש לרענן את הדף'};
    if (m.status !== ST_W) return {success:false, msg:'המכונה במצב "' + m.status + '" ולא ניתן להתקין אותה'};
    const c = cleanCustomer(p);
    if (!c.name) return {success:false, msg:'חובה להזין שם לקוח'};
    if (!c.address) return {success:false, msg:'חובה להזין כתובת'};
    const start = pickDate(p.date, isManager);
    const photo = savePhoto(p.photo, m.number + '_' + start);
    ensureCustomer(c);
    const rid = openNewRental(m, c, start, String(p.equipment || ''), photo, who.name, String(p.notes || ''));
    setMachineStatus(m.id, ST_R);
    addLog(m.number, 'התקנה', c.name, who.name, 'תחילת שכירות ' + showDate(start));
    return {success:true, rentalId:rid, start:start, monthly:isMonthly(c.name)};
  }

  // ---------- סיום שכירות ----------
  if (action === 'mEnd') {
    const m = findMachine(p.id);
    if (!m) return {success:false, msg:'מכונה לא נמצאה'};
    const r = openRental(m.id);
    if (!r) return {success:false, msg:'אין שכירות פתוחה למכונה הזאת'};
    const end = pickDate(p.date, isManager);
    if (end < r.start) return {success:false, msg:'תאריך הסיום לפני תאריך ההתחלה'};
    const d = closeRental(r, end, who.name, 'סיום');
    setMachineStatus(m.id, ST_W);
    addLog(m.number, 'סיום שכירות', r.customer, who.name, d + ' ימים (' + showDate(r.start) + ' עד ' + showDate(end) + ')');
    return {success:true, days:d, customer:r.customer, start:r.start, end:end, equipment:r.equipment};
  }

  // ---------- העברה ישירה ללקוח אחר ----------
  if (action === 'mTransfer') {
    const m = findMachine(p.id);
    if (!m) return {success:false, msg:'מכונה לא נמצאה'};
    const r = openRental(m.id);
    if (!r) return {success:false, msg:'אין שכירות פתוחה למכונה הזאת'};
    const c = cleanCustomer(p);
    if (!c.name) return {success:false, msg:'חובה להזין שם לקוח חדש'};
    if (!c.address) return {success:false, msg:'חובה להזין כתובת'};
    const date = pickDate(p.date, isManager);
    if (date < r.start) return {success:false, msg:'התאריך לפני תחילת השכירות הקודמת'};
    const photo = savePhoto(p.photo, m.number + '_' + date);
    const d = closeRental(r, date, who.name, 'העברה ל' + c.name);
    ensureCustomer(c);
    openNewRental(m, c, date, String(p.equipment || ''), photo, who.name, String(p.notes || ''));
    setMachineStatus(m.id, ST_R);
    addLog(m.number, 'העברה', c.name, who.name, 'מ' + r.customer + ' (' + d + ' ימים) אל ' + c.name);
    return {success:true, days:d, from:r.customer, to:c.name};
  }

  // ---------- החלפת מכונה תקולה ----------
  if (action === 'mSwap') {
    const oldM = findMachine(p.id);
    if (!oldM) return {success:false, msg:'מכונה לא נמצאה'};
    const r = openRental(oldM.id);
    if (!r) return {success:false, msg:'אין שכירות פתוחה למכונה הזאת'};
    const newM = findMachine(p.newId);
    if (!newM) return {success:false, msg:'המכונה החלופית לא נמצאה'};
    if (String(newM.id) === String(oldM.id)) return {success:false, msg:'יש לבחור מכונה אחרת'};
    if (newM.status !== ST_W || openRental(newM.id)) return {success:false, msg:'המכונה החלופית לא פנויה'};
    const date = pickDate(p.date, isManager);
    if (date < r.start) return {success:false, msg:'התאריך לפני תחילת השכירות'};
    const d = closeRental(r, date, who.name, 'הוחלפה ב־' + newM.number);
    setMachineStatus(oldM.id, ST_F);
    openNewRental(newM, {name:r.customer, address:r.address, contact:r.contact}, date, r.equipment, '', who.name, 'מחליפה את ' + oldM.number);
    setMachineStatus(newM.id, ST_R);
    addLog(oldM.number, 'הוחלפה (תקולה)', r.customer, who.name, 'הוחלפה במכונה ' + newM.number + ' אחרי ' + d + ' ימים');
    addLog(newM.number, 'התקנה כמחליפה', r.customer, who.name, 'במקום מכונה ' + oldM.number);
    return {success:true, days:d, customer:r.customer, newNumber:newM.number, monthly:isMonthly(r.customer)};
  }

  // מכאן והלאה רק מנהל
  if (!isManager) return {success:false, msg:'פעולה למנהל בלבד'};

  // ---------- שינוי מצב מכונה ----------
  if (action === 'mSetStatus') {
    const m = findMachine(p.id);
    if (!m) return {success:false, msg:'מכונה לא נמצאה'};
    if (openRental(m.id)) return {success:false, msg:'המכונה בשכירות. קודם לסיים שכירות'};
    if ([ST_W, ST_M, ST_F].indexOf(p.status) < 0) return {success:false, msg:'מצב לא תקין'};
    setMachineStatus(m.id, p.status);
    addLog(m.number, 'שינוי מצב', '', who.name, m.status + ' ← ' + p.status);
    return {success:true};
  }

  // ---------- תיקון פרטי שכירות ----------
  if (action === 'mEditRental') {
    const r = rentalsList().filter(function (x) { return String(x.id) === String(p.rentalId); })[0];
    if (!r) return {success:false, msg:'שכירות לא נמצאה'};
    const c = cleanCustomer(p);
    const name = c.name || r.customer;
    const address = p.address !== undefined ? c.address : r.address;
    const contact = p.contact !== undefined ? c.contact : r.contact;
    const start = validIso(p.start) ? p.start : r.start;
    let end = r.end;
    if (r.end && validIso(p.end)) end = p.end;
    if (!r.end && validIso(p.end)) return {success:false, msg:'כדי לסגור שכירות יש להשתמש ב"סיום שכירות"'};
    if (end && end < start) return {success:false, msg:'תאריך הסיום לפני תאריך ההתחלה'};
    const equipment = p.equipment !== undefined ? String(p.equipment) : r.equipment;
    const notes = p.notes !== undefined ? String(p.notes) : r.notes;
    if (name !== r.customer) ensureCustomer({name:name, address:address, contact:contact});
    const sh = mSheet('rentals');
    sh.getRange(r.row, 4, 1, 6).setValues([[T(name), T(address), T(contact), T(start), T(end), end ? daysBetween(start, end) : '']]);
    sh.getRange(r.row, 10).setValue(T(equipment));
    sh.getRange(r.row, 15).setValue(T(notes));
    const changes = [];
    if (name !== r.customer) changes.push('לקוח: ' + r.customer + ' ← ' + name);
    if (start !== r.start) changes.push('התחלה: ' + showDate(r.start) + ' ← ' + showDate(start));
    if (end !== r.end) changes.push('סיום: ' + showDate(r.end) + ' ← ' + showDate(end));
    if (address !== r.address) changes.push('כתובת עודכנה');
    if (contact !== r.contact) changes.push('איש קשר עודכן');
    if (equipment !== r.equipment) changes.push('ציוד עודכן');
    addLog(r.number, 'תיקון', name, who.name, changes.join(' · ') || 'ללא שינוי');
    return {success:true};
  }

  // מכאן והלאה רק מאפליקציית הניהול
  if (!isAdmin) return {success:false, msg:'פעולה מאפליקציית הניהול בלבד'};

  // ---------- צי המכונות ----------
  if (action === 'mAddMachine') {
    const number = String(p.number || '').trim();
    if (!number) return {success:false, msg:'חובה להזין מספר מכונה'};
    if (machinesList().some(function (x) { return normNum(x.number) === normNum(number); })) return {success:false, msg:'מכונה עם המספר הזה כבר קיימת'};
    const id = nextId(machinesList());
    mSheet('machines').appendRow([id, T(number), T(p.model || ''), ST_W, T(p.notes || '')]);
    addLog(number, 'מכונה נוספה לצי', '', who.name, p.model || '');
    return {success:true, id:id};
  }

  if (action === 'mUpdateMachine') {
    const m = findMachine(p.id);
    if (!m) return {success:false, msg:'מכונה לא נמצאה'};
    const number = String(p.number || '').trim() || m.number;
    if (normNum(number) !== normNum(m.number) && machinesList().some(function (x) { return normNum(x.number) === normNum(number); })) {
      return {success:false, msg:'מכונה עם המספר הזה כבר קיימת'};
    }
    mSheet('machines').getRange(m.row, 2, 1, 2).setValues([[T(number), T(p.model !== undefined ? p.model : m.model)]]);
    mSheet('machines').getRange(m.row, 5).setValue(T(p.notes !== undefined ? p.notes : m.notes));
    if (number !== m.number) {
      const sh = mSheet('rentals');
      rentalsList().forEach(function (r) { if (String(r.machineId) === String(m.id)) sh.getRange(r.row, 3).setValue(T(number)); });
      addLog(number, 'מספר מכונה שונה', '', who.name, m.number + ' ← ' + number);
    }
    return {success:true};
  }

  if (action === 'mDeleteMachine') {
    const m = findMachine(p.id);
    if (!m) return {success:false, msg:'מכונה לא נמצאה'};
    if (openRental(m.id)) return {success:false, msg:'המכונה בשכירות. קודם לסיים שכירות'};
    mSheet('machines').deleteRow(m.row);
    addLog(m.number, 'מכונה הוסרה מהצי', '', who.name, '');
    return {success:true};
  }

  // ---------- אנשים וקודים ----------
  if (action === 'mAddPerson') {
    const name = String(p.name || '').trim();
    if (!name) return {success:false, msg:'חובה להזין שם'};
    const code = newCode();
    mSheet('people').appendRow([T(name), roleLabel(p.role), T(code)]);
    return {success:true, code:code};
  }

  if (action === 'mResetCode' || action === 'mDeletePerson' || action === 'mSetRole') {
    const sh = mSheet('people');
    const rows = rowsOf('people');
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][2]) === String(p.personCode)) {
        if (action === 'mDeletePerson') { sh.deleteRow(i + 2); return {success:true}; }
        if (action === 'mSetRole') { sh.getRange(i + 2, 2).setValue(roleLabel(p.role)); return {success:true}; }
        const code = newCode();
        sh.getRange(i + 2, 3).setValue(T(code));
        return {success:true, code:code};
      }
    }
    return {success:false, msg:'האדם לא נמצא'};
  }

  // ---------- לקוחות ----------
  if (action === 'mUpdateCustomer') {
    const oldName = String(p.oldName || '').trim();
    const newName = String(p.name || '').trim() || oldName;
    const sh = mSheet('customers');
    const rows = rowsOf('customers');
    let idx = -1;
    for (let i = 0; i < rows.length; i++) if (String(rows[i][0]) === oldName) idx = i;
    if (idx < 0) return {success:false, msg:'לקוח לא נמצא'};
    if (newName !== oldName && rows.some(function (r) { return String(r[0]) === newName; })) return {success:false, msg:'כבר קיים לקוח בשם הזה'};
    sh.getRange(idx + 2, 1, 1, 4).setValues([[T(newName), T(p.address || ''), T(p.contact || ''), p.monthly === '1' || p.monthly === true ? 'כן' : '']]);
    if (newName !== oldName) {
      const rsh = mSheet('rentals');
      rentalsList().forEach(function (r) { if (r.customer === oldName) rsh.getRange(r.row, 4).setValue(T(newName)); });
    }
    return {success:true};
  }

  if (action === 'mDeleteCustomer') {
    const name = String(p.name || '').trim();
    if (rentalsList().some(function (r) { return r.customer === name && !r.end; })) return {success:false, msg:'ללקוח יש מכונה בשכירות'};
    const rows = rowsOf('customers');
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === name) { mSheet('customers').deleteRow(i + 2); return {success:true}; }
    }
    return {success:false, msg:'לקוח לא נמצא'};
  }

  // ---------- ציוד נלווה ----------
  if (action === 'mAddEquipment') {
    const name = String(p.name || '').trim();
    if (!name) return {success:false, msg:'חובה להזין שם'};
    if (equipmentList().indexOf(name) >= 0) return {success:false, msg:'כבר קיים'};
    mSheet('equipment').appendRow([T(name)]);
    return {success:true};
  }

  if (action === 'mDeleteEquipment') {
    const rows = rowsOf('equipment');
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(p.name)) { mSheet('equipment').deleteRow(i + 2); return {success:true}; }
    }
    return {success:false, msg:'לא נמצא'};
  }

  return {success:false, msg:'פעולה לא מוכרת'};
}

// ================= עזרים =================

let _mss = null;
function mSheet(key) {
  const def = M_SHEETS[key];
  if (!_mss) _mss = SpreadsheetApp.openById(SS_ID);
  let sh = _mss.getSheetByName(def[0]);
  if (!sh) {
    sh = _mss.insertSheet(def[0]);
    sh.appendRow(def[1]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, def[1].length).setFontWeight('bold');
    sh.setRightToLeft(true);
  }
  return sh;
}

function rowsOf(key) {
  const v = mSheet(key).getDataRange().getValues();
  return v.slice(1).filter(function (r) { return r.join('') !== ''; });
}

// שמירה כטקסט, כדי שגוגל שיטס לא יהפוך מספרים ותאריכים לדברים אחרים
function T(v) {
  const s = (v === undefined || v === null) ? '' : String(v);
  return s === '' ? '' : "'" + s;
}

function iso(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Jerusalem', 'yyyy-MM-dd');
  return v === undefined || v === null ? '' : String(v);
}

function todayIso() { return Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd'); }
function nowTime() { return Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'HH:mm'); }
function validIso(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')); }

function showDate(s) {
  s = iso(s);
  if (!validIso(s)) return s;
  const a = s.split('-');
  return a[2] + '/' + a[1] + '/' + a[0];
}

// טכנאי: תמיד היום. מנהל: אפשר לקבוע תאריך אחר
function pickDate(d, isManager) {
  if (isManager && validIso(d)) return d;
  return todayIso();
}

// ספירה כולל יום ההתחלה ויום הסיום
function daysBetween(a, b) {
  a = iso(a); b = iso(b);
  if (!validIso(a) || !validIso(b)) return '';
  const pa = a.split('-'), pb = b.split('-');
  const da = Date.UTC(+pa[0], +pa[1] - 1, +pa[2]);
  const db = Date.UTC(+pb[0], +pb[1] - 1, +pb[2]);
  return Math.round((db - da) / 86400000) + 1;
}

function normNum(s) {
  return String(s || '').toUpperCase().replace(/\s+/g, '').replace(/_/g, '-');
}

function roleLabel(r) { return r === 'manager' ? 'מנהל' : 'טכנאי'; }
function roleCode(r) { return String(r) === 'מנהל' ? 'manager' : 'tech'; }

function whoIs(p) {
  if (p.key && p.key === ADMIN_KEY) return {name:'ניהול', role:'manager'};
  const code = String(p.code || '').trim();
  if (!code) return null;
  const rows = rowsOf('people');
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][2]) === code) return {name:String(rows[i][0]), role:roleCode(rows[i][1])};
  }
  return null;
}

function newCode() {
  const used = rowsOf('people').map(function (r) { return String(r[2]); });
  for (let i = 0; i < 200; i++) {
    const c = String(Math.floor(1000 + Math.random() * 9000));
    if (used.indexOf(c) < 0) return c;
  }
  throw new Error('לא נמצא קוד פנוי');
}

function nextId(list) {
  let max = 0;
  list.forEach(function (x) { const n = parseInt(x.id, 10); if (n > max) max = n; });
  return max + 1;
}

function machinesList() {
  return rowsOf('machines').map(function (r, i) {
    return {row:i + 2, id:String(r[0]), number:String(r[1]), model:String(r[2]), status:String(r[3]) || ST_W, notes:String(r[4])};
  });
}

function findMachine(id) {
  if (id === undefined || id === null || id === '') return null;
  return machinesList().filter(function (x) { return String(x.id) === String(id); })[0] || null;
}

function setMachineStatus(id, status) {
  const m = findMachine(id);
  if (m) mSheet('machines').getRange(m.row, 4).setValue(status);
}

function rentalsList() {
  return rowsOf('rentals').map(function (r, i) {
    return {
      row:i + 2, id:String(r[0]), machineId:String(r[1]), number:String(r[2]),
      customer:String(r[3]), address:String(r[4]), contact:String(r[5]),
      start:iso(r[6]), end:iso(r[7]), days:r[8], equipment:String(r[9]), photo:String(r[10]),
      openedBy:String(r[11]), closedBy:String(r[12]), closeType:String(r[13]), notes:String(r[14])
    };
  });
}

function openRental(machineId) {
  const list = rentalsList().filter(function (r) { return String(r.machineId) === String(machineId) && !r.end; });
  return list.length ? list[list.length - 1] : null;
}

function rentalOut(r, customers) {
  const o = {};
  Object.keys(r).forEach(function (k) { if (k !== 'row') o[k] = r[k]; });
  o.daysSoFar = r.end ? r.days : daysBetween(r.start, todayIso());
  const c = (customers || customersList()).filter(function (x) { return x.name === r.customer; })[0];
  o.monthly = !!(c && c.monthly);
  return o;
}

function openNewRental(m, c, start, equipment, photo, by, notes) {
  const id = nextId(rentalsList());
  mSheet('rentals').appendRow([id, m.id, T(m.number), T(c.name), T(c.address), T(c.contact), T(start), '', '', T(equipment), photo || '', T(by), '', '', T(notes)]);
  return id;
}

function closeRental(r, end, by, type) {
  const d = daysBetween(r.start, end);
  const sh = mSheet('rentals');
  sh.getRange(r.row, 8, 1, 2).setValues([[T(end), d]]);
  sh.getRange(r.row, 13, 1, 2).setValues([[T(by), T(type)]]);
  return d;
}

function cleanCustomer(p) {
  return {
    name: String(p.customer || '').trim(),
    address: String(p.address || '').trim(),
    contact: String(p.contact || '').trim()
  };
}

function customersList() {
  return rowsOf('customers').map(function (r) {
    return {name:String(r[0]), address:String(r[1]), contact:String(r[2]), monthly:String(r[3]) === 'כן'};
  });
}

function isMonthly(name) {
  return customersList().some(function (c) { return c.name === name && c.monthly; });
}

// לקוח חדש נוסף לרשימה. לקוח קיים לא משתנה כאן (רק מנהל משנה מאפליקציית הניהול)
function ensureCustomer(c) {
  if (!c.name) return;
  if (customersList().some(function (x) { return x.name === c.name; })) return;
  mSheet('customers').appendRow([T(c.name), T(c.address), T(c.contact), '']);
}

function equipmentList() {
  return rowsOf('equipment').map(function (r) { return String(r[0]); }).filter(String);
}

function logList() {
  return rowsOf('mlog').map(function (r) {
    return {date:iso(r[0]), time:String(r[1] instanceof Date ? Utilities.formatDate(r[1], 'Asia/Jerusalem', 'HH:mm') : r[1]),
      number:String(r[2]), action:String(r[3]), customer:String(r[4]), who:String(r[5]), details:String(r[6])};
  });
}

function addLog(number, action, customer, who, details) {
  mSheet('mlog').appendRow([T(todayIso()), T(nowTime()), T(number), T(action), T(customer), T(who), T(details)]);
}

function savePhoto(b64, label) {
  if (!b64) return '';
  const data = String(b64).indexOf(',') >= 0 ? String(b64).split(',')[1] : String(b64);
  const safe = String(label).replace(/[\/\\:*?"<>|]/g, '-');
  const blob = Utilities.newBlob(Utilities.base64Decode(data), 'image/jpeg', safe + '.jpg');
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
  const f = folder.createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1200';
}
