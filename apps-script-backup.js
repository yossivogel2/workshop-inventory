function doGet(e) {
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
