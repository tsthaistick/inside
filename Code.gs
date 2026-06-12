const SHEET_ID = '16tLF448AfoNqaHJX0FkbneAfdJKTN1jyhlb9-cA_fLQ';

function doGet(e) {
  const sheetName = e.parameter.sheet;
  const callback  = e.parameter.callback;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ws = ss.getSheetByName(sheetName);

  if (!ws) {
    return ContentService
      .createTextOutput(callback + '({"error":"Sheet not found"})')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  const data    = ws.getDataRange().getValues();
  const headers = data[0];
  const rows    = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });

  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(rows) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.openById(SHEET_ID);
    const ts   = new Date();

    if (data.sheet === 'Enquiries') {
      let ws = ss.getSheetByName('Enquiries');
      if (!ws) {
        ws = ss.insertSheet('Enquiries');
        ws.appendRow(['timestamp','first_name','last_name','company','email','phone','message','meeting_date']);
      }
      const eRow = ws.getLastRow() + 1;
      ws.appendRow([ts, data.first_name||'', data.last_name||'', data.company||'', data.email||'', data.phone||'', data.message||'', data.meeting_date||'']);
      ws.getRange(eRow, 6).setNumberFormat('@');

    } else if (data.sheet === 'VisitorLog') {
      let ws = ss.getSheetByName('Visitor Log');
      if (!ws) {
        ws = ss.insertSheet('Visitor Log');
        ws.appendRow(['timestamp','name','company','email','phone']);
      }
      const vRow = ws.getLastRow() + 1;
      ws.appendRow([ts, data.name||'', data.company||'', data.email||'', data.phone||'']);
      ws.getRange(vRow, 5).setNumberFormat('@');

    } else {
      let ws = ss.getSheetByName('Rep_Log');
      if (!ws) {
        ws = ss.insertSheet('Rep_Log');
        ws.appendRow(['timestamp','rep_name','clinic_name','contact_name','email','contact_number','interested_sku','next_meeting_date','next_meeting_time','notes','outcome']);
      }
      const rRow = ws.getLastRow() + 1;
      ws.appendRow([ts, data.rep_name||'', data.clinic_name||'', data.contact_name||'', data.email||'', data.contact_number||'', data.interested_sku||'', data.next_meeting_date||'', data.next_meeting_time||'', data.notes||'', data.outcome||'']);
      ws.getRange(rRow, 6).setNumberFormat('@');
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
