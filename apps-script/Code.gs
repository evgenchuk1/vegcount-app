const KEY = 'veg118';

// TEST mode: send only to etsipis@013.net. After final approval -> TEST_MODE = false
const TEST_MODE = true;
const TEST_RECIPIENTS = ['etsipis@013.net'];
const RECIPIENTS = ['snif118m@shufersal.co.il', 'snif118@shufersal.co.il', 'etsipis@013.net'];

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.key !== KEY) return out({ ok: false, error: 'bad key' });
    const blob = Utilities.newBlob(Utilities.base64Decode(p.b64),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', p.filename);
    const to = (TEST_MODE ? TEST_RECIPIENTS : RECIPIENTS).join(',');
    GmailApp.sendEmail(to, p.subject, p.body || '', { attachments: [blob], name: 'Vegcount 118' });
    return out({ ok: true, to });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
