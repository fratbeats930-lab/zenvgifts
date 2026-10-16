/**
 * Dream India Auditions — automatic Google Sheets setup.
 * Open your Google Sheet > Extensions > Apps Script and paste this entire file.
 * Save, select setup, then Run and authorize your own Google account.
 * Sheet ID and a private shared secret are configured automatically.
 * Finally deploy as a Web app: Execute as Me, access Anyone.
 */

function setup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Apni Google Sheet se Extensions > Apps Script kholein aur setup dobara Run karein.');
  }
  const properties = PropertiesService.getScriptProperties();
  const previousSheetId = properties.getProperty('SHEET_ID');
  if (previousSheetId && previousSheetId !== spreadsheet.getId()) {
    throw new Error('Ye script pehle doosri Google Sheet se connected hai. Us Sheet ke Extensions > Apps Script se setup kholein.');
  }
  let secret = properties.getProperty('SHARED_SECRET');
  if (!secret || secret.length < 32 || secret.startsWith('REPLACE_') || /[\x00-\x1f\x7f]/.test(secret)) {
    secret = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40);
  }
  const sheet = spreadsheet.getSheetByName('Auditions') || spreadsheet.insertSheet('Auditions');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Registration ID', 'Received Date & Time (IST)', 'Timezone', 'Name', 'Age', 'Gender', 'WhatsApp', 'Alternate Mobile', 'Email', 'City', 'Audition Mode', 'Consent', 'WordPress Received (IST)']);
  }
  sheet.getRange(1, 11).setValue('Audition Mode');
  sheet.setFrozenRows(1);
  properties.setProperties({SHEET_ID: spreadsheet.getId(), SHARED_SECRET: secret});
  onOpen();
  showWordPressConnection();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Dream India')
    .addItem('WordPress connection', 'showWordPressConnection')
    .addToUi();
}

function showWordPressConnection() {
  const secret = PropertiesService.getScriptProperties().getProperty('SHARED_SECRET');
  if (!secret || secret.length < 32) {
    throw new Error('Pehle function dropdown se setup select karke Run karein.');
  }
  let deploymentUrl = '';
  try {
    deploymentUrl = ScriptApp.getService().getUrl() || '';
  } catch (error) {
    // The owner can always copy the Web app URL directly from the Deploy screen.
  }
  const validUrl = /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(deploymentUrl);
  const escapeHtml = value => String(value).replace(/[&<>"']/g, character => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[character]);
  const html = `<!doctype html><html><head><meta charset="utf-8"><base target="_blank">
  <style>
    body{font:14px/1.6 Arial,sans-serif;color:#241a2c;padding:18px;margin:0;background:#faf6ef}
    h2{font-size:21px;line-height:1.25;margin:0 0 10px}p{margin:10px 0}
    label{display:block;font-weight:700;margin-top:18px}input{box-sizing:border-box;width:100%;padding:11px;border:1px solid #c6bdb4;border-radius:6px;font:13px monospace;background:#fff}
    button{padding:10px 16px;border:0;border-radius:6px;background:#d34c20;color:#fff;cursor:pointer;margin-top:8px;font-weight:700}
    small{display:block;color:#6d5f67;margin:5px 0}#copy-status{min-height:24px;font-weight:700;color:#23673a}
    ol{padding-left:22px}code{background:#f0e8de;padding:1px 4px}
  </style></head><body>
  <h2>WordPress connection ready</h2>
  <p>Sheet ID aur secret khud taiyar ho gaye hain. Code mein kuch badalne ki zaroorat nahi hai.</p>
  <label for="dip-secret">1. Shared secret</label>
  <input id="dip-secret" readonly autocomplete="off" value="${escapeHtml(secret)}" onclick="this.select()">
  <button type="button" data-copy="dip-secret">Copy shared secret</button>
  <small>WordPress &gt; Settings &gt; Dream India Auditions &gt; Shared secret mein paste karein. Quotes nahi lagane hain.</small>
  ${validUrl ? `<label for="dip-endpoint">2. Apps Script /exec URL</label>
  <input id="dip-endpoint" readonly value="${escapeHtml(deploymentUrl)}" onclick="this.select()">
  <button type="button" data-copy="dip-endpoint">Copy /exec URL</button>
  <small>WordPress ke Apps Script /exec URL field mein paste karein.</small>` : `<p><b>2. Web app deploy karein:</b> Apps Script mein Deploy &gt; New deployment &gt; Web app. Execute as <b>Me</b> aur Who has access <b>Anyone</b> rakhein. Deploy se mila <code>/exec</code> URL WordPress ke Apps Script /exec URL field mein paste karein.</p>`}
  <p id="copy-status" role="status" aria-live="polite"></p>
  <p><b>WordPress mein Save Changes</b> dabayein. Phir ek test registration karke Google Sheet ke Auditions tab mein uski row check karein.</p>
  <small>Ye window baad mein Google Sheet ke Dream India &gt; WordPress connection menu se dobara khol sakte hain. Secret ko private rakhein.</small>
  <script>
    document.querySelectorAll('[data-copy]').forEach(function(button){
      button.addEventListener('click',async function(){
        var field=document.getElementById(button.getAttribute('data-copy'));
        var status=document.getElementById('copy-status');
        var copied=false;
        try{
          if(navigator.clipboard&&navigator.clipboard.writeText){await navigator.clipboard.writeText(field.value);copied=true;}
        }catch(error){}
        if(!copied){field.focus();field.select();try{copied=document.execCommand('copy');}catch(error){}}
        status.textContent=copied?'Copied! Ab WordPress ke sahi field mein paste karein.':'Code select ho gaya hai. Ctrl+C ya phone ke Copy option se copy karein.';
      });
    });
  </script></body></html>`;
  const output = HtmlService.createHtmlOutput(html).setWidth(560).setHeight(610);
  SpreadsheetApp.getUi().showModalDialog(output, 'Dream India — WordPress connection');
}


function doPost(e) {
  const reply = object => ContentService.createTextOutput(JSON.stringify(object)).setMimeType(ContentService.MimeType.JSON);
  let lock;
  try {
    const properties = PropertiesService.getScriptProperties();
    const data = JSON.parse(e.postData.contents);
    if (!data || !properties.getProperty('SHARED_SECRET') || data.secret !== properties.getProperty('SHARED_SECRET')) {
      return reply({ok: false, error: 'Unauthorized'});
    }
    const text = (value, min, max) => typeof value === 'string' && value.length >= min && value.length <= max;
    if (!/^DIP-[A-F0-9]{12}$/.test(data.registration_id || '')
        || !text(data.full_name, 2, 100) || !/^\d+$/.test(data.age || '') || Number(data.age) < 5 || Number(data.age) > 100
        || !['Female', 'Male', 'Other', 'Prefer not to say'].includes(data.gender)
        || !/^[6-9][0-9]{9}$/.test(data.whatsapp || '')
        || (data.alternate && !/^[6-9][0-9]{9}$/.test(data.alternate))
        || !text(data.city, 2, 100) || !text(data.audition_mode, 1, 500)
        || data.consent !== '1' || data.timezone !== 'Asia/Kolkata'
        || !/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(data.timestamp_ist || '')) {
      return reply({ok: false, error: 'Invalid registration'});
    }
    lock = LockService.getScriptLock();
    lock.waitLock(20000);
    const sheet = SpreadsheetApp.openById(properties.getProperty('SHEET_ID')).getSheetByName('Auditions');
    if (!sheet) {
      return reply({ok: false, error: 'Run setup first'});
    }
    // Retries reuse the same canonical ID, so a delayed response cannot add another row.
    if (sheet.getLastRow() > 1 && sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).createTextFinder(data.registration_id).matchEntireCell(true).findNext()) {
      return reply({ok: true, registration_id: data.registration_id});
    }
    const safe = value => {
      const string = String(value == null ? '' : value).slice(0, 500);
      return /^\s*[=+\-@]/.test(string) ? "'" + string : string;
    };
    const row = [data.registration_id, Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss'), 'Asia/Kolkata', data.full_name, data.age, data.gender, data.whatsapp, data.alternate, data.email, data.city, data.audition_mode, 'Yes', data.timestamp_ist].map(safe);
    const range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length);
    range.setNumberFormat('@');
    range.setValues([row]);
    SpreadsheetApp.flush();
    return reply({ok: true, registration_id: data.registration_id});
  } catch (error) {
    return reply({ok: false, error: 'Could not save. Please retry.'});
  } finally {
    if (lock && lock.hasLock()) {
      lock.releaseLock();
    }
  }
}
