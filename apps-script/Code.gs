// Writes a timestamped line to a "Debug" sheet tab (auto-created) in the same
// spreadsheet, since Executions/Cloud Logging aren't easily reachable for
// this project. Safe to delete this function and its call sites once things
// are confirmed working.
function logDebug(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Debug');
  if (!sheet) sheet = ss.insertSheet('Debug');
  sheet.appendRow([new Date(), message]);
}

// Lightweight flood/spam guard, on top of the client-side honeypot and
// submit-timing check. Apps Script has no per-request IP, so this throttles
// by submitted email plus an overall burst cap — enough to stop a script
// hammering this endpoint from burning through the daily Mail/Sheets quota.
// Silently reports success to a throttled caller so a bot has nothing to
// adapt to.
function isRateLimited(email) {
  const cache = CacheService.getScriptCache();

  const globalKey = 'rsvp_burst_count';
  const count = parseInt(cache.get(globalKey) || '0', 10);
  if (count >= 20) return true; // >20 submissions in the current 60s window
  cache.put(globalKey, String(count + 1), 60);

  const emailKey = 'rsvp_cooldown_' + (email || '').toLowerCase();
  if (cache.get(emailKey)) return true; // same email submitted in the last 30s
  cache.put(emailKey, '1', 30);

  return false;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (isRateLimited(data.email)) {
      return ContentService
        .createTextOutput(JSON.stringify({ result: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raspunsuri');

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Data trimiterii', 'Prenume', 'Nume', 'Email',
        'Participă', 'Nr. invitați', 'Cerințe alimentare',
        'Detalii alimentare', 'Cazare', 'Mesaj', 'Limbă'
      ]);
    }

    sheet.appendRow([
      Utilities.formatDate(new Date(data.submittedAt), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
      data.firstName,
      data.lastName,
      data.email,
      data.attendance === 'yes' ? 'Da' : 'Nu',
      data.attendance === 'yes' ? data.guests : '',
      data.dietary === 'yes' ? 'Da' : 'Nu',
      data.dietaryText || '',
      data.accommodation || '',
      data.message || '',
      data.lang || '',
    ]);

    // A notification-email failure shouldn't erase a successfully saved RSVP,
    // so it's isolated in its own try/catch and just logged for later review
    // (Apps Script editor → Executions) instead of failing the whole request.
    try {
      MailApp.sendEmail({
        to: 'cristitone96@gmail.com, petrescu.diana2@yahoo.com',
        subject: `Răspuns nou: ${data.firstName} ${data.lastName}`,
        body: `
          Răspuns nou de la formular:

          Nume: ${data.firstName} ${data.lastName}
          Email: ${data.email}
          Participă: ${data.attendance === 'yes' ? 'Da' : 'Nu'}
          Nr. invitați: ${data.attendance === 'yes' ? data.guests : 'N/A'}
          Cerințe alimentare: ${data.dietary === 'yes' ? 'Da' : 'Nu'}
          Detalii: ${data.dietaryText || '-'}
          Cazare: ${data.accommodation || '-'}
          Mesaj: ${data.message || '-'}
          Limbă: ${data.lang || '-'}
          Trimis: ${Utilities.formatDate(new Date(data.submittedAt), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')}
        `.trim()
      });
    } catch (mailErr) {
      logDebug('MailApp.sendEmail failed: ' + mailErr.message
        + ' | remaining daily quota: ' + MailApp.getRemainingDailyQuota());
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    try { logDebug('doPost threw: ' + err.message); } catch (_) { /* Debug sheet itself unreachable */ }
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
