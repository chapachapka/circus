// Nastya's Circus: receives RSVPs from the invitation site and keeps one row per guest.
// Paste this into the Guests spreadsheet (Extensions → Apps Script), then
// Deploy → New deployment → Web app, "Execute as: Me", "Who has access: Anyone".

const SHEET = "Guests";
const ANSWERS = { yes: "Coming", no: "Can't make it" };

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const id = String(data.id || "").slice(0, 64);
    if (!id) return reply({ ok: false, error: "missing id" });

    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET);
    const last = sheet.getLastRow();
    const ids = last > 1 ? sheet.getRange(2, 4, last - 1, 1).getValues().map(r => String(r[0])) : [];
    const row = ids.indexOf(id) + 2; // 1 when the guest has no row yet

    if (data.status === "withdrawn") {
      if (row > 1) sheet.deleteRow(row);
      return reply({ ok: true });
    }

    const answer = ANSWERS[data.status];
    if (!answer) return reply({ ok: false, error: "unknown status" });
    const values = [[plain(data.name), answer, new Date(), id]];
    if (row > 1) sheet.getRange(row, 1, 1, 4).setValues(values);
    else sheet.appendRow(values[0]);
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// guests type free text; keep it from being read as a formula
function plain(name) {
  const s = String(name || "").trim().slice(0, 40);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
