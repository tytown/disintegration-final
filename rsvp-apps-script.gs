/**
 * DISINTEGRATION — RSVP backend (Google Apps Script web app)
 *
 * One guest per time slot, plus declines, stored in the bound Google Sheet.
 * Called from the invite page via simple GET requests (no CORS preflight).
 * Deploy: Extensions ▸ Apps Script ▸ paste this ▸ Deploy ▸ New deployment ▸
 * Web app ▸ Execute as: Me ▸ Who has access: Anyone. After edits, redeploy a
 * NEW VERSION (Manage deployments ▸ Edit ▸ New version) — same /exec URL.
 */

var SHEET_NAME = 'RSVPs';
var SLOTS = ['7:00','7:20','7:40','8:00','8:20','8:40','9:00','9:20','9:40'];
var DECLINED = 'DECLINED'; // sentinel stored in the Slot column for a "can't make it"

function doGet(e) {
  var action = ((e && e.parameter && e.parameter.action) || '').toLowerCase();
  if (action === 'rsvp')    return handleRsvp(e.parameter.name, e.parameter.slot);
  if (action === 'cancel')  return handleCancel(e.parameter.name);
  if (action === 'decline') return handleDecline(e.parameter.name);
  return json(listState());
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Slot', 'Name', 'Reserved at']);
    sh.setFrozenRows(1);
  }
  sh.getRange('A1:A').setNumberFormat('@'); // keep Slot column as plain text
  return sh;
}

// Slots can come back from the sheet as a Date if a cell was ever time-typed;
// normalise everything to "H:MM".
function slotStr(v) {
  if (v instanceof Date) {
    var h = v.getHours(), m = v.getMinutes();
    return h + ':' + (m < 10 ? '0' + m : m);
  }
  return String(v).trim();
}

function listState() {
  var rows = getSheet().getDataRange().getValues();
  var bookings = {};
  var declined = [];
  for (var i = 1; i < rows.length; i++) {
    var slot = slotStr(rows[i][0]);
    var nm = String(rows[i][1]).trim();
    if (!nm) continue;
    if (slot === DECLINED) declined.push(nm);
    else if (slot) bookings[slot] = nm;
  }
  return { ok: true, bookings: bookings, declined: declined };
}

function handleRsvp(name, slot) {
  name = String(name || '').trim();
  slot = String(slot || '').trim();
  if (!name || SLOTS.indexOf(slot) < 0) {
    return json({ ok: false, reason: 'invalid' });
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) { return json({ ok: false, reason: 'busy' }); }

  try {
    var sh = getSheet();
    var rows = sh.getDataRange().getValues();
    var myRow = -1, slotOwner = '';
    for (var i = 1; i < rows.length; i++) {
      if (slotStr(rows[i][0]) === slot) slotOwner = String(rows[i][1]).trim();
      if (String(rows[i][1]).trim().toLowerCase() === name.toLowerCase()) myRow = i;
    }

    if (slotOwner && slotOwner.toLowerCase() !== name.toLowerCase()) {
      return json(extend({ ok: false, reason: 'taken' }, listState()));
    }

    // Reserving overwrites any prior reservation OR decline for this guest.
    var now = new Date();
    var row = (myRow > -1) ? (myRow + 1) : (sh.getLastRow() + 1);
    sh.getRange(row, 1).setNumberFormat('@').setValue(slot);
    sh.getRange(row, 2).setValue(name);
    sh.getRange(row, 3).setValue(now);
    return json(extend({ ok: true, slot: slot, name: name }, listState()));
  } finally { lock.releaseLock(); }
}

function handleCancel(name) {
  name = String(name || '').trim();
  if (!name) return json({ ok: false, reason: 'invalid' });
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) { return json({ ok: false, reason: 'busy' }); }
  try {
    removeName(name);
    return json(extend({ ok: true, cancelled: true }, listState()));
  } finally { lock.releaseLock(); }
}

function handleDecline(name) {
  name = String(name || '').trim();
  if (!name) return json({ ok: false, reason: 'invalid' });
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) { return json({ ok: false, reason: 'busy' }); }
  try {
    var sh = getSheet();
    removeName(name);                       // free any held slot / dupes first
    var r = sh.getLastRow() + 1;
    sh.getRange(r, 1).setNumberFormat('@').setValue(DECLINED);
    sh.getRange(r, 2).setValue(name);
    sh.getRange(r, 3).setValue(new Date());
    return json(extend({ ok: true, didDecline: true }, listState()));
  } finally { lock.releaseLock(); }
}

// delete every row for a name (bottom-up so indices stay valid)
function removeName(name) {
  var sh = getSheet();
  var rows = sh.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][1]).trim().toLowerCase() === name.toLowerCase()) {
      sh.deleteRow(i + 1);
    }
  }
}

function extend(target, src) {
  for (var k in src) if (src.hasOwnProperty(k)) target[k] = src[k];
  return target;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
