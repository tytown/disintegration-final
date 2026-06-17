/**
 * DISINTEGRATION — RSVP backend (Google Apps Script web app)
 *
 * Stores one guest per time slot in the bound Google Sheet and returns the
 * current bookings. Called from the invite page via simple GET requests
 * (no CORS preflight). Deploy: Extensions ▸ Apps Script ▸ paste this ▸
 * Deploy ▸ New deployment ▸ Web app ▸ Execute as: Me ▸ Who has access: Anyone.
 */

var SHEET_NAME = 'RSVPs';
var SLOTS = ['7:00','7:20','7:40','8:00','8:20','8:40','9:00','9:20','9:40'];

function doGet(e) {
  var action = ((e && e.parameter && e.parameter.action) || '').toLowerCase();
  if (action === 'rsvp') {
    return handleRsvp(e.parameter.name, e.parameter.slot);
  }
  if (action === 'cancel') {
    return handleCancel(e.parameter.name);
  }
  return json(listBookings());
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

function listBookings() {
  var rows = getSheet().getDataRange().getValues();
  var bookings = {};
  for (var i = 1; i < rows.length; i++) {
    var slot = slotStr(rows[i][0]);
    var nm = String(rows[i][1]).trim();
    if (slot && nm) bookings[slot] = nm;
  }
  return { ok: true, bookings: bookings };
}

function handleRsvp(name, slot) {
  name = String(name || '').trim();
  slot = String(slot || '').trim();
  if (!name || SLOTS.indexOf(slot) < 0) {
    return json({ ok: false, reason: 'invalid' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return json({ ok: false, reason: 'busy' });
  }

  try {
    var sh = getSheet();
    var rows = sh.getDataRange().getValues();
    var myRow = -1, slotOwner = '';
    for (var i = 1; i < rows.length; i++) {
      if (slotStr(rows[i][0]) === slot) slotOwner = String(rows[i][1]).trim();
      if (String(rows[i][1]).trim().toLowerCase() === name.toLowerCase()) myRow = i;
    }

    if (slotOwner && slotOwner.toLowerCase() !== name.toLowerCase()) {
      return json({ ok: false, reason: 'taken', bookings: listBookings().bookings });
    }

    var now = new Date();
    var row = (myRow > -1) ? (myRow + 1) : (sh.getLastRow() + 1);
    sh.getRange(row, 1).setNumberFormat('@').setValue(slot); // force text "9:40"
    sh.getRange(row, 2).setValue(name);
    sh.getRange(row, 3).setValue(now);
    return json({ ok: true, slot: slot, name: name, bookings: listBookings().bookings });
  } finally {
    lock.releaseLock();
  }
}

function handleCancel(name) {
  name = String(name || '').trim();
  if (!name) return json({ ok: false, reason: 'invalid' });

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return json({ ok: false, reason: 'busy' });
  }

  try {
    var sh = getSheet();
    var rows = sh.getDataRange().getValues();
    // delete bottom-up so row indices stay valid
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][1]).trim().toLowerCase() === name.toLowerCase()) {
        sh.deleteRow(i + 1);
      }
    }
    return json({ ok: true, cancelled: true, bookings: listBookings().bookings });
  } finally {
    lock.releaseLock();
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
