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
  return sh;
}

function listBookings() {
  var rows = getSheet().getDataRange().getValues();
  var bookings = {};
  for (var i = 1; i < rows.length; i++) {
    var slot = String(rows[i][0]).trim();
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
      if (String(rows[i][0]).trim() === slot) slotOwner = String(rows[i][1]).trim();
      if (String(rows[i][1]).trim().toLowerCase() === name.toLowerCase()) myRow = i;
    }

    // Slot already held by someone else → reject.
    if (slotOwner && slotOwner.toLowerCase() !== name.toLowerCase()) {
      return json({ ok: false, reason: 'taken', bookings: listBookings().bookings });
    }

    var now = new Date();
    if (myRow > -1) {
      // Guest already in the sheet — move them to the new slot.
      sh.getRange(myRow + 1, 1).setValue(slot);
      sh.getRange(myRow + 1, 3).setValue(now);
    } else {
      sh.appendRow([slot, name, now]);
    }
    return json({ ok: true, slot: slot, name: name, bookings: listBookings().bookings });
  } finally {
    lock.releaseLock();
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
