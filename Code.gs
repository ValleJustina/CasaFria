/**
 * Casa Fria Private Resort
 * Google Apps Script backend for the website.
 *
 * Does two jobs:
 *   GET   -> hands the website the availability calendar as JSON
 *   POST  -> receives a reservation request, logs it to the Bookings tab,
 *            and emails the resort
 *
 * SETUP, once:
 *   1. Open the master sheet, Extensions > Apps Script
 *   2. Delete whatever is there and paste this whole file
 *   3. Save, then in the function dropdown pick `setup` and click Run.
 *
 *      `setup` is the ONLY function you ever run by hand. Everything else in
 *      this file is called by the website and needs information passed to it,
 *      so running one on its own just reports that back to you.
 *      Google will ask you to authorise it. Choose your account, then
 *      "Advanced" > "Go to ... (unsafe)" > Allow. That warning is normal
 *      for your own scripts. It now also asks for Drive, because deposit
 *      slips are saved into a folder called "Valle Justina deposit slips".
 *   4. Deploy > New deployment > gear icon > Web app
 *        Execute as:      Me
 *        Who has access:  Anyone            <-- this one matters most
 *      Deploy, copy the /exec URL, and put it in CONFIG.dataUrl on the site.
 *
 *   IMPORTANT: every time you change this file you must Deploy > Manage
 *   deployments > edit > Version: New version > Deploy. Saving alone does
 *   not update the live URL.
 *
 *   IF YOU SEE A PERMISSION ERROR ABOUT DriveApp
 *   Google only asks for the permissions a script actually uses, and it asks
 *   once. An older version of this file never touched Drive, so Drive was
 *   never granted. Run `setup` again and accept the new prompt. Until you do,
 *   deposit slips are pinned into the sheet instead of saved to Drive, so
 *   nothing is lost either way.
 */

/* ====================== SETTINGS ====================== */

var SHEET_ID        = '1l8vRFpiyAIrcY51TFY0DGeUrJ83uKsJqfyMm4E0nRVM';
var CALENDAR_TAB    = 'Update Here';    // the Date | Room | Status tab (was "Sheet1")
var BOOKINGS_TAB    = 'Bookings';  // created automatically
var DAYPASS_TAB     = 'Day Pass';  // entrance and swimming tickets
var CAMPING_TAB     = 'Camping';   // camping tickets
var SHUTTLE_TAB     = 'Shuttle';   // round trip shuttle tickets
/* Optional: paste a dedicated Casa Fria Drive folder id here.
   When blank, setup creates or reuses the named folder below. */
var SLIP_FOLDER_ID  = '';
var SLIP_FOLDER     = 'Casa Fria deposit slips';
var NOTIFY_TO       = 'vallejustina25@gmail.com';
// everyone else who should see each request, comma separated, no spaces
var NOTIFY_CC       = 'apayaresorts@gmail.com,vallejustina@gmail.com';
var RESORT_NAME     = 'Casa Fria Private Resort';
var TZ              = 'Asia/Manila';

/* Finds the Date | Room | Status tab by name, ignoring case and stray spaces,
   and falling back to a tab still called "Sheet1" if the rename to
   CALENDAR_TAB has not happened yet on this copy of the sheet. This exists
   because a sheet gets renamed by hand from time to time, and a single
   case mismatch (e.g. "update here" typed where the code expects
   "Update Here") should never be able to silently break the whole
   calendar for the website. */
function calSheet(ss) {
  var direct = ss.getSheetByName(CALENDAR_TAB);
  if (direct) return direct;
  var all = ss.getSheets();
  var want = CALENDAR_TAB.trim().toLowerCase();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getName().trim().toLowerCase() === want) return all[i];
  }
  for (var j = 0; j < all.length; j++) {
    if (all[j].getName().trim().toLowerCase() === 'sheet1') return all[j];
  }
  return null;
}

/* ====================== ONE-TIME SETUP ====================== */

function setup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // make sure the calendar tab has the right headers
  var cal = calSheet(ss) || ss.insertSheet(CALENDAR_TAB);
  if (cal.getLastRow() === 0) {
    cal.getRange(1, 1, 1, 3).setValues([['Date', 'Room', 'Status']]).setFontWeight('bold');
    cal.setFrozenRows(1);
  }

  // create the bookings tab
  var bk = ss.getSheetByName(BOOKINGS_TAB);
  if (!bk) {
    bk = ss.insertSheet(BOOKINGS_TAB);
    bk.getRange(1, 1, 1, 16).setValues([[
      'Received', 'Reference', 'Room', 'Check in', 'Check out', 'Nights',
      'Guests', 'Total', 'Deposit slip', 'Name', 'Mobile', 'Email', 'Notes',
      'Check-in (ISO)', 'Check-out (ISO)', 'Room sheets'
    ]]).setFontWeight('bold');
    bk.setFrozenRows(1);
    bk.setColumnWidth(1, 150);
    bk.setColumnWidth(13, 320);
  }

  // create the ticket tabs
  ticketTab(ss, DAYPASS_TAB);
  ticketTab(ss, CAMPING_TAB);
  ticketTab(ss, SHUTTLE_TAB);

  // an existing Bookings tab predates the Deposit slip column, so widen it
  var bookNote = upgradeBookings(bk || ss.getSheetByName(BOOKINGS_TAB));
  upgradeBookingsIso(bk || ss.getSheetByName(BOOKINGS_TAB));

  /* Touching Drive here is deliberate. Google only asks for the permissions a
     script actually uses, and it asks once, at authorisation. If Drive is
     refused, say so plainly rather than letting it surface later as a lost
     deposit slip. */
  var driveNote;
  try {
    var f = slipFolder();
    driveNote = 'Drive OK. Deposit slips will be filed in "' + f.getName() + '".';
  } catch (driveErr) {
    driveNote = 'DRIVE NOT GRANTED. Deposit slips will be pinned into the sheet '
              + 'instead of saved to Drive, so nothing is lost, but to get proper '
              + 'links: Project Settings, tick "Show appsscript.json manifest file '
              + 'in editor", paste the appsscript.json that came with this file, '
              + 'save, then run setup again and accept the permission screen. '
              + '(' + driveErr + ')';
  }

  // sending one mail here forces the mail permission prompt now rather than
  // failing silently on the first real booking
  var mailNote;
  try {
    MailApp.sendEmail({
      to: NOTIFY_TO,
      subject: RESORT_NAME + ' website is connected',
      cc: NOTIFY_CC,
      body: 'Setup finished. Reservation requests from the website will arrive at this address, '
          + 'copied to ' + NOTIFY_CC + ', and be logged in the Bookings tab of the master sheet.'
    });
    mailNote = 'Mail OK. A test message has gone to ' + NOTIFY_TO + '.';
  } catch (mailErr) {
    mailNote = 'MAIL NOT GRANTED (' + mailErr + ')';
  }

  var report = [
    'Tabs OK: ' + CALENDAR_TAB + ', ' + BOOKINGS_TAB + ', ' + DAYPASS_TAB + ', ' + CAMPING_TAB + ', ' + SHUTTLE_TAB + '.',
    bookNote,
    driveNote,
    mailNote,
    'Last step: Deploy > Manage deployments > edit > Version: New version > Deploy.'
  ].join('\n');

  Logger.log(report);
  return report;
}

/* ====================== GET: availability ====================== */

function doGet(e) {
  try {
    var sh = calSheet(SpreadsheetApp.openById(SHEET_ID));
    if (!sh) return json({ error: 'Calendar tab not found: ' + CALENDAR_TAB });

    var rows = sh.getDataRange().getValues();
    if (rows.length < 2) return json([]);           // header only, nothing booked

    var head = rows.shift().map(function (h) { return String(h).trim().toLowerCase(); });
    var iD = head.indexOf('date');
    var iR = head.indexOf('room') >= 0 ? head.indexOf('room') : head.indexOf('unit');
    var iS = head.indexOf('status');
    if (iD < 0 || iR < 0 || iS < 0) {
      return json({ error: 'Columns must be Date, Room, Status' });
    }

    var out = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var d = r[iD], room = String(r[iR] || '').trim(), st = String(r[iS] || '').trim();
      if (!room || !st) continue;

      // the sheet may hold a real Date or plain text, normalise both
      var iso;
      if (Object.prototype.toString.call(d) === '[object Date]') {
        iso = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
      } else {
        iso = String(d || '').trim();
        if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) {
          var p = new Date(iso);
          if (isNaN(p.getTime())) continue;         // skip junk rows
          iso = Utilities.formatDate(p, TZ, 'yyyy-MM-dd');
        }
      }
      out.push({ Date: iso, Room: room, Status: st.toLowerCase() });
    }
    return json(out);

  } catch (err) {
    return json({ error: String(err) });
  }
}

/* ====================== POST: a reservation request ====================== */

function ticketTab(ss, name) {
  // Only `setup` is meant to be run by hand. Everything else is called by the
  // web app with arguments, so say so rather than throwing a TypeError.
  if (!ss || !name) return mustNotRunByHand('ticketTab');
  var tk = ss.getSheetByName(name);
  if (!tk) {
    tk = ss.insertSheet(name);
    tk.getRange(1, 1, 1, 13).setValues([[
      'Issued', 'Ticket', 'Date', 'Guests', 'What they bought', 'Total',
      'Paid by', 'Their reference', 'Deposit slip', 'Name', 'Mobile', 'Email', 'Notes'
    ]]).setFontWeight('bold');
    tk.setFrozenRows(1);
    tk.setColumnWidth(1, 150);
    tk.setColumnWidth(2, 190);
    tk.setColumnWidth(5, 330);
    tk.setColumnWidth(13, 280);
  }
  return tk;
}

/* A Bookings tab made before deposit slips existed has twelve columns and no
   slip column, so a new thirteen-value row would land one place out and put a
   Drive link under the Name heading. Widen it in place instead, which keeps
   every booking already in there exactly where it is. */
/* Writes Hold rows into the calendar tab for every night of a request.

   Careful on purpose:
   - the check-out night is NOT held, because the guest leaves that morning
   - a night already marked Booked is left exactly as it is, so an automatic
     hold can never quietly downgrade a real booking
   - a night already held is not written twice
   - the room name must match one you actually have, or nothing is written */
function holdNights(ss, b) {
  try {
    /* Use the sheet's own room name and the plain YYYY-MM-DD dates the website
       sends alongside the pretty ones. Falling back to the display values only
       if the plain ones are missing. */
    var rooms = [];
    if (b.roomSheets && b.roomSheets.length) rooms = b.roomSheets;
    else if (b.roomSheet) rooms = [b.roomSheet];
    else if (b.room) rooms = String(b.room).split(',');
    rooms = rooms.map(function (r) { return String(r).trim(); })
                 .filter(function (r) { return r; });

    var inStr  = b.checkInISO  || b.checkIn;
    var outStr = b.checkOutISO || b.checkOut;
    if (!b || !rooms.length || !inStr) return '';
    var cal = calSheet(ss);
    if (!cal) return '';

    var start = new Date(inStr);
    if (isNaN(start.getTime())) return '';
    var end = outStr ? new Date(outStr) : null;
    if (end && isNaN(end.getTime())) end = null;
    // no check-out, or a bad one, means hold the single night
    if (!end || end <= start) end = new Date(start.getTime() + 86400000);

    // what is already in the calendar, so nothing is written twice
    var have = {};
    var rows = cal.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var d = rows[i][0], rm = String(rows[i][1] || '').trim().toLowerCase(),
          st = String(rows[i][2] || '').trim().toLowerCase();
      if (!rm) continue;
      var iso = Object.prototype.toString.call(d) === '[object Date]'
        ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : String(d || '').trim();
      have[iso + '|' + rm] = st;
    }

    var added = 0, skipped = 0;
    for (var r = 0; r < rooms.length; r++) {
      for (var t = new Date(start); t < end; t = new Date(t.getTime() + 86400000)) {
        var iso2 = Utilities.formatDate(t, TZ, 'yyyy-MM-dd');
        var existing = have[iso2 + '|' + rooms[r].toLowerCase()];
        if (existing === 'booked' || existing === 'hold') { skipped++; continue; }
        cal.appendRow([iso2, rooms[r], 'Hold']);
        added++;
      }
    }
    if (!added && !skipped) return '';
    return added + ' room-night(s) put on Hold across ' + rooms.length + ' room(s)'
         + (skipped ? ', ' + skipped + ' already blocked' : '') + '.';
  } catch (err) {
    // a failed hold must never lose the booking itself
    return 'could not place the hold: ' + err;
  }
}

function upgradeBookings(bk) {
  if (!bk) return 'Bookings tab missing.';
  var lastCol = Math.max(1, bk.getLastColumn());
  var head = bk.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  if (head.indexOf('Deposit slip') > -1) return 'Bookings already has the Deposit slip column.';
  if (head[8] !== 'Name') {
    return 'Bookings columns are not the layout I expected, left alone. '
         + 'Expected Name in column I, found "' + (head[8] || '') + '".';
  }
  bk.insertColumnBefore(9);
  bk.getRange(1, 9).setValue('Deposit slip').setFontWeight('bold');
  bk.setColumnWidth(9, 220);
  return 'Bookings upgraded: a Deposit slip column was inserted before Name. '
       + 'Your existing rows are untouched.';
}

/* A Bookings tab made before "Confirm booking by reference" existed has 13
   columns and no ISO dates or sheet room names. Add the three columns it
   needs, at the end, so every existing row is left exactly where it is. */
function upgradeBookingsIso(bk) {
  if (!bk) return 'Bookings tab missing.';
  var lastCol = Math.max(1, bk.getLastColumn());
  var head = bk.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  if (head.indexOf('Room sheets') > -1) return 'Bookings already has the ISO columns.';
  var start = lastCol + 1;
  bk.getRange(1, start, 1, 3).setValues([['Check-in (ISO)', 'Check-out (ISO)', 'Room sheets']]).setFontWeight('bold');
  return 'Bookings upgraded: Check-in (ISO), Check-out (ISO), Room sheets added at the end.';
}

/* ====================== confirm a booking (Sheet menu) ====================== */

/* Adds "Valle Justina" to the Sheet's own menu bar. Runs automatically every
   time the sheet is opened; nothing to run by hand. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Casa Fria')
    .addItem('Confirm booking by reference…', 'confirmBookingByRef')
    .addSeparator()
    .addItem('Run setup', 'setup')
    .addToUi();
}

/* Staff-facing helper: type a reservation's reference number, and every
   night that request put on Hold across every room it covers is flipped to
   Booked in one go. Nothing is changed unless the reference is found and at
   least one matching Hold row exists.

   This never touches Day Pass, Camping, or Shuttle tickets — those are
   already paid, so there is nothing to "confirm" about them. It only
   applies to room reservations, which are logged as a request first and
   held pending a deposit. */
function confirmBookingByRef() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Confirm booking', 'Reference number (e.g. CF-20260803-AB12):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var ref = String(resp.getResponseText() || '').trim();
  if (!ref) { ui.alert('No reference entered.'); return; }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var bk = ss.getSheetByName(BOOKINGS_TAB);
  var cal = calSheet(ss);
  if (!bk || !cal) { ui.alert('Could not find the Bookings or ' + CALENDAR_TAB + ' tab.'); return; }

  var head = bk.getRange(1, 1, 1, bk.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  var iRef = head.indexOf('Reference'), iInIso = head.indexOf('Check-in (ISO)'),
      iOutIso = head.indexOf('Check-out (ISO)'), iRooms = head.indexOf('Room sheets'),
      iRoomDisp = head.indexOf('Room');
  if (iRef < 0) { ui.alert('The Bookings tab is missing a Reference column.'); return; }

  var rows = bk.getDataRange().getValues();
  var found = null;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][iRef] || '').trim().toLowerCase() === ref.toLowerCase()) { found = rows[i]; break; }
  }
  if (!found) { ui.alert('No booking found with reference "' + ref + '".'); return; }

  var inStr  = iInIso  >= 0 ? String(found[iInIso]  || '').trim() : '';
  var outStr = iOutIso >= 0 ? String(found[iOutIso] || '').trim() : '';
  var roomsStr = iRooms >= 0 ? String(found[iRooms] || '').trim() : '';
  var rooms = roomsStr ? roomsStr.split(',').map(function (r) { return r.trim(); }).filter(function (r) { return r; }) : [];

  if (!rooms.length && iRoomDisp >= 0) {
    // an older row made before "Room sheets" existed: fall back to the
    // display name, which matches the sheet name for every room except the
    // Balcony, so say so rather than silently missing it
    var disp = String(found[iRoomDisp] || '');
    rooms = disp.split(',').map(function (r) { return r.trim().replace(/\s+Room$/i, ''); }).filter(function (r) { return r; });
  }

  if (!rooms.length || !inStr) {
    ui.alert('Found reference "' + ref + '" but it has no room/date detail saved '
      + '(likely an older booking, made before this tool existed). '
      + 'Please change the Status cells for those nights by hand in ' + CALENDAR_TAB + '.');
    return;
  }

  var start = new Date(inStr);
  var end = outStr ? new Date(outStr) : null;
  if (isNaN(start.getTime())) { ui.alert('Could not read the check-in date for this booking.'); return; }
  if (!end || isNaN(end.getTime()) || end <= start) end = new Date(start.getTime() + 86400000);

  var calRows = cal.getDataRange().getValues();
  var updated = 0, notFound = 0;
  for (var r = 0; r < rooms.length; r++) {
    for (var t = new Date(start); t < end; t = new Date(t.getTime() + 86400000)) {
      var iso = Utilities.formatDate(t, TZ, 'yyyy-MM-dd');
      var hit = -1;
      for (var k = 1; k < calRows.length; k++) {
        var d = calRows[k][0], rm = String(calRows[k][1] || '').trim();
        var dIso = Object.prototype.toString.call(d) === '[object Date]'
          ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : String(d || '').trim();
        if (dIso === iso && rm.toLowerCase() === rooms[r].toLowerCase()) { hit = k; break; }
      }
      if (hit < 0) { notFound++; continue; }
      cal.getRange(hit + 1, 3).setValue('Booked');
      calRows[hit][2] = 'Booked';   // keep the in-memory copy in step, in case of duplicate rows
      updated++;
    }
  }

  var msg = 'Reference ' + ref + ': marked ' + updated + ' room-night(s) as Booked across ' + rooms.length + ' room(s).';
  if (notFound) msg += ' ' + notFound + ' night(s) had no matching row in ' + CALENDAR_TAB + ' and were left alone.';
  ui.alert(msg);
}

function mustNotRunByHand(name) {
  var msg = name + ' is called by the website, not run by hand. '
          + 'The only function you ever need to run yourself is `setup`.';
  Logger.log(msg);
  return msg;
}

/* Turns the data URL the browser sent into a file Google can store. */
function slipBlob(slip, ref) {
  if (!slip || !slip.dataUrl) return null;
  var parts = String(slip.dataUrl).split(',');
  if (parts.length < 2) return null;
  var mime = (parts[0].match(/data:([^;]+)/) || [null, 'image/jpeg'])[1];
  var ext  = mime.indexOf('png') > -1 ? '.png' : '.jpg';
  return Utilities.newBlob(Utilities.base64Decode(parts[1]), mime, (ref || 'slip') + ext);
}

function slipFolder() {
  if (SLIP_FOLDER_ID) {
    try { return DriveApp.getFolderById(SLIP_FOLDER_ID); }
    catch (e) { /* id wrong or no access, fall through to the named folder */ }
  }
  var it = DriveApp.getFoldersByName(SLIP_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(SLIP_FOLDER);
}

/* Saves the deposit slip and writes where it went into the row just added.
   Never throws. A picture must never cost anyone a booking.

   First choice is Drive, because a link opens in one click and can be
   forwarded. If Drive was not authorised, the picture is pinned into the sheet
   instead, so the slip still survives and you can see it without having to fix
   anything first. */
function attachSlip(sheet, rowIndex, col, slip, ref) {
  var blob;
  try { blob = slipBlob(slip, ref); } catch (e) { blob = null; }
  if (!blob) return '';

  try {
    var url = slipFolder().createFile(blob).getUrl();
    sheet.getRange(rowIndex, col).setValue(url);
    return url;
  } catch (driveErr) {
    try {
      sheet.insertImage(blob, col, rowIndex);
      sheet.getRange(rowIndex, col).setValue('picture pinned to this row');
      sheet.setRowHeight(rowIndex, 120);
      return 'pinned into the ' + sheet.getName() + ' tab, row ' + rowIndex;
    } catch (sheetErr) {
      sheet.getRange(rowIndex, col).setValue('could not save: ' + driveErr);
      return '';
    }
  }
}

function doPost(e) {
  try {
    /* Run by hand from the editor there is no request, and an earlier version
       of this happily wrote a row of blanks into Bookings. A post with nothing
       in it is not a booking, so refuse it. */
    if (!e || !e.postData || !e.postData.contents) return json({ ok: false, error: mustNotRunByHand('doPost') });

    var b = null;
    try { b = JSON.parse(e.postData.contents); } catch (x) { b = null; }
    if (!b || typeof b !== 'object') return json({ ok: false, error: 'could not read that request' });
    if (!b.ref && !b.name && !b.room && !b.kind) return json({ ok: false, error: 'empty request ignored' });

    var now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');

    /* Day pass, camping, and shuttle tickets go in their own tab. They are
       already paid, so they are a record rather than a request, and the
       email says so. */
    if (b.kind === 'daypass' || b.kind === 'camping' || b.kind === 'shuttle') {
      var ss2 = SpreadsheetApp.openById(SHEET_ID);
      var tab = b.kind === 'camping' ? CAMPING_TAB : (b.kind === 'shuttle' ? SHUTTLE_TAB : DAYPASS_TAB);
      var tk = ticketTab(ss2, tab);
      tk.appendRow([
        now, b.ref || '', b.date || '', b.guests || '', b.lines || '',
        b.total || '', b.pay || '', b.payRef || '', '',
        b.name || '', b.mobile || '', b.email || '', b.notes || ''
      ]);
      b._slipUrl = attachSlip(tk, tk.getLastRow(), 9, b.slip, b.ref);
      notifyTicket(b);
      return json({ ok: true, ref: b.ref || '' });
    }
    /* The three columns at the end are not shown to staff in the usual course
       of things, they exist so "Confirm booking by reference" (the sheet's
       custom menu) can find the exact calendar rows a booking put on Hold,
       without having to guess a display name back into a sheet name. */
    var row = [
      now,
      b.ref     || '',
      b.room    || '',
      b.checkIn || '',
      b.checkOut|| '',
      b.nights  || '',
      b.guests  || '',
      b.total   || '',
      '',                       // deposit slip, filled in after the row exists
      b.name    || '',
      b.mobile  || '',
      b.email   || '',
      b.notes   || '',
      b.checkInISO  || '',
      b.checkOutISO || '',
      (b.roomSheets && b.roomSheets.length) ? b.roomSheets.join(', ') : ''
    ];

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var bk = ss.getSheetByName(BOOKINGS_TAB);
    if (!bk) {
      bk = ss.insertSheet(BOOKINGS_TAB);
      bk.getRange(1, 1, 1, 16).setValues([[
        'Received', 'Reference', 'Room', 'Check in', 'Check out', 'Nights',
        'Guests', 'Total', 'Deposit slip', 'Name', 'Mobile', 'Email', 'Notes',
        'Check-in (ISO)', 'Check-out (ISO)', 'Room sheets'
      ]]).setFontWeight('bold');
      bk.setFrozenRows(1);
    }
    upgradeBookingsIso(bk);
    bk.appendRow(row);
    b._slipUrl = attachSlip(bk, bk.getLastRow(), 9, b.slip, b.ref);

    /* Put the nights on hold straight away. A request is not a booking, but it
       must stop the same room being offered to somebody else while you check
       whether the deposit has landed. */
    b._heldNote = holdNights(ss, b);

    notify(b);
    return json({ ok: true, ref: b.ref || '' });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/* ====================== the notification email ====================== */

function notify(b) {
  if (!b) return mustNotRunByHand('notify');
  var esc = function (s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  };
  var line = function (k, v) {
    if (!v && v !== 0) return '';
    return '<tr>'
      + '<td style="padding:7px 16px 7px 0;color:#7A675C;font:13px Arial,sans-serif;'
      + 'white-space:nowrap;vertical-align:top">' + esc(k) + '</td>'
      + '<td style="padding:7px 0;color:#1a1a1a;font:15px Arial,sans-serif">'
      + esc(v) + '</td></tr>';
  };

  var subject = 'Reservation request ' + (b.ref || '') + ' - ' + (b.name || 'website');

  var html =
      '<div style="font:15px Arial,sans-serif;color:#1a1a1a;max-width:560px">'
    +   '<p style="font-size:12px;letter-spacing:.14em;text-transform:uppercase;color:#8A5A18;margin:0 0 6px">'
    +     RESORT_NAME + '</p>'
    +   '<h2 style="margin:0 0 4px;font-weight:normal;font-size:22px">New reservation request</h2>'
    +   '<p style="color:#7A675C;margin:0 0 20px">Sent from the website'
    +     (b.ref ? ', reference <b style="color:#1a1a1a">' + esc(b.ref) + '</b>' : '') + '</p>'
    +   '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;'
    +     'border-top:2px solid #C88A3C">'
    +     line('Room', b.room)
    +     line('Check in', b.checkIn)
    +     line('Check out', b.checkOut)
    +     line('Nights', b.nights)
    +     line('Guests', b.guests)
    +     line('Total', b.total)
    +     line('Name', b.name)
    +     line('Mobile', b.mobile)
    +     line('Email', b.email)
    +     line('Notes', b.notes)
    +     line('Deposit slip', b._slipUrl)
    +     line('Calendar', b._heldNote)
    +   '</table>'
    +   '<p style="color:#7A675C;font-size:13px;margin:22px 0 0;padding-top:14px;'
    +     'border-top:1px solid #e2ddd7">'
    +     'Logged in the Bookings tab, and the nights are on Hold in the calendar so nobody '
    +     'else can take them. Change the Status to Booked once the deposit lands, or delete '
    +     'those rows if the guest goes quiet.'
    +   '</p>'
    + '</div>';

  var plain = [
    RESORT_NAME + ' - new reservation request',
    '',
    'Reference: ' + (b.ref || ''),
    'Room:      ' + (b.room || ''),
    'Check in:  ' + (b.checkIn || ''),
    'Check out: ' + (b.checkOut || ''),
    'Nights:    ' + (b.nights || ''),
    'Guests:    ' + (b.guests || ''),
    'Total:     ' + (b.total || ''),
    '',
    'Name:   ' + (b.name || ''),
    'Mobile: ' + (b.mobile || ''),
    'Email:  ' + (b.email || ''),
    'Notes:  ' + (b.notes || ''),
    'Slip:   ' + (b._slipUrl || 'none attached'),
    'Calendar: ' + (b._heldNote || 'nothing held'),
    '',
    'The nights are on Hold. Change the Status to Booked once the deposit lands,',
    'or delete those rows if the guest goes quiet.'
  ].join('\n');

  var opts = { to: NOTIFY_TO, subject: subject, body: plain, htmlBody: html, name: RESORT_NAME };
  if (NOTIFY_CC) opts.cc = NOTIFY_CC;
  if (b.email && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(b.email)) opts.replyTo = b.email;

  MailApp.sendEmail(opts);
}

/* ====================== the ticket email ====================== */

function notifyTicket(b) {
  if (!b) return mustNotRunByHand('notifyTicket');
  var esc = function (s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  };
  var line = function (k, v) {
    if (!v && v !== 0) return '';
    return '<tr><td style="padding:7px 16px 7px 0;color:#7A675C;font:13px Arial,sans-serif;'
      + 'white-space:nowrap;vertical-align:top">' + esc(k) + '</td>'
      + '<td style="padding:7px 0;color:#1a1a1a;font:15px Arial,sans-serif">'
      + esc(v) + '</td></tr>';
  };

  var what = b.label || (b.kind === 'camping' ? 'Camping' : b.kind === 'shuttle' ? 'Shuttle' : 'Day Pass');
  var subject = what + ' ticket ' + (b.ref || '') + ' - ' + (b.name || 'website');

  var html =
      '<div style="font:15px Arial,sans-serif;color:#1a1a1a;max-width:560px">'
    +   '<p style="font-size:12px;letter-spacing:.14em;text-transform:uppercase;color:#8A5A18;margin:0 0 6px">'
    +     RESORT_NAME + '</p>'
    +   '<h2 style="margin:0 0 4px;font-weight:normal;font-size:22px">' + esc(what) + ' ticket issued</h2>'
    +   '<p style="color:#7A675C;margin:0 0 20px">Ticket <b style="color:#1a1a1a">' + esc(b.ref || '') + '</b>. '
    +     'The guest says they have already paid. Check their reference against the account before they come up.</p>'
    +   '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;'
    +     'border-top:2px solid #C88A3C">'
    +     line('Date', b.date)
    +     line('Guests', b.guests)
    +     line('What they bought', b.lines)
    +     line('Total', b.total)
    +     line('Paid by', b.pay)
    +     line('Their reference', b.payRef)
    +     line('Name', b.name)
    +     line('Mobile', b.mobile)
    +     line('Email', b.email)
    +     line('Notes', b.notes)
    +     line('Deposit slip', b._slipUrl)
    +   '</table>'
    +   '<p style="color:#7A675C;font-size:13px;margin:22px 0 0;padding-top:14px;'
    +     'border-top:1px solid #e2ddd7">Logged in the Tickets tab of the master sheet.</p>'
    + '</div>';

  var plain = [
    RESORT_NAME + ' - ' + what + ' ticket issued',
    '',
    'Ticket:    ' + (b.ref || ''),
    'Date:      ' + (b.date || ''),
    'Guests:    ' + (b.guests || ''),
    'Bought:    ' + (b.lines || ''),
    'Total:     ' + (b.total || ''),
    'Paid by:   ' + (b.pay || ''),
    'Their ref: ' + (b.payRef || ''),
    '',
    'Name:   ' + (b.name || ''),
    'Mobile: ' + (b.mobile || ''),
    'Email:  ' + (b.email || ''),
    'Notes:  ' + (b.notes || ''),
    'Slip:   ' + (b._slipUrl || 'none attached'),
    '',
    'The guest says they have already paid. Check the reference before they come up.'
  ].join('\n');

  var opts = { to: NOTIFY_TO, subject: subject, body: plain, htmlBody: html, name: RESORT_NAME };
  if (NOTIFY_CC) opts.cc = NOTIFY_CC;
  if (b.email && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(b.email)) opts.replyTo = b.email;
  MailApp.sendEmail(opts);
}

/* ====================== helper ====================== */

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* Run this to check the GET side without leaving the editor. */
function testGet() {
  Logger.log(doGet({}).getContent());
}
