// KidsUnited Woburn — free trial bookings.
// Deploy as Web app: Execute as "Me", Who has access "Anyone".
// Bookings land in the "Bookings" tab of the spreadsheet SHEET_ID.

const CAPACITY = 3;                          // trial spots per class session
const NOTIFY = 'info@kidsunitedwoburn.com';  // gets an email for every booking
const SHEET = 'Bookings';
const SHEET_ID = '176WSCsqTgfSWKCRnSX0gzhvjbN4aO-R5pmDwIYI_9FU';  // KidsUnited Woburn - Trial Bookings
const HEADERS = ['Booked at', 'Slot', 'Class date', 'Class time', 'Program', 'Child', 'Child DOB',
  'Parent', 'Phone', 'Email', 'Played before', 'Heard about us', 'Notes', 'Text consent', 'Synced to OctoSphere'];

function sheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function counts_() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const c = {};
  rows.forEach(r => { if (r[1]) c[r[1]] = (c[r[1]] || 0) + 1; });
  return c;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ ok: true, capacity: CAPACITY, counts: counts_() });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request' }); }
  if (b.website) return json_({ ok: true });  // honeypot: bots fill hidden field
  const need = ['slot', 'date', 'time', 'program', 'child', 'parent', 'phone', 'email'];
  for (const k of need) if (!String(b[k] || '').trim()) return json_({ ok: false, error: 'Missing ' + k });
  if (!b.consent) return json_({ ok: false, error: 'Consent required' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if ((counts_()[b.slot] || 0) >= CAPACITY) return json_({ ok: false, full: true, error: 'Sorry, that class just filled up. Please pick another time.' });
    sheet_().appendRow([new Date(), b.slot, b.date, b.time, b.program, b.child, b.dob || '', b.parent,
      "'" + b.phone, b.email, b.played || '', b.heard || '', b.notes || '', 'Yes', '']);
  } finally {
    lock.releaseLock();
  }

  const when = b.dateLabel + ' at ' + b.timeLabel;
  MailApp.sendEmail({
    to: NOTIFY,
    subject: 'New free trial: ' + b.child + ' — ' + b.program + ', ' + when,
    body: [
      'Add to OctoSphere within 2 hours (Source: Booking page, status Booked Intro).', '',
      'Class: ' + b.program + ' — ' + when,
      'Child: ' + b.child + (b.dob ? ' (DOB ' + b.dob + ')' : ''),
      'Parent: ' + b.parent, 'Phone: ' + b.phone, 'Email: ' + b.email,
      'Played before: ' + (b.played || '-'), 'Heard about us: ' + (b.heard || '-'), 'Notes: ' + (b.notes || '-')
    ].join('\n')
  });
  MailApp.sendEmail({
    to: b.email,
    name: 'KidsUnited Woburn',
    replyTo: NOTIFY,
    subject: "You're booked! " + b.child + "'s free class at KidsUnited Woburn",
    body: [
      'Hi ' + b.parent.split(' ')[0] + ',', '',
      "Thanks for booking! Here are the details for " + b.child + "'s free trial class:", '',
      b.program + '\n' + when + ' (45 minutes)',
      'KidsUnited Woburn, 240 W Cummings Park, Woburn, MA 01801',
      'Free parking · off I-95 Exit 54', '',
      'Please arrive 10 minutes early in comfy clothes and sneakers, and bring a water bottle.',
      'Parent & me classes (KUB, Little Champs): a grown-up joins on the field.', '',
      'Pre-sale families get $135/month (regularly $150) plus a FREE Player Pack ($120 value) when they join by Dec 3.', '',
      'Need to change your time? Just reply to this email or call/text (781) 404-3490.', '',
      'See you on the field!', 'The KidsUnited Woburn team'
    ].join('\n')
  });
  return json_({ ok: true });
}
