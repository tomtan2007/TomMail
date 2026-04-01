// Add this function to your existing Apps Script project.
// It fetches your Canvas iCal feed and writes assignments to an 'assignments' sheet tab.
// Set up a time trigger for syncAssignments (every 15 min).

var SHEET_ID = 'YOUR_SHEET_ID'; // reuse same sheet ID

function syncAssignments() {
  var ICAL_URL = 'https://umich.instructure.com/feeds/calendars/user_QWg4huzUeUvgRILDMViiMXV0kkeRaL0GbsueXqsd.ics';
  var resp = UrlFetchApp.fetch(ICAL_URL, { muteHttpExceptions: true });
  var text = resp.getContentText();
  var blocks = text.split('BEGIN:VEVENT').slice(1);
  var rows = [];
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  for (var i = 0; i < blocks.length; i++) {
    var block = blocks[i];
    var sm = block.match(/SUMMARY[^:]*:(.+)/);
    var dm = block.match(/DTSTART[^:]*:(\d{8})/);
    if (!sm || !dm) continue;
    var raw = sm[1].replace(/\r/g, '').trim();
    var title = raw.replace(/\[.*?\]/g, '').replace(/\+/g, ' ').trim();
    var ds = dm[1];
    var due = ds.slice(0, 4) + '-' + ds.slice(4, 6) + '-' + ds.slice(6, 8);
    var dueDate = new Date(due + 'T00:00:00');
    // Skip past assignments
    if (dueDate < today) continue;
    // Detect course
    var course = detectAssignmentCourse(raw);
    if (!course) continue;
    // Detect type
    var type = detectAssignmentType(title);
    // Skip async physics
    if (type === 'asynch' && course === 'PHYSICS 240') continue;
    rows.push([title, course, due, type]);
  }
  // Sort by due date
  rows.sort(function(a, b) { return a[2] < b[2] ? -1 : a[2] > b[2] ? 1 : 0; });

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('assignments');
  if (!sheet) sheet = ss.insertSheet('assignments');
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 4).setValues([['title', 'course', 'due', 'type']]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
}

function detectAssignmentCourse(s) {
  var courses = ['PHYSICS 240', 'PHYSICS 241', 'ASIAN 325', 'CHEM 215', 'CHEM 216', 'PSYCH 111'];
  var upper = s.toUpperCase();
  for (var i = 0; i < courses.length; i++) {
    if (upper.indexOf(courses[i]) >= 0) return courses[i];
  }
  return null;
}

function detectAssignmentType(title) {
  var s = title.toLowerCase();
  if (/exam|midterm|final/.test(s)) return 'exam';
  if (/checkpoint/.test(s)) return 'checkpoint';
  if (/essay|critical term|writing|paragraph/.test(s)) return 'writing';
  if (/asynch|async/.test(s)) return 'asynch';
  return 'reading';
}
