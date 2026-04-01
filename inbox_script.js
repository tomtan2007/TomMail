var SHEET_ID = 'YOUR_SHEET_ID';
var OPENAI_KEY = 'YOUR_OPENAI_KEY';
var AI_MODEL = 'gpt-4.1-nano'; // change to your preferred model

function syncInbox() {
  var threads = GmailApp.search('in:inbox newer_than:14d', 0, 200);

  // Load existing classifications from sheet (id -> {cat, sum})
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('inbox');
  if (!sheet) sheet = ss.insertSheet('inbox');
  var existing = {};
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][6]) {
      existing[data[i][0]] = { cat: data[i][6], sum: data[i][7] || '' };
    }
  }

  // Fetch all emails
  var emails = [];
  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    var m = msgs[msgs.length - 1];
    var body = m.getPlainBody();
    if (body) body = body.split('________')[0].trim();
    else body = '';
    emails.push([
      m.getId(), m.getSubject(), m.getFrom(),
      m.getDate().getTime(),
      body.substring(0, 800), m.isUnread()
    ]);
  }
  emails.sort(function(a, b) { return b[3] - a[3]; });

  // Split into new (needs AI) and already-classified
  var newEmails = [];
  var newIndices = [];
  for (var i = 0; i < emails.length; i++) {
    var id = emails[i][0];
    if (!existing[id]) {
      newEmails.push(emails[i]);
      newIndices.push(i);
    }
  }

  Logger.log('Total: ' + emails.length + ', New: ' + newEmails.length + ', Cached: ' + (emails.length - newEmails.length));

  // Only call AI for new emails
  var newResults = newEmails.length > 0 ? aiProcess(newEmails) : [];

  // Build final rows: use cached classification or new AI result
  var rows = [];
  var newIdx = 0;
  for (var i = 0; i < emails.length; i++) {
    var e = emails[i];
    var id = e[0];
    var cat, sum;
    if (existing[id]) {
      // Reuse existing classification, but update unread status
      cat = existing[id].cat;
      sum = existing[id].sum;
    } else {
      var r = newResults[newIdx] || {};
      cat = r.cat || 'campus';
      sum = r.sum || '';
      newIdx++;
    }
    rows.push(e.concat([cat, sum]));
  }

  // Write to sheet
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 8).setValues(
    [['id', 'subject', 'from', 'date', 'body', 'unread', 'category', 'summary']]
  );
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 8).setValues(rows);
  }
}

function aiCall(prompt) {
  var resp = UrlFetchApp.fetch(
    'https://api.openai.com/v1/chat/completions',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + OPENAI_KEY },
      payload: JSON.stringify({
        model: AI_MODEL,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    }
  );
  var raw = resp.getContentText();
  var j = JSON.parse(raw);
  if (!j.choices || !j.choices[0]) {
    Logger.log('AI response: ' + raw.substring(0, 500));
    throw new Error('No response from model');
  }
  return j.choices[0].message.content;
}

function aiProcess(emails) {
  var out = [];
  var BATCH = 10;
  for (var i = 0; i < emails.length; i += BATCH) {
    var batch = emails.slice(i, i + BATCH);
    if (i > 0) Utilities.sleep(4000);
    var lines = batch.map(function(e, j) {
      var b = (e[4] || '').substring(0, 200);
      return (j + 1) + '. FROM: ' + e[2] + ' | SUBJ: ' + e[1] + ' | BODY: ' + b;
    }).join('\n');
    var prompt = 'You classify and summarize emails for a UMich BME undergrad.\n\n'
      + 'For each email reply with exactly one line:\n'
      + 'NUMBER. CATEGORY | SUMMARY\n\n'
      + 'Example: 1. important | Dr. Douville sent anesthesia tech details, needs response by Friday.\n\n'
      + 'Categories (pick ONE):\n'
      + '- important: emails directed personally to the student that need attention — direct messages from professors, advisors, or contacts; registrar actions; financial aid; personal requests; deadlines requiring action. The key test: is this sent TO the student specifically, not to a mailing list?\n'
      + '- lab: lab PI/members, research group emails, lab meetings, reading groups\n'
      + '- dept: BME department blasts, seminar series, thesis defenses, career services, job postings\n'
      + '- orgs: student organizations, clubs, org newsletters\n'
      + '- campus: campus-wide announcements, safety alerts, IT notices, housing, dining\n'
      + '- promo: promotional emails, commercial services, companies, apps, subscriptions\n'
      + '- lowpri: surveys, voting, mass newsletters, events student likely wont attend, MCTP events, storage reminders, study recruitment, generic department blasts\n\n'
      + 'Key rules:\n'
      + '- If the email is from a person writing directly to the student (not a mailing list or automated), classify as important\n'
      + '- If the email is a mass send / mailing list / automated notification, it is NOT important — use the appropriate other category\n'
      + '- BME seminar series, lab equipment notices (like Rogel room updates), reading group schedules = lab or dept, NOT important\n'
      + '- When in doubt between important and another category, pick the other category\n\n'
      + 'Summary: 1 sentence, max 20 words. Focus on action needed or key info.\n\n'
      + 'Emails:\n' + lines;
    try {
      var text = aiCall(prompt);
      var valid = ['important', 'lab', 'dept', 'orgs', 'campus', 'promo', 'lowpri'];
      text.split('\n').forEach(function(line) {
        var m = line.match(/^(\d+)\.\s*(\w+)\s*\|\s*(.+)/);
        if (m) {
          var idx = parseInt(m[1]) - 1;
          var c = m[2].toLowerCase();
          if (valid.indexOf(c) < 0) c = 'campus';
          out[i + idx] = { cat: c, sum: m[3].trim().substring(0, 150) };
        }
      });
      for (var j = 0; j < batch.length; j++) {
        if (!out[i + j]) out[i + j] = { cat: 'campus', sum: '' };
      }
    } catch (err) {
      Logger.log('AI error: ' + err);
      for (var j = 0; j < batch.length; j++) {
        if (!out[i + j]) out[i + j] = { cat: 'campus', sum: '' };
      }
    }
  }
  return out;
}
