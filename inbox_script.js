var SHEET_ID = 'YOUR_SHEET_ID';
var OPENROUTER_KEY = 'YOUR_OPENROUTER_KEY';
var AI_MODEL = 'openai/gpt-oss-120b:free'; // change to your preferred model

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
    'https://openrouter.ai/api/v1/chat/completions',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + OPENROUTER_KEY },
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
      + 'Example: 1. lab | PI moved reading group to Wed 4pm, asking for confirmation.\n\n'
      + 'Categories:\n'
      + '- lab: lab PI/members, research\n'
      + '- dept: BME dept, career, seminars, thesis defenses, jobs/internships\n'
      + '- orgs: student organizations, clubs\n'
      + '- campus: important campus announcements that need action (registrar, financial aid, safety, housing, IT)\n'
      + '- lowpri: mass emails, surveys, voting, newsletters, promos, events student likely wont attend, MCTP, department blasts, storage, Lyft, recruitment\n\n'
      + 'Summary rules:\n'
      + '- 1 sentence, max 20 words\n'
      + '- Focus on what matters: action needed, key info, deadline\n'
      + '- Skip greetings and filler\n'
      + '- When in doubt classify as lowpri\n\n'
      + 'Emails:\n' + lines;
    try {
      var text = aiCall(prompt);
      var valid = ['lab', 'dept', 'orgs', 'campus', 'lowpri'];
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
