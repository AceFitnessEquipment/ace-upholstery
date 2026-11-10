// Minimal mocks of the Google services used by processResponse_ and friends.
const fs = require('fs');
const vm = require('vm');
const src = fs.readFileSync(process.argv[2], 'utf8');

function makeSheet(name) {
  const data = []; // rows of arrays (row 1 = index 0)
  const sh = {
    name, data, maxRows: 1000,
    getName: () => name,
    getLastRow: () => data.length,
    getLastColumn: () => Math.max(0, ...data.map(r => r.length)),
    getMaxRows: () => sh.maxRows,
    insertRowsAfter: (a, n) => { sh.maxRows += n; },
    appendRow: row => { data.push(row.slice()); },
    deleteRow: r => { data.splice(r - 1, 1); },
    getRange: (r, c, nr, nc) => ({
      getValues: () => { const out = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) { const v = (data[r - 1 + i] || [])[c - 1 + j]; row.push(v === undefined ? '' : v); } out.push(row); } return out; },
      setValues: vals => { vals.forEach((row, i) => { data[r - 1 + i] = data[r - 1 + i] || []; row.forEach((v, j) => { data[r - 1 + i][c - 1 + j] = v; }); }); },
      clearContent: () => { for (let i = 0; i < nr; i++) if (data[r - 1 + i]) for (let j = 0; j < nc; j++) data[r - 1 + i][c - 1 + j] = ''; },
    }),
  };
  return sh;
}
const sheets = { 'Log': makeSheet('Log'), 'Current Status': makeSheet('Current Status'), 'Settings': makeSheet('Settings') };
sheets['Log'].appendRow(['Timestamp']);
sheets['Current Status'].appendRow(['Job']);
sheets['Settings'].data.push(['Technicians', '', 'Email']);
sheets['Settings'].data.push(['Jesse', '', 'director@example.com']);
sheets['Settings'].data.push(['Maria', '', '']);
const ss = { getSheetByName: n => sheets[n], getUrl: () => 'https://sheet' };

const choices = {};
const form = { getItemById: id => ({ asListItem: () => ({ setChoiceValues: v => { choices[id] = v; } }) }) };
const mails = [];
const props = { ssId: 'ss', formId: 'f', statusId: '1', clientTextId: '2', jobListId: '3', otherNameId: '4', piecesId: '5', descId: '6', techId: '7', notesId: '8' };

const ctx = {
  console,
  SpreadsheetApp: { openById: () => ss, flush: () => {} },
  FormApp: { openById: () => form },
  MailApp: { sendEmail: m => mails.push(m) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  PropertiesService: { getScriptProperties: () => ({ getProperties: () => props }) },
  Utilities: { formatDate: (d, tz, f) => f === 'M/d' ? (d.getMonth() + 1) + '/' + d.getDate() : d.toISOString() },
};
vm.createContext(ctx);
vm.runInContext(src, ctx);

let n = 0;
function submit(t, answers) {
  const id = 'r' + (++n);
  const resp = {
    getId: () => id,
    getTimestamp: () => new Date(t),
    getItemResponses: () => Object.keys(answers).map(k => ({ getItem: () => ({ getId: () => Number(k) }), getResponse: () => answers[k] })),
  };
  return { resp, out: ctx.processResponse_(resp) };
}
const assert = (c, m) => { if (!c) { console.error('FAIL:', m); process.exitCode = 1; } else console.log('ok:', m); };

let r = submit('2026-11-01T10:00:00', { 1: 'Checked in at shop', 2: '  Smith  Gym ', 5: '4', 6: 'Bench pad; Seat pad; Back pad; Roller pad', 7: 'Jesse' });
const job = r.out.job;
assert(job === 'Smith Gym · 11/1', 'job label ' + job);
assert(JSON.stringify(choices[3]) === JSON.stringify(['Smith Gym · 11/1', 'Not in list (type the name below)']), 'dropdown has open job');
assert(JSON.stringify(choices[7]) === '["Jesse","Maria"]', 'tech list from Settings');

// duplicate processing of same response is skipped
const dup = ctx.processResponse_(r.resp);
assert(dup.duplicate && dup.job === job, 'dedupe by response id');
assert(sheets.Log.data.length === 2, 'only one log row after dup');

submit('2026-11-02T09:00:00', { 1: 'Checked in at shop', 2: 'Jones Fitness', 5: '2', 7: 'Jesse' });
// Not-in-list resolves to existing open job by name
r = submit('2026-11-20T15:00:00', { 1: 'Ready for reinstallation', 3: 'Not in list (type the name below)', 4: 'smith gym', 5: '3', 6: 'bench pad; Seat pad; Back pad', 7: 'Maria', 8: 'one pad torn' });
assert(r.out.job === job, 'not-in-list matched to open job');
const backMails = mails.filter(m => m.subject.indexOf('Ready for reinstallation') === 0); const inMails = mails.filter(m => m.subject.indexOf('Checked in at shop') === 0); assert(backMails.length === 1 && inMails.length === 2, 'check-in + back emails sent (' + inMails.length + ' in, ' + backMails.length + ' back)'); console.log('  in subject:', inMails[0].subject); assert(/Item 4: Roller pad/.test(inMails[0].body), 'check-in email lists items'); mails.splice(0, mails.length, backMails[0]);
console.log('  subject:', mails[0].subject);
assert(/Not back yet: Roller pad/.test(mails[0].body), 'missing piece named'); assert(/Item 3: Back pad/.test(mails[0].body), 'items listed');
assert(/Out for upholstery: 19 days/.test(mails[0].body), 'days at upholsterer');

const cs = sheets['Current Status'].data;
assert(cs[1][2] === 'Ready for reinstallation' && cs[2][2] === 'Checked in at shop', 'status sort: back first');
assert(String(cs[1][4]).startsWith('=IF(D2'), 'days formula');
console.log('  alert formula:', cs[1][10]);

submit('2026-11-21T08:00:00', { 1: 'Scheduled', 3: job, 5: '4', 7: 'Jesse' });
assert(JSON.stringify(choices[3]) === JSON.stringify(['Jones Fitness · 11/2', 'Not in list (type the name below)']), 'scheduled job closes and drops from dropdown');
// same client same day -> (2)
r = submit('2026-11-02T16:00:00', { 1: 'Checked in at shop', 2: 'Jones Fitness', 5: '1', 7: 'Jesse' });
assert(r.out.job === 'Jones Fitness · 11/2 (2)', 'same-day duplicate label ' + r.out.job);
console.log('\nCurrent Status:'); cs.slice(1).forEach(row => console.log('  ', row.slice(0, 4).map(String).join(' | ')));
const pd = ctx.pageData_();
console.log('openJobs:', JSON.stringify(pd.openJobs));
assert(pd.openJobs.every(j => Array.isArray(j.pieces)), 'page data includes piece lists');
assert(JSON.stringify(pd.statuses) === '["Checked in at shop","Ready for reinstallation","Scheduled"]', 'three statuses');
assert(pd.finalStatus === 'Scheduled', 'page knows final status');
assert(/not back from Tucson/.test(cs[2][10]) && !/Sent to Tucson/.test(cs[2][10]), 'alert formula uses new steps');
