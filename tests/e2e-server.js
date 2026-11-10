// Serves checkin-site/ with the API pointed at /api, backed by Code.gs running against in-memory sheets.
const http = require('http'), fs = require('fs'), path = require('path'), vm = require('vm');
const [codePath, siteDir, port] = process.argv.slice(2);
const src = fs.readFileSync(codePath, 'utf8');

function makeSheet(name) {
  const data = [];
  const sh = {
    data, maxRows: 1000,
    getName: () => name,
    getLastRow: () => data.length,
    getLastColumn: () => Math.max(0, ...data.map(r => r.length)),
    getMaxRows: () => sh.maxRows,
    insertRowsAfter: (a, n) => { sh.maxRows += n; },
    appendRow: row => { data.push(row.slice()); },
    deleteRow: r => { data.splice(r - 1, 1); },
    getRange: (r, c, nr, nc) => { if (typeof r === 'string') { const m = r.match(/^([A-Z])([0-9]+)(?::([A-Z])([0-9]+))?$/); const c1 = m[1].charCodeAt(0) - 64, r1 = +m[2], c2 = m[3] ? m[3].charCodeAt(0) - 64 : c1, r2 = m[4] ? +m[4] : r1; r = r1; c = c1; nr = r2 - r1 + 1; nc = c2 - c1 + 1; } return ({
      getValue: () => { const v = (data[r - 1] || [])[c - 1]; return v === undefined ? '' : v; },
      setNumberFormat() { return this; }, setFontWeight() { return this; },
      getValues: () => { const out = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) { const v = (data[r - 1 + i] || [])[c - 1 + j]; row.push(v === undefined ? '' : v); } out.push(row); } return out; },
      setValues: vals => { vals.forEach((row, i) => { data[r - 1 + i] = data[r - 1 + i] || []; row.forEach((v, j) => { data[r - 1 + i][c - 1 + j] = v; }); }); },
      clearContent: () => { for (let i = 0; i < nr; i++) if (data[r - 1 + i]) for (let j = 0; j < nc; j++) data[r - 1 + i][c - 1 + j] = ''; },
    }); },
  };
  return sh;
}
const sheets = { 'Log': makeSheet('Log'), 'Current Status': makeSheet('Current Status'), 'Settings': makeSheet('Settings') };
sheets.Log.appendRow(['Timestamp']);
sheets['Current Status'].appendRow(['Job']);
[['Technicians', '', 'Emails', '', 'Setting', 'Value'], ['Jesse', '', 'director@example.com', '', 'in', 21], ['Camron', '', '', '', 'ready', 3], ['Logan'], [], [], [], [], [], ['', '', '', '', 'PIN', '123456']].forEach(r => sheets.Settings.data.push(r));
const ss = { getSheetByName: n => sheets[n], getUrl: () => 'https://sheet' };
const mails = [];
const props = { ssId: 'ss', formId: 'f', statusId: '1', clientTextId: '2', jobListId: '3', otherNameId: '4', piecesId: '5', descId: '6', techId: '7', notesId: '8' };
const ctx = {
  console, JSON, Object, Date, Math, String, Number, Array,
  SpreadsheetApp: { openById: () => ss, flush: () => {} },
  FormApp: { openById: () => { throw new Error('FormApp should not be used by the web path'); } },
  MailApp: { sendEmail: m => { mails.push(m); console.log('EMAIL:', m.subject); } },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  PropertiesService: { getScriptProperties: () => ({ getProperties: () => props }) },
  Utilities: { formatDate: (d, tz, f) => f === 'M/d' ? (d.getMonth() + 1) + '/' + d.getDate() : d.toISOString(), getUuid: () => Math.random().toString(36).slice(2) },
  CacheService: (() => { const m = {}; return { getScriptCache: () => ({ get: k => m[k] || null, put: (k, v) => { m[k] = v; } }) }; })(),
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ s, setMimeType() { return this; }, getContent() { return this.s; } }) },
};
vm.createContext(ctx);
vm.runInContext(src, ctx);

const DAY = 86400000, now = Date.now();
const seed = (rid, daysAgo, a) => ctx.logEntry_(props, rid, new Date(now - daysAgo * DAY), a, false);
seed('s1', 10, { 1: 'Checked in at shop', 2: 'Smith Gym', 5: '3', 6: 'Seat pad; Back pad; Head pad', 7: 'Camron' });
seed('s2', 5, { 1: 'Ready for reinstallation', 3: 'Smith Gym · ' + ctx.Utilities.formatDate(new Date(now - 10 * DAY), '', 'M/d'), 5: '3', 6: 'Seat pad; Back pad; Head pad', 7: 'Logan', 8: 'Back pad has a small tear' });
seed('s3', 25, { 1: 'Checked in at shop', 2: 'Jones Fitness', 5: '2', 6: 'Bench pad; Roller pad', 7: 'Jesse' });
seed('s4', 2, { 1: 'Checked in at shop', 2: 'Lee Studio', 5: '1', 6: 'Chest pad', 7: 'Camron' });
seed('s5', 0, { 1: 'Ready for reinstallation', 3: 'Lee Studio · ' + ctx.Utilities.formatDate(new Date(now - 2 * DAY), '', 'M/d'), 5: '1', 6: 'Chest pad', 7: 'Logan' });
seed('s6', 3, { 1: 'Checked in at shop', 2: 'Park HOA Fitness Center', 5: '4', 6: 'Seat pad; Back pad; Leg pad; Arm pad', 7: 'Logan' });
mails.length = 0;
http.createServer((req, res) => {
  const url = new URL(req.url, 'http://x');
  if (url.pathname === '/api') {
    if (req.method === 'GET') {
      const out = ctx.doGet({ parameter: { action: url.searchParams.get('action') } });
      res.writeHead(200, { 'Content-Type': 'application/json' }); return res.end(out.getContent());
    }
    let body = '';
    req.on('data', c => body += c);
    return req.on('end', () => {
      setTimeout(() => {
        const out = ctx.doPost({ postData: { contents: body } });
        console.log('POST', body, '->', out.getContent().slice(0, 120));
        res.writeHead(200, { 'Content-Type': 'application/json' }); res.end(out.getContent());
      }, 800); // simulate server time
    });
  }
  if (url.pathname === '/state') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ log: sheets.Log.data.slice(1).map(r => r.slice(1, 8)), status: sheets['Current Status'].data.slice(1).map(r => r.slice(0, 3)), mails: mails.map(m => m.subject) }));
  }
  const file = url.pathname === '/' ? 'index.html' : url.pathname.slice(1);
  fs.readFile(path.join(siteDir, file), (e, buf) => {
    if (e) { res.writeHead(404); return res.end(); }
    if (file.endsWith('.html')) buf = Buffer.from(buf.toString().replace(/var API = '[^']+';/, "var API = '/api';"));
    res.writeHead(200, { 'Content-Type': file.endsWith('.png') ? 'image/png' : 'text/html; charset=utf-8' });
    res.end(buf);
  });
}).listen(Number(port), () => console.log('listening', port));
