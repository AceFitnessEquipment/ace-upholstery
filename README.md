# ACE Fitness – Upholstery Tracker

ACE Fitness technicians remove worn upholstery at client sites, bring it to the shop, send it to an upholsterer in Tucson, and reinstall it when it comes back. This app logs every hand-off in one place and emails the Director of Operations at each step, so reinstalls get scheduled.

| What | Where |
|---|---|
| **Check-in page** (techs, opened from the QR sign) | https://acefitnessequipment.github.io/ace-upholstery/ |
| **Director dashboard** (PIN protected) | https://acefitnessequipment.github.io/ace-upholstery/admin.html |
| **Printable QR sign + tech instructions** | https://acefitnessequipment.github.io/ace-upholstery/printouts/shop-printouts.html |
| **Tracker spreadsheet + backend script** | A Google Sheet owned by the Director's Google account (not in this repo; ask the Director for access) |

For changing, deploying, troubleshooting or rebuilding, see **[docs/MAINTAINING.md](docs/MAINTAINING.md)**.

---

## How it works

### The process (3 steps)

| # | When | Who | Logs | Result |
|---|---|---|---|---|
| 1 | Upholstery comes back to the shop **from a client** | The tech who removed it | **Checked in at shop**: client name (as on the orange tape) plus each piece as its own item | Director gets an email |
| 2 | Upholstery comes back to the shop **from Tucson** | The tech who picked it up | **Ready for reinstallation**: picks the client, unchecks any piece that didn't come back | Director gets an email (it flags missing pieces) |
| 3 | Reinstall is on the calendar | Director | **Mark scheduled** on the dashboard (optional reinstall date) | Job closes |

The Director can also **Delete** a job from the dashboard, for example a mistake or a test. That hides it everywhere, and the Log keeps a "Deleted" line.

Pieces are physically tagged with orange tape showing the client's name; the app uses the same name.

### Architecture

```
 Tech's phone ──QR──▶ index.html ─┐           (GitHub Pages, this repo)
 Director ──────────▶ admin.html ─┤
                                  │  fetch(): GET ?action=data / POST JSON
                                  ▼
                    Google Apps Script web app  (apps-script/Code.gs)
                    runs as the sheet owner; access: "Anyone"
                                  │
                ┌─────────────────┼──────────────────┐
                ▼                 ▼                  ▼
        Google Sheet:        MailApp emails      Backup Google Form
        Log, Current Status, to the Director     (legacy; feeds the hidden
        Settings, Raw form                       "Raw form responses" tab)
        responses (hidden)
```

* **Front end**: two static pages (`index.html`, `admin.html`) on GitHub Pages. No build step, no frameworks, no login. They call the Apps Script web app with `fetch`.
* **Back end**: one Google Apps Script project bound to the tracker spreadsheet (`apps-script/Code.gs`, plus `apps-script/CheckIn.html`, a legacy Apps-Script-hosted version of the check-in page that the old QR code still opens).
* **Data**: the Google Sheet. The `Log` tab is the source of truth; everything else is derived from it.

### The spreadsheet

| Tab | Purpose |
|---|---|
| **Current Status** | One row per job, rebuilt from the Log after every entry. Latest status, days in status, pieces, latest notes, and an **Alert** column. Colors: red = overdue, yellow = ready for reinstallation, grey = scheduled. Deleted jobs are left out. |
| **Log** | One row per submission, oldest first. Columns: Timestamp, Job, Client, Status, Pieces, Description (pieces joined with `; `), Technician, Notes, Response ID (hidden, used to ignore duplicate submissions). Protected with a warning. |
| **Settings** | Everything the Director can change without touching code (see below). |
| **Raw form responses** (hidden) | Responses to the legacy Google Form. Not used by the current pages. |

**Settings layout**

| Cell(s) | Meaning |
|---|---|
| A2:A… | Technician names (the "Your name" list on the check-in page) |
| C2:C… | Email addresses that get the Director emails |
| F2 | Days at "Checked in at shop" before a job is flagged overdue (default 21) |
| F3 | Days at "Ready for reinstallation" before it is flagged overdue (default 3) |
| F6 | Check-in page link (goes in the QR code) |
| F7 / F8 | Legacy Google Form edit link / public link |
| F10 | **Dashboard PIN** (plain-text cell; never commit it anywhere) |
| F11 | Dashboard link |

The sheet also has an **Upholstery Tracker** menu: show/write links, refresh the backup form's lists, rebuild Current Status, set up the dashboard PIN, and run or remove a test job.

### Jobs and statuses

* A job is created by **Checked in at shop**. Its label is `Client · M/D` (check-in date), with ` (2)`, ` (3)`… added if the same client is checked in twice on one day.
* Statuses: `Checked in at shop` → `Ready for reinstallation` → `Scheduled` (closed). `Deleted` (dashboard only) also closes a job.
* The piece list recorded at check-in becomes the job's checklist for later steps.
* The "Ready for reinstallation" email lists the pieces, how long the job has been out, and any piece from the original list that didn't come back.

### Backend API (Apps Script web app)

Base URL is the web app's `/exec` URL (the `API` constant in `index.html` and `admin.html`).

| Call | Body / params | Returns |
|---|---|---|
| `GET ?action=data` | – | `{ ok, data: { statuses, newStatus, backStatus, finalStatus, notListed, openJobs: [{label, pieces[]}], techs[] } }` |
| `POST` (log an entry) | `{ status, client \| job \| other, pieces[], tech, notes, rid }` | `{ ok, job, client, status, data }` or `{ ok:false, error }` |
| `POST {action:'jobs'}` | `{ action, key }` (key = PIN) | `{ ok, jobs[], limits, statuses, sheetUrl }` |
| `POST {action:'schedule'}` | `{ action, key, job, date?, rid }` | same as `jobs` |
| `POST {action:'delete'}` | `{ action, key, job, rid }` | same as `jobs` |
| `GET` (no action) | – | Legacy HTML check-in page (`CheckIn.html`) |

POST bodies are sent as plain text (`fetch(url, {method:'POST', body: JSON.stringify(...)})`), which avoids a CORS preflight. Apps Script answers with a redirect to `script.googleusercontent.com`, which allows any origin.

**Reliability details**: every write takes a script lock. `rid` is a client-generated ID that makes retries safe: the same `rid` is never logged twice. A wrong PIN is counted, and after 20 wrong tries the dashboard locks for 10 minutes (CacheService).

### Security and privacy

* This repo is **public**. It contains no secrets: no PIN, no spreadsheet ID, no personal email addresses. Keep it that way.
* The dashboard PIN lives only in the sheet (Settings F10). Each device stores it in its own browser after the first unlock. "Lock this device" removes it.
* By design, anyone with the check-in link can submit entries and see open job labels and technician first names. Detailed job info and actions on the dashboard require the PIN.

---

## Repo layout

```
index.html               Tech check-in page (served at the site root)
admin.html               Director dashboard
logo.png                 Logo used by both pages
.nojekyll                Serve files as-is (no Jekyll build)
apps-script/Code.gs      Backend: copy of the Apps Script project's Code.gs
apps-script/CheckIn.html Backend: legacy Apps Script check-in page
printouts/               QR sign + tech instructions (HTML source, logo, ready-made PDF)
tests/mocktest.js        Logic tests for Code.gs (Node, no dependencies)
tests/e2e-server.js      Local server that runs Code.gs against an in-memory sheet
docs/MAINTAINING.md      How to change, deploy, test, troubleshoot and rebuild
```

> The files in `apps-script/` are a **copy**. The live code runs in Google Apps Script. After changing them here, paste them into the Apps Script editor and redeploy (see the maintainer guide). Keep both in sync.
