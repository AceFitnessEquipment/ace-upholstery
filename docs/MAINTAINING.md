# Maintaining the Upholstery Tracker

Read the [README](../README.md) first for what the app does and how the pieces fit together.

## Access you need

| To change… | You need |
|---|---|
| Tech names, email recipients, overdue limits, dashboard PIN | Edit access to the **tracker Google Sheet** |
| Backend logic (emails, statuses, API) | The same sheet: **Extensions › Apps Script** opens the script. Changes must be deployed by an editor of the script. The web app runs **as the sheet owner**, so emails come from the owner's Gmail. |
| Check-in page or dashboard look/wording | Write access to this GitHub repo (`AceFitnessEquipment/ace-upholstery`) |

## Everyday changes (no code)

All of these are in the sheet's **Settings** tab and take effect immediately:

* **Add or remove a tech**: edit column A. The check-in page picks it up the next time it loads.
* **Change who gets emails**: edit column C, one address per row.
* **Change the overdue limits**: F2 (days at "Checked in at shop") and F3 (days at "Ready for reinstallation").
* **Change the dashboard PIN**: type a new one in F10. Devices with the old PIN get asked for the new one.
* **Remove a mistaken entry**: use **Delete** on the dashboard. To undo a delete, remove the job's "Deleted" row from the Log tab, then run **Upholstery Tracker › Rebuild Current Status**.

## Changing the check-in page or dashboard

1. Edit `index.html` or `admin.html`. They are plain HTML/CSS/JS with no build step.
2. Test locally (see [Testing](#testing)).
3. Commit and push to `main`. GitHub Pages republishes in about a minute. The URLs and the printed QR code stay the same.

Both pages keep a constant `API` pointing at the Apps Script `/exec` URL. Only change it if the web app gets a new deployment URL (see below).

## Changing the backend (Apps Script)

1. Edit `apps-script/Code.gs` (and/or `CheckIn.html`) here and test locally.
2. Open the sheet → **Extensions › Apps Script**, paste the new code over the old, then click **Save**.
3. **Deploy › Manage deployments › ✏️ (edit) › Version: New version › Deploy.**
   * ⚠️ Always edit the existing deployment. **Don't** use "New deployment". It creates a new `/exec` URL, and both pages (`API` constant) would have to be updated.
4. Commit the same files to this repo so the copy here matches what's live.

Menu functions (Upholstery Tracker menu) and triggers always run the latest **saved** code. The web app (the pages' backend) runs the **deployed version**, so step 3 is required for any change the pages depend on.

### Triggers (set up by `setup()`)

* `handleFormSubmit`: on submit of the legacy Google Form. Logs it like any other entry.
* `handleSettingsEdit`: on edit of the sheet. Keeps the legacy form's dropdowns in sync with Settings.

Check them under **Apps Script › Triggers** (clock icon). Failures show under **Executions**.

### Useful functions in Code.gs

| Function | What it does |
|---|---|
| `logEntry_` | The one place an entry is written: dedupe by `rid`, resolve the job, append to the Log, send email, rebuild Current Status |
| `submitEntry` | Validates a check-in page submission and calls `logEntry_` |
| `doGet` / `doPost` | Web app entry points (see the API table in the README) |
| `adminJobs_` / `adminSchedule_` / `adminDelete_` / `checkAdmin_` | Dashboard |
| `loadJobs_` | Replays the Log into one object per job |
| `rebuildStatus_` | Rewrites the Current Status tab |
| `sendCheckedInEmail_` / `sendBackInShopEmail_` | The two Director emails |
| `runTestJob` / `removeTestJob` | Menu: create/remove a `ZZ TEST` job (sends 2 emails) |

## Testing

Requires Node.js (no packages to install).

```bash
# Logic tests for Code.gs (in-memory mocks of SpreadsheetApp, MailApp, etc.)
node tests/mocktest.js apps-script/Code.gs

# Run both pages locally against Code.gs and an in-memory sheet with sample jobs
node tests/e2e-server.js apps-script/Code.gs . 5183
# then open http://localhost:5183/ (check-in) and http://localhost:5183/admin.html (PIN 123456)
# http://localhost:5183/state shows the in-memory Log, Current Status and sent emails
```

The local server rewrites the pages' `API` constant to `/api`, so nothing touches the real sheet.

**Testing against the live system**: use a client name starting with `ZZ TEST`, then run **Upholstery Tracker › Remove test job** to delete all `ZZ TEST` rows. Real entries are never touched.

## Printing the QR sign and instructions

Open `printouts/shop-printouts.html` (also live at `/printouts/shop-printouts.html`). It pre-fills the check-in link. Click **Print** (Letter, background graphics on). Page 1 is the wall sign; page 2 is the tech instructions. `printouts/shop-printouts.pdf` is a ready-made copy.

If the check-in URL ever changes, open the page with `?url=<new link>`, reprint, and replace the sign. Phones then treat the new address as a new site, so techs pick their name once and the Director re-enters the PIN once.

## Troubleshooting

| Symptom | Check |
|---|---|
| Page says "Couldn't reach the server" | Phone signal; Apps Script **Executions** for errors; is the deployment still "Anyone" access? |
| No email arrived | Junk folder; Settings column C; **Executions** for `handleFormSubmit`/`doPost` errors (an email failure is reported after the entry is logged). Consumer Gmail sends at most ~100 emails/day. |
| Dashboard says "Wrong PIN" | Compare with Settings F10. After 20 wrong tries it locks for 10 minutes. |
| Tech list or job list looks old | Pull down to reload the page; the list refreshes in the background on every load. |
| Current Status looks wrong | **Upholstery Tracker › Rebuild Current Status** (it's always rebuilt from the Log). |
| Changed Code.gs but nothing changed on the pages | You saved but didn't deploy a **new version** of the existing deployment. |

## Rebuilding from scratch (new Google account or new sheet)

1. Create a new Google Sheet with the account that should own the tracker (emails will come from it).
2. **Extensions › Apps Script**: paste `apps-script/Code.gs` into `Code.gs`; add an HTML file named `CheckIn` and paste `apps-script/CheckIn.html`. Save.
3. Run `setup` from the editor and authorize. It creates the tabs, the legacy form, the triggers and the menu.
4. In Settings, enter tech names (column A) and email recipients (column C).
5. **Deploy › New deployment › Web app**, Execute as **Me**, access **Anyone**. Copy the `/exec` URL.
6. Put that URL in the `API` constant of `index.html` and `admin.html`, and update `CHECKIN_URL` in Code.gs if the pages move. Commit and push.
7. Reload the sheet, then run **Upholstery Tracker › Set up dashboard** (creates the PIN in F10) and **Write links to Settings**.
8. Reprint the QR sign if the check-in URL changed.

Old data can be carried over by copying the old Log rows into the new Log tab, then running **Rebuild Current Status**.

## Handing off: checklist

- [ ] New owner has edit access to the tracker sheet (and therefore the script)
- [ ] New owner has write access to `AceFitnessEquipment/ace-upholstery` on GitHub
- [ ] New owner knows the dashboard PIN (Settings F10). Change it if a previous person shouldn't keep access
- [ ] Settings column C lists the right email recipients
- [ ] If the sheet changes owner, emails will come from the new owner's Gmail after redeploying the web app under their account (step 5 above)
