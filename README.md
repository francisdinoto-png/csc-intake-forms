# CSC Intake Forms

Pre-production intake forms for the Car Stereo Chick / Sounds Incredible Mobile video pipeline. Two forms + one Apps Script. Hosted as static HTML on GitHub Pages.

---

## What's in this bundle

| File | Who fills it out | When |
|---|---|---|
| **`index.html`** | — | Landing page with links to both forms |
| **`client-intake.html`** | Annie | When a new build is booked, before filming |
| **`customer-intake.html`** | The car owner | When they agree to be in the video, before filming |
| **`CSC-Apps-Script-v4.gs`** | — | Lives in Google Apps Script. Routes form submissions to the sheet. |

The existing shot list app and daily checklist app are **not touched** — they keep working exactly as they do today.

---

## How it works

1. Annie books a build. She fills out **client-intake.html**, which auto-saves her progress for 30 days. She submits when done.
2. Once a customer agrees to be in the video, Annie sends them the URL for **customer-intake.html**. They fill it out (3-5 min), submit. Optional 30-second video upload goes to a Drive folder named after them.
3. **Both intake forms append to a single tab in the sheet** — each new submission stacks below the last as a labeled, formatted section. The top of each tab has a clickable TOC (newest first) so you can jump straight to any build's brief.
4. The shot list app and daily checklist keep appending rows to `🎬 SHOT LOG` and `📓 DAILY NOTES` as they always have.

**Nothing overwrites anything. Everything is scrollable history.**

---

## Final sheet layout

After deploying, the sheet has four tabs:

| Tab | What it holds |
|---|---|
| **📋 CLIENT INTAKES** | TOC at top + every Annie pre-pro brief stacked below |
| **🎧 CUSTOMER INTAKES** | TOC at top + every customer submission stacked below |
| **📓 DAILY NOTES** | One row per daily checklist submission |
| **🎬 SHOT LOG** | One row per shot list EOD submission |

The TOC entries are clickable hyperlinks — click one and the sheet jumps you straight to that build's section.

Customer video uploads land in Drive folders: `[Video Parent Folder]/[Customer First Name]/[date]_filename.mp4`

---

## Live URLs (once GitHub Pages is enabled)

- Landing: `https://YOUR-USERNAME.github.io/csc-intake-forms/`
- Client form: `https://YOUR-USERNAME.github.io/csc-intake-forms/client-intake.html`
- Customer form: `https://YOUR-USERNAME.github.io/csc-intake-forms/customer-intake.html`

Send Annie the client URL. Send each customer their own URL when they're ready.

---

## Deploying — full walkthrough

### Step 1 — Prep the sheet manually first

Before touching the Apps Script, clean up the sheet by hand:

1. **Rename `📋 CLIENT INTAKE`** → right-click → Rename → `📋 Stinger Horizon 10` (preserves Annie's filled data as an archived build tab)
2. **Delete dead tabs** (right-click → Delete): 🎧 CUSTOMER INTAKE (empty), 📝 SCRIPT, 🎬 SHOT LIST, 🎙️ VO, 📰 BLOG DRAFT, 📓 BUILD LOG
3. The 🎬 SHOT LOG tab (with the 9 rows of 993 data) — keep or delete, your call

Leave **📓 DAILY NOTES** alone.

### Step 2 — Push the HTML to your GitHub repo

If you've already got the repo from earlier, upload these new files and let them overwrite:
1. Go to your `csc-intake-forms` repo
2. **Add file → Upload files** → drag in `index.html`, `client-intake.html`, `customer-intake.html`, `README.md`
3. Commit message: `Update intake forms v2`
4. **Commit changes**

GitHub Pages auto-rebuilds in ~30 seconds.

### Step 3 — Update the Apps Script

1. Open the production Google Sheet
2. **Extensions → Apps Script**
3. Select all the existing code (Ctrl/Cmd + A) → delete
4. Open `CSC-Apps-Script-v4.gs` from this bundle in any text editor → copy all
5. Paste into the Apps Script editor
6. Click **💾 Save**
7. **(Recommended)** Run `setupCSCProduction()` once. Function picker (top toolbar, next to ▶ Run) → select `setupCSCProduction` → ▶ Run. First time it'll prompt for permissions — Review → Advanced → Allow.

   This creates the four tabs above. It will NOT delete anything.

8. **Deploy → Manage deployments → ✏️ Edit** the active deployment → Version: **New version** → describe as "v4 stacked sections + TOC" → **Deploy**

The Web App URL stays the same. The forms keep posting to it.

### Step 4 — Test

1. Open the client form on the GitHub Pages URL. Fill out a few questions with a test video name like "TEST BUILD". Submit.
2. Check the sheet → 📋 CLIENT INTAKES tab. You should see:
   - A TOC entry at the top: `→ TEST BUILD · Vlog Style · [date]` (clickable, jumps to the brief)
   - The brief itself stacked below the last one (or as the first entry if it's empty)
3. Open the customer form in a different browser. Fill it out as "Test Customer". Submit.
4. Check the 🎧 CUSTOMER INTAKES tab — same pattern.
5. Submit a test EOD from the shot list app → new row in 🎬 SHOT LOG (unchanged behavior)
6. Delete the test entries when satisfied (just delete the rows; the TOC entries can also be deleted manually)

---

## Notes

### Webhook URL
The forms already point to this Apps Script URL:
```
https://script.google.com/macros/s/AKfycbwN3EpUwQpMpByvBM98BLttHX3PmhnfCJFTY7MaFTa3tfDrHlAF0o71G3n2sTlKw_Vl/exec
```
If that ever changes, update `SHEET_WEBHOOK_URL` at the top of the `<script>` block in both HTML files.

### 30-day progress save
Both intake forms save progress in the browser's local storage. If someone closes the tab mid-form and comes back within 30 days (same browser/device), they get a "Resume where you left off" prompt. After 30 days, the saved draft auto-clears.

### Customer video uploads
- Files up to ~25 MB upload directly to a Drive folder named after the customer's first name
- Parent folder ID is hardcoded at the top of the Apps Script (`VIDEO_PARENT_FOLDER_ID`)
- Larger files are skipped with a console warning
- 30-second phone video at typical quality is well under the limit

### Long-term organization
After 50+ builds, each intakes tab will be hundreds of rows long — but the TOC at the top stays current and is clickable. Newest builds always appear at the top of the TOC. Ctrl+F still works to search across the whole sheet.

### Editing the forms
To tweak a question, add a service category, or change wording:
- Open the relevant HTML file
- Both forms have a `QUESTIONS` array (client) or `BASE_QUESTIONS` / `SENSORY_QUESTIONS` (customer) near the top of the `<script>` block
- Each question is a JSON object with `id`, `text`, `hint`, `type`, etc.
- After editing, commit & push to GitHub. Pages auto-rebuilds in ~30 seconds.

If you change a question's `fieldId`, you'll also need to update the matching row label in the Apps Script (`clientVlogSections`, `clientReviewSections`, or the customer sections inside `appendCustomerIntake`).
