# CSC Intake Forms (v5)

Pre-production intake forms for the Car Stereo Chick / Sounds Incredible Mobile video pipeline. Two forms + one Apps Script. Hosted as static HTML on GitHub Pages.

---

## What's in this bundle

| File | Who fills it out | When |
|---|---|---|
| **`index.html`** | — | Landing page with links to both forms |
| **`client-intake.html`** | Annie | When a build is booked, before filming |
| **`customer-intake.html`** | The car owner | Once they agree to be in the video, before filming |
| **`CSC-Apps-Script-v5.gs`** | — | Lives in the Google Sheet's Apps Script. Routes form submissions. |

The existing shot list app and daily checklist app are **not touched** — they keep working exactly as today.

---

## How it works

**Each video gets its own tab in the sheet.** Three sections stacked top to bottom:

```
📋 [Video Name]
├── 1. ANNIE'S CLIENT INTAKE BRIEF
│   Her pre-production answers in video order
│
├── 2. CUSTOMER INTAKE
│   The car owner's story (or "pending" placeholder until they submit)
│
└── 3. ROUGH STORY BREAKDOWN
    Auto-generated pre-script outline pulling from both intakes
    Reads like a brief — section headers in video order with the answers slotted underneath
```

When you finish producing the video, just **hide the tab**. Right-click the tab → Hide sheet. Done. To unhide, right-click any tab → View hidden sheets.

### The submission flow

1. **Annie books a build.** Fills out the client intake form, names the video (e.g. "Stinger Horizon 10"). Submits.
2. **A new tab appears in the sheet** named `📋 Stinger Horizon 10`. Annie's brief fills sections 1 and 3. Section 2 shows "⏳ Waiting for customer."
3. **The client form's success screen shows a copy-link helper** — the customer form URL with the video name already attached. Annie clicks "Copy link" and pastes it into a text/email to the customer.
4. **Customer clicks the link, fills out their form, submits.** Their data lands in section 2 of the same tab. Section 3 (the story breakdown) regenerates to include their answers too.

If the customer fills out their form before Annie's client intake exists, their submission goes to a fallback tab called `🎧 Customer Submissions (orphans)` so the data isn't lost.

---

## Live URLs (after GitHub Pages is enabled)

- Landing: `https://francisdinoto-png.github.io/csc-intake-forms/`
- Client form: `https://francisdinoto-png.github.io/csc-intake-forms/client-intake.html`
- Customer form: `https://francisdinoto-png.github.io/csc-intake-forms/customer-intake.html`

You'll only ever send the customer form URL with a video name attached — the success screen on the client form generates that link for you.

---

## Deploying — full walkthrough

### Step 1 — Push the HTML to GitHub

You already have the repo at `github.com/francisdinoto-png/csc-intake-forms`. Replace the existing files:

1. Open your `csc-intake-forms` repo
2. Click **Add file → Upload files**
3. Drag in all four files: `index.html`, `client-intake.html`, `customer-intake.html`, `README.md`
4. GitHub will detect they already exist and prompt to replace — that's fine
5. Commit message: `Update forms — v5 with copy-link + video-param`
6. Click **Commit changes**

GitHub Pages rebuilds in ~30 seconds.

### Step 2 — Update the Apps Script

1. Open the production Google Sheet
2. **Extensions → Apps Script**
3. Select all the existing code (Ctrl/Cmd + A) → delete
4. Open `CSC-Apps-Script-v5.gs` from this bundle in any text editor → copy all
5. Paste into the Apps Script editor
6. Click **💾 Save** icon
7. **(Optional but recommended)** Run `setupCSCProduction()` once — Function picker (top toolbar) → select `setupCSCProduction` → ▶ Run. This safely creates 📓 DAILY NOTES and 🎬 SHOT LOG if missing. It does NOT delete anything.
8. **Deploy → Manage deployments → ✏️ Edit** active deployment → Version: **New version** → describe as *"v5 per-build tabs + story breakdown"* → **Deploy**

The Web App URL stays the same. The forms keep posting to it.

### Step 3 — Test (in this order)

1. **Open the client form.** Fill out a few questions with video name "TEST BUILD" and pick Vlog Style. Submit.
2. **Check the sheet** → new tab `📋 TEST BUILD` should exist with section 1 filled and section 2 showing "Waiting for customer."
3. **The success screen on the client form** should show a customer link like `.../customer-intake.html?video=TEST+BUILD` with a Copy button. Click "Open in new tab" to verify the link works.
4. **On the customer form** — you should see "For: TEST BUILD" in blue near the top. Fill it out with "Test Customer." Submit.
5. **Check the sheet again** → on the `📋 TEST BUILD` tab, section 2 is now filled and section 3 has regenerated to include the customer's answers.
6. **Submit from the shot list app** at EOD → new row in `🎬 SHOT LOG` as before. Unchanged.
7. Delete the `📋 TEST BUILD` tab when you're satisfied.

---

## Notes

### Webhook URL
The forms point to this URL (already in both HTML files):
```
https://script.google.com/macros/s/AKfycbwN3EpUwQpMpByvBM98BLttHX3PmhnfCJFTY7MaFTa3tfDrHlAF0o71G3n2sTlKw_Vl/exec
```
If you ever need to change it, edit the `SHEET_WEBHOOK_URL` constant at the top of the `<script>` block in both forms.

### 30-day progress save
Both intake forms save in-progress answers to browser local storage. Closes tab → comes back within 30 days → "Resume" prompt appears. After 30 days, the saved draft auto-clears.

### Customer video uploads
- Optional 30-second video/audio file at the end of the customer form
- Uploads up to ~25 MB go to a Drive folder named after the customer's first name
- Parent folder ID is hardcoded at the top of the Apps Script (`VIDEO_PARENT_FOLDER_ID`)
- 30 seconds of phone video is well under the limit

### Hiding old build tabs
When a video is done and published:
- Right-click the tab → **Hide sheet**
- The tab stays in the sheet but doesn't clutter the bottom strip
- To see hidden tabs: right-click any visible tab → **View hidden sheets**
- The data is fully preserved — you can unhide anytime if you need to revisit a build

### Editing the forms
To tweak a question, add a service category, or change wording:
- Open the relevant HTML file
- The client form has a `SHARED_START`, `VLOG_QUESTIONS`, and `REVIEW_QUESTIONS` array near the top of the `<script>` block
- The customer form has `BASE_QUESTIONS`, `SENSORY_QUESTIONS`, and `FINAL_QUESTIONS`
- Each question is a JSON object with `id`, `text`, `hint`, `type`, etc.
- After editing, commit & push to GitHub. Pages auto-rebuilds in ~30 seconds

If you change a question's `fieldId`, also update the matching row label in the Apps Script (`clientVlogSections`, `clientReviewSections`, the customer sections in `writeCustomerIntakeSection`, and the story breakdown builders).
