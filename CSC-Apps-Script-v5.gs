// ═══════════════════════════════════════════════════════════════════════════
// CSC PRODUCTION SHEET — Google Apps Script  (v5 — final)
// ═══════════════════════════════════════════════════════════════════════════
//
// WHAT THIS DOES
//
//   1. Client intake (Annie's pre-production brief) submitted →
//      Creates a new tab named "📋 [Video Name]" with three stacked sections:
//        1. ANNIE'S CLIENT INTAKE BRIEF
//        2. CUSTOMER INTAKE  [pending until customer submits]
//        3. ROUGH STORY BREAKDOWN  (pre-script outline)
//
//   2. Customer intake submitted →
//      Form sends a 'video_name' field (from ?video=… URL parameter).
//      Script finds the matching "📋 [Video Name]" tab, fills in the
//      customer section, and regenerates the story breakdown.
//      If no matching tab exists, customer's data lands on a fallback tab.
//
//   3. Shot list app + daily checklist (existing) → appends to 🎬 SHOT LOG
//      and 📓 DAILY NOTES as it does today. Untouched.
//
// SAFETY
//   - This script does NOT delete any tabs.
//   - setupCSCProduction() only creates the log tabs if missing.
//   - Every client submission creates a new per-build tab.
//   - Customer submissions land on the matching tab if found; otherwise
//     a fallback "🎧 Customer Submissions (orphans)" tab.
//
// SETUP
//   Extensions → Apps Script → paste this whole file → save.
//   Optionally run setupCSCProduction() to ensure the log tabs exist.
//   Then: Deploy → Manage deployments → Edit (pencil) → New version → Deploy
//   The Web App URL stays the same — every form already points to it.
//
// ═══════════════════════════════════════════════════════════════════════════


// ── COLORS ──────────────────────────────────────────────────────────────────
const C = {
  ink:        '#0A0B0E',
  orange:     '#E8460A',
  amber:      '#F0A500',
  green:      '#3A7A35',
  green2:     '#4E9E47',
  blue:       '#4A7EC8',
  purple:     '#9B6DD9',
  dark_red:   '#8B1A1A',
  grey_mid:   '#5C5E6A',
  white:      '#FFFFFF',
  header_bg:  '#111318',
  section_a:  '#FFF6E6',    // soft amber bg for section dividers
  section_b:  '#EAF1FB',    // soft blue bg
  section_c:  '#F0F8EF',    // soft green bg
  row_alt:    '#F8F8F8',
  row_white:  '#FFFFFF',
  divider:    '#222222',
  pending:    '#FFF59D',    // pale yellow for "pending customer"
};

// ── DRIVE FOLDER (customer video upload destination) ────────────────────────
const VIDEO_PARENT_FOLDER_ID = '1etjiS0yORaqtgaUXW6_dQyFLYdHKWqG0';
const AUDIO_FOLDER_ID        = '1RCscdw4F2k5H19IzYbV9bcqH1hr-zks9';

// ── LOG TAB NAMES ───────────────────────────────────────────────────────────
const TAB_NOTES    = '📓 DAILY NOTES';
const TAB_SHOTLOG  = '🎬 SHOT LOG';
const TAB_ORPHANS  = '🎧 Customer Submissions (orphans)';


// ═══════════════════════════════════════════════════════════════════════════
// SETUP — safe. Only creates missing log tabs. Does not delete anything.
// ═══════════════════════════════════════════════════════════════════════════
function setupCSCProduction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(TAB_NOTES);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_NOTES);
    sheet.setTabColor('#1A6B5A');
    buildDailyNotesTab(sheet);
  }

  sheet = ss.getSheetByName(TAB_SHOTLOG);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_SHOTLOG);
    sheet.setTabColor(C.blue);
    buildShotLogTab(sheet);
  }

  SpreadsheetApp.getUi().alert(
    '✅ Setup complete.\n\n' +
    'Each client intake submission will create its own tab named "📋 [Video Name]".\n' +
    'Customer intake submissions land on the matching tab automatically.\n\n' +
    'No existing data was touched.'
  );
}


// ═══════════════════════════════════════════════════════════════════════════
// doPost — RECEIVES FORM SUBMISSIONS
// ═══════════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contentType = (e.postData && e.postData.type) || '';

    // ── ROUTE 1: SHOT LIST APP + DAILY CHECKLIST (form-encoded) ──────────────
    if (contentType.indexOf('application/x-www-form-urlencoded') !== -1 ||
        (e.parameter && e.parameter.form_type === 'build_log')) {
      const p = e.parameter || {};
      if (p.form_type === 'build_log') {
        return writeToBuildLog(ss, p);
      }
    }

    // ── ROUTE 2: INTAKE FORMS (JSON) ─────────────────────────────────────────
    const raw  = e.postData.contents;
    const data = JSON.parse(raw);

    // Save uploaded customer video to a folder named after them
    if (data.customer_files && data.customer_files.length) {
      saveCustomerVideoToNamedFolder(data.customer_files, data);
    }

    // Save Annie's audio brain-dumps to the audio folder
    if (data.audio_file_data && data.audio_file_data.length) {
      saveFilesToDrive(data.audio_file_data, AUDIO_FOLDER_ID, data, 'annie');
    }

    if (data.form_type === 'customer_intake') {
      const tabName = handleCustomerIntake(ss, data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, tab: tabName }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      // default: client_intake
      const tabName = handleClientIntake(ss, data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, tab: tabName }))
        .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    console.error('doPost error:', err.message, err.stack);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ═══════════════════════════════════════════════════════════════════════════
// TAB NAMING
// ═══════════════════════════════════════════════════════════════════════════
function tabNameForVideo(videoName) {
  const cleaned = (videoName || 'Untitled Build').toString().trim();
  return cleanTabName('📋 ' + cleaned);
}

function cleanTabName(name) {
  return name
    .replace(/[\[\]\*\?\/\\:']/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 100);
}

function uniqueTabName(ss, desiredName) {
  let name = desiredName;
  let n    = 2;
  while (ss.getSheetByName(name)) {
    name = desiredName + ' (' + n + ')';
    n++;
    if (n > 50) break;
  }
  return name;
}


// ═══════════════════════════════════════════════════════════════════════════
// CLIENT INTAKE → creates a new per-build tab
// ═══════════════════════════════════════════════════════════════════════════
function handleClientIntake(ss, data) {
  const desiredName = tabNameForVideo(data.video_name);
  const tabName = uniqueTabName(ss, desiredName);

  const sheet = ss.insertSheet(tabName);
  sheet.setTabColor(C.orange);
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 700);

  // Cache client data on the sheet so we can rebuild the story when customer arrives
  // (Stored in developer metadata so we don't pollute the visible data)
  sheet.addDeveloperMetadata('client_data', JSON.stringify(data));

  // ── Section 1: Annie's client intake brief ──
  let row = writeClientIntakeSection(sheet, data, 1);

  // ── Section 2: Customer intake placeholder ──
  row += 1;
  row = writeCustomerPlaceholder(sheet, row);

  // ── Section 3: Story breakdown ──
  row += 1;
  writeStoryBreakdown(sheet, row, data, null); // null = no customer yet

  sheet.setFrozenRows(2);

  return tabName;
}

function writeClientIntakeSection(sheet, data, startRow) {
  let row = startRow;

  // Title bar
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('▼ 1 — ANNIE\'S CLIENT INTAKE BRIEF')
    .setBackground(C.ink).setFontColor(C.white).setFontSize(15)
    .setFontWeight('bold').setFontFamily('Arial').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 40);
  row++;

  // Subtitle
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy · h:mm a');
  const fmt = data.video_format === 'review' ? 'Review / How-To' : 'Vlog Style';
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('Submitted ' + ts + '  ·  ' + fmt)
    .setBackground(C.header_bg).setFontColor(C.grey_mid)
    .setFontSize(10).setFontFamily('Arial').setFontStyle('italic');
  sheet.setRowHeight(row, 26);
  row++;

  // FIELD / VALUE column headers
  ['FIELD', 'VALUE'].forEach(function(h, i) {
    sheet.getRange(row, i + 1).setValue(h)
      .setBackground(C.orange).setFontColor(C.white).setFontSize(9)
      .setFontWeight('bold').setFontFamily('Arial').setHorizontalAlignment('left');
  });
  sheet.setRowHeight(row, 28);
  row++;

  // Sections based on video format
  const isVlog = (data.video_format !== 'review');
  const sections = isVlog ? clientVlogSections() : clientReviewSections();

  sections.forEach(function(sec) {
    sheet.getRange(row, 1, 1, 2).merge()
      .setValue(sec.label)
      .setBackground(sec.color).setFontColor(C.white).setFontSize(9)
      .setFontWeight('bold').setFontFamily('Arial');
    sheet.setRowHeight(row, 26);
    row++;

    sec.rows.forEach(function(r, i) {
      const label   = r[0];
      const fieldId = r[1];
      const rawVal  = data[fieldId];
      const value   = (rawVal === undefined || rawVal === null) ? '' : rawVal.toString();
      const bg = i % 2 === 0 ? C.row_white : C.row_alt;

      sheet.getRange(row, 1).setValue(label).setBackground(bg)
        .setFontColor('#333333').setFontSize(10).setFontFamily('Arial')
        .setVerticalAlignment('top').setNote('Field ID: ' + fieldId);

      sheet.getRange(row, 2).setValue(value).setBackground(bg)
        .setFontColor('#111111').setFontSize(11).setFontFamily('Arial')
        .setVerticalAlignment('top').setWrap(true);

      const lines = (value.match(/\n/g) || []).length + 1;
      const wordWrapLines = Math.ceil(value.length / 90);
      const targetLines = Math.max(lines, wordWrapLines);
      sheet.setRowHeight(row, Math.max(36, Math.min(targetLines * 16, 400)));
      row++;
    });
  });

  return row;
}

function writeCustomerPlaceholder(sheet, startRow) {
  let row = startRow;

  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('▼ 2 — CUSTOMER INTAKE')
    .setBackground(C.ink).setFontColor(C.white).setFontSize(15)
    .setFontWeight('bold').setFontFamily('Arial').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 40);
  row++;

  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('⏳ Waiting for customer to fill out their form. Send them the customer link from the client form\'s success screen.')
    .setBackground(C.pending).setFontColor('#5A4500')
    .setFontSize(11).setFontFamily('Arial').setFontStyle('italic').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 44);
  row++;

  return row;
}


// ═══════════════════════════════════════════════════════════════════════════
// CUSTOMER INTAKE → finds the matching tab and fills the customer section
// ═══════════════════════════════════════════════════════════════════════════
function handleCustomerIntake(ss, data) {
  const videoName = (data.video_name || '').toString().trim();

  // Try to find the matching client tab
  let sheet = null;
  if (videoName) {
    const desiredTab = tabNameForVideo(videoName);
    sheet = ss.getSheetByName(desiredTab);

    // Also try variants — maybe Annie's tab has a (2) suffix or slight diff
    if (!sheet) {
      const allSheets = ss.getSheets();
      for (let i = 0; i < allSheets.length; i++) {
        const name = allSheets[i].getName();
        if (name.indexOf(videoName) !== -1 && name.indexOf('📋') !== -1) {
          sheet = allSheets[i];
          break;
        }
      }
    }
  }

  // No matching tab found — write to orphan tab so data isn't lost
  if (!sheet) {
    return writeCustomerToOrphans(ss, data);
  }

  // Found the matching tab — rewrite the customer section + story breakdown
  rebuildCustomerAndStory(sheet, data);
  return sheet.getName();
}

function rebuildCustomerAndStory(sheet, customerData) {
  // Find the row where "▼ 2 — CUSTOMER INTAKE" lives
  const data = sheet.getDataRange().getValues();
  let customerStart = -1;
  let storyStart = -1;
  for (let i = 0; i < data.length; i++) {
    const v = String(data[i][0] || '');
    if (v.indexOf('▼ 2 — CUSTOMER INTAKE') !== -1) customerStart = i + 1;
    if (v.indexOf('▼ 3 — ROUGH STORY BREAKDOWN') !== -1) storyStart = i + 1;
  }

  if (customerStart === -1) {
    // Couldn't find section 2 — append at end (defensive fallback)
    customerStart = sheet.getLastRow() + 2;
  }

  // Determine the range to clear: customerStart through end of sheet
  const lastRow = sheet.getLastRow();
  if (lastRow >= customerStart) {
    // Clear everything from section 2 onward (we're rebuilding both 2 and 3)
    sheet.getRange(customerStart, 1, lastRow - customerStart + 1, 2).clear();
  }

  // Retrieve cached client data to regenerate the story breakdown
  let clientData = {};
  try {
    const meta = sheet.createDeveloperMetadataFinder()
      .withKey('client_data').find();
    if (meta && meta.length > 0) {
      clientData = JSON.parse(meta[0].getValue());
    }
  } catch (err) {
    console.warn('Could not retrieve cached client data: ' + err.message);
  }

  // Cache the customer data too so future rebuilds work
  try {
    const existing = sheet.createDeveloperMetadataFinder()
      .withKey('customer_data').find();
    if (existing && existing.length > 0) existing[0].remove();
    sheet.addDeveloperMetadata('customer_data', JSON.stringify(customerData));
  } catch (err) {
    console.warn('Could not cache customer data: ' + err.message);
  }

  // Write section 2 — customer intake
  let row = writeCustomerIntakeSection(sheet, customerData, customerStart);

  // Write section 3 — story breakdown (using both client + customer data)
  row += 1;
  writeStoryBreakdown(sheet, row, clientData, customerData);
}

function writeCustomerIntakeSection(sheet, data, startRow) {
  let row = startRow;

  // Title
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('▼ 2 — CUSTOMER INTAKE')
    .setBackground(C.ink).setFontColor(C.white).setFontSize(15)
    .setFontWeight('bold').setFontFamily('Arial').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 40);
  row++;

  // Subtitle
  const customerName = (data.customer_name || 'Unnamed Customer').toString().trim();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy · h:mm a');
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('Submitted by ' + customerName + ' · ' + ts)
    .setBackground(C.header_bg).setFontColor(C.grey_mid)
    .setFontSize(10).setFontFamily('Arial').setFontStyle('italic');
  sheet.setRowHeight(row, 26);
  row++;

  // Column headers
  ['FIELD', 'VALUE'].forEach(function(h, i) {
    sheet.getRange(row, i + 1).setValue(h)
      .setBackground(C.green2).setFontColor(C.white).setFontSize(9)
      .setFontWeight('bold').setFontFamily('Arial').setHorizontalAlignment('left');
  });
  sheet.setRowHeight(row, 28);
  row++;

  // Format services + files for display
  if (Array.isArray(data.services)) {
    data.services_display = data.services.join(', ');
  } else {
    data.services_display = data.services || '';
  }
  if (data.customer_files && data.customer_files.length) {
    data.customer_files_display = data.customer_files
      .map(function(f) { return f.name + ' (saved to Drive folder: ' + customerName + ')'; })
      .join('\n');
  } else {
    data.customer_files_display = '';
  }

  const sections = [
    { label: '◆ ABOUT', color: C.blue, rows: [
      ['First name',                       'customer_name'],
      ['Services they\'re getting',        'services_display'],
      ['Wants to be on camera?',           'on_camera'],
    ]},
    { label: '◆ THE STORY', color: C.purple, rows: [
      ['What this car means to them',      'car_story'],
      ['Why now',                          'why_now'],
      ['What they hope changes',           'hopes'],
    ]},
    { label: '◆ SENSORY ANSWERS', color: C.amber, rows: [
      ['Audio — first song + why',                  'sensory_audio'],
      ['CarPlay/Android Auto — first thing',        'sensory_carplay'],
      ['OEM integration — why factory look',        'sensory_oem'],
      ['Remote starter — the morning',              'sensory_remote'],
      ['Heated seats — the moment',                 'sensory_heated'],
      ['Sound dampening — what drives them nuts',   'sensory_dampening'],
      ['Radar & laser — when this matters',         'sensory_radar'],
      ['Vehicle safety — moment they wished',       'sensory_safety'],
      ['Wrangler — kind of driving + build',        'sensory_wrangler'],
      ['Other — what + what changes',               'sensory_other'],
    ]},
    { label: '◆ MEDIA', color: C.green2, rows: [
      ['Customer video/audio',             'customer_files_display'],
    ]},
  ];

  sections.forEach(function(sec) {
    sheet.getRange(row, 1, 1, 2).merge()
      .setValue(sec.label)
      .setBackground(sec.color).setFontColor(C.white).setFontSize(9)
      .setFontWeight('bold').setFontFamily('Arial');
    sheet.setRowHeight(row, 26);
    row++;

    sec.rows.forEach(function(r, i) {
      const label   = r[0];
      const fieldId = r[1];
      const rawVal  = data[fieldId];

      if (sec.label.indexOf('SENSORY') !== -1 &&
          (rawVal === undefined || rawVal === null || rawVal === '')) {
        return;
      }

      const value = (rawVal === undefined || rawVal === null) ? '' : rawVal.toString();
      const bg = i % 2 === 0 ? C.row_white : C.row_alt;

      sheet.getRange(row, 1).setValue(label).setBackground(bg)
        .setFontColor('#333333').setFontSize(10).setFontFamily('Arial')
        .setVerticalAlignment('top').setNote('Field ID: ' + fieldId);

      sheet.getRange(row, 2).setValue(value).setBackground(bg)
        .setFontColor('#111111').setFontSize(11).setFontFamily('Arial')
        .setVerticalAlignment('top').setWrap(true);

      const lines = (value.match(/\n/g) || []).length + 1;
      const wordWrapLines = Math.ceil(value.length / 90);
      const targetLines = Math.max(lines, wordWrapLines);
      sheet.setRowHeight(row, Math.max(36, Math.min(targetLines * 16, 400)));
      row++;
    });
  });

  return row;
}

// Orphan customer tab — if there's no matching client intake
function writeCustomerToOrphans(ss, data) {
  let sheet = ss.getSheetByName(TAB_ORPHANS);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_ORPHANS);
    sheet.setTabColor(C.dark_red);
    sheet.setColumnWidth(1, 280);
    sheet.setColumnWidth(2, 700);

    sheet.getRange(1, 1, 1, 2).merge()
      .setValue('CUSTOMER SUBMISSIONS — NO MATCHING BUILD')
      .setBackground(C.ink).setFontColor(C.white).setFontSize(15)
      .setFontWeight('bold').setFontFamily('Arial').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 40);

    sheet.getRange(2, 1, 1, 2).merge()
      .setValue('Customer submissions that didn\'t match a client intake tab end up here. Usually because Annie sent the wrong URL, or the customer filled it out before Annie did the client intake. Move them manually if needed.')
      .setBackground(C.header_bg).setFontColor(C.grey_mid)
      .setFontSize(10).setFontFamily('Arial').setFontStyle('italic');
    sheet.setRowHeight(2, 36);
    sheet.setFrozenRows(2);
  }

  const startRow = Math.max(sheet.getLastRow(), 2) + 2;
  writeCustomerIntakeSection(sheet, data, startRow);

  // End divider
  const endRow = sheet.getLastRow() + 1;
  sheet.getRange(endRow, 1, 1, 2).merge()
    .setValue('— end of orphan submission —')
    .setBackground(C.divider).setFontColor(C.grey_mid).setFontSize(9)
    .setFontStyle('italic').setFontFamily('Arial').setHorizontalAlignment('center');
  sheet.setRowHeight(endRow, 22);

  return TAB_ORPHANS;
}


// ═══════════════════════════════════════════════════════════════════════════
// STORY BREAKDOWN — templated pre-script outline
// ═══════════════════════════════════════════════════════════════════════════
function writeStoryBreakdown(sheet, startRow, clientData, customerData) {
  let row = startRow;

  // Title
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('▼ 3 — ROUGH STORY BREAKDOWN')
    .setBackground(C.ink).setFontColor(C.white).setFontSize(15)
    .setFontWeight('bold').setFontFamily('Arial').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 40);
  row++;

  // Subtitle
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy · h:mm a');
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('Auto-generated outline · pulls from client + customer intake · ' + ts)
    .setBackground(C.header_bg).setFontColor(C.grey_mid)
    .setFontSize(10).setFontFamily('Arial').setFontStyle('italic');
  sheet.setRowHeight(row, 26);
  row++;

  const isVlog = (clientData.video_format !== 'review');
  const sections = isVlog ? buildVlogBreakdown(clientData, customerData) : buildReviewBreakdown(clientData);

  sections.forEach(function(sec) {
    // Section header
    sheet.getRange(row, 1, 1, 2).merge()
      .setValue(sec.header)
      .setBackground(sec.color).setFontColor(C.white).setFontSize(11)
      .setFontWeight('bold').setFontFamily('Arial');
    sheet.setRowHeight(row, 30);
    row++;

    // Body content
    if (sec.body) {
      sheet.getRange(row, 1, 1, 2).merge()
        .setValue(sec.body)
        .setBackground(sec.bgTint || C.row_white)
        .setFontColor('#111111').setFontSize(11).setFontFamily('Arial')
        .setVerticalAlignment('top').setWrap(true);
      const lines = (sec.body.match(/\n/g) || []).length + 1;
      const wordWrapLines = Math.ceil(sec.body.length / 100);
      const targetLines = Math.max(lines, wordWrapLines);
      sheet.setRowHeight(row, Math.max(40, Math.min(targetLines * 17, 800)));
      row++;
    }
  });
}


// Format a list of bullets, skipping empties
function bullets(items) {
  return items
    .filter(function(x) { return x && x.label && x.value && x.value.toString().trim(); })
    .map(function(x) { return '• ' + x.label + ': ' + x.value.toString().trim(); })
    .join('\n');
}

// Just the value, prefixed with a label inline
function labeled(label, value) {
  if (!value || !value.toString().trim()) return '';
  return label + ': ' + value.toString().trim();
}

function emptyMark(v) {
  return (v && v.toString().trim()) ? v.toString().trim() : '[ — not yet filled in — ]';
}


function buildVlogBreakdown(c, cust) {
  // c = client intake data; cust = customer intake data (may be null)
  const hasCustomer = !!cust;

  const sections = [];

  // ── 0 — VIDEO META ──
  sections.push({
    header: '◆ VIDEO META',
    color: C.grey_mid,
    bgTint: C.row_white,
    body: bullets([
      { label: 'Video Name',  value: c.video_name },
      { label: 'Format',      value: 'Vlog Style' },
      { label: 'Install Window', value: (c.install_start || '?') + ' → ' + (c.install_end || '?') },
      { label: 'Customer Intake', value: hasCustomer ? '✓ Received from ' + (cust.customer_name || 'customer') : '⏳ Pending' },
    ])
  });

  // ── §1 HOOK ──
  sections.push({
    header: '§1 · HOOK  ·  0:00–0:10',
    color: C.orange,
    bgTint: C.section_a,
    body: emptyMark(c.hook),
  });

  // ── §2 COLD OPEN ──
  let coldOpenBody = bullets([
    { label: 'Client + history',   value: c.client_context },
    { label: 'Car role in life',   value: c.car_role },
    { label: 'Their words',        value: c.client_said ? '"' + c.client_said + '"' : '' },
  ]);
  if (hasCustomer) {
    const custLines = bullets([
      { label: 'Customer on what this car means', value: cust.car_story },
      { label: 'Customer\'s why-now',             value: cust.why_now },
    ]);
    if (custLines) coldOpenBody += '\n\n— FROM CUSTOMER —\n' + custLines;
  }
  sections.push({
    header: '§2 · COLD OPEN  ·  0:10–0:45',
    color: C.purple,
    bgTint: C.section_b,
    body: coldOpenBody || '[ — cold open data not yet filled in — ]',
  });

  // ── §3 THE CAR ──
  sections.push({
    header: '§3 · THE CAR  ·  0:45–1:30',
    color: C.orange,
    bgTint: C.section_a,
    body: bullets([
      { label: 'Vehicle', value: [c.year, c.make, c.model, c.color].filter(Boolean).join(' ') },
      { label: 'What makes it special', value: c.vehicle_special },
      { label: 'Factory audio situation', value: c.factory_audio },
    ]) || '[ — car data not yet filled in — ]',
  });

  // ── §4 CLIENT BRIEF ──
  let briefBody = bullets([
    { label: 'What they want (feeling)', value: c.client_wants },
    { label: 'Emotional stakes',         value: c.client_stakes },
  ]);
  if (hasCustomer) {
    const custLines = bullets([
      { label: 'Customer\'s hopes',  value: cust.hopes },
    ]);
    if (custLines) briefBody += '\n\n— FROM CUSTOMER —\n' + custLines;
  }
  sections.push({
    header: '§4 · CLIENT BRIEF  ·  1:30–2:30',
    color: C.blue,
    bgTint: C.section_b,
    body: briefBody || '[ — brief data not yet filled in — ]',
  });

  // ── §5 PRODUCTS + PLAN ──
  sections.push({
    header: '§5 · PRODUCTS + THE PLAN  ·  2:30–5:00',
    color: C.amber,
    bgTint: C.section_a,
    body: bullets([
      { label: 'Products going in', value: c.products_list },
      { label: 'Why these products', value: c.products_why },
      { label: 'Alternatives passed on', value: c.alternatives },
    ]) || '[ — products not yet filled in — ]',
  });

  // ── §6 THE CHALLENGE ──
  sections.push({
    header: '§6 · THE CHALLENGE  ·  5:00–6:00',
    color: C.purple,
    bgTint: C.section_b,
    body: bullets([
      { label: 'Unique challenge of this build', value: c.challenge },
      { label: 'What\'s at risk',                value: c.at_risk },
    ]) || '[ — challenge not yet filled in — ]',
  });

  // ── §7 INSTALL PLAN ──
  sections.push({
    header: '§7 · INSTALL  ·  6:00–14:00  (day-by-day)',
    color: C.green2,
    bgTint: C.section_c,
    body: bullets([
      { label: 'Day-by-day plan',          value: c.install_plan },
      { label: 'Named tools (close-ups)',  value: c.install_tools },
      { label: 'Hacks / tips / How-Tos',   value: c.install_hacks },
    ]) || '[ — install plan not yet filled in — ]',
  });

  // ── §8 REVEAL & FIRST LISTEN ──
  let revealBody = bullets([
    { label: 'Reveal plan + first song', value: c.reveal_plan },
    { label: 'Hoped-for reaction',       value: c.reveal_reaction },
  ]);
  if (hasCustomer) {
    // Pull the relevant sensory answer based on what services were checked
    const sensoryBits = [];
    if (cust.sensory_audio)     sensoryBits.push({ label: 'Customer\'s first song',           value: cust.sensory_audio });
    if (cust.sensory_carplay)   sensoryBits.push({ label: 'Customer\'s first CarPlay moment', value: cust.sensory_carplay });
    if (cust.sensory_oem)       sensoryBits.push({ label: 'Why factory look matters',         value: cust.sensory_oem });
    if (cust.sensory_remote)    sensoryBits.push({ label: 'The morning this changes',         value: cust.sensory_remote });
    if (cust.sensory_heated)    sensoryBits.push({ label: 'The first heated-seat moment',     value: cust.sensory_heated });
    if (cust.sensory_dampening) sensoryBits.push({ label: 'The sound to silence',             value: cust.sensory_dampening });
    if (cust.sensory_radar)     sensoryBits.push({ label: 'Where radar matters',              value: cust.sensory_radar });
    if (cust.sensory_safety)    sensoryBits.push({ label: 'When they wished they had safety', value: cust.sensory_safety });
    if (cust.sensory_wrangler)  sensoryBits.push({ label: 'Wrangler driving + build',         value: cust.sensory_wrangler });
    if (cust.sensory_other)     sensoryBits.push({ label: 'Other service + payoff',           value: cust.sensory_other });

    const custLines = bullets(sensoryBits);
    if (custLines) revealBody += '\n\n— FROM CUSTOMER —\n' + custLines;

    if (cust.on_camera) {
      revealBody += '\n\nOn camera for reveal? ' + cust.on_camera;
    }
  }
  sections.push({
    header: '§8 · REVEAL & FIRST LISTEN  ·  14:00–16:00',
    color: C.orange,
    bgTint: C.section_a,
    body: revealBody || '[ — reveal plan not yet filled in — ]',
  });

  // ── §9 OUTRO ──
  sections.push({
    header: '§9 · OUTRO + CTAs  ·  16:00–17:00',
    color: C.blue,
    bgTint: C.section_b,
    body: bullets([
      { label: 'Callback to hook + lesson', value: c.outro_callback },
      { label: 'Blog CTA + next video tease', value: c.blog_cta_and_tease },
    ]) || '[ — outro not yet filled in — ]',
  });

  return sections;
}


function buildReviewBreakdown(c) {
  const sections = [];

  sections.push({
    header: '◆ VIDEO META',
    color: C.grey_mid,
    bgTint: C.row_white,
    body: bullets([
      { label: 'Video Name', value: c.video_name },
      { label: 'Format',     value: 'Review / How-To' },
      { label: 'Install Window', value: (c.install_start || '?') + ' → ' + (c.install_end || '?') },
    ])
  });

  sections.push({
    header: '§1 · HOOK + VERDICT  ·  0:00–0:20',
    color: C.orange,
    bgTint: C.section_a,
    body: bullets([
      { label: 'The question this video answers', value: c.review_question },
      { label: 'Verdict (lead the hook with this)', value: c.verdict },
    ]) || '[ — hook / verdict not yet filled in — ]',
  });

  sections.push({
    header: '§2 · WHY VIEWERS ARE ASKING  ·  0:20–0:45',
    color: C.purple,
    bgTint: C.section_b,
    body: emptyMark(c.why_now),
  });

  sections.push({
    header: '§3 · THE PRODUCT  ·  0:45–2:00',
    color: C.amber,
    bgTint: C.section_a,
    body: bullets([
      { label: 'Product being reviewed', value: c.product_being_reviewed },
      { label: 'Manufacturer\'s claim',  value: c.mfr_claim },
      { label: 'Starting hypothesis',    value: c.hypothesis },
    ]) || '[ — product not yet filled in — ]',
  });

  sections.push({
    header: '§4 · THE TEST  ·  2:00–6:00',
    color: C.blue,
    bgTint: C.section_b,
    body: bullets([
      { label: 'Test plan',           value: c.test_plan },
      { label: 'Vehicle / setup',     value: c.test_vehicle },
      { label: 'Comparison target',   value: c.comparison_target },
      { label: 'Controlled variable', value: c.controlled_var },
    ]) || '[ — test plan not yet filled in — ]',
  });

  sections.push({
    header: '§5 · WHAT\'S EXPECTED  ·  during test',
    color: C.purple,
    bgTint: C.section_b,
    body: bullets([
      { label: 'Most likely surprise',     value: c.expect_surprise },
      { label: 'Where hypothesis flips',   value: c.hypothesis_flip },
      { label: 'What might fail',          value: c.might_fail },
      { label: 'John\'s likely reaction',  value: c.johns_reaction },
    ]) || '[ — expectations not yet filled in — ]',
  });

  sections.push({
    header: '§6 · VERDICT EXPANDED  ·  6:00–8:00',
    color: C.green2,
    bgTint: C.section_c,
    body: bullets([
      { label: 'Verdict',                  value: c.verdict },
      { label: 'Who FOR / who NOT for',    value: c.who_for },
      { label: 'What would flip the answer', value: c.verdict_flip },
    ]) || '[ — verdict not yet filled in — ]',
  });

  sections.push({
    header: '§7 · THE COMPARISON  ·  8:00–10:00',
    color: C.amber,
    bgTint: C.section_a,
    body: bullets([
      { label: 'Alternatives + why this wins/loses', value: c.compare_alternatives },
      { label: 'Is the price honest',                value: c.price_honest },
    ]) || '[ — comparison not yet filled in — ]',
  });

  if (c.common_mistakes && c.common_mistakes.toString().trim()) {
    sections.push({
      header: '§8 · COMMON MISTAKES  ·  10:00–11:00',
      color: C.orange,
      bgTint: C.section_a,
      body: c.common_mistakes,
    });
  }

  sections.push({
    header: '§9 · OUTRO + CTAs',
    color: C.blue,
    bgTint: C.section_b,
    body: bullets([
      { label: 'Engagement CTA + next video tease', value: c.review_cta },
      { label: 'Blog CTA',                          value: c.review_blog_cta },
    ]) || '[ — outro not yet filled in — ]',
  });

  return sections;
}


// ═══════════════════════════════════════════════════════════════════════════
// CLIENT INTAKE SECTION DEFINITIONS (the structured field-by-field display)
// ═══════════════════════════════════════════════════════════════════════════
function clientVlogSections() {
  return [
    { label: '◆ START', color: C.blue, rows: [
      ['Video Name',                  'video_name'],
      ['Video Format',                'video_format'],
      ['Install Start Date',          'install_start'],
      ['Install End Date',            'install_end'],
    ]},
    { label: '◆ §1 — THE HOOK', color: C.orange, rows: [
      ['The hook (cliffhanger line)', 'hook'],
    ]},
    { label: '◆ §2 — COLD OPEN', color: C.purple, rows: [
      ['Client + history',            'client_context'],
      ['Car role in their life',      'car_role'],
      ['What client said (quote)',    'client_said'],
    ]},
    { label: '◆ §3 — THE CAR', color: C.orange, rows: [
      ['Year',                        'year'],
      ['Make',                        'make'],
      ['Model + Trim',                'model'],
      ['Color',                       'color'],
      ['What makes the car special',  'vehicle_special'],
      ['Factory audio situation',     'factory_audio'],
    ]},
    { label: '◆ §4 — CLIENT BRIEF', color: C.blue, rows: [
      ['What client wants (feeling)', 'client_wants'],
      ['Emotional stakes',            'client_stakes'],
    ]},
    { label: '◆ §5 — PRODUCTS + THE PLAN', color: C.amber, rows: [
      ['Products list',               'products_list'],
      ['Why these products',          'products_why'],
      ['Alternatives passed on',      'alternatives'],
    ]},
    { label: '◆ §6 — THE CHALLENGE', color: C.purple, rows: [
      ['Unique challenge of this build', 'challenge'],
      ['What\'s at risk',                'at_risk'],
    ]},
    { label: '◆ §7 — INSTALL PLAN', color: C.green2, rows: [
      ['Day-by-day plan',             'install_plan'],
      ['Named tools (close-ups)',     'install_tools'],
      ['Hacks / tips / How-Tos',      'install_hacks'],
    ]},
    { label: '◆ §8 — REVEAL & FIRST LISTEN', color: C.orange, rows: [
      ['Reveal plan + first song',    'reveal_plan'],
      ['Hoped-for reaction',          'reveal_reaction'],
    ]},
    { label: '◆ §9 — OUTRO', color: C.blue, rows: [
      ['Callback to hook + lesson',   'outro_callback'],
      ['Blog CTA + next video tease', 'blog_cta_and_tease'],
    ]},
  ];
}

function clientReviewSections() {
  return [
    { label: '◆ START', color: C.blue, rows: [
      ['Video Name',                  'video_name'],
      ['Video Format',                'video_format'],
      ['Install Start Date',          'install_start'],
      ['Install End Date',            'install_end'],
    ]},
    { label: '◆ §1 — THE QUESTION', color: C.orange, rows: [
      ['Question this video answers', 'review_question'],
      ['Why viewers ask this now',    'why_now'],
    ]},
    { label: '◆ §2 — THE PRODUCT', color: C.amber, rows: [
      ['Product being reviewed',      'product_being_reviewed'],
      ['Manufacturer\'s claim',       'mfr_claim'],
      ['Starting hypothesis',         'hypothesis'],
    ]},
    { label: '◆ §3 — THE TEST PLAN', color: C.blue, rows: [
      ['Test plan',                   'test_plan'],
      ['Vehicle / setup',             'test_vehicle'],
      ['Comparison target',           'comparison_target'],
      ['Controlled variable',         'controlled_var'],
    ]},
    { label: '◆ §4 — WHAT YOU\'RE EXPECTING', color: C.purple, rows: [
      ['Most likely surprise',        'expect_surprise'],
      ['Where hypothesis could flip', 'hypothesis_flip'],
      ['What might fail',             'might_fail'],
      ['John\'s reaction',            'johns_reaction'],
    ]},
    { label: '◆ §5 — THE VERDICT', color: C.green2, rows: [
      ['Gut call',                    'verdict'],
      ['Who FOR / NOT for',           'who_for'],
      ['What would change answer',    'verdict_flip'],
    ]},
    { label: '◆ §6 — THE COMPARISON', color: C.amber, rows: [
      ['Alternatives + win/lose',     'compare_alternatives'],
      ['Is the price honest',         'price_honest'],
    ]},
    { label: '◆ §7 — COMMON MISTAKES', color: C.orange, rows: [
      ['DIY mistakes to warn against', 'common_mistakes'],
    ]},
    { label: '◆ §8 — CTAs', color: C.blue, rows: [
      ['Engagement CTA + tease',      'review_cta'],
      ['Blog CTA reason',             'review_blog_cta'],
    ]},
  ];
}


// ═══════════════════════════════════════════════════════════════════════════
// SHOT LIST + DAILY CHECKLIST ROUTING (existing behavior, untouched)
// ═══════════════════════════════════════════════════════════════════════════
function writeToBuildLog(ss, p) {
  if (p.source === 'CHECKLIST') {
    return writeToDailyNotes(ss, p);
  } else {
    return writeToShotLog(ss, p);
  }
}

function writeToDailyNotes(ss, p) {
  const TAB = TAB_NOTES;
  let sheet = ss.getSheetByName(TAB);
  if (!sheet) {
    sheet = ss.insertSheet(TAB);
    sheet.setTabColor('#1A6B5A');
    buildDailyNotesTab(sheet);
  }

  if (sheet.getLastRow() === 2 &&
      String(sheet.getRange(2, 1).getValue()).includes('Waiting')) {
    sheet.deleteRow(2);
  }

  const ts       = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
  const shootDay = p.shoot_day || p.day || '—';
  const dayName  = p.day_name  || '';
  const dayLabel = (dayName && dayName !== shootDay) ? shootDay + ' — ' + dayName : shootDay;

  sheet.appendRow([ts, p.vehicle || '', dayLabel, p.completed || '', p.missed || '', p.note || '']);

  const lastRow = sheet.getLastRow();
  const r = sheet.getRange(lastRow, 1, 1, 6);
  r.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('top').setWrap(true);
  r.setBackground(lastRow % 2 === 0 ? '#F0FBF0' : '#FFFFFF');
  sheet.setRowHeight(lastRow, 80);
  if (p.note && p.note.trim()) sheet.getRange(lastRow, 6).setBackground('#FFF9C4');

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, tab: TAB, row: lastRow }))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildDailyNotesTab(sheet) {
  sheet.setColumnWidth(1, 140); sheet.setColumnWidth(2, 160); sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 280); sheet.setColumnWidth(5, 220); sheet.setColumnWidth(6, 360);
  sheet.setFrozenRows(1);

  sheet.getRange(1, 1, 1, 6)
    .setValues([['DATE', 'VEHICLE', 'DAY', 'COMPLETED', 'MISSED / NOT DONE', "ANNIE'S NOTE"]])
    .setBackground('#111318').setFontColor('#4E9E47').setFontWeight('bold')
    .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  sheet.getRange(2, 1, 1, 6)
    .setValues([['Waiting for first checklist submission', '—', '—', '—', '—', '—']])
    .setFontColor('#525460').setFontStyle('italic').setFontFamily('Arial').setFontSize(10);
}

function writeToShotLog(ss, p) {
  const TAB = TAB_SHOTLOG;
  let sheet = ss.getSheetByName(TAB);
  if (!sheet) {
    sheet = ss.insertSheet(TAB);
    sheet.setTabColor(C.blue);
    buildShotLogTab(sheet);
  }

  if (sheet.getLastRow() === 2 &&
      String(sheet.getRange(2, 1).getValue()).includes('Waiting')) {
    sheet.deleteRow(2);
  }

  const ts       = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
  const shootDay = p.shoot_day || p.day || '—';
  const dayName  = p.day_name  || '';
  const dayLabel = (dayName && dayName !== shootDay) ? shootDay + ' — ' + dayName : shootDay;

  sheet.appendRow([ts, p.vehicle || '', dayLabel, p.completed || '', p.missed || '', p.note || '']);

  const lastRow = sheet.getLastRow();
  const r = sheet.getRange(lastRow, 1, 1, 6);
  r.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('top').setWrap(true);
  r.setBackground(lastRow % 2 === 0 ? '#EAF4FB' : '#FFFFFF');
  sheet.setRowHeight(lastRow, 80);
  if (p.note && p.note.trim()) sheet.getRange(lastRow, 6).setBackground('#FFF9C4');

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, tab: TAB, row: lastRow }))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildShotLogTab(sheet) {
  sheet.setColumnWidth(1, 140); sheet.setColumnWidth(2, 160); sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 360); sheet.setColumnWidth(5, 220); sheet.setColumnWidth(6, 300);
  sheet.setFrozenRows(1);

  sheet.getRange(1, 1, 1, 6)
    .setValues([['DATE', 'VEHICLE', 'DAY', 'SHOTS COMPLETED', 'SHOTS MISSED', "ANNIE'S NOTE"]])
    .setBackground('#111318').setFontColor('#4A7EC8').setFontWeight('bold')
    .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  sheet.getRange(2, 1, 1, 6)
    .setValues([['Waiting for first shot list EOD submission', '—', '—', '—', '—', '—']])
    .setFontColor('#525460').setFontStyle('italic').setFontFamily('Arial').setFontSize(10);
}


// ═══════════════════════════════════════════════════════════════════════════
// DRIVE — file uploads
// ═══════════════════════════════════════════════════════════════════════════
function saveCustomerVideoToNamedFolder(fileArray, submissionData) {
  if (!fileArray || !fileArray.length) return;

  const customerName = (submissionData.customer_name || 'Unnamed').toString().trim()
    .replace(/[\/\\:*?"<>|]/g, '_');

  let parent;
  try { parent = DriveApp.getFolderById(VIDEO_PARENT_FOLDER_ID); }
  catch (err) {
    console.error('Could not open parent video folder: ' + err.message);
    return;
  }

  let customerFolder;
  const existing = parent.getFoldersByName(customerName);
  if (existing.hasNext()) {
    customerFolder = existing.next();
  } else {
    customerFolder = parent.createFolder(customerName);
  }

  const dateStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  fileArray.forEach(function(file) {
    try {
      const approxBytes = (file.data || '').length * 0.75;
      if (approxBytes > 25 * 1024 * 1024) {
        console.warn('File too large, skipping: ' + file.name);
        return;
      }
      const decoded = Utilities.base64Decode(file.data);
      const blob    = Utilities.newBlob(
        decoded,
        file.type || 'application/octet-stream',
        dateStamp + '_' + (file.name || 'file')
      );
      const created = customerFolder.createFile(blob);
      console.log('Saved to Drive: ' + customerName + '/' + created.getName());
    } catch (err) {
      console.error('Error saving file "' + (file.name || 'unknown') + '": ' + err.message);
    }
  });
}

function saveFilesToDrive(fileArray, folderId, submissionData, sourceLabel) {
  if (!fileArray || !fileArray.length) return;
  let folder;
  try { folder = DriveApp.getFolderById(folderId); }
  catch (err) { console.error('Could not open Drive folder ' + folderId + ': ' + err.message); return; }

  const name = (submissionData.video_name || sourceLabel || 'intake')
    .toString().replace(/\s+/g, '-').replace(/[^a-zA-Z0-9\-_]/g, '');
  const dateStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  fileArray.forEach(function(file) {
    try {
      const approxBytes = (file.data || '').length * 0.75;
      if (approxBytes > 25 * 1024 * 1024) {
        console.warn('File too large, skipping: ' + file.name);
        return;
      }
      const decoded = Utilities.base64Decode(file.data);
      const blob    = Utilities.newBlob(
        decoded,
        file.type || 'application/octet-stream',
        name + '_' + dateStamp + '_' + (file.name || 'file')
      );
      const created = folder.createFile(blob);
      console.log('Saved to Drive: ' + created.getName());
    } catch (err) {
      console.error('Error saving file "' + (file.name || 'unknown') + '": ' + err.message);
    }
  });
}


// ═══════════════════════════════════════════════════════════════════════════
// DEPLOYMENT
//
// To update:
//   Deploy → Manage deployments → ✏️ edit the active deployment →
//   Version: New version → Deploy
//
// The URL stays the same. All forms keep pointing at it.
// ═══════════════════════════════════════════════════════════════════════════
