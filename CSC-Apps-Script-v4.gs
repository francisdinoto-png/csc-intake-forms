// ═══════════════════════════════════════════════════════════════════════════
// CSC PRODUCTION SHEET — Google Apps Script  (v4 — final)
// ═══════════════════════════════════════════════════════════════════════════
//
// WHAT THIS DOES
//
//   1. Client intake (Annie's pre-production brief) → appends a new section
//      to the 📋 CLIENT INTAKES tab. Each submission stacks below the last.
//      TOC at the top auto-updates with a jump link to each build.
//
//   2. Customer intake (car owner) → appends a new section to the 🎧 CUSTOMER
//      INTAKES tab. Same TOC pattern. Optional video upload goes into a
//      Drive folder named after the customer's first name.
//
//   3. Shot list app + daily checklist (existing) → appends to 🎬 SHOT LOG
//      and 📓 DAILY NOTES as it does today. Untouched.
//
// FINAL TAB LAYOUT
//   📋 CLIENT INTAKES   — TOC + stacked client intake briefs
//   🎧 CUSTOMER INTAKES — TOC + stacked customer intakes
//   📓 DAILY NOTES      — row per daily checklist
//   🎬 SHOT LOG         — row per shot list EOD
//
// SAFETY
//   - This script does NOT delete any tabs.
//   - setupCSCProduction() only creates tabs if missing. It will not wipe data.
//   - Every submission appends; nothing overwrites prior submissions.
//
// SETUP
//   Extensions → Apps Script → paste this whole file → save.
//   Optionally run setupCSCProduction() once to create the four tabs.
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
  row_alt:    '#F8F8F8',
  row_white:  '#FFFFFF',
  toc_bg:     '#FFF8E1',
  divider:    '#222222',
};

// ── DRIVE FOLDER — where customer video folders get created ─────────────────
// This is the parent folder. Each customer's video lands in a subfolder
// named after their first name (e.g. "Marcus", "Howard", etc.).
const VIDEO_PARENT_FOLDER_ID = '1etjiS0yORaqtgaUXW6_dQyFLYdHKWqG0';
const AUDIO_FOLDER_ID        = '1RCscdw4F2k5H19IzYbV9bcqH1hr-zks9'; // Annie's audio brain-dumps

// ── TAB NAMES ───────────────────────────────────────────────────────────────
const TAB_CLIENT   = '📋 CLIENT INTAKES';
const TAB_CUSTOMER = '🎧 CUSTOMER INTAKES';
const TAB_NOTES    = '📓 DAILY NOTES';
const TAB_SHOTLOG  = '🎬 SHOT LOG';


// ═══════════════════════════════════════════════════════════════════════════
// SETUP — run once. Safe. Only creates tabs if missing.
// ═══════════════════════════════════════════════════════════════════════════
function setupCSCProduction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureClientIntakesTab(ss);
  ensureCustomerIntakesTab(ss);

  // Log tabs (unchanged from previous version)
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
    'Four tabs are ready:\n' +
    '  📋 CLIENT INTAKES — Annie\'s pre-production briefs\n' +
    '  🎧 CUSTOMER INTAKES — customer story submissions\n' +
    '  📓 DAILY NOTES — daily checklist log\n' +
    '  🎬 SHOT LOG — shot list EOD submissions\n\n' +
    'Each new submission appends below the previous one. Nothing overwrites.'
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
      appendCustomerIntake(ss, data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, tab: TAB_CUSTOMER }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      // default: client_intake
      appendClientIntake(ss, data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, tab: TAB_CLIENT }))
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
// CLIENT INTAKE TAB — TOC + stacked briefs
// ═══════════════════════════════════════════════════════════════════════════
function ensureClientIntakesTab(ss) {
  let sheet = ss.getSheetByName(TAB_CLIENT);
  if (sheet) return sheet;

  sheet = ss.insertSheet(TAB_CLIENT);
  sheet.setTabColor(C.orange);
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 700);

  // Title row
  sheet.getRange(1, 1, 1, 2).merge()
    .setValue('CLIENT INTAKES — PRE-PRODUCTION BRIEFS')
    .setBackground(C.ink).setFontColor(C.white).setFontSize(16)
    .setFontWeight('bold').setFontFamily('Arial').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 48);

  // TOC header
  sheet.getRange(2, 1, 1, 2).merge()
    .setValue('INDEX  ·  click a build to jump to its brief')
    .setBackground(C.header_bg).setFontColor(C.amber).setFontSize(10)
    .setFontWeight('bold').setFontFamily('Arial').setFontStyle('italic');
  sheet.setRowHeight(2, 28);

  // Empty TOC body — submissions add lines below row 2
  // First brief starts after a divider (auto-managed below)

  sheet.setFrozenRows(2);
  return sheet;
}

function appendClientIntake(ss, data) {
  const sheet = ensureClientIntakesTab(ss);

  // ── Compute label for TOC ──
  const videoName = (data.video_name || '').toString().trim() || 'Untitled Build';
  const format    = data.video_format === 'review' ? 'Review / How-To' : 'Vlog Style';
  const ts        = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy');
  const tocLabel  = videoName + '  ·  ' + format + '  ·  ' + ts;

  // ── Find where to insert: end of sheet, but leave the TOC section intact ──
  // TOC lives between row 3 and the first divider row (or end of file)
  // We append to the bottom of the sheet, then add a TOC entry at row 3 (newest first)

  // Find the bottom of the sheet to write the new brief
  const startRow = Math.max(sheet.getLastRow(), 2) + 2; // 2-row gap before new brief

  // ── Write the brief block ──
  let row = startRow;

  // Brief header (the anchor for the TOC link)
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('▼ ' + videoName + '  ·  ' + format + '  ·  ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy · h:mm a'))
    .setBackground(C.orange).setFontColor(C.white).setFontSize(14)
    .setFontWeight('bold').setFontFamily('Arial').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 36);
  const briefAnchorRow = row;
  row++;

  // Sub-row: format + timestamp
  // (already in header — skip)

  // Sections — vlog vs review
  const isVlog = (data.video_format !== 'review');
  const sections = isVlog ? clientVlogSections() : clientReviewSections();

  // FIELD / VALUE column headers
  ['FIELD', 'VALUE'].forEach(function(h, i) {
    sheet.getRange(row, i + 1).setValue(h)
      .setBackground(C.orange).setFontColor(C.white).setFontSize(9)
      .setFontWeight('bold').setFontFamily('Arial').setHorizontalAlignment('left');
  });
  sheet.setRowHeight(row, 28);
  row++;

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

      // Tall row for long answers
      const lines = (value.match(/\n/g) || []).length + 1;
      const wordWrapLines = Math.ceil(value.length / 90);
      const targetLines = Math.max(lines, wordWrapLines);
      sheet.setRowHeight(row, Math.max(36, Math.min(targetLines * 16, 400)));
      row++;
    });
  });

  // End-of-brief divider
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('— end of brief —')
    .setBackground(C.divider).setFontColor(C.grey_mid).setFontSize(9)
    .setFontStyle('italic').setFontFamily('Arial').setHorizontalAlignment('center');
  sheet.setRowHeight(row, 22);

  // ── Add TOC entry at the top (newest first) ──
  sheet.insertRowAfter(2);
  sheet.getRange(3, 1, 1, 2).merge()
    .setBackground(C.toc_bg)
    .setFontFamily('Arial').setFontSize(11).setVerticalAlignment('middle');
  sheet.setRowHeight(3, 24);

  // Build clickable rich-text link to the brief anchor row
  // Anchor moves down by 1 because we just inserted a row above it
  const finalAnchorRow = briefAnchorRow + 1;
  const sheetId = sheet.getSheetId();
  const linkText = '→ ' + tocLabel;
  const richText = SpreadsheetApp.newRichTextValue()
    .setText(linkText)
    .setLinkUrl('#gid=' + sheetId + '&range=A' + finalAnchorRow)
    .setTextStyle(SpreadsheetApp.newTextStyle()
      .setForegroundColor(C.blue)
      .setFontFamily('Arial')
      .setFontSize(11)
      .build())
    .build();
  sheet.getRange(3, 1).setRichTextValue(richText);
}


function clientVlogSections() {
  return [
    { label: '◆ START', color: C.blue, rows: [
      ['Video Name',                  'video_name'],
      ['Video Format',                'video_format'],
      ['Install Start Date',          'install_start'],
      ['Install End Date',            'install_end'],
    ]},
    { label: '◆ §2 — THE HOOK', color: C.orange, rows: [
      ['The hook (cliffhanger line)', 'hook'],
    ]},
    { label: '◆ §3 — COLD OPEN', color: C.purple, rows: [
      ['Client + history',            'client_context'],
      ['Car role in their life',      'car_role'],
      ['What client said (quote)',    'client_said'],
    ]},
    { label: '◆ §4 — THE CAR', color: C.orange, rows: [
      ['Year',                        'year'],
      ['Make',                        'make'],
      ['Model + Trim',                'model'],
      ['Color',                       'color'],
      ['What makes the car special',  'vehicle_special'],
      ['Factory audio situation',     'factory_audio'],
    ]},
    { label: '◆ §5 — CLIENT BRIEF', color: C.blue, rows: [
      ['What client wants (feeling)', 'client_wants'],
      ['Emotional stakes',            'client_stakes'],
    ]},
    { label: '◆ §6 — PRODUCTS + THE PLAN', color: C.amber, rows: [
      ['Products list',               'products_list'],
      ['Why these products',          'products_why'],
      ['Alternatives passed on',      'alternatives'],
    ]},
    { label: '◆ §7 — THE CHALLENGE', color: C.purple, rows: [
      ['Unique challenge of this build', 'challenge'],
      ['What\'s at risk',                'at_risk'],
    ]},
    { label: '◆ §8 — INSTALL PLAN', color: C.green2, rows: [
      ['Day-by-day plan',             'install_plan'],
      ['Named tools (close-ups)',     'install_tools'],
      ['Hacks / tips / How-Tos',      'install_hacks'],
    ]},
    { label: '◆ §9 — REVEAL & FIRST LISTEN', color: C.orange, rows: [
      ['Reveal plan + first song',    'reveal_plan'],
      ['Hoped-for reaction',          'reveal_reaction'],
    ]},
    { label: '◆ §10 — OUTRO', color: C.blue, rows: [
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
    { label: '◆ §2 — THE QUESTION', color: C.orange, rows: [
      ['Question this video answers', 'review_question'],
      ['Why viewers ask this now',    'why_now'],
    ]},
    { label: '◆ §3 — THE PRODUCT', color: C.amber, rows: [
      ['Product being reviewed',      'product_being_reviewed'],
      ['Manufacturer\'s claim',       'mfr_claim'],
      ['Starting hypothesis',         'hypothesis'],
    ]},
    { label: '◆ §4 — THE TEST PLAN', color: C.blue, rows: [
      ['Test plan',                   'test_plan'],
      ['Vehicle / setup',             'test_vehicle'],
      ['Comparison target',           'comparison_target'],
      ['Controlled variable',         'controlled_var'],
    ]},
    { label: '◆ §5 — WHAT YOU\'RE EXPECTING', color: C.purple, rows: [
      ['Most likely surprise',        'expect_surprise'],
      ['Where hypothesis could flip', 'hypothesis_flip'],
      ['What might fail',             'might_fail'],
      ['John\'s reaction',            'johns_reaction'],
    ]},
    { label: '◆ §6 — THE VERDICT', color: C.green2, rows: [
      ['Gut call',                    'verdict'],
      ['Who FOR / NOT for',           'who_for'],
      ['What would change answer',    'verdict_flip'],
    ]},
    { label: '◆ §7 — THE COMPARISON', color: C.amber, rows: [
      ['Alternatives + win/lose',     'compare_alternatives'],
      ['Is the price honest',         'price_honest'],
    ]},
    { label: '◆ §8 — COMMON MISTAKES', color: C.orange, rows: [
      ['DIY mistakes to warn against', 'common_mistakes'],
    ]},
    { label: '◆ §9 — CTAs', color: C.blue, rows: [
      ['Engagement CTA + tease',      'review_cta'],
      ['Blog CTA reason',             'review_blog_cta'],
    ]},
  ];
}


// ═══════════════════════════════════════════════════════════════════════════
// CUSTOMER INTAKE TAB — TOC + stacked sections
// ═══════════════════════════════════════════════════════════════════════════
function ensureCustomerIntakesTab(ss) {
  let sheet = ss.getSheetByName(TAB_CUSTOMER);
  if (sheet) return sheet;

  sheet = ss.insertSheet(TAB_CUSTOMER);
  sheet.setTabColor(C.green2);
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 700);

  sheet.getRange(1, 1, 1, 2).merge()
    .setValue('CUSTOMER INTAKES')
    .setBackground(C.ink).setFontColor(C.white).setFontSize(16)
    .setFontWeight('bold').setFontFamily('Arial').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 48);

  sheet.getRange(2, 1, 1, 2).merge()
    .setValue('INDEX  ·  click a customer to jump to their submission')
    .setBackground(C.header_bg).setFontColor(C.amber).setFontSize(10)
    .setFontWeight('bold').setFontFamily('Arial').setFontStyle('italic');
  sheet.setRowHeight(2, 28);

  sheet.setFrozenRows(2);
  return sheet;
}

function appendCustomerIntake(ss, data) {
  const sheet = ensureCustomerIntakesTab(ss);

  const customerName = (data.customer_name || 'Unnamed Customer').toString().trim();
  const services = Array.isArray(data.services) ? data.services.join(', ') : (data.services || '');
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy');
  const tocLabel = customerName + '  ·  ' + ts + (services ? '  ·  ' + services : '');

  const startRow = Math.max(sheet.getLastRow(), 2) + 2;
  let row = startRow;

  // Brief header
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('▼ ' + customerName + '  ·  ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy · h:mm a'))
    .setBackground(C.green2).setFontColor(C.white).setFontSize(14)
    .setFontWeight('bold').setFontFamily('Arial').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 36);
  const anchorRow = row;
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

  ['FIELD', 'VALUE'].forEach(function(h, i) {
    sheet.getRange(row, i + 1).setValue(h)
      .setBackground(C.green2).setFontColor(C.white).setFontSize(9)
      .setFontWeight('bold').setFontFamily('Arial').setHorizontalAlignment('left');
  });
  sheet.setRowHeight(row, 28);
  row++;

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
    { label: '◆ META', color: C.grey_mid, rows: [
      ['Submitted at',                     'submitted_at'],
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

      // For sensory section, skip empty (only show the ones answered)
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

  // Divider
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('— end of submission —')
    .setBackground(C.divider).setFontColor(C.grey_mid).setFontSize(9)
    .setFontStyle('italic').setFontFamily('Arial').setHorizontalAlignment('center');
  sheet.setRowHeight(row, 22);

  // Add TOC entry at the top
  sheet.insertRowAfter(2);
  sheet.getRange(3, 1, 1, 2).merge()
    .setBackground(C.toc_bg)
    .setFontFamily('Arial').setFontSize(11).setVerticalAlignment('middle');
  sheet.setRowHeight(3, 24);

  const finalAnchorRow = anchorRow + 1;
  const sheetId = sheet.getSheetId();
  const linkText = '→ ' + tocLabel;
  const richText = SpreadsheetApp.newRichTextValue()
    .setText(linkText)
    .setLinkUrl('#gid=' + sheetId + '&range=A' + finalAnchorRow)
    .setTextStyle(SpreadsheetApp.newTextStyle()
      .setForegroundColor(C.blue)
      .setFontFamily('Arial')
      .setFontSize(11)
      .build())
    .build();
  sheet.getRange(3, 1).setRichTextValue(richText);
}


// ═══════════════════════════════════════════════════════════════════════════
// SHOT LIST EOD + DAILY CHECKLIST ROUTING  (existing behavior preserved)
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
// DRIVE — customer video to a folder named after them
// ═══════════════════════════════════════════════════════════════════════════
function saveCustomerVideoToNamedFolder(fileArray, submissionData) {
  if (!fileArray || !fileArray.length) return;

  // Get/create a subfolder named after the customer
  const customerName = (submissionData.customer_name || 'Unnamed').toString().trim()
    .replace(/[\/\\:*?"<>|]/g, '_'); // strip illegal folder chars

  let parent;
  try { parent = DriveApp.getFolderById(VIDEO_PARENT_FOLDER_ID); }
  catch (err) {
    console.error('Could not open parent video folder: ' + err.message);
    return;
  }

  // Find or create subfolder
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

// Annie's audio brain-dumps — go to the audio folder, flat
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
// DEPLOYMENT NOTES
//
// To update:
//   Deploy → Manage deployments → ✏️ edit the active deployment →
//   Version: New version → Deploy
//
// The URL stays the same. All forms keep pointing at it.
// ═══════════════════════════════════════════════════════════════════════════
