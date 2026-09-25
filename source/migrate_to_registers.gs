const STUDENTS_SHEET = "ученици";
const MENTORS_SHEET = "ментори";
const SETTINGS_SHEET = "Settings";
const CITY_COLUMN = "В кой град предпочитате";
const ONLINE = "Онлайн";                        // register used when the city is empty
const EMAIL_COLUMN = "Email";                   // unique ID of a person

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ABLE Tools")
    //.addItem("Set up registers", "setupRegisters")
    .addItem("Migrate to registers", "migrateToRegisters")
    .addToUi();
}

// ---------- Settings ----------

// Returns { "София": "1AbC...", ... }. Accepts a bare ID or a full spreadsheet URL.
function readSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
  if (!sheet) return {};

  const map = {};
  sheet.getDataRange().getValues().slice(1).forEach(([city, idOrUrl]) => {
    city = String(city).trim();
    idOrUrl = String(idOrUrl).trim();
    if (!city) return;
    const m = idOrUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    map[city] = m ? m[1] : idOrUrl;
  });
  return map;
}

// ---------- Migration ----------

// Copies every student and mentor row into the register for their city.
// - Empty city -> the "Онлайн" register.
// - Email is the unique ID: a person whose email is already in the register's tab is not added again.
// - Nothing is written until you confirm the summary.
function migrateToRegisters() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    ui.alert("Another migration is already running. Try again in a minute.");
    return;
  }

  try {
    const plan = buildPlan();
    const summary = plan.lines.join("\n");
    console.log(summary);

    if (plan.totalNew === 0) {
      ui.alert("Migrate to registers", summary + "\n\nNothing new to add.", ui.ButtonSet.OK);
      return;
    }

    // Plain server-side dialog: OK = migrate, Cancel / X = write nothing.
    const answer = ui.alert(
      `Migrate ${plan.totalNew} new rows?`,
      summary + "\n\nClick OK to migrate or Cancel to stop without writing anything.",
      ui.ButtonSet.OK_CANCEL);
    if (answer !== ui.Button.OK) {
      SpreadsheetApp.getActiveSpreadsheet().toast("Migration cancelled. Nothing was written.");
      return;
    }

    const results = plan.jobs.map(job => {
      try {
        appendRows(job);
        return `✓ ${job.city} / ${job.tab}: added ${job.rows.length}`;
      } catch (e) {
        return `✗ ${job.city} / ${job.tab}: ${e.message}`;
      }
    });

    const text = results.join("\n");
    console.log(text);
    ui.alert("Migration finished", text, ui.ButtonSet.OK);
  } finally {
    lock.releaseLock();
  }
}

// Reads the source tabs and every register, and works out which rows are new.
// Writes nothing.
function buildPlan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = readSettings();

  // city -> tab -> [{ email, keys, row }]
  const byCity = {};
  const notes = [];

  [STUDENTS_SHEET, MENTORS_SHEET].forEach(tab => {
    const src = readTab(ss.getSheetByName(tab), tab);
    const iEmail = src.col(EMAIL_COLUMN);
    const iCity = src.col(CITY_COLUMN);
    const firstRow = {};      // email -> row number of the first submission
    const dupes = {};         // email -> [row numbers of skipped repeats]
    const noEmail = [];

    src.rows.forEach((row, i) => {
      const rowNum = src.rowNums[i];
      const email = normEmail(row[iEmail]);
      if (!email) { noEmail.push(rowNum); return; }
      if (email in firstRow) {                    // same email twice in the form: keep the first
        (dupes[email] = dupes[email] || []).push(rowNum);
        return;
      }
      firstRow[email] = rowNum;

      const city = String(row[iCity]).trim() || ONLINE;
      if (!byCity[city]) byCity[city] = {};
      if (!byCity[city][tab]) byCity[city][tab] = [];
      byCity[city][tab].push({ email, keys: src.keys, labels: src.labels, row });
    });

    if (noEmail.length) {
      notes.push(`⚠ ${tab}: ${noEmail.length} rows without an email were skipped (rows ${noEmail.join(", ")})`);
    }
    const dupEmails = Object.keys(dupes);
    if (dupEmails.length) {
      const skipped = dupEmails.reduce((n, e) => n + dupes[e].length, 0);
      notes.push(`⚠ ${tab}: ${dupEmails.length} emails appear more than once — kept the first row, skipped ${skipped}:`);
      dupEmails.forEach(e => notes.push(`    ${e} — kept row ${firstRow[e]}, skipped ${dupes[e].length > 1 ? "rows" : "row"} ${dupes[e].join(", ")}`));
    }
  });

  const lines = [];
  const jobs = [];
  let totalNew = 0;

  Object.keys(byCity).sort().forEach(city => {
    const id = settings[city];
    const counts = Object.keys(byCity[city]).map(tab => `${byCity[city][tab].length} ${tab}`).join(", ");
    if (!id) {
      lines.push(`✗ ${city}: no register in Settings — skipped (${counts})`);
      return;
    }

    let dst;
    try {
      dst = openRegister(id);
    } catch (e) {
      lines.push(`✗ ${city}: ${e.message} — skipped (${counts})`);
      return;
    }

    const parts = [];
    Object.keys(byCity[city]).forEach(tab => {
      const sheet = dst.getSheetByName(tab);
      if (!sheet) {
        parts.push(`${tab}: tab missing, run "Set up registers" first`);
        return;
      }

      let target;
      try {
        target = readTab(sheet, `${city}/${tab}`);
        target.col(EMAIL_COLUMN);
      } catch (e) {
        parts.push(`${tab}: ${e.message}`);
        return;
      }

      const people = byCity[city][tab];
      const map = columnMap(people[0].keys, people[0].labels, target.keys);
      if (map.missing.length) {
        notes.push(`⚠ ${city} / ${tab}: register has no column for: ${map.missing.join(" | ")} — those values are not copied`);
      }

      const iEmail = target.col(EMAIL_COLUMN);
      const existing = new Set(target.rows.map(r => normEmail(r[iEmail])).filter(Boolean));
      const fresh = people.filter(p => !existing.has(p.email));
      if (fresh.length) {
        jobs.push({
          city, tab, sheet,
          startRow: lastFilledRow(target, map.cols) + 1,
          cols: map.cols.map(c => c.dst),
          rows: fresh.map(p => map.cols.map(c => p.row[c.src])),
        });
        totalNew += fresh.length;
      }
      parts.push(`${tab}: ${fresh.length} new, ${people.length - fresh.length} already there`);
    });
    lines.push(`• ${city}: ${parts.join("; ")}`);
  });

  return { lines: lines.concat(notes.length ? [""].concat(notes) : []), jobs, totalNew };
}

// Matches form columns to register columns by header name (not position).
// Register columns with no matching form column are custom columns and are never written.
function columnMap(srcKeys, srcLabels, dstKeys) {
  const srcIndex = {};
  srcKeys.forEach((k, i) => { if (k) srcIndex[k] = i; });
  const dstKeySet = new Set(dstKeys.filter(Boolean));

  const cols = [];
  dstKeys.forEach((k, i) => { if (k && k in srcIndex) cols.push({ src: srcIndex[k], dst: i }); });
  const missing = srcKeys.map((k, i) => (k && !dstKeySet.has(k)) ? srcLabels[i] : null).filter(Boolean);
  return { cols, missing };
}

// Last row that has data in any of the form columns. Custom columns are ignored, so
// checkboxes or formulas filled down a custom column don't push new rows to the bottom.
function lastFilledRow(target, cols) {
  let last = 1;  // header
  target.rows.forEach((r, i) => {
    if (cols.some(c => r[c.dst] !== "" && r[c.dst] !== null)) last = target.rowNums[i];
  });
  return last;
}

// Writes only the form columns, one block per run of adjacent columns,
// so custom columns (notes, checkboxes, formulas) in the same rows are left untouched.
function appendRows(job) {
  const { sheet, startRow, cols, rows } = job;
  const lastRow = startRow + rows.length - 1;
  if (lastRow > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), lastRow - sheet.getMaxRows());

  let i = 0;
  while (i < cols.length) {
    let j = i;
    while (j + 1 < cols.length && cols[j + 1] === cols[j] + 1) j++;
    const block = rows.map(r => r.slice(i, j + 1));
    sheet.getRange(startRow, cols[i] + 1, rows.length, j - i + 1).setValues(block);
    i = j + 1;
  }
}

// ---------- Register setup ----------

// Creates empty "ученици" and "ментори" tabs in every register listed in Settings,
// with the same headers, column widths, formatting, frozen rows and validation as here.
// Tabs that already exist in a register are left untouched.
function setupRegisters() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = readSettings();
  const cities = Object.keys(settings).filter(c => settings[c]);
  const missing = Object.keys(settings).filter(c => !settings[c]);

  const ok = ui.alert(
    "Set up registers",
    `Create empty "${STUDENTS_SHEET}" and "${MENTORS_SHEET}" tabs in ${cities.length} registers?\n` +
    "Existing tabs with those names are not changed.",
    ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;

  // Build empty templates inside THIS spreadsheet first, so no personal data
  // is ever copied into (or left in the version history of) the registers.
  const templates = [STUDENTS_SHEET, MENTORS_SHEET].map(name => makeTemplate(ss, name));

  const results = missing.map(city => `– ${city}: skipped, no register in Settings`);
  try {
    cities.forEach(city => {
      try {
        const dst = openRegister(settings[city]);
        const done = [];
        templates.forEach(t => {
          if (dst.getSheetByName(t.name)) { done.push(`${t.name}: already exists`); return; }
          t.sheet.copyTo(dst).setName(t.name);
          done.push(`${t.name}: created`);
        });
        results.push(`✓ ${city} (${dst.getName()}): ${done.join(", ")}`);
      } catch (e) {
        results.push(`✗ ${city}: ${e.message}`);
      }
    });
  } finally {
    templates.forEach(t => ss.deleteSheet(t.sheet));
  }

  const text = results.join("\n");
  console.log(text);
  ui.alert("Set up registers", text, ui.ButtonSet.OK);
}

// Opens a register by ID, turning Google's vague errors into readable ones.
function openRegister(id) {
  if (!/^[a-zA-Z0-9-_]{25,}$/.test(id)) {
    throw new Error(`"${id}" doesn't look like a spreadsheet ID or URL`);
  }
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error(`can't open ${id}: it doesn't exist or you don't have access to it (${e.message})`);
  }
}

// Duplicates a tab and clears everything except the header row.
function makeTemplate(ss, name) {
  const src = ss.getSheetByName(name);
  if (!src) throw new Error(`Sheet "${name}" not found`);
  const copy = src.copyTo(ss).setName(`__template_${name}`);
  const lastRow = copy.getMaxRows();
  if (lastRow > 1) {
    copy.getRange(2, 1, lastRow - 1, copy.getMaxColumns()).clearContent().clearNote();
  }
  return { name, sheet: copy };
}

// ---------- Helpers ----------

// Reads a tab: header keys + non-empty data rows (with their sheet row numbers).
// Header keys are "text#n" so repeated headers (the students tab has one twice) stay distinct;
// the 2nd "X" in the form goes to the 2nd "X" column in the register.
function readTab(sheet, label) {
  if (!sheet) throw new Error(`Sheet "${label}" not found`);
  const values = sheet.getDataRange().getValues();
  const header = values.shift() || [];
  const keys = headerKeys(header);
  const labels = header.map(h => String(h).replace(/\s+/g, " ").trim());
  const rows = [], rowNums = [];
  values.forEach((r, i) => {
    if (r.some(v => v !== "" && v !== null)) { rows.push(r); rowNums.push(i + 2); }  // +2: header + 1-based
  });
  const col = name => {
    const i = keys.indexOf(`${normHeader(name)}#1`);
    if (i < 0) throw new Error(`Missing "${name}" column in "${label}"`);
    return i;
  };
  return { keys, labels, rows, rowNums, col };
}

function headerKeys(header) {
  const seen = {};
  return header.map(h => {
    h = normHeader(h);
    if (!h) return "";
    seen[h] = (seen[h] || 0) + 1;
    return `${h}#${seen[h]}`;
  });
}

// "Email", " email ", "EMAIL\n" all match. Line breaks and repeated spaces count as one space.
function normHeader(h) {
  return String(h).replace(/\s+/g, " ").trim().toLowerCase();
}

function normEmail(v) {
  return String(v || "").trim().toLowerCase();
}