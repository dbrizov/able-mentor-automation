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
        appendRows(job.sheet, job.rows);
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
      byCity[city][tab].push({ email, keys: src.keys, row });
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
      const target = readTab(sheet, `${city}/${tab}`);
      const iEmail = target.col(EMAIL_COLUMN);
      const existing = new Set(target.rows.map(r => normEmail(r[iEmail])).filter(Boolean));

      const people = byCity[city][tab];
      const fresh = people.filter(p => !existing.has(p.email));
      if (fresh.length) {
        jobs.push({ city, tab, sheet, rows: fresh.map(p => mapRow(p, target.keys)) });
        totalNew += fresh.length;
      }
      parts.push(`${tab}: ${fresh.length} new, ${people.length - fresh.length} already there`);
    });
    lines.push(`• ${city}: ${parts.join("; ")}`);
  });

  return { lines: lines.concat(notes.length ? [""].concat(notes) : []), jobs, totalNew };
}

// Puts a source row into the register's column order, matching by header text.
function mapRow(person, targetKeys) {
  const srcIndex = {};
  person.keys.forEach((k, i) => { if (k) srcIndex[k] = i; });
  return targetKeys.map(k => (k && k in srcIndex) ? person.row[srcIndex[k]] : "");
}

function appendRows(sheet, rows) {
  const start = sheet.getLastRow() + 1;
  const width = rows[0].length;
  const needRows = start + rows.length - 1 - sheet.getMaxRows();
  if (needRows > 0) sheet.insertRowsAfter(sheet.getMaxRows(), needRows);
  const needCols = width - sheet.getMaxColumns();
  if (needCols > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), needCols);
  sheet.getRange(start, 1, rows.length, width).setValues(rows);
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

// Reads a tab: header keys + non-empty data rows.
// Header keys are "text#n" so repeated headers (the students tab has one twice) stay distinct.
function readTab(sheet, label) {
  if (!sheet) throw new Error(`Sheet "${label}" not found`);
  const values = sheet.getDataRange().getValues();
  const keys = headerKeys(values.shift() || []);
  const rows = [], rowNums = [];
  values.forEach((r, i) => {
    if (r.some(v => v !== "" && v !== null)) { rows.push(r); rowNums.push(i + 2); }  // +2: header + 1-based
  });
  const col = name => {
    const i = keys.indexOf(`${name}#1`);
    if (i < 0) throw new Error(`Missing "${name}" column in "${label}"`);
    return i;
  };
  return { keys, rows, rowNums, col };
}

function headerKeys(header) {
  const seen = {};
  return header.map(h => {
    h = String(h).trim();
    if (!h) return "";
    seen[h] = (seen[h] || 0) + 1;
    return `${h}#${seen[h]}`;
  });
}

function normEmail(v) {
  return String(v || "").trim().toLowerCase();
}