// ABLE Mentor – email drafts from a register.
// Paste this whole file into Extensions → Apps Script of each register.

const TEMPLATES_SHEET = "Email Templates";
const EMAIL_COLUMN = "Email";

// Which address the drafts are sent from, per tab.
// The address must be set up as a "Send mail as" address in your Gmail.
const SENDERS = {
  "ученици": "students@ablementor.bg",
  "ментори": "mentors@ablementor.bg",
};
const FOR_ALL = "всички";        // template "For" value that works on every tab
const MAX_TEMPLATES = 30;        // how many templates the menu can show

// ---------- Menu ----------

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const send = ui.createMenu("Send email");

  const templates = readTemplates();
  if (templates === null) {
    send.addItem("(no templates yet – run Setup email templates)", "showNoTemplatesHelp");
  } else if (templates.length === 0) {
    send.addItem(`(the "${TEMPLATES_SHEET}" tab is empty)`, "showNoTemplatesHelp");
  } else {
    templates.slice(0, MAX_TEMPLATES).forEach((t, i) => send.addItem(t.name, `sendTemplate_${i}`));
  }

  const emails = ui.createMenu("Emails")
    .addItem("Setup email templates", "setupEmailTemplates")
    .addItem("Refresh templates list", "onOpen")
    .addSeparator()
    .addSubMenu(send);

  ui.createMenu("ABLE Tools")
    .addSubMenu(emails)
    .addToUi();
}

// Menu items can only call global functions by name, so create sendTemplate_0 … sendTemplate_29.
for (let i = 0; i < MAX_TEMPLATES; i++) {
  globalThis[`sendTemplate_${i}`] = () => sendTemplate(i);
}

function showNoTemplatesHelp() {
  SpreadsheetApp.getUi().alert(`Run "ABLE Tools → Emails → Setup email templates" first, add a template, then "Refresh templates list".`);
}

// ---------- Templates ----------

// Creates the "Email Templates" tab with two examples. Does nothing if the tab exists.
function setupEmailTemplates() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(TEMPLATES_SHEET)) {
    ui.alert(`The "${TEMPLATES_SHEET}" tab already exists. Nothing was changed.`);
    return;
  }

  const sheet = ss.insertSheet(TEMPLATES_SHEET);
  const rows = [
    ["Name", "For", "Subject", "Body"],
    [
      "Покана за среща",
      "ученици",
      "ABLE Mentor – покана за среща",
      "Здравей, {{Име}},\n\nБихме искали да те поканим на среща за програмата ABLE Mentor.\n\n" +
      "Поздрави,\nЕкипът на ABLE Mentor",
    ],
    [
      "Благодарим на ментора",
      "ментори",
      "Благодарим Ви, {{Име}}!",
      "Здравейте, {{Име}} {{Фамилия}},\n\nБлагодарим Ви, че се присъединихте към ABLE Mentor като ментор.\n\n" +
      "Поздрави,\nЕкипът на ABLE Mentor",
    ],
  ];
  sheet.getRange(1, 1, rows.length, 4).setValues(rows);

  sheet.setFrozenRows(1);
  sheet.getRange("A1:D1").setFontWeight("bold");
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 500);
  sheet.getRange("A:D").setVerticalAlignment("top");
  sheet.getRange("D:D").setWrap(true);

  const forRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([...Object.keys(SENDERS), FOR_ALL], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).setDataValidation(forRule);

  sheet.getRange("F1").setValue(
    "How to use:\n" +
    "• One row = one template. Name is what appears in ABLE Tools → Emails → Send email.\n" +
    `• For: ${Object.keys(SENDERS).join(" / ")} / ${FOR_ALL}.\n` +
    "• Use {{Column name}} in Subject or Body to insert that person's value, e.g. {{Име}}.\n" +
    "• New line in a cell: Ctrl+Enter (Cmd+Enter on Mac).\n" +
    "• After changes, run ABLE Tools → Emails → Refresh templates list.")
    .setWrap(true).setVerticalAlignment("top").setFontColor("#666666");
  sheet.setColumnWidth(6, 380);

  ss.setActiveSheet(sheet);
  onOpen();  // refresh the menu with the new templates
  ss.toast("Email templates tab created.");
}

// Returns [{ name, forTab, subject, body }], or null if the tab doesn't exist.
function readTemplates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEMPLATES_SHEET);
  if (!sheet) return null;
  return sheet.getDataRange().getValues().slice(1)
    .map(r => ({
      name: String(r[0]).trim(),
      forTab: String(r[1]).trim().toLowerCase(),
      subject: String(r[2]),
      body: String(r[3]),
    }))
    .filter(t => t.name);
}

// ---------- Send (create drafts) ----------

function sendTemplate(index) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const tab = sheet.getName();

  // 1. Template
  const templates = readTemplates() || [];
  const t = templates[index];
  if (!t) {
    ui.alert("That template no longer exists. Run ABLE Tools → Emails → Refresh templates list.");
    return;
  }

  // 2. Tab and sender
  const sender = SENDERS[tab];
  if (!sender) {
    ui.alert(`Select rows in the ${Object.keys(SENDERS).map(s => `"${s}"`).join(" or ")} tab first.`);
    return;
  }
  if (t.forTab && t.forTab !== FOR_ALL && t.forTab !== tab) {
    ui.alert(`"${t.name}" is a template for "${t.forTab}", but you are in "${tab}".\n` +
      `Change its "For" column to "${tab}" or "${FOR_ALL}" if it should work here.`);
    return;
  }

  // 3. People from the selected rows
  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const keys = header.map(normHeader);
  const iEmail = keys.indexOf(normHeader(EMAIL_COLUMN));
  if (iEmail < 0) {
    ui.alert(`No "${EMAIL_COLUMN}" column in "${tab}".`);
    return;
  }

  const people = [];
  const noEmail = [];
  selectedRows(sheet, values.length).forEach(rowNum => {
    const row = values[rowNum - 1];
    if (!row || row.every(v => v === "" || v === null)) return;   // empty row
    const email = String(row[iEmail]).trim();
    if (!email) { noEmail.push(rowNum); return; }
    people.push({ rowNum, email, row });
  });

  if (people.length === 0) {
    ui.alert("Select one or more rows with people in them (any cell in the row is enough).");
    return;
  }

  // 4. Placeholders that don't match any column
  const unknown = unknownPlaceholders(t.subject + " " + t.body, keys);

  // 5. Sender alias available?
  const aliases = GmailApp.getAliases().map(a => a.toLowerCase());
  let from = sender;
  let fromNote = "";
  if (!aliases.includes(sender.toLowerCase())) {
    from = null;
    fromNote = `\n\n⚠ ${sender} is not a "Send mail as" address in your Gmail, ` +
      `so the drafts will be from your own address (${Session.getActiveUser().getEmail()}).`;
  }

  // 6. Confirm
  const list = people.slice(0, 15).map(p => `  row ${p.rowNum}: ${p.email}`).join("\n") +
    (people.length > 15 ? `\n  … and ${people.length - 15} more` : "");
  const warnings = [];
  if (noEmail.length) warnings.push(`⚠ Skipped rows without an email: ${noEmail.join(", ")}`);
  if (unknown.length) warnings.push(`⚠ These placeholders match no column and will stay as they are: ${unknown.join(", ")}`);

  const answer = ui.alert(
    `Create ${people.length} draft${people.length > 1 ? "s" : ""}?`,
    `Template: ${t.name}\nFrom: ${from || Session.getActiveUser().getEmail()}\n\n${list}` +
    (warnings.length ? "\n\n" + warnings.join("\n") : "") + fromNote +
    "\n\nThe drafts go to Gmail → Drafts. Nothing is sent until you send them.",
    ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;

  // 7. Create drafts
  const failed = [];
  people.forEach(p => {
    const subject = fillPlaceholders(t.subject, keys, p.row);
    const body = fillPlaceholders(t.body, keys, p.row);
    const options = from ? { from } : {};
    try {
      GmailApp.createDraft(p.email, subject, body, options);
    } catch (e) {
      failed.push(`row ${p.rowNum} (${p.email}): ${e.message}`);
    }
  });

  const created = people.length - failed.length;
  ui.alert("Drafts created",
    `Created ${created} draft${created === 1 ? "" : "s"} in Gmail → Drafts.` +
    (failed.length ? `\n\nFailed:\n${failed.join("\n")}` : ""),
    ui.ButtonSet.OK);
}

// Row numbers (1-based) of every selected row below the header, skipping rows
// hidden by a filter or by hand, so a selection over a filtered view only uses visible people.
function selectedRows(sheet, lastDataRow) {
  const rows = new Set();
  const ranges = sheet.getActiveRangeList() ? sheet.getActiveRangeList().getRanges() : [];
  ranges.forEach(range => {
    const from = Math.max(range.getRow(), 2);
    const to = Math.min(range.getLastRow(), lastDataRow);
    for (let r = from; r <= to; r++) rows.add(r);
  });
  return [...rows]
    .sort((a, b) => a - b)
    .filter(r => !sheet.isRowHiddenByFilter(r) && !sheet.isRowHiddenByUser(r));
}

// Replaces {{Column name}} with that row's value. Column names ignore case and extra spaces.
function fillPlaceholders(text, keys, row) {
  return text.replace(/\{\{([^}]+)\}\}/g, (match, name) => {
    const i = keys.indexOf(normHeader(name));
    if (i < 0) return match;
    const v = row[i];
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd.MM.yyyy");
    return String(v).trim();
  });
}

function unknownPlaceholders(text, keys) {
  const found = new Set();
  (text.match(/\{\{([^}]+)\}\}/g) || []).forEach(m => {
    if (keys.indexOf(normHeader(m.slice(2, -2))) < 0) found.add(m);
  });
  return [...found];
}

function normHeader(h) {
  return String(h).replace(/\s+/g, " ").trim().toLowerCase();
}