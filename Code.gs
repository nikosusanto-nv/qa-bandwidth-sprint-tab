/**
 * QA Bandwidth Calculation – create a tab for the active Jira sprint.
 *
 * Adds a "QA Tools" menu to the spreadsheet:
 *   - Set Jira token ................ stores your Jira Personal Access Token (per user)
 *   - Create tab for active sprint .. copies the template tab, names it after the
 *                                     active sprint, and lists the QA tickets
 *                                     from B21 down (QA story point in column C)
 *
 * A ticket counts as a QA ticket when it is not a sub-task and either its title
 * contains "[QA]" or its "QA Estimate" field has a value.
 */

const CONFIG = {
  JIRA_BASE: 'https://jira.ninjavan.net',
  BOARD_ID: 220,                       // XB Scrum Board Daily
  QA_ESTIMATE_FIELD: 'customfield_12005', // "QA Estimate" = QA story point
  TEMPLATE_SHEET_ID: 1704848751,        // the "Tepmplate" tab
  LIST_START_ROW: 21,                   // first Unassigned ticket row
  LIST_SLOTS: 10,                       // template rows 21–30 before the Total row
  TICKET_COL: 2,                        // column B
  POINT_COL: 3,                         // column C
  TOKEN_PROPERTY: 'JIRA_PAT',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('QA Tools')
    .addItem('Create tab for active sprint', 'createActiveSprintTab')
    .addSeparator()
    .addItem('Set Jira token', 'setJiraToken')
    .addToUi();
}

function setJiraToken() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Jira Personal Access Token',
    'Create one in Jira: profile picture → Profile → Personal Access Tokens.\n' +
      'It is stored only for your Google account in this script.',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const token = res.getResponseText().trim();
  if (!token) return;
  PropertiesService.getUserProperties().setProperty(CONFIG.TOKEN_PROPERTY, token);
  ui.alert('Jira token saved.');
}

function createActiveSprintTab() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sprint = getActiveSprint_();
  if (ss.getSheetByName(sprint.name)) {
    ui.alert(`A tab named "${sprint.name}" already exists. Rename or delete it first.`);
    return;
  }

  const template = ss.getSheets().find((s) => s.getSheetId() === CONFIG.TEMPLATE_SHEET_ID);
  if (!template) throw new Error(`Template tab (gid ${CONFIG.TEMPLATE_SHEET_ID}) not found.`);

  const tickets = getQaTickets_(sprint.id);

  const sheet = template.copyTo(ss).setName(sprint.name);
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);

  // Grow the list inside the SUM(C20:C30) range so the Total row keeps counting everything.
  const extra = tickets.length - CONFIG.LIST_SLOTS;
  if (extra > 0) {
    sheet.insertRowsAfter(CONFIG.LIST_START_ROW + CONFIG.LIST_SLOTS - 2, extra);
  }

  if (tickets.length) {
    const links = tickets.map((t) => [
      SpreadsheetApp.newRichTextValue()
        .setText(`${t.key} ${t.summary}`)
        .setLinkUrl(`${CONFIG.JIRA_BASE}/browse/${t.key}`)
        .build(),
    ]);
    const points = tickets.map((t) => [t.qaPoint === null ? '' : t.qaPoint]);
    sheet.getRange(CONFIG.LIST_START_ROW, CONFIG.TICKET_COL, tickets.length, 1).setRichTextValues(links);
    sheet.getRange(CONFIG.LIST_START_ROW, CONFIG.POINT_COL, tickets.length, 1).setValues(points);
  }

  ui.alert(`Created "${sprint.name}" with ${tickets.length} QA ticket(s) from row ${CONFIG.LIST_START_ROW}.`);
}

function getActiveSprint_() {
  const data = jiraGet_(`/rest/agile/1.0/board/${CONFIG.BOARD_ID}/sprint?state=active`);
  if (!data.values || !data.values.length) throw new Error('No active sprint on the board.');
  return data.values[0];
}

function getQaTickets_(sprintId) {
  const fields = ['summary', 'issuetype', CONFIG.QA_ESTIMATE_FIELD].join(',');
  const issues = [];
  let startAt = 0;
  while (true) {
    const page = jiraGet_(
      `/rest/agile/1.0/board/${CONFIG.BOARD_ID}/sprint/${sprintId}/issue` +
        `?startAt=${startAt}&maxResults=100&fields=${fields}`
    );
    issues.push(...page.issues);
    startAt += page.issues.length;
    if (!page.issues.length || startAt >= page.total) break;
  }

  return issues
    .filter((i) => !(i.fields.issuetype && i.fields.issuetype.subtask))
    .map((i) => ({
      key: i.key,
      summary: i.fields.summary,
      qaPoint: i.fields[CONFIG.QA_ESTIMATE_FIELD],
    }))
    .filter((t) => /\[QA\]/i.test(t.summary) || (t.qaPoint !== null && t.qaPoint !== undefined));
}

function jiraGet_(path) {
  const token = PropertiesService.getUserProperties().getProperty(CONFIG.TOKEN_PROPERTY);
  if (!token) throw new Error('No Jira token yet. Use QA Tools → Set Jira token first.');
  const res = UrlFetchApp.fetch(CONFIG.JIRA_BASE + path, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code === 401) throw new Error('Jira rejected the token (401). Set a new one via QA Tools → Set Jira token.');
  if (code !== 200) throw new Error(`Jira request failed (${code}): ${res.getContentText().slice(0, 300)}`);
  return JSON.parse(res.getContentText());
}
