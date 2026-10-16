/**
 * QA Bandwidth Calculation – create a tab for the active Jira sprint.
 *
 * Adds a "QA Tools" menu to the spreadsheet:
 *   - Set Jira token ................ stores your Jira Personal Access Token (per user)
 *   - Create tab for active sprint .. copies the template tab, names it after the
 *                                     active sprint, and lists the QA tickets
 *                                     from B21 down (QA story point in column C)
 *   - Assign QA tickets in Jira ..... reads the BREAKDOWN section (QA names on row 7)
 *                                     of the active tab and assigns each ticket in Jira
 *
 * A ticket counts as a QA ticket when it is not a sub-task and either its title
 * contains "[QA]" or its "QA Estimate" field has a value.
 *
 * Assignment rules (one ticket can be listed under several QAs):
 *   - title has [QA] and one QA ....... assign the ticket itself, plus its [QA] sub-task
 *   - every ticket, for each QA ....... assign a [QA] sub-task: keep one already
 *                                       assigned to them, else reuse a [QA] sub-task
 *                                       not held by a listed QA, else create
 *                                       "[QA] Execute testing" and assign it
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

  BREAKDOWN_HEADER_ROW: 7,              // "Niko | Assigned Ticket | QA Story Point | Elsa | ..."
  BREAKDOWN_TICKET_HEADER: 'Assigned Ticket',
  QA_SUBTASK_SUMMARY: '[QA] Execute testing',
  // QA name as written on row 7 → Jira username.
  QA_USERS: {
    Niko: 'niko.susanto@ninjavan.co',
    Elsa: 'elsa.fauzia@ninjavan.co',
    Andhar: 'andhar.dezan@ninjavan.co',
  },
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('QA Tools')
    .addItem('Create tab for active sprint', 'createActiveSprintTab')
    .addItem('Assign QA tickets in Jira', 'assignQaTickets')
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

function assignQaTickets() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  const entries = readBreakdown_(sheet);
  if (!entries.length) {
    ui.alert(`No tickets found in the BREAKDOWN section of "${sheet.getName()}".`);
    return;
  }
  const unknown = [...new Set(entries.filter((e) => !e.user).map((e) => e.qa))];
  if (unknown.length) {
    ui.alert(
      `No Jira username for: ${unknown.join(', ')}.\n` +
        'Put their Jira email under their name on row 7, or add them to CONFIG.QA_USERS in the script.'
    );
    return;
  }

  const keys = [...new Set(entries.map((e) => e.key))];
  const plan = planAssignments_(entries, fetchTicketInfo_(keys));
  const warnings = plan.filter((a) => a.type === 'missing').map(describeAction_);
  const actions = plan.filter((a) => a.type !== 'missing');
  if (!actions.length) {
    ui.alert(['Everything is already assigned as listed. Nothing to do.', ...warnings].join('\n'));
    return;
  }

  const preview = [...actions.map(describeAction_), ...(warnings.length ? ['', 'Skipped:', ...warnings] : [])];
  const ok = ui.alert(
    `Apply ${actions.length} change(s) in Jira?`,
    preview.join('\n'),
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  const subtaskTypeId = actions.some((a) => a.type === 'create') ? getSubtaskTypeId_() : null;
  const results = actions.map((a) => {
    try {
      if (a.type === 'assign') {
        jiraRequest_('put', `/rest/api/2/issue/${a.key}/assignee`, { name: a.user });
        return `✓ ${describeAction_(a)}`;
      }
      const created = jiraRequest_('post', '/rest/api/2/issue', {
        fields: {
          project: { key: a.parent.split('-')[0] },
          parent: { key: a.parent },
          issuetype: { id: subtaskTypeId },
          summary: CONFIG.QA_SUBTASK_SUMMARY,
          assignee: { name: a.user },
        },
      });
      return `✓ ${describeAction_(a)} → ${created.key}`;
    } catch (err) {
      return `✗ ${describeAction_(a)}: ${err.message}`;
    }
  });
  ui.alert('Jira assignment finished', results.join('\n'), ui.ButtonSet.OK);
}

/** Returns [{qa, user, key}] for every ticket listed under a QA in the BREAKDOWN section. */
function readBreakdown_(sheet) {
  const headerRow = CONFIG.BREAKDOWN_HEADER_ROW;
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];

  // The breakdown ends just above the "Unassigned" row.
  const below = sheet.getRange(headerRow + 1, 1, sheet.getLastRow() - headerRow, lastCol).getDisplayValues();
  const unassignedIdx = below.findIndex((row) => row.some((v) => v.trim() === 'Unassigned'));
  const rowCount = unassignedIdx === -1 ? below.length : unassignedIdx;
  if (rowCount <= 0) return [];

  const entries = [];
  header.forEach((title, i) => {
    if (title.trim() !== CONFIG.BREAKDOWN_TICKET_HEADER || i === 0) return;
    const { qa, user } = parseQaHeader_(header[i - 1]);
    const cells = sheet.getRange(headerRow + 1, i + 1, rowCount, 1).getRichTextValues();
    cells.forEach(([cell]) => {
      const key = ticketKeyOf_(cell);
      if (qa && key) entries.push({ qa, user, key });
    });
  });
  return entries;
}

/** "Niko niko.susanto@ninjavan.co" → {qa: "Niko", user: "niko.susanto@ninjavan.co"}. */
function parseQaHeader_(text) {
  const email = (text.match(/[\w.+-]+@[\w-]+(\.[\w-]+)+/) || [])[0];
  const qa = text.replace(email || '', '').replace(/\s+/g, ' ').trim();
  const fallback = CONFIG.QA_USERS[qa];
  return { qa, user: email || (fallback && fallback !== 'TODO' ? fallback : null) };
}

function ticketKeyOf_(richText) {
  const sources = [richText.getLinkUrl(), ...richText.getRuns().map((r) => r.getLinkUrl()), richText.getText()];
  for (const s of sources) {
    const m = s && s.match(/[A-Z][A-Z0-9]+-\d+/);
    if (m) return m[0];
  }
  return null;
}

/** Returns {KEY: {summary, assignee, qaSubtasks: [{key, assignee, done}]}}. */
function fetchTicketInfo_(keys) {
  const list = keys.join(',');
  const parents = jiraSearch_(`key in (${list})`, 'summary,assignee');
  const subtasks = jiraSearch_(`parent in (${list})`, 'summary,assignee,status,parent');

  const info = {};
  parents.forEach((i) => {
    info[i.key] = {
      summary: i.fields.summary,
      assignee: i.fields.assignee ? i.fields.assignee.name : null,
      qaSubtasks: [],
    };
  });
  subtasks
    .filter((s) => /\[QA\]/i.test(s.fields.summary) && info[s.fields.parent.key])
    .forEach((s) => {
      info[s.fields.parent.key].qaSubtasks.push({
        key: s.key,
        assignee: s.fields.assignee ? s.fields.assignee.name : null,
        done: s.fields.status.statusCategory.key === 'done',
      });
    });
  return info;
}

/** Pure planning step: what to assign or create so Jira matches the sheet. */
function planAssignments_(entries, info) {
  const qasByKey = {};
  entries.forEach(({ qa, user, key }) => {
    qasByKey[key] = qasByKey[key] || [];
    if (!qasByKey[key].some((q) => q.user === user)) qasByKey[key].push({ qa, user });
  });

  const actions = [];
  Object.keys(qasByKey).forEach((key) => {
    const ticket = info[key];
    const qas = qasByKey[key];
    if (!ticket) {
      actions.push({ type: 'missing', key, qa: qas.map((q) => q.qa).join(', ') });
      return;
    }

    // A [QA] ticket with one QA: assign the ticket itself, then its [QA] sub-task below.
    if (/\[QA\]/i.test(ticket.summary) && qas.length === 1) {
      const { qa, user } = qas[0];
      if (ticket.assignee !== user) actions.push({ type: 'assign', key, qa, user, what: 'ticket' });
    }

    const users = qas.map((q) => q.user);
    // Reusable [QA] sub-tasks: open and not held by a QA listed for this ticket, unassigned first.
    const pool = ticket.qaSubtasks
      .filter((s) => !s.done && !users.includes(s.assignee))
      .sort((a, b) => (a.assignee ? 1 : 0) - (b.assignee ? 1 : 0));
    qas.forEach(({ qa, user }) => {
      if (ticket.qaSubtasks.some((s) => s.assignee === user)) return;
      const reuse = pool.shift();
      if (reuse) actions.push({ type: 'assign', key: reuse.key, parent: key, qa, user, what: 'sub-task' });
      else actions.push({ type: 'create', parent: key, qa, user });
    });
  });
  return actions;
}

function describeAction_(a) {
  if (a.type === 'missing') return `${a.key}: not found in Jira (listed for ${a.qa})`;
  if (a.type === 'create') return `${a.parent}: create "${CONFIG.QA_SUBTASK_SUMMARY}" sub-task → ${a.qa}`;
  if (a.what === 'sub-task') return `${a.parent}: assign sub-task ${a.key} → ${a.qa}`;
  return `${a.key}: assign ticket → ${a.qa}`;
}

function getSubtaskTypeId_() {
  const types = jiraRequest_('get', '/rest/api/2/issuetype').filter((t) => t.subtask);
  const type = types.find((t) => t.name === 'Sub-task') || types[0];
  if (!type) throw new Error('No sub-task issue type found in Jira.');
  return type.id;
}

function jiraSearch_(jql, fields) {
  const issues = [];
  let startAt = 0;
  while (true) {
    // validateQuery: false so a mistyped key in the sheet is reported as missing instead of failing the search.
    const page = jiraRequest_('post', '/rest/api/2/search', {
      jql,
      fields: fields.split(','),
      startAt,
      maxResults: 100,
      validateQuery: false,
    });
    issues.push(...page.issues);
    startAt += page.issues.length;
    if (!page.issues.length || startAt >= page.total) break;
  }
  return issues;
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
  return jiraRequest_('get', path);
}

function jiraRequest_(method, path, payload) {
  const token = PropertiesService.getUserProperties().getProperty(CONFIG.TOKEN_PROPERTY);
  if (!token) throw new Error('No Jira token yet. Use QA Tools → Set Jira token first.');
  const options = {
    method,
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    muteHttpExceptions: true,
  };
  if (payload !== undefined) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  const res = UrlFetchApp.fetch(CONFIG.JIRA_BASE + path, options);
  const code = res.getResponseCode();
  if (code === 401) throw new Error('Jira rejected the token (401). Set a new one via QA Tools → Set Jira token.');
  if (code < 200 || code >= 300) throw new Error(`Jira request failed (${code}): ${res.getContentText().slice(0, 300)}`);
  const body = res.getContentText();
  return body ? JSON.parse(body) : null;
}
