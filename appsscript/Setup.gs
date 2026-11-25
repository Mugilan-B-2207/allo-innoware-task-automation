/**
 * Setup.gs
 * Creates and formats all 6 operational Google Sheets tabs, data validations,
 * conditional formatting, protected ranges, sample data, and resets demo data safely.
 */

function initializeSystem(referenceToday) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.setSpreadsheetTimeZone) {
    ss.setSpreadsheetTimeZone('Asia/Kolkata');
  }
  const today = toDateOnly(referenceToday || new Date());

  const tabs = [
    { name: SHEET_NAMES.EMP_DASHBOARD, color: '#1a73e8' },
    { name: SHEET_NAMES.MGR_DASHBOARD, color: '#37474f' },
    { name: SHEET_NAMES.TASK_MASTER, color: '#3f51b5' },
    { name: SHEET_NAMES.CYCLE_LOG, color: '#00897b' },
    { name: SHEET_NAMES.ALERT_LOG, color: '#e65100' },
    { name: SHEET_NAMES.CONFIG, color: '#5f6368' }
  ];

  tabs.forEach((tab, idx) => {
    let sheet = ss.getSheetByName(tab.name);
    if (!sheet) {
      sheet = ss.insertSheet(tab.name, idx);
    } else {
      sheet.clear();
      sheet.clearConditionalFormatRules();
      const f = sheet.getFilter();
      if (f) f.remove();
      clearProtections(sheet);
    }
    sheet.setTabColor(tab.color);
  });

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  setupConfigurationSheet(ss.getSheetByName(SHEET_NAMES.CONFIG));
  setupTaskMasterSheet(ss.getSheetByName(SHEET_NAMES.TASK_MASTER), today);
  setupEmployeeDashboardSheet(ss.getSheetByName(SHEET_NAMES.EMP_DASHBOARD));
  setupCycleLogSheet(ss.getSheetByName(SHEET_NAMES.CYCLE_LOG));
  setupManagerDashboardSheet(ss.getSheetByName(SHEET_NAMES.MGR_DASHBOARD));
  setupAlertLogSheet(ss.getSheetByName(SHEET_NAMES.ALERT_LOG));

  applySheetProtections(ss);

  generateDueCycles(today);
  refreshAllDashboards(today);

  return 'Allo Operations system initialized across 6 tabs.';
}

/**
 * Resets demo operational data (Cycle Log, Alert Log, Dashboards) WITHOUT deleting Configuration or Task Master rules.
 */
function resetDemoData(referenceToday) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = toDateOnly(referenceToday || new Date());

  const logSheet = ss.getSheetByName(SHEET_NAMES.CYCLE_LOG);
  const alertSheet = ss.getSheetByName(SHEET_NAMES.ALERT_LOG);
  if (logSheet && logSheet.getLastRow() >= CL_START_ROW) {
    logSheet.getRange(CL_START_ROW, 1, logSheet.getLastRow() - CL_START_ROW + 1, CL_COL_COUNT).clearContent();
  }
  if (alertSheet && alertSheet.getLastRow() >= AL_START_ROW) {
    alertSheet.getRange(AL_START_ROW, 1, alertSheet.getLastRow() - AL_START_ROW + 1, AL_COL_COUNT).clearContent();
  }

  const tmSheet = ss.getSheetByName(SHEET_NAMES.TASK_MASTER);
  if (tmSheet && tmSheet.getLastRow() >= TM_START_ROW) {
    const count = tmSheet.getLastRow() - TM_START_ROW + 1;
    const dueDates = [];
    for (let i = 0; i < count; i++) {
      dueDates.push([formatYMD(today)]);
    }
    tmSheet.getRange(TM_START_ROW, TM_COLS.INITIAL_DUE_DATE, count, 1).setValues(dueDates);
  }

  generateDueCycles(today);
  refreshAllDashboards(today);
  return 'Demo cycle & alert history reset while preserving Configuration and Task Master.';
}

function clearProtections(sheet) {
  if (!sheet.getProtections) return;
  const prots = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .concat(sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE));
  prots.forEach(p => {
    try { p.remove(); } catch (e) {}
  });
}

function styleSheetBanner(sheet, title, subtitle, bgHex, colCount) {
  sheet.getRange(1, 1, 1, colCount).merge()
    .setValue(title)
    .setBackground(bgHex)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(13)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);

  sheet.getRange(2, 1, 1, colCount).merge()
    .setValue(subtitle)
    .setBackground('#f8f9fa')
    .setFontColor('#3c4043')
    .setFontSize(10)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 22);
}

function styleTableHeaders(sheet, rowIdx, headers, bgHex) {
  sheet.getRange(rowIdx, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(bgHex)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(rowIdx, 28);
  sheet.setFrozenRows(rowIdx);
}

function addStatusConditionalRules(sheet, rangeA1) {
  const range = sheet.getRange(rangeA1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(CYCLE_STATUS.PENDING).setBackground('#e8f0fe').setFontColor('#1967d2').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(CYCLE_STATUS.COMPLETED).setBackground('#e6f4ea').setFontColor('#137333').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(CYCLE_STATUS.OVERDUE).setBackground('#fce8e6').setFontColor('#c5221f').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(CYCLE_STATUS.ACKNOWLEDGED).setBackground('#fef7e0').setFontColor('#b06000').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(CYCLE_STATUS.ESCALATED).setBackground('#a50e0e').setFontColor('#ffffff').setBold(true).setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
}

function setupConfigurationSheet(sheet) {
  styleSheetBanner(
    sheet,
    'ALLO INNOWARE — SYSTEM CONFIGURATION & HOLIDAYS',
    'Centralized SLA timings, email recipients, working days, and holidays. Never hardcode real emails in code.',
    '#455a64',
    6
  );

  styleTableHeaders(sheet, CFG_HEADER_ROW, ['Setting Key', 'Value', 'Description', '', 'Holiday Date', 'Holiday Description'], '#37474f');

  const rows = [
    ['Company Name', DEFAULT_CONFIG['Company Name'], 'Organization name used in alerts and headers'],
    ['Spreadsheet Timezone', DEFAULT_CONFIG['Spreadsheet Timezone'], 'Default Asia/Kolkata'],
    ['Manager Email', DEFAULT_CONFIG['Manager Email'], 'Placeholder manager email (replace before going live)'],
    ['Leadership Email', DEFAULT_CONFIG['Leadership Email'], 'Placeholder leadership email (comma-separated allowed)'],
    ['Working Days', DEFAULT_CONFIG['Working Days'], 'Active operating days for Daily tasks and working-day SLAs'],
    ['Working Hours Start', DEFAULT_CONFIG['Working Hours Start'], 'Business hours start (HH:mm)'],
    ['Working Hours End', DEFAULT_CONFIG['Working Hours End'], 'Business hours end (HH:mm)'],
    ['Daily Task Deadline', DEFAULT_CONFIG['Daily Task Deadline'], 'Default daily task deadline (5:00 PM)'],
    ['Portal Task Deadline', DEFAULT_CONFIG['Portal Task Deadline'], 'Default recurring portal task deadline (5:00 PM)'],
    ['Employee Reminder Time', DEFAULT_CONFIG['Employee Reminder Time'], 'Time on due date to send employee reminder (9:00 AM)'],
    ['Manager First Alert Delay (Mins)', DEFAULT_CONFIG['Manager First Alert Delay (Mins)'], 'Minutes after deadline before sending Manager First Alert (15m -> 5:15 PM)'],
    ['Manager Follow-up Time', DEFAULT_CONFIG['Manager Follow-up Time'], 'Time on next working day to send unacked Manager Follow-up (10:00 AM)'],
    ['Manager Ack Deadline Time', DEFAULT_CONFIG['Manager Ack Deadline Time'], 'Manager acknowledgement deadline on next working day (10:00 AM)'],
    ['Unacked Early Escalation Hours', DEFAULT_CONFIG['Unacked Early Escalation Hours'], 'Hours after Manager Ack Deadline before early Leadership Escalation if still unacked'],
    ['Leadership Escalation Delay (Hrs)', DEFAULT_CONFIG['Leadership Escalation Delay (Hrs)'], 'Calendar hours after Manager First Alert before Leadership Escalation (24h)'],
    ['Leadership Escalation Enabled', DEFAULT_CONFIG['Leadership Escalation Enabled'], 'TRUE/FALSE toggle for leadership escalation emails'],
    ['Email Notifications Enabled', DEFAULT_CONFIG['Email Notifications Enabled'], 'TRUE/FALSE master switch for GmailApp/MailApp sending'],
    ['Test Mode Enabled', DEFAULT_CONFIG['Test Mode Enabled'], 'When TRUE, alerts are logged to Alert Log as TEST_MODE_LOGGED without emailing'],
    ['Daily Automation Enabled', DEFAULT_CONFIG['Daily Automation Enabled'], 'TRUE/FALSE master switch for daily cycle generation'],
    ['Daily Backlog Policy', DEFAULT_CONFIG['Daily Backlog Policy'], 'KEEP_ONE_OVERDUE_BLOCK_NEW prevents unlimited daily backlog']
  ];

  sheet.getRange(CFG_START_ROW, 1, rows.length, 3).setValues(rows);

  const sampleHolidays = [
    ['2026-01-26', 'Republic Day (Sample Holiday)'],
    ['2026-08-15', 'Independence Day (Sample Holiday)'],
    ['2026-10-02', 'Gandhi Jayanti (Sample Holiday)']
  ];
  sheet.getRange(CFG_START_ROW, 5, sampleHolidays.length, 2).setValues(sampleHolidays);
  sheet.getRange(CFG_START_ROW, 5, 20, 1).setNumberFormat('yyyy-mm-dd');

  sheet.setColumnWidths(1, 1, 250);
  sheet.setColumnWidths(2, 1, 280);
  sheet.setColumnWidths(3, 1, 440);
  sheet.setColumnWidths(4, 1, 30);
  sheet.setColumnWidths(5, 1, 130);
  sheet.setColumnWidths(6, 1, 240);
}

function setupTaskMasterSheet(sheet, today) {
  styleSheetBanner(
    sheet,
    'ALLO INNOWARE — TASK MASTER (ADMIN PROTECTED)',
    'Master configuration for recurring operational tasks. Default portal tasks use Completion Based recurrence.',
    '#3f51b5',
    TM_COL_COUNT
  );

  const headers = [
    'Task ID', 'Task Name', 'Description', 'Assigned Employee', 'Employee Email',
    'Frequency Type', 'Frequency Days', 'Recurrence Mode', 'Initial Due Date', 'Due Time',
    'Working Days Only', 'Proof Type', 'Proof Required', 'Active',
    'Manager Email', 'Leadership Email', 'Created At', 'Updated At'
  ];
  styleTableHeaders(sheet, TM_HEADER_ROW, headers, '#303f9f');

  const nowStr = formatTimestamp(new Date());
  const todayStr = formatYMD(today);

  const sampleTasks = [
    [
      'TASK-001', 'Daily Operations Update', 'Share daily operational update in team WhatsApp group and confirm here.',
      'Employee A', 'employee.a.placeholder@example.com',
      FREQ_TYPES.DAILY, 1, RECURRENCE_MODES.FIXED_SCHEDULE, todayStr, '17:00',
      true, PROOF_TYPES.WHATSAPP_CONFIRM, true, true,
      DEFAULT_CONFIG['Manager Email'], DEFAULT_CONFIG['Leadership Email'], nowStr, nowStr
    ],
    [
      'TASK-002', 'Portal Task A', 'Complete weekly portal workflow and enter Ticket ID.',
      'Employee B', 'employee.b.placeholder@example.com',
      FREQ_TYPES.WEEKLY, 7, RECURRENCE_MODES.COMPLETION_BASED, todayStr, '17:00',
      false, PROOF_TYPES.TICKET_ID, true, true,
      DEFAULT_CONFIG['Manager Email'], DEFAULT_CONFIG['Leadership Email'], nowStr, nowStr
    ],
    [
      'TASK-003', 'Portal Task B', 'Complete bi-weekly portal review and enter Ticket ID.',
      'Employee C', 'employee.c.placeholder@example.com',
      FREQ_TYPES.EVERY_15, 15, RECURRENCE_MODES.COMPLETION_BASED, todayStr, '17:00',
      false, PROOF_TYPES.TICKET_ID, true, true,
      DEFAULT_CONFIG['Manager Email'], DEFAULT_CONFIG['Leadership Email'], nowStr, nowStr
    ],
    [
      'TASK-004', 'Portal Task C', 'Complete monthly compliance portal ticket and enter Ticket ID.',
      'Employee D', 'employee.d.placeholder@example.com',
      FREQ_TYPES.EVERY_30, 30, RECURRENCE_MODES.COMPLETION_BASED, todayStr, '17:00',
      false, PROOF_TYPES.TICKET_ID, true, true,
      DEFAULT_CONFIG['Manager Email'], DEFAULT_CONFIG['Leadership Email'], nowStr, nowStr
    ]
  ];

  sheet.getRange(TM_START_ROW, 1, sampleTasks.length, TM_COL_COUNT).setValues(sampleTasks);

  sheet.getRange(TM_START_ROW, TM_COLS.WORKING_DAYS_ONLY, 50, 1).insertCheckboxes();
  sheet.getRange(TM_START_ROW, TM_COLS.PROOF_REQUIRED, 50, 1).insertCheckboxes();
  sheet.getRange(TM_START_ROW, TM_COLS.ACTIVE, 50, 1).insertCheckboxes();
  sheet.getRange(TM_START_ROW, TM_COLS.INITIAL_DUE_DATE, 50, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(TM_START_ROW, TM_COLS.DUE_TIME, 50, 1).setNumberFormat('@');

  if (SpreadsheetApp.newDataValidation) {
    const freqRule = SpreadsheetApp.newDataValidation().requireValueInList(Object.values(FREQ_TYPES), true).build();
    const modeRule = SpreadsheetApp.newDataValidation().requireValueInList(Object.values(RECURRENCE_MODES), true).build();
    const proofRule = SpreadsheetApp.newDataValidation().requireValueInList(Object.values(PROOF_TYPES), true).build();
    sheet.getRange(TM_START_ROW, TM_COLS.FREQUENCY_TYPE, 50, 1).setDataValidation(freqRule);
    sheet.getRange(TM_START_ROW, TM_COLS.RECURRENCE_MODE, 50, 1).setDataValidation(modeRule);
    sheet.getRange(TM_START_ROW, TM_COLS.PROOF_TYPE, 50, 1).setDataValidation(proofRule);
  }

  const activeRange = sheet.getRange(TM_START_ROW, TM_COLS.ACTIVE, 50, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FALSE').setBackground('#fce8e6').setFontColor('#c5221f').setRanges([activeRange]).build()
  ]);

  sheet.setColumnWidths(1, 1, 95);
  sheet.setColumnWidths(2, 1, 190);
  sheet.setColumnWidths(3, 1, 290);
  sheet.setColumnWidths(4, 1, 140);
  sheet.setColumnWidths(5, 1, 230);
  sheet.setColumnWidths(6, 3, 140);
  sheet.setColumnWidths(9, 6, 130);
  sheet.setColumnWidths(15, 4, 190);
}

function setupEmployeeDashboardSheet(sheet) {
  styleSheetBanner(
    sheet,
    'ALLO INNOWARE — EMPLOYEE DASHBOARD (SINGLE-POINT COMPLETION INTERFACE)',
    'View your active pending/overdue tasks. Enter Ticket ID (Portal tasks) or check Completion Confirmation + optional WhatsApp Link (Daily), then set Submission Action to "Submit Completion".',
    '#1a73e8',
    ED_COL_COUNT
  );

  sheet.getRange('A3').setValue('PENDING:').setFontWeight('bold').setFontColor('#1967d2');
  sheet.getRange('B3').setFormula(`=COUNTIF(G${ED_START_ROW}:G,"Pending")`).setFontWeight('bold').setFontColor('#1967d2');
  sheet.getRange('C3').setValue('OVERDUE:').setFontWeight('bold').setFontColor('#c5221f');
  sheet.getRange('D3').setFormula(`=COUNTIF(G${ED_START_ROW}:G,"Overdue")`).setFontWeight('bold').setFontColor('#c5221f');
  sheet.getRange('E3').setValue('ACKNOWLEDGED:').setFontWeight('bold').setFontColor('#b06000');
  sheet.getRange('F3').setFormula(`=COUNTIF(G${ED_START_ROW}:G,"Acknowledged")`).setFontWeight('bold').setFontColor('#b06000');
  sheet.getRange('G3').setValue('ESCALATED:').setFontWeight('bold').setFontColor('#a50e0e');
  sheet.getRange('H3').setFormula(`=COUNTIF(G${ED_START_ROW}:G,"Escalated")`).setFontWeight('bold').setFontColor('#a50e0e');
  sheet.getRange('I3:P3').merge().setValue('Workflow: Enter proof in Col J–M -> Select "Submit Completion" in Col P (or use Allo Operations menu).')
    .setFontStyle('italic').setFontColor('#5f6368');

  const headers = [
    'Cycle ID', 'Task ID', 'Task Name', 'Assigned Employee', 'Due Date', 'Due Time',
    'Status', 'Days Overdue', 'Required Proof', 'Ticket ID', 'Completion Confirmation',
    'Proof Link', 'Completion Notes', 'Completed By', 'Completed At', 'Submission Action'
  ];
  styleTableHeaders(sheet, ED_HEADER_ROW, headers, '#1557b0');

  sheet.getRange(ED_START_ROW, ED_COLS.DUE_DATE, 100, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(ED_START_ROW, ED_COLS.DUE_TIME, 100, 1).setNumberFormat('@');
  sheet.getRange(ED_START_ROW, ED_COLS.CONFIRMATION, 100, 1).insertCheckboxes();
  sheet.getRange(ED_START_ROW, ED_COLS.TICKET_ID, 100, 4).setBackground('#fffde7');
  sheet.getRange(ED_START_ROW, ED_COLS.SUBMISSION_ACTION, 100, 1).setBackground('#e8f0fe');

  if (SpreadsheetApp.newDataValidation) {
    const actionRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([SUBMISSION_ACTIONS.READY, SUBMISSION_ACTIONS.SUBMIT], true)
      .build();
    sheet.getRange(ED_START_ROW, ED_COLS.SUBMISSION_ACTION, 100, 1).setDataValidation(actionRule);
  }

  addStatusConditionalRules(sheet, `G${ED_START_ROW}:G150`);
  if (sheet.createFilter) {
    sheet.getRange(ED_HEADER_ROW, 1, 100, ED_COL_COUNT).createFilter();
  }

  sheet.setColumnWidths(1, 1, 185);
  sheet.setColumnWidths(2, 1, 90);
  sheet.setColumnWidths(3, 1, 195);
  sheet.setColumnWidths(4, 1, 135);
  sheet.setColumnWidths(5, 4, 105);
  sheet.setColumnWidths(9, 1, 205);
  sheet.setColumnWidths(10, 1, 130);
  sheet.setColumnWidths(11, 1, 155);
  sheet.setColumnWidths(12, 2, 200);
  sheet.setColumnWidths(14, 3, 155);
}

function setupCycleLogSheet(sheet) {
  styleSheetBanner(
    sheet,
    'ALLO INNOWARE — CYCLE LOG (PERMANENT AUDIT HISTORY)',
    'Permanent ledger of all recurring task cycles. Only one active cycle exists per task at a time; closed cycles are never overwritten.',
    '#00897b',
    CL_COL_COUNT
  );

  const headers = [
    'Cycle ID', 'Task ID', 'Task Name', 'Assigned Employee', 'Cycle Number',
    'Cycle Start Date', 'Due Date', 'Due Time', 'Status', 'Ticket ID',
    'Proof Link', 'Completion Notes', 'Completed By', 'Completed At', 'Next Due Date',
    'Manager Alert Sent At', 'Manager Acknowledged At', 'Leadership Escalated At',
    'Created At', 'Updated At'
  ];
  styleTableHeaders(sheet, CL_HEADER_ROW, headers, '#00695c');
  addStatusConditionalRules(sheet, `I${CL_START_ROW}:I500`);
  sheet.setColumnWidths(1, CL_COL_COUNT, 150);
  sheet.setColumnWidths(1, 1, 185);
  sheet.setColumnWidths(3, 1, 190);
}

function setupManagerDashboardSheet(sheet) {
  styleSheetBanner(
    sheet,
    'ALLO INNOWARE — MANAGER DASHBOARD (CENTRALIZED RISK & ACKNOWLEDGEMENT)',
    'Live visibility across all employees, overdue tasks, manager acknowledgements, and leadership escalations.',
    '#37474f',
    MD_COL_COUNT
  );

  const kpiHeaders = [
    'Total Active Tasks', 'Tasks Pending', 'Tasks Overdue',
    'Completed Today', 'Awaiting Manager Ack', 'Escalated to Leadership'
  ];
  sheet.getRange(4, 1, 1, 6).setValues([kpiHeaders])
    .setBackground('#eceff1').setFontColor('#263238').setFontWeight('bold');
  sheet.getRange(5, 1, 1, 6).setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(5, 30);

  sheet.getRange(7, 1, 1, 6).setValues([[
    'Employee', 'Pending Tasks', 'Overdue Tasks', 'Completed Today', 'Awaiting Ack', 'Escalated'
  ]]).setBackground('#455a64').setFontColor('#ffffff').setFontWeight('bold');

  const detailHeaders = [
    'Employee', 'Task', 'Cycle ID', 'Due Date', 'Status', 'Days Overdue',
    'Ticket ID / Proof', 'Manager Acknowledgement', 'Escalation Status',
    'Acknowledge', 'Acknowledgement Notes', 'Acknowledged By', 'Acknowledged At', 'Expected Resolution Date'
  ];
  sheet.getRange(MD_DETAIL_HEADER_ROW, 1, 1, MD_COL_COUNT)
    .setValues([detailHeaders])
    .setBackground('#263238')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setWrap(true);
  sheet.setRowHeight(MD_DETAIL_HEADER_ROW, 28);
  sheet.setFrozenRows(MD_DETAIL_HEADER_ROW);

  sheet.getRange(MD_DETAIL_START_ROW, MD_COLS.ACK_CHECKBOX, 100, 1).insertCheckboxes();
  sheet.getRange(MD_DETAIL_START_ROW, MD_COLS.ACK_NOTES, 100, 1).setBackground('#fffde7');
  sheet.getRange(MD_DETAIL_START_ROW, MD_COLS.EXPECTED_RES_DATE, 100, 1).setBackground('#fffde7').setNumberFormat('yyyy-mm-dd');
  addStatusConditionalRules(sheet, `E${MD_DETAIL_START_ROW}:E150`);

  sheet.setColumnWidths(1, MD_COL_COUNT, 145);
  sheet.setColumnWidths(2, 1, 190);
  sheet.setColumnWidths(3, 1, 185);
  sheet.setColumnWidths(11, 1, 210);
}

function setupAlertLogSheet(sheet) {
  styleSheetBanner(
    sheet,
    'ALLO INNOWARE — ALERT AUDIT LOG',
    'Permanent deduplicated audit trail of Employee Reminders, Manager First Alerts, Manager Follow-ups, Leadership Escalations, and Resolution Notifications.',
    '#e65100',
    AL_COL_COUNT
  );

  const headers = [
    'Alert ID', 'Cycle ID', 'Task ID', 'Employee', 'Alert Type', 'Recipient Email',
    'Alert Subject', 'Alert Sent At', 'Alert Status', 'Acknowledged', 'Acknowledged At', 'Error Message'
  ];
  styleTableHeaders(sheet, AL_HEADER_ROW, headers, '#bf360c');
  sheet.getRange(AL_START_ROW, AL_COLS.ACKNOWLEDGED, 200, 1).insertCheckboxes();
  sheet.setColumnWidths(1, AL_COL_COUNT, 150);
  sheet.setColumnWidths(6, 2, 230);
  sheet.setColumnWidths(12, 1, 220);
}

function applySheetProtections(ss) {
  try {
    const tm = ss.getSheetByName(SHEET_NAMES.TASK_MASTER);
    const cl = ss.getSheetByName(SHEET_NAMES.CYCLE_LOG);
    const al = ss.getSheetByName(SHEET_NAMES.ALERT_LOG);
    const cfg = ss.getSheetByName(SHEET_NAMES.CONFIG);
    [tm, cl, al, cfg].forEach(s => {
      if (s && s.protect) {
        s.protect().setDescription(`${s.getName()} is admin-protected`).setWarningOnly(true);
      }
    });
  } catch (e) {
    // Non-fatal in restricted/mock environments
  }
}
