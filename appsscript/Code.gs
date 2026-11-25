/**
 * Code.gs
 * Spreadsheet entrypoints: "Allo Operations" custom menu, installable triggers,
 * scheduled job runners, and safe trigger management.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Allo Operations')
    .addItem('Initialize System', 'menuInitializeSystem')
    .addItem('Refresh Dashboards', 'menuRefreshDashboards')
    .addItem('Generate Due Cycles', 'menuGenerateDueCycles')
    .addItem('Run Pending Check', 'menuRunPendingCheck')
    .addItem('Run Alert Check', 'menuRunAlertCheck')
    .addSeparator()
    .addItem('Submit Selected Completion', 'submitSelectedEmployeeRow')
    .addItem('Open Employee Dashboard', 'menuOpenEmployeeDashboard')
    .addItem('Open Manager Dashboard', 'menuOpenManagerDashboard')
    .addSeparator()
    .addItem('Run Test Cases', 'menuRunTestCases')
    .addItem('Install Automation Triggers', 'installProjectTriggers')
    .addItem('Remove Project Triggers', 'removeProjectTriggers')
    .addItem('Reset Demo Data', 'menuResetDemoData')
    .addToUi();
}

function menuInitializeSystem() {
  const msg = initializeSystem();
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Allo Operations', 5);
}

function menuRefreshDashboards() {
  refreshAllDashboards();
  SpreadsheetApp.getActiveSpreadsheet().toast('Employee & Manager Dashboards refreshed.', 'Allo Operations', 4);
}

function menuGenerateDueCycles() {
  const res = generateDueCycles();
  refreshAllDashboards();
  SpreadsheetApp.getActiveSpreadsheet().toast(`Cycle check complete. Created ${res.created} new cycle(s).`, 'Allo Operations', 4);
}

function menuRunPendingCheck() {
  const res = runPendingCheck();
  SpreadsheetApp.getActiveSpreadsheet().toast(`Pending check complete. Newly overdue: ${res.newlyOverdue}.`, 'Allo Operations', 4);
}

function menuRunAlertCheck() {
  const res = runAlertAndEscalationCheck();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Alert check complete. Reminders: ${res.reminders}, Mgr Alerts: ${res.mgrFirstAlerts}, Follow-ups: ${res.mgrFollowups}, Leadership: ${res.leadershipEscalations}`,
    'Allo Operations',
    6
  );
}

function menuOpenEmployeeDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.EMP_DASHBOARD);
  if (sheet) ss.setActiveSheet(sheet);
}

function menuOpenManagerDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MGR_DASHBOARD);
  if (sheet) ss.setActiveSheet(sheet);
}

function menuRunTestCases() {
  const results = runAllTestCases();
  const passed = results.filter(r => r[4] === 'PASS').length;
  SpreadsheetApp.getActiveSpreadsheet().toast(`Tests completed: ${passed}/${results.length} PASSED. See "Test Results" tab.`, 'Test Runner', 6);
}

function menuResetDemoData() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Confirm Demo Data Reset',
    'This will clear Cycle Log and Alert Log demo records and reset active cycles to today. Configuration and Task Master definitions will be preserved. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  const msg = resetDemoData();
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Allo Operations', 5);
}

/**
 * Installable onEdit trigger handler.
 */
function onSheetEditTrigger(e) {
  if (!e || !e.range) return;
  const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.range.getSheet();
  const name = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (name === SHEET_NAMES.EMP_DASHBOARD && row >= ED_START_ROW) {
    handleEmployeeDashboardEdit(ss, sheet, row, col);
  } else if (name === SHEET_NAMES.MGR_DASHBOARD && row >= MD_DETAIL_START_ROW) {
    handleManagerDashboardEdit(ss, sheet, row, col);
  }
}

/**
 * Scheduled job: Daily at 8:45 AM (creates daily cycles & prepares employee reports).
 */
function scheduledMorningCycleJob() {
  generateDueCycles();
  refreshAllDashboards();
}

/**
 * Scheduled job: Daily at 9:00 AM (sends due-today employee reminders).
 */
function scheduledEmployeeReminderJob() {
  runAlertAndEscalationCheck();
}

/**
 * Scheduled job: Every 15 minutes (overdue detection, manager alerts, escalation checks).
 */
function scheduledQuarterHourMonitorJob() {
  generateDueCycles();
  runAlertAndEscalationCheck();
}

/**
 * Safely installs project triggers without creating duplicates.
 */
function installProjectTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ScriptApp.getProjectTriggers();
  const handlerNames = new Set(existing.map(t => t.getHandlerFunction()));
  let added = 0;

  if (!handlerNames.has('onSheetEditTrigger')) {
    ScriptApp.newTrigger('onSheetEditTrigger').forSpreadsheet(ss).onEdit().create();
    added++;
  }
  if (!handlerNames.has('scheduledMorningCycleJob')) {
    ScriptApp.newTrigger('scheduledMorningCycleJob').timeBased().everyDays(1).atHour(8).nearMinute(45).create();
    added++;
  }
  if (!handlerNames.has('scheduledEmployeeReminderJob')) {
    ScriptApp.newTrigger('scheduledEmployeeReminderJob').timeBased().everyDays(1).atHour(9).create();
    added++;
  }
  if (!handlerNames.has('scheduledQuarterHourMonitorJob')) {
    ScriptApp.newTrigger('scheduledQuarterHourMonitorJob').timeBased().everyMinutes(15).create();
    added++;
  }

  return { totalTriggers: ScriptApp.getProjectTriggers().length, added };
}

/**
 * Safely removes all project triggers.
 */
function removeProjectTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  return { removed: triggers.length };
}
