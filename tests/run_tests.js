/**
 * Node.js verification suite executing the exact modular Google Apps Script files.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');
const { createGasContext } = require('./gas_mock');

const appsScriptDir = path.join(__dirname, '..', 'appsscript');
const files = [
  'Config.gs',
  'Utils.gs',
  'Setup.gs',
  'TaskEngine.gs',
  'DashboardService.gs',
  'AlertService.gs',
  'EscalationService.gs',
  'CompletionService.gs',
  'Code.gs',
  'TestRunner.gs'
];

const { ss, triggers, context } = createGasContext();
vm.createContext(context);

for (const file of files) {
  const code = fs.readFileSync(path.join(appsScriptDir, file), 'utf8');
  vm.runInContext(code, context, { filename: file });
}

console.log('=== Running Allo Innoware Master Test Suite ===');
const results = vm.runInContext('runAllTestCases()', context);

let failures = 0;
for (const [id, name, expected, actual, status, errMsg] of results) {
  console.log(`[${status}] ${id}: ${name}`);
  console.log(`       Expected: ${expected}`);
  console.log(`       Actual:   ${actual}`);
  if (errMsg) console.log(`       Error:    ${errMsg}`);
  if (status !== 'PASS') failures++;
}

// Verify live MailApp call path when Test Mode Enabled = FALSE & Email Notifications Enabled = TRUE
console.log('\n=== Extra Verification: Live MailApp Dispatch & Trigger Reset ===');
vm.runInContext(`
  initializeSystem(new Date(2026, 10, 25, 10, 0, 0));
  setConfigSetting(SpreadsheetApp.getActiveSpreadsheet(), 'Test Mode Enabled', 'FALSE');
  setConfigSetting(SpreadsheetApp.getActiveSpreadsheet(), 'Email Notifications Enabled', 'TRUE');
  runAlertAndEscalationCheck(new Date(2026, 10, 25, 17, 20, 0));
`, context);

assert.strictEqual(ss.sentEmails.length, 4, `Expected 4 live emails (TASK-001..TASK-004 Manager First Alerts), got ${ss.sentEmails.length}`);
console.log(`[PASS] Live MailApp Path: ${ss.sentEmails.length} HTML alerts dispatched when Test Mode = FALSE & Email Notifications = TRUE.`);

const removed = vm.runInContext('removeProjectTriggers()', context);
assert.strictEqual(triggers.length, 0, 'removeProjectTriggers should remove all triggers');
console.log(`[PASS] Trigger Removal: Removed ${removed.removed} triggers cleanly.`);

// Verify Reset Demo Data preserves Configuration & Task Master
vm.runInContext(`resetDemoData(new Date(2026, 10, 25, 10, 0, 0))`, context);
assert.strictEqual(ss.getSheetByName('Task Master').getLastRow(), 8, 'Task Master rows must remain intact after resetDemoData');
assert.strictEqual(ss.getSheetByName('Cycle Log').getLastRow(), 8, 'Cycle Log should contain only the 4 fresh active cycles after reset');
console.log('[PASS] Reset Demo Data: Preserved Task Master & Configuration while resetting Cycle Log & Dashboards.');

if (failures > 0) {
  console.error(`\nFAILED: ${failures} test case(s) failed.`);
  process.exit(1);
}
console.log(`\nALL VERIFICATIONS PASSED (${results.length}/${results.length} Suite Tests + 3 Live/Ops Checks).`);
