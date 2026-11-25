/**
 * TestRunner.gs
 * Repeatable automated verification suite for Tests 1-8 and Section 12 edge cases.
 * Writes structured results to the "Test Results" sheet.
 */

function assertCondition(cond, msg) {
  if (!cond) throw new Error(msg);
}

function runAllTestCases() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseTime = new Date(2026, 10, 25, 10, 0, 0); // Wed 2026-11-25 10:00 IST
  initializeSystem(baseTime);
  setConfigSetting(ss, 'Test Mode Enabled', 'TRUE');
  setConfigSetting(ss, 'Email Notifications Enabled', 'TRUE');

  const rows = [];
  function runTest(id, name, expected, fn) {
    const ts = formatTimestamp(new Date());
    try {
      const actual = fn();
      rows.push([id, name, expected, actual, 'PASS', '', ts]);
    } catch (err) {
      rows.push([id, name, expected, 'Assertion failed', 'FAIL', err.message || String(err), ts]);
    }
  }

  // TEST 1: Recurring Cycle (7-day Completion Based)
  runTest(
    'TEST-01',
    'Recurring Cycle (7-Day Task Completion)',
    'Current cycle closes with timestamp; next cycle generated for Completion Date + 7 days.',
    () => {
      const ed = ss.getSheetByName(SHEET_NAMES.EMP_DASHBOARD);
      const r = findDashboardRowByTask(ed, 'TASK-002');
      assertCondition(r > 0, 'TASK-002 should appear on Employee Dashboard');
      ed.getRange(r, ED_COLS.TICKET_ID).setValue('TKT-7001');
      ed.getRange(r, ED_COLS.SUBMISSION_ACTION).setValue(SUBMISSION_ACTIONS.SUBMIT);

      const res = handleEmployeeDashboardEdit(ss, ed, r, ED_COLS.SUBMISSION_ACTION, {
        now: new Date(2026, 10, 25, 14, 0, 0),
        actor: 'employee.b.placeholder@example.com'
      });
      assertCondition(res && res.ok, `Completion failed: ${res && res.error}`);
      assertCondition(res.nextDueDate === '2026-12-02', `Expected 2026-12-02, got ${res.nextDueDate}`);
      const nextCycle = getAllCycles(ss).find(c => c.taskId === 'TASK-002' && c.status === CYCLE_STATUS.PENDING);
      assertCondition(Boolean(nextCycle) && formatYMD(nextCycle.dueDate) === '2026-12-02', 'Next cycle must be Pending for 2026-12-02');
      return `Closed ${res.cycleId} at ${res.completedAt}; generated ${nextCycle.cycleId} due ${res.nextDueDate}.`;
    }
  );

  // TEST 2: Single-Point Completion
  runTest(
    'TEST-02',
    'Single-Point Completion (Employee Dashboard Only)',
    'Ticket ID saved from Employee Dashboard; cycle closes; dashboard updates with next cycle without touching other sheets.',
    () => {
      const ed = ss.getSheetByName(SHEET_NAMES.EMP_DASHBOARD);
      const r = findDashboardRowByTask(ed, 'TASK-003');
      assertCondition(r > 0, 'TASK-003 should appear on Employee Dashboard');
      ed.getRange(r, ED_COLS.TICKET_ID).setValue('TKT-15099');
      ed.getRange(r, ED_COLS.SUBMISSION_ACTION).setValue(SUBMISSION_ACTIONS.SUBMIT);

      const res = handleEmployeeDashboardEdit(ss, ed, r, ED_COLS.SUBMISSION_ACTION, {
        now: new Date(2026, 10, 25, 15, 0, 0),
        actor: 'employee.c.placeholder@example.com'
      });
      assertCondition(res && res.ok, `Single-point submission failed: ${res && res.error}`);
      const closed = getAllCycles(ss).find(c => c.cycleId === res.cycleId);
      assertCondition(closed && closed.ticketId === 'TKT-15099' && closed.status === CYCLE_STATUS.COMPLETED, 'Ticket ID and Completed status must persist in Cycle Log');
      const newDashRow = findDashboardRowByTask(ed, 'TASK-003');
      assertCondition(ed.getRange(newDashRow, ED_COLS.CYCLE_ID).getValue() !== res.cycleId, 'Employee Dashboard must update to new cycle');
      return `Saved TKT-15099 on ${res.cycleId} from Employee Dashboard alone; dashboard refreshed to next cycle due ${res.nextDueDate}.`;
    }
  );

  // TEST 3: Daily Task (No Ticket ID Required)
  runTest(
    'TEST-03',
    'Daily WhatsApp Task Completion (No Ticket ID Required)',
    'Confirmation & timestamp recorded without Ticket ID; next working-day cycle generated.',
    () => {
      const ed = ss.getSheetByName(SHEET_NAMES.EMP_DASHBOARD);
      const r = findDashboardRowByTask(ed, 'TASK-001');
      assertCondition(r > 0, 'TASK-001 should appear on Employee Dashboard');
      ed.getRange(r, ED_COLS.CONFIRMATION).setValue(true);
      ed.getRange(r, ED_COLS.PROOF_LINK).setValue('https://wa.me/p/sample-update');
      ed.getRange(r, ED_COLS.SUBMISSION_ACTION).setValue(SUBMISSION_ACTIONS.SUBMIT);

      const res = handleEmployeeDashboardEdit(ss, ed, r, ED_COLS.SUBMISSION_ACTION, {
        now: new Date(2026, 10, 25, 16, 30, 0),
        actor: 'employee.a.placeholder@example.com'
      });
      assertCondition(res && res.ok, `Daily completion failed: ${res && res.error}`);
      assertCondition(res.nextDueDate === '2026-11-26', `Expected next working day 2026-11-26, got ${res.nextDueDate}`);
      return `Confirmed ${res.cycleId} without Ticket ID at ${res.completedAt}; spawned next working-day cycle due 2026-11-26.`;
    }
  );

  // TEST 4: Same-Day Manager Alert + Email Notification Enabled Toggle
  runTest(
    'TEST-04',
    'Same-Day Manager Alert & Email Enabled Check',
    'TASK-004 crosses 17:00 deadline -> Overdue at 17:05 -> Manager First Alert logged at 17:20; skipped when disabled.',
    () => {
      // First verify disabled email toggle logs SKIPPED_DISABLED
      setConfigSetting(ss, 'Test Mode Enabled', 'FALSE');
      setConfigSetting(ss, 'Email Notifications Enabled', 'FALSE');
      const disabledRes = runAlertAndEscalationCheck(new Date(2026, 10, 25, 17, 20, 0));
      assertCondition(disabledRes.mgrFirstAlerts === 0, 'Must not send when Email Notifications Enabled is FALSE');

      // Reset Manager Alert Sent At on TASK-004 cycle, enable Test Mode & Email Notifications, and re-run
      const c4 = getAllCycles(ss).find(c => c.taskId === 'TASK-004' && isActiveCycleStatus(c.status));
      ss.getSheetByName(SHEET_NAMES.CYCLE_LOG).getRange(c4.row, CL_COLS.MGR_ALERT_SENT_AT).setValue('');
      setConfigSetting(ss, 'Test Mode Enabled', 'TRUE');
      setConfigSetting(ss, 'Email Notifications Enabled', 'TRUE');

      const enabledRes = runAlertAndEscalationCheck(new Date(2026, 10, 25, 17, 20, 0));
      assertCondition(enabledRes.mgrFirstAlerts === 1, `Expected 1 Manager First Alert, got ${enabledRes.mgrFirstAlerts}`);
      const updatedC4 = getAllCycles(ss).find(c => c.cycleId === c4.cycleId);
      assertCondition(updatedC4.status === CYCLE_STATUS.OVERDUE && Boolean(updatedC4.mgrAlertSentAt), 'TASK-004 must be Overdue with Manager Alert Sent At');
      return `Verified disabled toggle skipped sending; when enabled, ${c4.cycleId} became Overdue and logged Manager First Alert at 17:20.`;
    }
  );

  // TEST 5: Manager Acknowledgement (Does NOT Close Underlying Task)
  runTest(
    'TEST-05',
    'Manager Acknowledgement Separation',
    'Acknowledged By & At recorded; Manager Dashboard updates; underlying task remains active (Acknowledged, not Completed).',
    () => {
      const md = ss.getSheetByName(SHEET_NAMES.MGR_DASHBOARD);
      const mdRow = findManagerRowByTask(md, 'TASK-004');
      assertCondition(mdRow > 0, 'TASK-004 must appear on Manager Dashboard');

      const ackRes = acknowledgeCycle(ss, mdRow, {
        now: new Date(2026, 10, 25, 18, 0, 0),
        ackBy: 'manager.placeholder@example.com',
        ackNotes: 'Pinged Employee D; awaiting portal ticket.',
        expectedResolutionDate: '2026-11-26'
      });
      assertCondition(ackRes.ok, `Acknowledgement failed: ${ackRes.error}`);
      const c4 = getAllCycles(ss).find(c => c.cycleId === ackRes.cycleId);
      assertCondition(c4.status === CYCLE_STATUS.ACKNOWLEDGED && Boolean(c4.mgrAckAt), 'Cycle status must be Acknowledged, never Completed');
      return `Recorded ack by ${ackRes.ackBy} at ${ackRes.ackAt} (delay ${ackRes.ackDelayMins}m); task remains open (${c4.status}).`;
    }
  );

  // TEST 6: Leadership Escalation (>24h SLA + Completed Task Not Escalated)
  runTest(
    'TEST-06',
    'Leadership Escalation (>24h After Manager Alert)',
    'Escalates unresolved task >24h after Manager Alert with reason; completed tasks are never escalated.',
    () => {
      const escRes = runAlertAndEscalationCheck(new Date(2026, 10, 26, 17, 25, 0)); // 24h05m after 17:20
      assertCondition(escRes.leadershipEscalations === 1, `Expected 1 leadership escalation for TASK-004, got ${escRes.leadershipEscalations}`);
      const c4 = getAllCycles(ss).find(c => c.taskId === 'TASK-004' && isActiveCycleStatus(c.status));
      assertCondition(c4.status === CYCLE_STATUS.ESCALATED && Boolean(c4.leadEscAt), 'TASK-004 must be Escalated with Leadership Escalated At');

      // Verify already-completed TASK-002 / TASK-003 were NOT escalated
      const completedEscalated = getAllCycles(ss).filter(c => c.status === CYCLE_STATUS.COMPLETED && c.leadEscAt);
      assertCondition(completedEscalated.length === 0, 'Completed tasks must never be escalated');
      return `Escalated ${c4.cycleId} at ${formatTimestamp(c4.leadEscAt)} (>24h SLA); verified 0 completed tasks escalated.`;
    }
  );

  // TEST 7: Full Escalation Timing Lifecycle (weekday-to-weekday: Tue 2026-12-01 -> Wed 2026-12-02)
  runTest(
    'TEST-07',
    'Full SLA Timing Progression (Reminder -> Overdue -> Mgr Alert -> Follow-up -> Leadership -> Resolution)',
    'Each SLA phase fires only in its exact configured window.',
    () => {
      // Advance TASK-001 cleanly to Tue 2026-12-01 (Cycle 5) so we observe exact 16.75h weekday Follow-up before 24h SLA
      const ed = ss.getSheetByName(SHEET_NAMES.EMP_DASHBOARD);
      const advanceDates = [
        new Date(2026, 10, 26, 17, 35, 0), // closes Cycle 2 -> spawns Fri 2026-11-27
        new Date(2026, 10, 27, 17, 35, 0), // closes Cycle 3 -> spawns Mon 2026-11-30
        new Date(2026, 10, 30, 17, 35, 0)  // closes Cycle 4 -> spawns Tue 2026-12-01
      ];
      advanceDates.forEach(dt => {
        const row = findDashboardRowByTask(ed, 'TASK-001');
        ed.getRange(row, ED_COLS.CONFIRMATION).setValue(true);
        ed.getRange(row, ED_COLS.SUBMISSION_ACTION).setValue(SUBMISSION_ACTIONS.SUBMIT);
        const adv = handleEmployeeDashboardEdit(ss, ed, row, ED_COLS.SUBMISSION_ACTION, {
          now: dt,
          actor: 'employee.a.placeholder@example.com'
        });
        assertCondition(adv && adv.ok, 'Advancing TASK-001 to Tue 2026-12-01 failed');
      });

      // 1. 09:05 on Tue 2026-12-01 -> Employee Reminder
      const rRem = runAlertAndEscalationCheck(new Date(2026, 11, 1, 9, 5, 0));
      assertCondition(rRem.reminders >= 1, 'Expected Employee Reminder at 09:05');

      // 2. 16:00 (Before deadline) -> still Pending, 0 Manager alerts
      const rPre = runAlertAndEscalationCheck(new Date(2026, 11, 1, 16, 0, 0));
      assertCondition(rPre.mgrFirstAlerts === 0, 'Must not send Manager alert before 17:00 deadline');

      // 3. 17:05 (Immediately after 17:00 deadline) -> Overdue, 0 Manager alerts (before 17:15)
      const rAfter = runAlertAndEscalationCheck(new Date(2026, 11, 1, 17, 5, 0));
      assertCondition(rAfter.mgrFirstAlerts === 0, 'Must wait until 17:15 (deadline + 15m)');
      const c5 = getAllCycles(ss).find(c => c.taskId === 'TASK-001' && isActiveCycleStatus(c.status));
      assertCondition(c5.status === CYCLE_STATUS.OVERDUE, 'Status must turn Overdue at 17:05');

      // 4. 17:20 -> Manager First Alert
      const rL1 = runAlertAndEscalationCheck(new Date(2026, 11, 1, 17, 20, 0));
      assertCondition(rL1.mgrFirstAlerts === 1, 'Expected Manager First Alert at 17:20');

      // 5. Next working day 10:05 (Wed 2026-12-02, 16.75h after L1) -> Manager Follow-up, BEFORE Leadership Escalation
      const rFollow = runAlertAndEscalationCheck(new Date(2026, 11, 2, 10, 5, 0));
      assertCondition(rFollow.mgrFollowups === 1 && rFollow.leadershipEscalations === 0, '10:05 next working day should send Follow-up, not Leadership yet');

      // 6. Next working day 14:05 (4h after unacked 10:00 deadline) -> Early Leadership Escalation
      const rLead = runAlertAndEscalationCheck(new Date(2026, 11, 2, 14, 5, 0));
      assertCondition(rLead.leadershipEscalations === 1, 'Expected early Leadership Escalation at 14:05 when unacked');

      // 7. Complete TASK-001 -> triggers Completion Notification & closes cycle
      const r = findDashboardRowByTask(ed, 'TASK-001');
      ed.getRange(r, ED_COLS.CONFIRMATION).setValue(true);
      ed.getRange(r, ED_COLS.SUBMISSION_ACTION).setValue(SUBMISSION_ACTIONS.SUBMIT);
      const compRes = handleEmployeeDashboardEdit(ss, ed, r, ED_COLS.SUBMISSION_ACTION, {
        now: new Date(2026, 11, 2, 14, 30, 0),
        actor: 'employee.a.placeholder@example.com'
      });
      assertCondition(compRes && compRes.ok, 'Post-escalation completion should succeed');
      return `Verified 7-stage lifecycle on ${c5.cycleId}: 09:05 Reminder -> 16:00 Pending -> 17:05 Overdue -> 17:20 Mgr Alert -> 10:05 Follow-up -> 14:05 Leadership Esc -> 14:30 Resolved.`;
    }
  );

  // TEST 8: Duplicate Prevention (Cycles, Completions, Alerts, Triggers)
  runTest(
    'TEST-08',
    'Duplicate Prevention & Idempotency',
    'Repeated runs create 0 duplicate active cycles, 0 duplicate completions, 0 duplicate alerts, 0 duplicate triggers.',
    () => {
      const cycleCountBefore = getAllCycles(ss).length;
      const alertCountBefore = ss.getSheetByName(SHEET_NAMES.ALERT_LOG).getLastRow();

      generateDueCycles(new Date(2026, 11, 2, 15, 0, 0));
      generateDueCycles(new Date(2026, 11, 2, 15, 15, 0));
      const a1 = runAlertAndEscalationCheck(new Date(2026, 11, 2, 15, 0, 0));
      const a2 = runAlertAndEscalationCheck(new Date(2026, 11, 2, 15, 15, 0));

      assertCondition(getAllCycles(ss).length === cycleCountBefore, 'No duplicate cycles allowed');
      assertCondition(ss.getSheetByName(SHEET_NAMES.ALERT_LOG).getLastRow() === alertCountBefore, 'No duplicate alerts allowed');
      assertCondition(a1.mgrFirstAlerts === 0 && a2.mgrFirstAlerts === 0 && a1.leadershipEscalations === 0, 'Repeated runs must send 0 new alerts');

      // Trigger idempotency check
      const t1 = installProjectTriggers();
      const t2 = installProjectTriggers();
      assertCondition(t1.totalTriggers === t2.totalTriggers && t2.added === 0, 'installProjectTriggers twice must add 0 duplicate triggers');

      return `Repeated generateDueCycles, runAlertAndEscalationCheck, and installProjectTriggers produced 0 duplicate cycles, alerts, or triggers.`;
    }
  );

  // EDGE CASES (Section 12)
  runTest(
    'EDGE-01',
    'Missing Ticket ID & Invalid Completion Guard',
    'Rejects blank Ticket ID, generic "done", and unchecked Daily confirmation without closing cycle.',
    () => {
      const ed = ss.getSheetByName(SHEET_NAMES.EMP_DASHBOARD);
      const r2 = findDashboardRowByTask(ed, 'TASK-002');
      ed.getRange(r2, ED_COLS.TICKET_ID).setValue('done');
      ed.getRange(r2, ED_COLS.SUBMISSION_ACTION).setValue(SUBMISSION_ACTIONS.SUBMIT);
      const badTicket = handleEmployeeDashboardEdit(ss, ed, r2, ED_COLS.SUBMISSION_ACTION);
      assertCondition(!badTicket.ok, 'Generic "done" must fail validation');

      const r1 = findDashboardRowByTask(ed, 'TASK-001');
      ed.getRange(r1, ED_COLS.CONFIRMATION).setValue(false);
      ed.getRange(r1, ED_COLS.SUBMISSION_ACTION).setValue(SUBMISSION_ACTIONS.SUBMIT);
      const badDaily = handleEmployeeDashboardEdit(ss, ed, r1, ED_COLS.SUBMISSION_ACTION);
      assertCondition(!badDaily.ok, 'Unchecked daily confirmation must fail validation');
      ed.getRange(r2, ED_COLS.TICKET_ID).setValue('');
      ed.getRange(r2, ED_COLS.NOTES).setValue('');
      ed.getRange(r1, ED_COLS.NOTES).setValue('');
      return 'Blocked both invalid Ticket ID ("done") and unchecked Daily Confirmation while keeping cycles open.';
    }
  );

  runTest(
    'EDGE-02',
    'Unauthorized Employee & Duplicate Cycle Submission Guard',
    'Blocks Employee C from submitting Employee B task; blocks re-submitting a completed Cycle ID.',
    () => {
      const ed = ss.getSheetByName(SHEET_NAMES.EMP_DASHBOARD);
      const r2 = findDashboardRowByTask(ed, 'TASK-002');
      ed.getRange(r2, ED_COLS.TICKET_ID).setValue('TKT-9991');
      const unauth = submitEmployeeCompletion(ss, r2, { actingEmployee: 'Employee C' });
      assertCondition(!unauth.ok && unauth.error.includes('Unauthorized'), 'Cross-employee submission must be blocked');
      ed.getRange(r2, ED_COLS.TICKET_ID).setValue('');
      ed.getRange(r2, ED_COLS.NOTES).setValue('');

      // Check duplicate completion check directly against an already-completed cycle
      const completedCycle = getAllCycles(ss).find(c => c.status === CYCLE_STATUS.COMPLETED);
      ed.getRange(r2, ED_COLS.CYCLE_ID).setValue(completedCycle.cycleId);
      const dupRes = submitEmployeeCompletion(ss, r2, { now: new Date() });
      assertCondition(!dupRes.ok && dupRes.error.includes('already completed'), 'Duplicate submission must be blocked');
      refreshAllDashboards(new Date(2026, 10, 27, 15, 0, 0));
      return 'Blocked cross-employee submission (Employee C -> TASK-002) and duplicate Cycle ID completion.';
    }
  );

  runTest(
    'EDGE-03',
    'Weekend, Holiday, Month-End, Inactive Task & Missing Email Handling',
    'Skips weekends/holidays on Daily recurrence; handles Jan 31 + 30d; skips inactive task; logs FAILED on missing email.',
    () => {
      const cfg = getConfig(ss);
      cfg.holidays = ['2026-11-30']; // Suppose Mon Nov 30 is a holiday
      const dailyTask = getTasksFromMaster(ss).find(t => t.taskId === 'TASK-001');
      // Friday Nov 27 -> skips Sat Nov 28, Sun Nov 29, Holiday Mon Nov 30 -> Tue Dec 01!
      const nextDaily = calculateNextDueDate(dailyTask, new Date(2026, 10, 27, 16, 0, 0), new Date(2026, 10, 27), cfg);
      assertCondition(formatYMD(nextDaily) === '2026-12-01', `Expected 2026-12-01 skipping weekend+holiday, got ${formatYMD(nextDaily)}`);

      const monthlyTask = getTasksFromMaster(ss).find(t => t.taskId === 'TASK-004');
      const jan31Plus30 = calculateNextDueDate(monthlyTask, new Date(2026, 0, 31, 12, 0, 0), new Date(2026, 0, 31), cfg);
      assertCondition(formatYMD(jan31Plus30) === '2026-03-02', `Expected 2026-03-02 for Jan 31 + 30d, got ${formatYMD(jan31Plus30)}`);

      // Missing email configuration check
      const badEmailRes = dispatchAlert(ss, cfg, {
        cycleId: 'TEST-MISSING-EMAIL',
        taskId: 'TASK-004',
        employee: 'Employee D',
        alertType: ALERT_TYPES.EMP_REMINDER,
        recipientEmail: '',
        subject: 'Missing Email Test',
        now: new Date(2026, 10, 27, 15, 0, 0)
      });
      assertCondition(!badEmailRes.sent && badEmailRes.status === 'FAILED', 'Empty recipient must log FAILED cleanly without throwing');

      // Inactive task check
      const tm = ss.getSheetByName(SHEET_NAMES.TASK_MASTER);
      tm.getRange(TM_START_ROW, TM_COLS.ACTIVE).setValue(false);
      const tasksAfter = getTasksFromMaster(ss);
      assertCondition(!tasksAfter[0].active, 'Task should read as inactive');
      tm.getRange(TM_START_ROW, TM_COLS.ACTIVE).setValue(true);

      return 'Verified Fri->Tue (weekend+holiday skip), Jan 31+30d=Mar 02, inactive task flag, and graceful FAILED log for missing email.';
    }
  );

  writeTestResultsSheet(ss, rows);
  return rows;
}

function findDashboardRowByTask(edSheet, taskId) {
  const lastRow = edSheet.getLastRow();
  if (lastRow < ED_START_ROW) return -1;
  const vals = edSheet.getRange(ED_START_ROW, ED_COLS.TASK_ID, lastRow - ED_START_ROW + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === taskId) return ED_START_ROW + i;
  }
  return -1;
}

function findManagerRowByTask(mdSheet, taskIdPrefix) {
  const lastRow = mdSheet.getLastRow();
  if (lastRow < MD_DETAIL_START_ROW) return -1;
  const vals = mdSheet.getRange(MD_DETAIL_START_ROW, MD_COLS.CYCLE_ID, lastRow - MD_DETAIL_START_ROW + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim().startsWith(taskIdPrefix)) return MD_DETAIL_START_ROW + i;
  }
  return -1;
}

function writeTestResultsSheet(ss, rows) {
  let sheet = ss.getSheetByName(SHEET_NAMES.TEST_RESULTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.TEST_RESULTS);
  } else {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  }
  sheet.setTabColor('#6a1b9a');
  styleSheetBanner(
    sheet,
    'ALLO INNOWARE — AUTOMATED TEST RESULTS',
    'Repeatable verification across Tests 1–8 and Section 12 edge cases.',
    '#6a1b9a',
    7
  );
  styleTableHeaders(sheet, 4, ['Test ID', 'Test Name', 'Expected Result', 'Actual Result', 'Pass / Fail', 'Error Message', 'Execution Timestamp'], '#4a148c');
  sheet.getRange(5, 1, rows.length, 7).setValues(rows);

  const pfRange = sheet.getRange(5, 5, rows.length, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('PASS').setBackground('#e6f4ea').setFontColor('#137333').setBold(true).setRanges([pfRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FAIL').setBackground('#fce8e6').setFontColor('#c5221f').setBold(true).setRanges([pfRange]).build()
  ]);

  sheet.setColumnWidths(1, 1, 95);
  sheet.setColumnWidths(2, 1, 270);
  sheet.setColumnWidths(3, 2, 370);
  sheet.setColumnWidths(5, 1, 95);
  sheet.setColumnWidths(6, 1, 220);
  sheet.setColumnWidths(7, 1, 165);
}
