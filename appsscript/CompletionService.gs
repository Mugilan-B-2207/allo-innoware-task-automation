/**
 * CompletionService.gs
 * Single-point employee completion submission, proof validation,
 * cycle closure, duplicate prevention, and automatic next-cycle generation.
 */

function handleEmployeeDashboardEdit(ss, sheet, row, col, options) {
  if (col !== ED_COLS.SUBMISSION_ACTION) return null;
  const actionVal = String(sheet.getRange(row, ED_COLS.SUBMISSION_ACTION).getValue() || '').trim();
  if (actionVal !== SUBMISSION_ACTIONS.SUBMIT) return null;

  return submitEmployeeCompletion(ss, row, options);
}

/**
 * Submits completion for an active cycle on Employee Dashboard.
 */
function submitEmployeeCompletion(ss, dashboardRow, options) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
    const opts = options || {};
    const now = opts.now ? new Date(opts.now) : new Date();
    const cfg = getConfig(spreadsheet);
    const edSheet = spreadsheet.getSheetByName(SHEET_NAMES.EMP_DASHBOARD);
    const logSheet = spreadsheet.getSheetByName(SHEET_NAMES.CYCLE_LOG);

    const rowVals = edSheet.getRange(dashboardRow, 1, 1, ED_COL_COUNT).getValues()[0];
    const cycleId = String(rowVals[ED_COLS.CYCLE_ID - 1] || '').trim();
    const taskId = String(rowVals[ED_COLS.TASK_ID - 1] || '').trim();
    const ticketId = String(rowVals[ED_COLS.TICKET_ID - 1] || '').trim();
    const confirmed = toBool(rowVals[ED_COLS.CONFIRMATION - 1]);
    const proofLink = String(rowVals[ED_COLS.PROOF_LINK - 1] || '').trim();
    const notes = String(rowVals[ED_COLS.NOTES - 1] || '').trim();

    if (!cycleId || !taskId) {
      return failSubmission(edSheet, dashboardRow, 'Invalid or empty Cycle ID / Task ID.');
    }

    const tasks = getTasksFromMaster(spreadsheet);
    const task = tasks.find(t => t.taskId === taskId);
    if (!task) {
      return failSubmission(edSheet, dashboardRow, `Task ${taskId} not found in Task Master.`);
    }
    if (!task.active && !opts.allowInactive) {
      return failSubmission(edSheet, dashboardRow, `Task ${taskId} is inactive.`);
    }

    const cycles = getAllCycles(spreadsheet);
    const cycle = cycles.find(c => c.cycleId === cycleId);
    if (!cycle) {
      return failSubmission(edSheet, dashboardRow, `Cycle ${cycleId} not found in Cycle Log.`);
    }
    if (cycle.status === CYCLE_STATUS.COMPLETED) {
      return failSubmission(edSheet, dashboardRow, `Cycle ${cycleId} is already completed (duplicate blocked).`);
    }

    // Employee assignment validation
    if (opts.actingEmployee && opts.actingEmployee !== task.assignedEmployee) {
      return failSubmission(edSheet, dashboardRow, `Unauthorized: ${opts.actingEmployee} cannot submit task assigned to ${task.assignedEmployee}.`);
    }
    if (!opts.actingEmployee && opts.actingEmail) {
      const ownerOfEmail = tasks.find(t => t.employeeEmail.toLowerCase() === opts.actingEmail.toLowerCase());
      if (ownerOfEmail && ownerOfEmail.assignedEmployee !== task.assignedEmployee) {
        return failSubmission(edSheet, dashboardRow, `Unauthorized: ${opts.actingEmail} cannot submit ${task.assignedEmployee}'s task.`);
      }
    }

    // Proof validation
    const proofCheck = validateCompletionProof(task, ticketId, confirmed, proofLink);
    if (!proofCheck.valid) {
      return failSubmission(edSheet, dashboardRow, proofCheck.error);
    }

    // All validations passed -> close current cycle cleanly
    const actor = opts.actor || resolveUserEmail(task.employeeEmail || task.assignedEmployee);
    const completedStr = formatTimestamp(now);
    const nextDue = calculateNextDueDate(task, now, cycle.dueDate, cfg);
    const nextDueStr = formatYMD(nextDue);
    const wasEscalatedOrAlerted = Boolean(cycle.mgrAlertSentAt || cycle.leadEscAt);

    logSheet.getRange(cycle.row, CL_COLS.STATUS, 1, 7).setValues([[
      CYCLE_STATUS.COMPLETED,
      ticketId,
      proofLink,
      notes,
      actor,
      completedStr,
      nextDueStr
    ]]);
    logSheet.getRange(cycle.row, CL_COLS.UPDATED_AT).setValue(completedStr);

    // Immediately spawn next cycle & refresh dashboards
    generateDueCycles(now);
    refreshAllDashboards(now);

    if (wasEscalatedOrAlerted) {
      const updatedCycle = Object.assign({}, cycle, {
        status: CYCLE_STATUS.COMPLETED,
        ticketId,
        proofLink,
        completionNotes: notes,
        completedBy: actor,
        completedAt: now
      });
      sendCompletionResolutionNotice(spreadsheet, cfg, updatedCycle, task, now);
    }

    return {
      ok: true,
      cycleId,
      taskId,
      completedAt: completedStr,
      completedBy: actor,
      nextDueDate: nextDueStr
    };
  } catch (err) {
    logSystemEvent('ERROR', 'submitEmployeeCompletion', err.message, { dashboardRow });
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function validateCompletionProof(task, ticketId, confirmed, proofLink) {
  if (!task.proofRequired) return { valid: true };

  if (task.proofType === PROOF_TYPES.TICKET_ID) {
    const clean = String(ticketId || '').trim();
    const invalidWords = new Set(['done', 'ok', 'yes', 'completed', 'complete', 'na', 'n/a', 'test']);
    if (!clean) {
      return { valid: false, error: 'Ticket ID is mandatory for portal tasks.' };
    }
    if (clean.length < 3 || invalidWords.has(clean.toLowerCase()) || !/[0-9]/.test(clean)) {
      return { valid: false, error: `Invalid Ticket ID "${clean}". Enter a real Ticket ID (e.g., TKT-1042).` };
    }
    return { valid: true };
  }

  // Daily WhatsApp task: checkbox confirmation required; Ticket ID not required; Proof Link optional
  if (!confirmed) {
    return { valid: false, error: 'Check "Completion Confirmation" to confirm the WhatsApp update was shared.' };
  }
  if (proofLink && !/^https?:\/\/\S+/i.test(proofLink)) {
    return { valid: false, error: 'Optional Proof Link must be a valid http(s) URL if provided.' };
  }
  return { valid: true };
}

function failSubmission(edSheet, row, message) {
  edSheet.getRange(row, ED_COLS.SUBMISSION_ACTION).setValue(SUBMISSION_ACTIONS.READY);
  edSheet.getRange(row, ED_COLS.NOTES).setValue(`VALIDATION ERROR: ${message}`);
  logSystemEvent('WARN', 'CompletionService', message, { row });
  return { ok: false, error: message };
}

/**
 * Menu helper: submits whichever row is currently selected on Employee Dashboard.
 */
function submitSelectedEmployeeRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.EMP_DASHBOARD) {
    SpreadsheetApp.getUi().alert('Please select a task row on the "Employee Dashboard" sheet first.');
    return;
  }
  const row = sheet.getActiveRange().getRow();
  if (row < ED_START_ROW) {
    SpreadsheetApp.getUi().alert('Please select a valid task row (row 6 or below).');
    return;
  }
  const res = submitEmployeeCompletion(ss, row);
  ss.toast(res.ok ? `Completed ${res.cycleId}! Next due ${res.nextDueDate}` : res.error, 'Completion Status', 5);
}
