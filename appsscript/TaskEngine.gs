/**
 * TaskEngine.gs
 * Reads Task Master & Cycle Log, calculates recurring next due dates,
 * generates active cycles idempotently, and detects overdue cycles.
 */

function getTasksFromMaster(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.TASK_MASTER);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < TM_START_ROW) return [];

  const vals = sheet.getRange(TM_START_ROW, 1, lastRow - TM_START_ROW + 1, TM_COL_COUNT).getValues();
  const tasks = [];
  vals.forEach((r, idx) => {
    const taskId = String(r[TM_COLS.TASK_ID - 1] || '').trim();
    if (!taskId) return;
    tasks.push({
      row: TM_START_ROW + idx,
      taskId,
      taskName: String(r[TM_COLS.TASK_NAME - 1] || '').trim(),
      description: String(r[TM_COLS.DESCRIPTION - 1] || '').trim(),
      assignedEmployee: String(r[TM_COLS.ASSIGNED_EMPLOYEE - 1] || '').trim(),
      employeeEmail: String(r[TM_COLS.EMPLOYEE_EMAIL - 1] || '').trim(),
      frequencyType: String(r[TM_COLS.FREQUENCY_TYPE - 1] || FREQ_TYPES.CUSTOM).trim(),
      frequencyDays: Math.max(1, Number(r[TM_COLS.FREQUENCY_DAYS - 1]) || 1),
      recurrenceMode: String(r[TM_COLS.RECURRENCE_MODE - 1] || RECURRENCE_MODES.COMPLETION_BASED).trim(),
      initialDueDate: toDateOnly(r[TM_COLS.INITIAL_DUE_DATE - 1]),
      dueTime: formatHM(r[TM_COLS.DUE_TIME - 1], '17:00'),
      workingDaysOnly: toBool(r[TM_COLS.WORKING_DAYS_ONLY - 1]),
      proofType: String(r[TM_COLS.PROOF_TYPE - 1] || PROOF_TYPES.TICKET_ID).trim(),
      proofRequired: toBool(r[TM_COLS.PROOF_REQUIRED - 1]),
      active: toBool(r[TM_COLS.ACTIVE - 1]),
      managerEmail: String(r[TM_COLS.MANAGER_EMAIL - 1] || '').trim(),
      leadershipEmail: String(r[TM_COLS.LEADERSHIP_EMAIL - 1] || '').trim()
    });
  });
  return tasks;
}

function getAllCycles(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.CYCLE_LOG);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < CL_START_ROW) return [];

  const vals = sheet.getRange(CL_START_ROW, 1, lastRow - CL_START_ROW + 1, CL_COL_COUNT).getValues();
  const cycles = [];
  vals.forEach((r, idx) => {
    const cycleId = String(r[CL_COLS.CYCLE_ID - 1] || '').trim();
    if (!cycleId) return;
    cycles.push({
      row: CL_START_ROW + idx,
      cycleId,
      taskId: String(r[CL_COLS.TASK_ID - 1] || '').trim(),
      taskName: String(r[CL_COLS.TASK_NAME - 1] || '').trim(),
      assignedEmployee: String(r[CL_COLS.ASSIGNED_EMPLOYEE - 1] || '').trim(),
      cycleNumber: Number(r[CL_COLS.CYCLE_NUMBER - 1]) || 1,
      cycleStartDate: toDateOnly(r[CL_COLS.CYCLE_START_DATE - 1]),
      dueDate: toDateOnly(r[CL_COLS.DUE_DATE - 1]),
      dueTime: formatHM(r[CL_COLS.DUE_TIME - 1], '17:00'),
      status: String(r[CL_COLS.STATUS - 1] || CYCLE_STATUS.PENDING).trim(),
      ticketId: String(r[CL_COLS.TICKET_ID - 1] || '').trim(),
      proofLink: String(r[CL_COLS.PROOF_LINK - 1] || '').trim(),
      completionNotes: String(r[CL_COLS.COMPLETION_NOTES - 1] || '').trim(),
      completedBy: String(r[CL_COLS.COMPLETED_BY - 1] || '').trim(),
      completedAt: parseDateTime(r[CL_COLS.COMPLETED_AT - 1]),
      nextDueDate: toDateOnly(r[CL_COLS.NEXT_DUE_DATE - 1]),
      mgrAlertSentAt: parseDateTime(r[CL_COLS.MGR_ALERT_SENT_AT - 1]),
      mgrAckAt: parseDateTime(r[CL_COLS.MGR_ACK_AT - 1]),
      leadEscAt: parseDateTime(r[CL_COLS.LEAD_ESC_AT - 1]),
      createdAt: parseDateTime(r[CL_COLS.CREATED_AT - 1]),
      updatedAt: parseDateTime(r[CL_COLS.UPDATED_AT - 1])
    });
  });
  return cycles;
}

function isActiveCycleStatus(status) {
  return status !== CYCLE_STATUS.COMPLETED;
}

/**
 * Calculates the next due date according to task frequency, recurrence mode,
 * working-day/holiday configuration, and month-end safety.
 */
function calculateNextDueDate(task, completedAt, previousDueDate, cfg) {
  const compDate = toDateOnly(completedAt);
  const prevDue = toDateOnly(previousDueDate) || compDate;

  if (task.frequencyType === FREQ_TYPES.DAILY || task.workingDaysOnly) {
    const base = task.recurrenceMode === RECURRENCE_MODES.COMPLETION_BASED
      ? compDate
      : (compDate > prevDue ? compDate : prevDue);
    if (task.frequencyType === FREQ_TYPES.DAILY) {
      return nextWorkingDay(base, cfg, false);
    }
    const rawTarget = addCalendarDays(base, task.frequencyDays);
    return nextWorkingDay(rawTarget, cfg, true);
  }

  if (task.recurrenceMode === RECURRENCE_MODES.FIXED_SCHEDULE) {
    let candidate = addCalendarDays(prevDue, task.frequencyDays);
    while (candidate <= compDate) {
      candidate = addCalendarDays(candidate, task.frequencyDays);
    }
    return candidate;
  }

  // Default: Completion Based (Next Due Date = Completion Date + Frequency Days)
  return addCalendarDays(compDate, task.frequencyDays);
}

/**
 * Generates active cycles for any active Task Master item lacking an active cycle.
 * Enforces: only 1 active cycle per task at a time + no unlimited daily backlog.
 */
function generateDueCycles(referenceNow) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = referenceNow ? new Date(referenceNow) : new Date();
    const today = toDateOnly(now);
    const cfg = getConfig(ss);
    const tasks = getTasksFromMaster(ss);
    const cycles = getAllCycles(ss);
    const logSheet = ss.getSheetByName(SHEET_NAMES.CYCLE_LOG);

    const byTask = {};
    cycles.forEach(c => {
      if (!byTask[c.taskId]) byTask[c.taskId] = [];
      byTask[c.taskId].push(c);
    });

    const newRows = [];
    const nowStr = formatTimestamp(now);

    tasks.forEach(task => {
      if (!task.active) return;
      if (task.frequencyType === FREQ_TYPES.DAILY && !cfg.dailyAutoEnabled) return;

      const taskCycles = byTask[task.taskId] || [];
      const activeCycle = taskCycles.find(c => isActiveCycleStatus(c.status));
      // Policy: KEEP_ONE_OVERDUE_BLOCK_NEW -> if active/overdue cycle exists, never create duplicate
      if (activeCycle) return;

      const maxNum = taskCycles.reduce((m, c) => Math.max(m, c.cycleNumber), 0);
      const nextNum = maxNum + 1;
      const lastCompleted = taskCycles
        .filter(c => c.status === CYCLE_STATUS.COMPLETED)
        .sort((a, b) => b.cycleNumber - a.cycleNumber)[0];

      let dueDate;
      let startDate = today;
      if (lastCompleted && lastCompleted.nextDueDate) {
        dueDate = lastCompleted.nextDueDate;
        startDate = toDateOnly(lastCompleted.completedAt) || today;
      } else if (lastCompleted) {
        dueDate = calculateNextDueDate(task, lastCompleted.completedAt || now, lastCompleted.dueDate, cfg);
        startDate = toDateOnly(lastCompleted.completedAt) || today;
      } else {
        dueDate = task.initialDueDate || today;
        if (task.workingDaysOnly) {
          dueDate = nextWorkingDay(dueDate, cfg, true);
        }
      }

      const cycleId = buildCycleId(task.taskId, dueDate, nextNum);
      if (cycles.some(c => c.cycleId === cycleId) || newRows.some(r => r[0] === cycleId)) {
        return;
      }

      const dueTime = task.dueTime || (task.frequencyType === FREQ_TYPES.DAILY ? cfg.dailyDeadline : cfg.portalDeadline);
      newRows.push([
        cycleId,
        task.taskId,
        task.taskName,
        task.assignedEmployee,
        nextNum,
        formatYMD(startDate),
        formatYMD(dueDate),
        dueTime,
        CYCLE_STATUS.PENDING,
        '', '', '', '', '', '', '', '', '',
        nowStr,
        nowStr
      ]);
    });

    if (newRows.length > 0) {
      const startRow = Math.max(CL_START_ROW, logSheet.getLastRow() + 1);
      logSheet.getRange(startRow, 1, newRows.length, CL_COL_COUNT).setValues(newRows);
      logSheet.getRange(startRow, CL_COLS.CYCLE_START_DATE, newRows.length, 2).setNumberFormat('yyyy-mm-dd');
    }

    return { created: newRows.length };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Transitions Pending cycles whose deadline has passed into Overdue (without moving due date forward).
 */
function runPendingCheck(referenceNow) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = referenceNow ? new Date(referenceNow) : new Date();
    const logSheet = ss.getSheetByName(SHEET_NAMES.CYCLE_LOG);
    const cycles = getAllCycles(ss);

    let newlyOverdue = 0;
    const nowStr = formatTimestamp(now);

    cycles.forEach(c => {
      if (c.status === CYCLE_STATUS.PENDING && c.dueDate) {
        const deadline = combineDateAndTime(c.dueDate, c.dueTime);
        if (deadline && now > deadline) {
          logSheet.getRange(c.row, CL_COLS.STATUS).setValue(CYCLE_STATUS.OVERDUE);
          logSheet.getRange(c.row, CL_COLS.UPDATED_AT).setValue(nowStr);
          newlyOverdue++;
        }
      }
    });

    refreshAllDashboards(now);
    return { newlyOverdue };
  } finally {
    lock.releaseLock();
  }
}
