/**
 * EscalationService.gs
 * Orchestrates SLA checks (Employee Reminder, Manager First Alert, Manager Follow-up,
 * Leadership Escalation) and Manager Acknowledgement processing.
 */

function runAlertAndEscalationCheck(referenceNow) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = referenceNow ? new Date(referenceNow) : new Date();
    const today = toDateOnly(now);
    const cfg = getConfig(ss);

    // Ensure statuses are up to date before alerting
    runPendingCheck(now);

    const tasks = getTasksFromMaster(ss);
    const taskMap = {};
    tasks.forEach(t => { taskMap[t.taskId] = t; });

    const cycles = getAllCycles(ss);
    const logSheet = ss.getSheetByName(SHEET_NAMES.CYCLE_LOG);
    const mdSheet = ss.getSheetByName(SHEET_NAMES.MGR_DASHBOARD);
    const mgrDraftMeta = mdSheet ? readManagerDraftMeta(mdSheet) : {};

    const summary = {
      reminders: 0,
      mgrFirstAlerts: 0,
      mgrFollowups: 0,
      leadershipEscalations: 0
    };

    const firstAlertDelayMs = cfg.mgrFirstAlertDelayMins * 60 * 1000;
    const leadDelayMs = cfg.leadershipEscDelayHrs * 3600 * 1000;

    cycles.forEach(c => {
      if (!isActiveCycleStatus(c.status) || !c.dueDate) return;
      const task = taskMap[c.taskId];
      if (!task || !task.active) return;

      const deadline = combineDateAndTime(c.dueDate, c.dueTime);
      const reminderMoment = combineDateAndTime(c.dueDate, cfg.reminderTime);
      const firstAlertMoment = new Date(deadline.getTime() + firstAlertDelayMs);

      // 1. Employee Reminder (at/after 09:00 on due date, before deadline)
      if (c.status === CYCLE_STATUS.PENDING && now >= reminderMoment && now <= deadline) {
        const res = sendEmployeeReminderAlert(ss, cfg, c, task, now);
        if (res.sent) summary.reminders++;
      }

      // 2. Same-Day Manager First Alert (at/after 17:15 on due date)
      if (now >= firstAlertMoment && !c.mgrAlertSentAt) {
        const res = sendManagerFirstAlertEmail(ss, cfg, c, task, now);
        if (res.sent || res.status === 'SKIPPED_DISABLED') {
          const ts = formatTimestamp(now);
          logSheet.getRange(c.row, CL_COLS.MGR_ALERT_SENT_AT).setValue(ts);
          logSheet.getRange(c.row, CL_COLS.UPDATED_AT).setValue(ts);
          c.mgrAlertSentAt = now;
          if (res.sent) summary.mgrFirstAlerts++;
        }
      }

      // 3. Manager Follow-up (next working day at 10:00 AM if unacknowledged)
      if (c.mgrAlertSentAt && !c.mgrAckAt) {
        const nextWorkDay = nextWorkingDay(c.mgrAlertSentAt, cfg, false);
        const followupMoment = combineDateAndTime(nextWorkDay, cfg.mgrFollowupTime);
        if (now >= followupMoment) {
          const res = sendManagerFollowupEmail(ss, cfg, c, task, now);
          if (res.sent) summary.mgrFollowups++;
        }
      }

      // 4. Leadership Escalation (24h calendar SLA OR unacked beyond ack deadline + early grace)
      if (cfg.leadershipEscEnabled && c.mgrAlertSentAt && !c.leadEscAt) {
        const sla24hMoment = new Date(c.mgrAlertSentAt.getTime() + leadDelayMs);
        const nextWorkDay = nextWorkingDay(c.mgrAlertSentAt, cfg, false);
        const ackDeadlineMoment = combineDateAndTime(nextWorkDay, cfg.mgrAckDeadlineTime);
        const earlyUnackedMoment = new Date(ackDeadlineMoment.getTime() + cfg.unackedEarlyEscHours * 3600 * 1000);

        let reason = '';
        if (now >= sla24hMoment) {
          reason = c.mgrAckAt
            ? `Unresolved ${cfg.leadershipEscDelayHrs}h after Manager First Alert (Manager acknowledged, task still pending)`
            : `Unresolved ${cfg.leadershipEscDelayHrs}h after Manager First Alert (Manager unacknowledged)`;
        } else if (!c.mgrAckAt && now >= earlyUnackedMoment) {
          reason = `Early escalation: Manager failed to acknowledge within SLA (${formatTimestamp(ackDeadlineMoment)} + ${cfg.unackedEarlyEscHours}h)`;
        }

        if (reason) {
          const notes = (mgrDraftMeta[c.cycleId] && mgrDraftMeta[c.cycleId].ackNotes) || '';
          const res = sendLeadershipEscalationEmail(ss, cfg, c, task, reason, notes, now);
          if (res.sent || res.status === 'SKIPPED_DISABLED') {
            const ts = formatTimestamp(now);
            logSheet.getRange(c.row, CL_COLS.LEAD_ESC_AT).setValue(ts);
            logSheet.getRange(c.row, CL_COLS.STATUS).setValue(CYCLE_STATUS.ESCALATED);
            logSheet.getRange(c.row, CL_COLS.UPDATED_AT).setValue(ts);
            if (res.sent) summary.leadershipEscalations++;
          }
        }
      }
    });

    refreshAllDashboards(now);
    return summary;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Processes Manager Acknowledgement from Manager Dashboard (never marks task Completed).
 */
function handleManagerDashboardEdit(ss, sheet, row, col, options) {
  if (row < MD_DETAIL_START_ROW || col !== MD_COLS.ACK_CHECKBOX) return null;
  const checked = toBool(sheet.getRange(row, MD_COLS.ACK_CHECKBOX).getValue());
  if (!checked) return null;
  return acknowledgeCycle(ss, row, options);
}

function acknowledgeCycle(ss, mdRow, options) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
    const opts = options || {};
    const now = opts.now ? new Date(opts.now) : new Date();
    const nowStr = formatTimestamp(now);

    const mdSheet = spreadsheet.getSheetByName(SHEET_NAMES.MGR_DASHBOARD);
    const logSheet = spreadsheet.getSheetByName(SHEET_NAMES.CYCLE_LOG);
    const alertSheet = spreadsheet.getSheetByName(SHEET_NAMES.ALERT_LOG);

    const rowVals = mdSheet.getRange(mdRow, 1, 1, MD_COL_COUNT).getValues()[0];
    const cycleId = String(rowVals[MD_COLS.CYCLE_ID - 1] || '').trim();
    if (!cycleId) return { ok: false, error: 'Missing Cycle ID on Manager Dashboard row.' };

    const cycles = getAllCycles(spreadsheet);
    const cycle = cycles.find(c => c.cycleId === cycleId);
    if (!cycle) return { ok: false, error: `Cycle ${cycleId} not found in Cycle Log.` };
    if (cycle.status === CYCLE_STATUS.COMPLETED) {
      return { ok: false, error: 'Completed cycles do not require acknowledgement.' };
    }

    const ackBy = opts.ackBy || resolveUserEmail(getConfig(spreadsheet).managerEmail);
    const ackNotes = String(opts.ackNotes || rowVals[MD_COLS.ACK_NOTES - 1] || '').trim();
    const expResDate = opts.expectedResolutionDate || rowVals[MD_COLS.EXPECTED_RES_DATE - 1] || '';

    mdSheet.getRange(mdRow, MD_COLS.ACK_CHECKBOX).setValue(true);
    mdSheet.getRange(mdRow, MD_COLS.ACK_BY).setValue(ackBy);
    mdSheet.getRange(mdRow, MD_COLS.ACK_AT).setValue(nowStr);
    if (ackNotes) mdSheet.getRange(mdRow, MD_COLS.ACK_NOTES).setValue(ackNotes);
    if (expResDate) mdSheet.getRange(mdRow, MD_COLS.EXPECTED_RES_DATE).setValue(formatYMD(toDateOnly(expResDate)));

    // Record acknowledgement against Cycle ID in Cycle Log; NEVER mark task Completed
    if (!cycle.mgrAckAt) {
      logSheet.getRange(cycle.row, CL_COLS.MGR_ACK_AT).setValue(nowStr);
    }
    if (cycle.status !== CYCLE_STATUS.ESCALATED) {
      logSheet.getRange(cycle.row, CL_COLS.STATUS).setValue(CYCLE_STATUS.ACKNOWLEDGED);
    }
    logSheet.getRange(cycle.row, CL_COLS.UPDATED_AT).setValue(nowStr);

    // Sync Alert Log acknowledgement columns
    if (alertSheet && alertSheet.getLastRow() >= AL_START_ROW) {
      const aVals = alertSheet.getRange(AL_START_ROW, 1, alertSheet.getLastRow() - AL_START_ROW + 1, AL_COL_COUNT).getValues();
      aVals.forEach((r, idx) => {
        if (String(r[AL_COLS.CYCLE_ID - 1] || '').trim() === cycleId &&
            (r[AL_COLS.ALERT_TYPE - 1] === ALERT_TYPES.MGR_FIRST_ALERT || r[AL_COLS.ALERT_TYPE - 1] === ALERT_TYPES.MGR_FOLLOWUP)) {
          alertSheet.getRange(AL_START_ROW + idx, AL_COLS.ACKNOWLEDGED).setValue(true);
          alertSheet.getRange(AL_START_ROW + idx, AL_COLS.ACK_AT).setValue(nowStr);
        }
      });
    }

    refreshAllDashboards(now);
    const ackDelayMins = cycle.mgrAlertSentAt
      ? Math.round((now.getTime() - cycle.mgrAlertSentAt.getTime()) / 60000)
      : 0;

    return {
      ok: true,
      cycleId,
      ackBy,
      ackAt: nowStr,
      ackDelayMins,
      underlyingTaskStatus: cycle.status === CYCLE_STATUS.ESCALATED ? CYCLE_STATUS.ESCALATED : CYCLE_STATUS.ACKNOWLEDGED
    };
  } finally {
    lock.releaseLock();
  }
}
