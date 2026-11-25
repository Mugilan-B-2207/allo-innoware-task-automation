/**
 * DashboardService.gs
 * Synchronizes Employee Dashboard (single-point workspace) and Manager Dashboard
 * (KPIs, employee breakdown, overdue risk table, and manager acknowledgement controls).
 */

function refreshAllDashboards(referenceNow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = referenceNow ? new Date(referenceNow) : new Date();
  const today = toDateOnly(now);

  const tasks = getTasksFromMaster(ss);
  const cycles = getAllCycles(ss);
  const taskMap = {};
  tasks.forEach(t => { taskMap[t.taskId] = t; });

  refreshEmployeeDashboard(ss, cycles, taskMap, today);
  refreshManagerDashboard(ss, tasks, cycles, taskMap, today);
}

function refreshEmployeeDashboard(ss, cycles, taskMap, today) {
  const edSheet = ss.getSheetByName(SHEET_NAMES.EMP_DASHBOARD);
  if (!edSheet) return;

  // Preserve any unsaved user inputs currently in Employee Dashboard keyed by Cycle ID
  const draftInputs = {};
  const oldLastRow = edSheet.getLastRow();
  if (oldLastRow >= ED_START_ROW) {
    const oldVals = edSheet.getRange(ED_START_ROW, 1, oldLastRow - ED_START_ROW + 1, ED_COL_COUNT).getValues();
    oldVals.forEach(r => {
      const cId = String(r[ED_COLS.CYCLE_ID - 1] || '').trim();
      if (cId) {
        draftInputs[cId] = {
          ticketId: r[ED_COLS.TICKET_ID - 1],
          confirmed: r[ED_COLS.CONFIRMATION - 1],
          proofLink: r[ED_COLS.PROOF_LINK - 1],
          notes: r[ED_COLS.NOTES - 1]
        };
      }
    });
    edSheet.getRange(ED_START_ROW, 1, oldLastRow - ED_START_ROW + 1, ED_COL_COUNT).clearContent();
  }

  const priorityOrder = {
    [CYCLE_STATUS.ESCALATED]: 1,
    [CYCLE_STATUS.OVERDUE]: 2,
    [CYCLE_STATUS.ACKNOWLEDGED]: 3,
    [CYCLE_STATUS.PENDING]: 4
  };

  const activeCycles = cycles
    .filter(c => isActiveCycleStatus(c.status))
    .sort((a, b) => (priorityOrder[a.status] || 9) - (priorityOrder[b.status] || 9) || a.dueDate - b.dueDate);

  if (activeCycles.length === 0) return;

  const rows = activeCycles.map(c => {
    const task = taskMap[c.taskId] || {};
    const draft = draftInputs[c.cycleId] || {};
    const daysOverdue = c.dueDate ? Math.max(0, diffCalendarDays(today, c.dueDate)) : 0;
    const reqProof = task.proofType === PROOF_TYPES.TICKET_ID
      ? 'Mandatory Ticket ID'
      : 'Checkbox Confirmation + Optional WhatsApp Link';

    return [
      c.cycleId,
      c.taskId,
      c.taskName,
      c.assignedEmployee,
      formatYMD(c.dueDate),
      c.dueTime,
      c.status,
      daysOverdue,
      reqProof,
      draft.ticketId || c.ticketId || '',
      draft.confirmed === true,
      draft.proofLink || c.proofLink || '',
      draft.notes || c.completionNotes || '',
      c.completedBy || '',
      c.completedAt ? formatTimestamp(c.completedAt) : '',
      SUBMISSION_ACTIONS.READY
    ];
  });

  edSheet.getRange(ED_START_ROW, 1, rows.length, ED_COL_COUNT).setValues(rows);
  edSheet.getRange(ED_START_ROW, ED_COLS.CONFIRMATION, rows.length, 1).insertCheckboxes();
  edSheet.getRange(ED_START_ROW, ED_COLS.DUE_DATE, rows.length, 1).setNumberFormat('yyyy-mm-dd');
}

function refreshManagerDashboard(ss, tasks, cycles, taskMap, today) {
  const mdSheet = ss.getSheetByName(SHEET_NAMES.MGR_DASHBOARD);
  if (!mdSheet) return;

  // Preserve manager notes & expected resolution dates keyed by Cycle ID
  const mgrMeta = readManagerDraftMeta(mdSheet);

  const activeCycles = cycles.filter(c => isActiveCycleStatus(c.status));
  const pendingCount = activeCycles.filter(c => c.status === CYCLE_STATUS.PENDING).length;
  const overdueCycles = activeCycles.filter(c =>
    c.status === CYCLE_STATUS.OVERDUE ||
    c.status === CYCLE_STATUS.ACKNOWLEDGED ||
    c.status === CYCLE_STATUS.ESCALATED
  );
  const completedToday = cycles.filter(c =>
    c.status === CYCLE_STATUS.COMPLETED && c.completedAt && formatYMD(c.completedAt) === formatYMD(today)
  ).length;
  const awaitingAck = activeCycles.filter(c => Boolean(c.mgrAlertSentAt) && !c.mgrAckAt).length;
  const escalatedCount = activeCycles.filter(c => c.status === CYCLE_STATUS.ESCALATED).length;

  // Update KPI Cards (Row 5)
  mdSheet.getRange(5, 1, 1, 6).setValues([[
    tasks.filter(t => t.active).length,
    pendingCount,
    overdueCycles.length,
    completedToday,
    awaitingAck,
    escalatedCount
  ]]);

  // Update Employee Grouped Summary (Rows 8+)
  const empSet = Array.from(new Set(tasks.map(t => t.assignedEmployee).filter(Boolean)));
  const groupRows = empSet.map(emp => {
    const empActive = activeCycles.filter(c => c.assignedEmployee === emp);
    const empPending = empActive.filter(c => c.status === CYCLE_STATUS.PENDING).length;
    const empOverdue = empActive.filter(c =>
      c.status === CYCLE_STATUS.OVERDUE || c.status === CYCLE_STATUS.ACKNOWLEDGED || c.status === CYCLE_STATUS.ESCALATED
    ).length;
    const empCompToday = cycles.filter(c =>
      c.assignedEmployee === emp && c.status === CYCLE_STATUS.COMPLETED && c.completedAt && formatYMD(c.completedAt) === formatYMD(today)
    ).length;
    const empAwaitAck = empActive.filter(c => Boolean(c.mgrAlertSentAt) && !c.mgrAckAt).length;
    const empEsc = empActive.filter(c => c.status === CYCLE_STATUS.ESCALATED).length;
    return [emp, empPending, empOverdue, empCompToday, empAwaitAck, empEsc];
  });

  mdSheet.getRange(8, 1, 5, 6).clearContent();
  if (groupRows.length > 0) {
    mdSheet.getRange(8, 1, groupRows.length, 6).setValues(groupRows);
  }

  // Update Detailed Pending/Active Task Table (Row 15+)
  const oldLastRow = mdSheet.getLastRow();
  if (oldLastRow >= MD_DETAIL_START_ROW) {
    mdSheet.getRange(MD_DETAIL_START_ROW, 1, oldLastRow - MD_DETAIL_START_ROW + 1, MD_COL_COUNT).clearContent();
  }

  if (activeCycles.length === 0) return;

  const detailRows = activeCycles.map(c => {
    const saved = mgrMeta[c.cycleId] || {};
    const daysOverdue = c.dueDate ? Math.max(0, diffCalendarDays(today, c.dueDate)) : 0;
    const proofSummary = c.ticketId || c.proofLink || 'Awaiting Proof';
    const ackSummary = c.mgrAckAt
      ? `Acknowledged (${formatTimestamp(c.mgrAckAt)})`
      : (c.mgrAlertSentAt ? 'Awaiting Manager Ack' : 'Not Alerted');
    const escStatus = c.leadEscAt
      ? `Escalated (${formatTimestamp(c.leadEscAt)})`
      : (c.mgrAlertSentAt ? 'Monitoring L1' : 'Within SLA');

    return [
      c.assignedEmployee,
      c.taskName,
      c.cycleId,
      formatYMD(c.dueDate),
      c.status,
      daysOverdue,
      proofSummary,
      ackSummary,
      escStatus,
      Boolean(c.mgrAckAt),
      saved.ackNotes || '',
      saved.ackBy || '',
      c.mgrAckAt ? formatTimestamp(c.mgrAckAt) : '',
      saved.expectedResDate || ''
    ];
  });

  mdSheet.getRange(MD_DETAIL_START_ROW, 1, detailRows.length, MD_COL_COUNT).setValues(detailRows);
  mdSheet.getRange(MD_DETAIL_START_ROW, MD_COLS.ACK_CHECKBOX, detailRows.length, 1).insertCheckboxes();
  mdSheet.getRange(MD_DETAIL_START_ROW, MD_COLS.DUE_DATE, detailRows.length, 1).setNumberFormat('yyyy-mm-dd');
}

function readManagerDraftMeta(mdSheet) {
  const map = {};
  const lastRow = mdSheet.getLastRow();
  if (lastRow < MD_DETAIL_START_ROW) return map;
  const vals = mdSheet.getRange(MD_DETAIL_START_ROW, 1, lastRow - MD_DETAIL_START_ROW + 1, MD_COL_COUNT).getValues();
  vals.forEach(r => {
    const cycleId = String(r[MD_COLS.CYCLE_ID - 1] || '').trim();
    if (cycleId) {
      map[cycleId] = {
        ackNotes: r[MD_COLS.ACK_NOTES - 1] || '',
        ackBy: r[MD_COLS.ACK_BY - 1] || '',
        expectedResDate: r[MD_COLS.EXPECTED_RES_DATE - 1] ? formatYMD(toDateOnly(r[MD_COLS.EXPECTED_RES_DATE - 1])) : ''
      };
    }
  });
  return map;
}
