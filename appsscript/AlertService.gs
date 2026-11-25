/**
 * AlertService.gs
 * Handles deduplicated email notifications, templates, Test Mode logging,
 * missing-email safeguards, and Alert Log auditing.
 */

function hasAlertAlreadyBeenLogged(ss, cycleId, alertType, recipientEmail) {
  const sheet = ss.getSheetByName(SHEET_NAMES.ALERT_LOG);
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < AL_START_ROW) return false;

  const vals = sheet.getRange(AL_START_ROW, 1, lastRow - AL_START_ROW + 1, AL_COL_COUNT).getValues();
  const targetRec = String(recipientEmail || '').trim().toLowerCase();
  return vals.some(r =>
    String(r[AL_COLS.CYCLE_ID - 1] || '').trim() === cycleId &&
    String(r[AL_COLS.ALERT_TYPE - 1] || '').trim() === alertType &&
    String(r[AL_COLS.RECIPIENT_EMAIL - 1] || '').trim().toLowerCase() === targetRec &&
    (r[AL_COLS.STATUS - 1] === 'SENT' || r[AL_COLS.STATUS - 1] === 'TEST_MODE_LOGGED')
  );
}

function dispatchAlert(ss, cfg, payload) {
  const alertSheet = ss.getSheetByName(SHEET_NAMES.ALERT_LOG);
  const nowStr = formatTimestamp(payload.now || new Date());
  const recipient = String(payload.recipientEmail || '').trim();

  // Deduplication check: same cycle + alert type + recipient
  if (hasAlertAlreadyBeenLogged(ss, payload.cycleId, payload.alertType, recipient)) {
    return { sent: false, status: 'DUPLICATE_SKIPPED' };
  }

  const alertId = generateShortId('ALT');
  let status = 'SENT';
  let errorMsg = '';

  if (!recipient || !isValidEmailList(recipient)) {
    status = 'FAILED';
    errorMsg = 'Missing or invalid recipient email address.';
  } else if (cfg.testMode) {
    status = 'TEST_MODE_LOGGED';
  } else if (!cfg.emailEnabled) {
    status = 'SKIPPED_DISABLED';
    errorMsg = 'Email Notifications Enabled is FALSE in Configuration.';
  } else {
    try {
      MailApp.sendEmail({
        to: recipient,
        subject: payload.subject,
        htmlBody: payload.htmlBody,
        body: payload.textBody || payload.subject
      });
    } catch (err) {
      status = 'FAILED';
      errorMsg = err.message || String(err);
    }
  }

  alertSheet.appendRow([
    alertId,
    payload.cycleId,
    payload.taskId,
    payload.employee,
    payload.alertType,
    recipient || '(missing)',
    payload.subject,
    nowStr,
    status,
    false,
    '',
    errorMsg
  ]);

  const newRow = alertSheet.getLastRow();
  alertSheet.getRange(newRow, AL_COLS.ACKNOWLEDGED).insertCheckboxes();

  return {
    sent: status === 'SENT' || status === 'TEST_MODE_LOGGED',
    status,
    alertId,
    errorMsg
  };
}

function isValidEmailList(str) {
  const parts = String(str || '').split(',').map(s => s.trim()).filter(Boolean);
  return parts.length > 0 && parts.every(isValidEmail);
}

function sendEmployeeReminderAlert(ss, cfg, cycle, task, now) {
  const recipient = task.employeeEmail;
  const sheetUrl = ss.getUrl ? ss.getUrl() : '';
  const proofLabel = task.proofType === PROOF_TYPES.TICKET_ID ? 'Portal Ticket ID' : 'WhatsApp Update Confirmation';
  const subject = `[Reminder] Due Today (${formatYMD(cycle.dueDate)} ${cycle.dueTime}): ${cycle.taskName}`;

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:640px;color:#202124;">
      <h3 style="color:#1a73e8;">Task Reminder — ${escapeHtml(cfg.companyName)}</h3>
      <p>Hi ${escapeHtml(cycle.assignedEmployee)},</p>
      <p>This recurring task is due today:</p>
      <ul>
        <li><strong>Task:</strong> ${escapeHtml(cycle.taskName)} (${escapeHtml(cycle.cycleId)})</li>
        <li><strong>Due:</strong> ${escapeHtml(formatYMD(cycle.dueDate))} at ${escapeHtml(cycle.dueTime)} (${escapeHtml(cfg.timezone)})</li>
        <li><strong>Required Proof:</strong> ${escapeHtml(proofLabel)}</li>
      </ul>
      ${sheetUrl ? `<p><a href="${escapeHtml(sheetUrl)}">Open Employee Dashboard</a></p>` : ''}
    </div>`;

  return dispatchAlert(ss, cfg, {
    now,
    cycleId: cycle.cycleId,
    taskId: cycle.taskId,
    employee: cycle.assignedEmployee,
    alertType: ALERT_TYPES.EMP_REMINDER,
    recipientEmail: recipient,
    subject,
    htmlBody,
    textBody: `${subject} | Required Proof: ${proofLabel}`
  });
}

function sendManagerFirstAlertEmail(ss, cfg, cycle, task, now) {
  const recipient = task.managerEmail || cfg.managerEmail;
  const sheetUrl = ss.getUrl ? ss.getUrl() : '';
  const daysOverdue = Math.max(0, diffCalendarDays(toDateOnly(now), cycle.dueDate));
  const subject = `[Manager Alert] Overdue Task: ${cycle.taskName} (${cycle.assignedEmployee}) - ${cycle.cycleId}`;

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:660px;color:#202124;">
      <div style="background:#fbbc04;padding:12px 18px;border-radius:6px 6px 0 0;">
        <h3 style="margin:0;">Same-Day Manager Alert — Pending Task Past Deadline</h3>
      </div>
      <div style="border:1px solid #dadce0;padding:18px;border-radius:0 0 6px 6px;">
        <p><strong>Employee:</strong> ${escapeHtml(cycle.assignedEmployee)}</p>
        <p><strong>Task:</strong> ${escapeHtml(cycle.taskName)} | <strong>Cycle ID:</strong> ${escapeHtml(cycle.cycleId)}</p>
        <p><strong>Due:</strong> ${escapeHtml(formatYMD(cycle.dueDate))} ${escapeHtml(cycle.dueTime)} | <strong>Days Overdue:</strong> ${daysOverdue}</p>
        <p><strong>Current Status:</strong> ${escapeHtml(cycle.status)} | <strong>Ticket / Proof:</strong> ${escapeHtml(cycle.ticketId || cycle.proofLink || 'Not submitted')}</p>
        <p><strong>Required Manager Action:</strong> Open the Manager Dashboard, review the task, check <em>Acknowledge</em>, and follow up with the employee before ${escapeHtml(cfg.mgrAckDeadlineTime)} next working day.</p>
        ${sheetUrl ? `<p><a href="${escapeHtml(sheetUrl)}">Open Manager Dashboard</a></p>` : ''}
      </div>
    </div>`;

  return dispatchAlert(ss, cfg, {
    now,
    cycleId: cycle.cycleId,
    taskId: cycle.taskId,
    employee: cycle.assignedEmployee,
    alertType: ALERT_TYPES.MGR_FIRST_ALERT,
    recipientEmail: recipient,
    subject,
    htmlBody,
    textBody: `${subject} | Due ${formatYMD(cycle.dueDate)} ${cycle.dueTime}`
  });
}

function sendManagerFollowupEmail(ss, cfg, cycle, task, now) {
  const recipient = task.managerEmail || cfg.managerEmail;
  const sheetUrl = ss.getUrl ? ss.getUrl() : '';
  const subject = `[Manager Follow-Up] Unacknowledged Overdue Task: ${cycle.taskName} (${cycle.cycleId})`;

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:660px;color:#202124;">
      <h3 style="color:#e65100;">Manager Follow-Up — Acknowledgement Still Pending</h3>
      <p>Task <strong>${escapeHtml(cycle.taskName)}</strong> assigned to <strong>${escapeHtml(cycle.assignedEmployee)}</strong> (${escapeHtml(cycle.cycleId)}) remains unresolved and unacknowledged since ${escapeHtml(formatTimestamp(cycle.mgrAlertSentAt))}.</p>
      <p>If not acknowledged and resolved within SLA, it will automatically escalate to Leadership.</p>
      ${sheetUrl ? `<p><a href="${escapeHtml(sheetUrl)}">Acknowledge in Manager Dashboard</a></p>` : ''}
    </div>`;

  return dispatchAlert(ss, cfg, {
    now,
    cycleId: cycle.cycleId,
    taskId: cycle.taskId,
    employee: cycle.assignedEmployee,
    alertType: ALERT_TYPES.MGR_FOLLOWUP,
    recipientEmail: recipient,
    subject,
    htmlBody,
    textBody: subject
  });
}

function sendLeadershipEscalationEmail(ss, cfg, cycle, task, reason, mgrNotes, now) {
  const recipient = task.leadershipEmail || cfg.leadershipEmail;
  const sheetUrl = ss.getUrl ? ss.getUrl() : '';
  const overdueHrs = cycle.dueDate
    ? Math.max(0, (now.getTime() - combineDateAndTime(cycle.dueDate, cycle.dueTime).getTime()) / 3600000).toFixed(1)
    : '0';
  const subject = `[LEADERSHIP ESCALATION] ${cycle.taskName} (${cycle.assignedEmployee}) - ${cycle.cycleId}`;

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:680px;color:#202124;">
      <div style="background:#a50e0e;color:#ffffff;padding:12px 18px;border-radius:6px 6px 0 0;">
        <h3 style="margin:0;">Leadership Escalation — Unresolved Recurring Task</h3>
      </div>
      <div style="border:1px solid #dadce0;padding:18px;border-radius:0 0 6px 6px;">
        <ul>
          <li><strong>Task:</strong> ${escapeHtml(cycle.taskName)} (${escapeHtml(cycle.cycleId)})</li>
          <li><strong>Employee:</strong> ${escapeHtml(cycle.assignedEmployee)}</li>
          <li><strong>Due:</strong> ${escapeHtml(formatYMD(cycle.dueDate))} ${escapeHtml(cycle.dueTime)} (${overdueHrs} hrs overdue)</li>
          <li><strong>Manager Alert Sent:</strong> ${escapeHtml(formatTimestamp(cycle.mgrAlertSentAt) || 'N/A')}</li>
          <li><strong>Manager Acknowledged:</strong> ${cycle.mgrAckAt ? `Yes (${escapeHtml(formatTimestamp(cycle.mgrAckAt))})` : 'No'}</li>
          <li><strong>Manager Notes:</strong> ${escapeHtml(mgrNotes || 'None')}</li>
          <li><strong>Ticket ID / Proof:</strong> ${escapeHtml(cycle.ticketId || cycle.proofLink || 'None')}</li>
          <li><strong>Escalation Reason:</strong> ${escapeHtml(reason)}</li>
        </ul>
        ${sheetUrl ? `<p><a href="${escapeHtml(sheetUrl)}">Review in Manager Dashboard</a></p>` : ''}
      </div>
    </div>`;

  return dispatchAlert(ss, cfg, {
    now,
    cycleId: cycle.cycleId,
    taskId: cycle.taskId,
    employee: cycle.assignedEmployee,
    alertType: ALERT_TYPES.LEADERSHIP_ESC,
    recipientEmail: recipient,
    subject,
    htmlBody,
    textBody: `${subject} | Reason: ${reason}`
  });
}

function sendCompletionResolutionNotice(ss, cfg, cycle, task, now) {
  const recipients = [task.managerEmail || cfg.managerEmail, cycle.leadEscAt ? (task.leadershipEmail || cfg.leadershipEmail) : '']
    .filter(Boolean)
    .join(',');
  const subject = `[Resolved] Overdue Task Completed: ${cycle.taskName} (${cycle.cycleId})`;
  const htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:640px;color:#202124;">
      <h3 style="color:#137333;">Overdue Task Resolved</h3>
      <p><strong>${escapeHtml(cycle.taskName)}</strong> (${escapeHtml(cycle.cycleId)}) assigned to <strong>${escapeHtml(cycle.assignedEmployee)}</strong> has now been completed.</p>
      <ul>
        <li><strong>Completed By:</strong> ${escapeHtml(cycle.completedBy)}</li>
        <li><strong>Completed At:</strong> ${escapeHtml(formatTimestamp(cycle.completedAt))}</li>
        <li><strong>Proof / Ticket ID:</strong> ${escapeHtml(cycle.ticketId || cycle.proofLink || 'Confirmed')}</li>
      </ul>
    </div>`;

  return dispatchAlert(ss, cfg, {
    now,
    cycleId: cycle.cycleId,
    taskId: cycle.taskId,
    employee: cycle.assignedEmployee,
    alertType: ALERT_TYPES.COMPLETION_NOTICE,
    recipientEmail: recipients,
    subject,
    htmlBody,
    textBody: subject
  });
}
