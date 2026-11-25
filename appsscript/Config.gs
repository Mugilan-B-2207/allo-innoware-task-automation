/**
 * Config.gs
 * Central constants, sheet names, 1-based column mappings, and Configuration reader.
 */

const SHEET_NAMES = {
  TASK_MASTER: 'Task Master',
  EMP_DASHBOARD: 'Employee Dashboard',
  CYCLE_LOG: 'Cycle Log',
  MGR_DASHBOARD: 'Manager Dashboard',
  ALERT_LOG: 'Alert Log',
  CONFIG: 'Configuration',
  TEST_RESULTS: 'Test Results'
};

const CYCLE_STATUS = {
  PENDING: 'Pending',
  OVERDUE: 'Overdue',
  ACKNOWLEDGED: 'Acknowledged',
  ESCALATED: 'Escalated',
  COMPLETED: 'Completed'
};

const FREQ_TYPES = {
  DAILY: 'Daily',
  WEEKLY: 'Weekly',
  EVERY_15: 'Every 15 Days',
  EVERY_30: 'Every 30 Days',
  CUSTOM: 'Custom'
};

const RECURRENCE_MODES = {
  COMPLETION_BASED: 'Completion Based',
  FIXED_SCHEDULE: 'Fixed Schedule'
};

const PROOF_TYPES = {
  TICKET_ID: 'Ticket ID',
  WHATSAPP_CONFIRM: 'Confirmation / WhatsApp Link'
};

const ALERT_TYPES = {
  EMP_REMINDER: 'Employee Reminder',
  MGR_FIRST_ALERT: 'Manager First Alert',
  MGR_FOLLOWUP: 'Manager Follow-up',
  LEADERSHIP_ESC: 'Leadership Escalation',
  COMPLETION_NOTICE: 'Completion Notification'
};

const SUBMISSION_ACTIONS = {
  READY: 'Pending Entry',
  SUBMIT: 'Submit Completion'
};

// Task Master (Header Row 4, Data Row 5+)
const TM_HEADER_ROW = 4;
const TM_START_ROW = 5;
const TM_COLS = {
  TASK_ID: 1,
  TASK_NAME: 2,
  DESCRIPTION: 3,
  ASSIGNED_EMPLOYEE: 4,
  EMPLOYEE_EMAIL: 5,
  FREQUENCY_TYPE: 6,
  FREQUENCY_DAYS: 7,
  RECURRENCE_MODE: 8,
  INITIAL_DUE_DATE: 9,
  DUE_TIME: 10,
  WORKING_DAYS_ONLY: 11,
  PROOF_TYPE: 12,
  PROOF_REQUIRED: 13,
  ACTIVE: 14,
  MANAGER_EMAIL: 15,
  LEADERSHIP_EMAIL: 16,
  CREATED_AT: 17,
  UPDATED_AT: 18
};
const TM_COL_COUNT = 18;

// Employee Dashboard (Header Row 5, Data Row 6+)
const ED_HEADER_ROW = 5;
const ED_START_ROW = 6;
const ED_COLS = {
  CYCLE_ID: 1,
  TASK_ID: 2,
  TASK_NAME: 3,
  ASSIGNED_EMPLOYEE: 4,
  DUE_DATE: 5,
  DUE_TIME: 6,
  STATUS: 7,
  DAYS_OVERDUE: 8,
  REQUIRED_PROOF: 9,
  TICKET_ID: 10,
  CONFIRMATION: 11,
  PROOF_LINK: 12,
  NOTES: 13,
  COMPLETED_BY: 14,
  COMPLETED_AT: 15,
  SUBMISSION_ACTION: 16
};
const ED_COL_COUNT = 16;

// Cycle Log (Header Row 4, Data Row 5+)
const CL_HEADER_ROW = 4;
const CL_START_ROW = 5;
const CL_COLS = {
  CYCLE_ID: 1,
  TASK_ID: 2,
  TASK_NAME: 3,
  ASSIGNED_EMPLOYEE: 4,
  CYCLE_NUMBER: 5,
  CYCLE_START_DATE: 6,
  DUE_DATE: 7,
  DUE_TIME: 8,
  STATUS: 9,
  TICKET_ID: 10,
  PROOF_LINK: 11,
  COMPLETION_NOTES: 12,
  COMPLETED_BY: 13,
  COMPLETED_AT: 14,
  NEXT_DUE_DATE: 15,
  MGR_ALERT_SENT_AT: 16,
  MGR_ACK_AT: 17,
  LEAD_ESC_AT: 18,
  CREATED_AT: 19,
  UPDATED_AT: 20
};
const CL_COL_COUNT = 20;

// Manager Dashboard (KPI Rows 4-5, Employee Group Table Rows 8-12, Active Detail Header Row 14, Data Row 15+)
const MD_DETAIL_HEADER_ROW = 14;
const MD_DETAIL_START_ROW = 15;
const MD_COLS = {
  EMPLOYEE: 1,
  TASK: 2,
  CYCLE_ID: 3,
  DUE_DATE: 4,
  STATUS: 5,
  DAYS_OVERDUE: 6,
  PROOF: 7,
  MGR_ACK_SUMMARY: 8,
  ESC_STATUS: 9,
  ACK_CHECKBOX: 10,
  ACK_NOTES: 11,
  ACK_BY: 12,
  ACK_AT: 13,
  EXPECTED_RES_DATE: 14
};
const MD_COL_COUNT = 14;

// Alert Log (Header Row 4, Data Row 5+)
const AL_HEADER_ROW = 4;
const AL_START_ROW = 5;
const AL_COLS = {
  ALERT_ID: 1,
  CYCLE_ID: 2,
  TASK_ID: 3,
  EMPLOYEE: 4,
  ALERT_TYPE: 5,
  RECIPIENT_EMAIL: 6,
  SUBJECT: 7,
  SENT_AT: 8,
  STATUS: 9,
  ACKNOWLEDGED: 10,
  ACK_AT: 11,
  ERROR_MESSAGE: 12
};
const AL_COL_COUNT = 12;

// Configuration (Header Row 4, Data Row 5+)
const CFG_HEADER_ROW = 4;
const CFG_START_ROW = 5;

const DEFAULT_CONFIG = {
  'Company Name': 'Allo Innoware',
  'Spreadsheet Timezone': 'Asia/Kolkata',
  'Manager Email': 'manager.placeholder@example.com',
  'Leadership Email': 'leadership.placeholder@example.com',
  'Working Days': 'Mon,Tue,Wed,Thu,Fri',
  'Working Hours Start': '09:00',
  'Working Hours End': '18:00',
  'Daily Task Deadline': '17:00',
  'Portal Task Deadline': '17:00',
  'Employee Reminder Time': '09:00',
  'Manager First Alert Delay (Mins)': '15',
  'Manager Follow-up Time': '10:00',
  'Manager Ack Deadline Time': '10:00',
  'Unacked Early Escalation Hours': '4',
  'Leadership Escalation Delay (Hrs)': '24',
  'Leadership Escalation Enabled': 'TRUE',
  'Email Notifications Enabled': 'FALSE',
  'Test Mode Enabled': 'TRUE',
  'Daily Automation Enabled': 'TRUE',
  'Daily Backlog Policy': 'KEEP_ONE_OVERDUE_BLOCK_NEW'
};

/**
 * Reads Configuration sheet + Holidays list into a normalized config object.
 */
function getConfig(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const cfg = Object.assign({}, DEFAULT_CONFIG);
  cfg.holidays = [];

  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.CONFIG);
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= CFG_START_ROW) {
      const vals = sheet.getRange(CFG_START_ROW, 1, lastRow - CFG_START_ROW + 1, 6).getValues();
      vals.forEach(r => {
        const key = String(r[0] || '').trim();
        const val = r[1];
        if (key) {
          cfg[key] = (val === '' || val === null || val === undefined) ? (cfg[key] || '') : formatConfigVal(val);
        }
        const holDate = r[4];
        if (holDate) {
          const ymd = formatYMD(toDateOnly(holDate));
          if (ymd) cfg.holidays.push(ymd);
        }
      });
    }
  }

  return {
    raw: cfg,
    companyName: String(cfg['Company Name'] || 'Allo Innoware'),
    timezone: String(cfg['Spreadsheet Timezone'] || 'Asia/Kolkata'),
    managerEmail: String(cfg['Manager Email'] || '').trim(),
    leadershipEmail: String(cfg['Leadership Email'] || '').trim(),
    workingDays: String(cfg['Working Days'] || 'Mon,Tue,Wed,Thu,Fri').split(',').map(s => s.trim()).filter(Boolean),
    workingHoursStart: String(cfg['Working Hours Start'] || '09:00'),
    workingHoursEnd: String(cfg['Working Hours End'] || '18:00'),
    dailyDeadline: String(cfg['Daily Task Deadline'] || '17:00'),
    portalDeadline: String(cfg['Portal Task Deadline'] || '17:00'),
    reminderTime: String(cfg['Employee Reminder Time'] || '09:00'),
    mgrFirstAlertDelayMins: Number(cfg['Manager First Alert Delay (Mins)']) || 15,
    mgrFollowupTime: String(cfg['Manager Follow-up Time'] || '10:00'),
    mgrAckDeadlineTime: String(cfg['Manager Ack Deadline Time'] || '10:00'),
    unackedEarlyEscHours: Number(cfg['Unacked Early Escalation Hours']) || 4,
    leadershipEscDelayHrs: Number(cfg['Leadership Escalation Delay (Hrs)']) || 24,
    leadershipEscEnabled: toBool(cfg['Leadership Escalation Enabled']),
    emailEnabled: toBool(cfg['Email Notifications Enabled']),
    testMode: toBool(cfg['Test Mode Enabled']),
    dailyAutoEnabled: toBool(cfg['Daily Automation Enabled']),
    dailyBacklogPolicy: String(cfg['Daily Backlog Policy'] || 'KEEP_ONE_OVERDUE_BLOCK_NEW'),
    holidays: cfg.holidays
  };
}

function formatConfigVal(val) {
  if (val instanceof Date) {
    const hh = String(val.getHours()).padStart(2, '0');
    const mm = String(val.getMinutes()).padStart(2, '0');
    return `${hh}:${mm}`;
  }
  return String(val).trim();
}

function setConfigSetting(ss, key, value) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow >= CFG_START_ROW) {
    const keys = sheet.getRange(CFG_START_ROW, 1, lastRow - CFG_START_ROW + 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0] || '').trim() === key) {
        sheet.getRange(CFG_START_ROW + i, 2).setValue(value);
        return;
      }
    }
  }
  sheet.getRange(Math.max(CFG_START_ROW, lastRow + 1), 1, 1, 3).setValues([[key, value, '']]);
}
