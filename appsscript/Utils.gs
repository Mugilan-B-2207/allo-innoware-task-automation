/**
 * Utils.gs
 * Shared date/time, working-day/holiday, validation, ID generation, and structured logging utilities.
 */

function toBool(v) {
  if (typeof v === 'boolean') return v;
  const s = String(v || '').trim().toUpperCase();
  return s === 'TRUE' || s === 'YES' || s === '1';
}

function toDateOnly(val) {
  if (!val) return null;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }
  const str = String(val).trim();
  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(str);
  if (iso) {
    return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  }
  const parsed = new Date(str);
  if (isNaN(parsed.getTime())) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function parseDateTime(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : new Date(val.getTime());
  const str = String(val).trim();
  const m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?/.exec(str);
  if (m) {
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), Number(m[6] || 0));
  }
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function formatYMD(date) {
  const d = toDateOnly(date);
  if (!d) return '';
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function formatTimestamp(date) {
  const d = date instanceof Date ? date : parseDateTime(date);
  if (!d || isNaN(d.getTime())) return '';
  const hh = String(d.getHours()).padStart(2, '0');
  const mi = String(d.getMinutes()).padStart(2, '0');
  const ss = String(d.getSeconds()).padStart(2, '0');
  return `${formatYMD(d)} ${hh}:${mi}:${ss}`;
}

function parseHM(timeVal, fallback) {
  const fb = fallback || '17:00';
  if (timeVal instanceof Date) {
    return { hours: timeVal.getHours(), minutes: timeVal.getMinutes() };
  }
  const m = /^(\d{1,2}):(\d{2})/.exec(String(timeVal || fb).trim());
  if (m) return { hours: Math.min(23, Number(m[1])), minutes: Math.min(59, Number(m[2])) };
  const fm = /^(\d{1,2}):(\d{2})/.exec(fb);
  return { hours: Number(fm[1]), minutes: Number(fm[2]) };
}

function formatHM(timeVal, fallback) {
  const { hours, minutes } = parseHM(timeVal, fallback);
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
}

function combineDateAndTime(dateVal, timeVal) {
  const base = toDateOnly(dateVal);
  if (!base) return null;
  const { hours, minutes } = parseHM(timeVal, '17:00');
  return new Date(base.getFullYear(), base.getMonth(), base.getDate(), hours, minutes, 0, 0);
}

function isWorkingDay(date, cfg) {
  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const workingSet = new Set(cfg.workingDays || ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']);
  const holidaySet = new Set(cfg.holidays || []);
  const d = toDateOnly(date);
  if (!d) return false;
  return workingSet.has(dayNames[d.getDay()]) && !holidaySet.has(formatYMD(d));
}

function nextWorkingDay(date, cfg, allowSameDay) {
  let cursor = toDateOnly(date);
  if (!allowSameDay) {
    cursor = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate() + 1);
  }
  for (let i = 0; i < 45; i++) {
    if (isWorkingDay(cursor, cfg)) return cursor;
    cursor = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate() + 1);
  }
  return cursor;
}

function addCalendarDays(date, days) {
  const d = toDateOnly(date);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + Number(days));
}

function diffCalendarDays(laterDate, earlierDate) {
  const a = toDateOnly(laterDate);
  const b = toDateOnly(earlierDate);
  if (!a || !b) return 0;
  return Math.round((a.getTime() - b.getTime()) / 86400000);
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function buildCycleId(taskId, dueDate, cycleNumber) {
  const datePart = formatYMD(dueDate).replace(/-/g, '');
  const seqPart = String(cycleNumber).padStart(3, '0');
  return `${taskId}-${datePart}-${seqPart}`;
}

function generateShortId(prefix) {
  return `${prefix}-${Utilities.getUuid().slice(0, 8).toUpperCase()}`;
}

function resolveUserEmail(fallback) {
  try {
    const e = Session.getActiveUser().getEmail();
    if (e) return e;
  } catch (err) {
    // Simple trigger / anonymous contexts may not expose email
  }
  return fallback || 'system@allo.local';
}

function logSystemEvent(level, context, message, meta) {
  const payload = {
    ts: formatTimestamp(new Date()),
    level,
    context,
    message,
    meta: meta || {}
  };
  try {
    console.log(JSON.stringify(payload));
  } catch (e) {
    // No-op if console unavailable
  }
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
