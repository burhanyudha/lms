function normalizeAttendanceIdentity(value) {
  const identity = String(value == null ? "" : value).trim().toLowerCase();
  return /^\d+$/.test(identity) ? identity.replace(/^0+(?=\d)/, "") : identity;
}

function attendanceDateKey(value) {
  const raw = String(value == null ? "" : value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return raw.slice(0, 10);
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Jakarta", year: "numeric", month: "2-digit", day: "2-digit"
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map(part => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
}

if (typeof module === "object" && module.exports) module.exports = {normalizeAttendanceIdentity, attendanceDateKey};
