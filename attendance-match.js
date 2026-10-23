function normalizeAttendanceIdentity(value) {
  const identity = String(value == null ? "" : value).trim().toLowerCase();
  return /^\d+$/.test(identity) ? identity.replace(/^0+(?=\d)/, "") : identity;
}

if (typeof module === "object" && module.exports) module.exports = normalizeAttendanceIdentity;
