// Shims for the Apps Script globals that OutlookSync.gs's pure layer uses
// (Utilities.formatDate for timezone math, Utilities.computeDigest for
// hashing), plus a loader that evaluates the .gs file and returns its
// functions. Functions that touch Google services (Calendar, UrlFetchApp,
// ScriptApp, LockService, PropertiesService) are not exercised here; the
// tests cover the parse/transform layer, which is where sync bugs live.
'use strict';
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function tzParts(date, tz) {
  const dtf = new Intl.DateTimeFormat('en-US', {
    timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
  });
  const parts = {};
  dtf.formatToParts(date).forEach(p => { parts[p.type] = p.value; });
  if (parts.hour === '24') parts.hour = '00';
  return parts;
}

const Utilities = {
  DigestAlgorithm: { MD5: 'md5' },
  Charset: { UTF_8: 'utf8' },
  computeDigest(alg, s) {
    const buf = crypto.createHash(alg).update(s, 'utf8').digest();
    // Apps Script returns signed bytes; mimic that so the hex logic matches.
    return Array.from(buf).map(b => (b > 127 ? b - 256 : b));
  },
  formatDate(date, tz, pattern) {
    const p = tzParts(date, tz);
    if (pattern === 'Z') {
      const asUTC = Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour, +p.minute, +p.second);
      const offMin = Math.round((asUTC - date.getTime()) / 60000);
      const sign = offMin < 0 ? '-' : '+';
      const a = Math.abs(offMin);
      return sign + String(Math.floor(a / 60)).padStart(2, '0') + String(a % 60).padStart(2, '0');
    }
    if (pattern === "yyyyMMdd'T'HHmmss'Z'") {
      return `${p.year}${p.month}${p.day}T${p.hour}${p.minute}${p.second}Z`;
    }
    if (pattern === 'yyyy-MM-dd') return `${p.year}-${p.month}-${p.day}`;
    if (pattern === "yyyy-MM-dd'T'HH:mm:ss") {
      return `${p.year}-${p.month}-${p.day}T${p.hour}:${p.minute}:${p.second}`;
    }
    throw new Error('unshimmed formatDate pattern: ' + pattern);
  }
};

const PropertiesService = {
  getScriptProperties() {
    return { getProperty: () => null };
  }
};

function loadGs() {
  const source = fs.readFileSync(path.join(__dirname, '..', 'OutlookSync.gs'), 'utf8');
  const loader = new Function(
    'Utilities', 'console', 'PropertiesService',
    source + `
    return { parseIcs, groupByUid, buildEventResource, buildRecurrence,
             icsTimeToGoogle, mapTzid, normalizeUntil_, contentHash_,
             toRfc3339_, addDuration_, unescapeIcsText, unfoldIcsLines,
             localToUtc_, parseIcsDateTime_, parseUtcOffsetMinutes_,
             eventEndedBefore_, isCancelled_, getProp, isOwnedMaster_ };`
  );
  return loader(Utilities, console, PropertiesService);
}

module.exports = { loadGs, Utilities };
