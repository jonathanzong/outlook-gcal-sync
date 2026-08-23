'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { loadGs } = require('./shims');

const gs = loadGs();

// Outlook serves CRLF line endings; the fixture is stored with LF, so
// re-expand to CRLF to exercise the same input the script sees in production.
const feed = fs
  .readFileSync(path.join(__dirname, 'fixture.ics'), 'utf8')
  .replace(/\r?\n/g, '\r\n');

const parsed = gs.parseIcs(feed);
const groups = gs.groupByUid(parsed.events);

test('parses 9 VEVENTs into 7 UID groups', () => {
  assert.equal(parsed.events.length, 9);
  assert.equal(Object.keys(groups).length, 7);
});

test('folded and escaped description round-trips', () => {
  const g = groups['AAMkAGRlY2Ix-recurring-1'];
  const res = gs.buildEventResource(g.master, parsed.tzMap, 'AAMkAGRlY2Ix-recurring-1', 'h');
  assert.equal(res.description,
    'Agenda: review the draft, then plan next steps. Bring your laptop and arrive early.\n');
});

test('Windows TZID maps to IANA on start and end', () => {
  const g = groups['AAMkAGRlY2Ix-recurring-1'];
  const res = gs.buildEventResource(g.master, parsed.tzMap, 'u', 'h');
  assert.deepEqual(res.start, { dateTime: '2026-08-14T11:00:00', timeZone: 'America/Denver' });
  assert.deepEqual(res.end, { dateTime: '2026-08-14T11:30:00', timeZone: 'America/Denver' });
});

test('RRULE passes through and EXDATE is re-emitted with an IANA TZID', () => {
  const g = groups['AAMkAGRlY2Ix-recurring-1'];
  const res = gs.buildEventResource(g.master, parsed.tzMap, 'u', 'h');
  assert.deepEqual(res.recurrence, [
    'RRULE:FREQ=WEEKLY;UNTIL=20261127T180000Z;INTERVAL=1;BYDAY=FR;WKST=SU',
    'EXDATE;TZID=America/Denver:20260904T110000,20260911T110000'
  ]);
});

test('RECURRENCE-ID override groups under the master UID and converts to RFC3339', () => {
  const g = groups['AAMkAGRlY2Ix-recurring-1'];
  assert.equal(g.overrides.length, 2); // one moved occurrence, one cancelled
  const moved = g.overrides.find(o => /moved/.test(gs.getProp(o, 'SUMMARY').value));
  const rid = moved.props['RECURRENCE-ID'][0];
  const orig = gs.icsTimeToGoogle(rid.value, rid.params, parsed.tzMap);
  // August in Denver is UTC-6 (daylight time).
  assert.equal(gs.toRfc3339_(orig.dateTime, orig.timeZone), '2026-08-21T11:00:00-06:00');
});

test('custom Outlook TZID resolves through its VTIMEZONE offset to a DST zone', () => {
  const res = gs.buildEventResource(groups['custom-tz-single'].master, parsed.tzMap, 'u', 'h');
  assert.equal(res.start.timeZone, 'America/Denver');
});

test('custom TZID without a DAYLIGHT rule resolves to a fixed Etc/GMT zone', () => {
  const mini = [
    'BEGIN:VCALENDAR',
    'BEGIN:VTIMEZONE',
    'TZID:Weird Fixed Zone',
    'BEGIN:STANDARD',
    'TZOFFSETFROM:-0700',
    'TZOFFSETTO:-0700',
    'END:STANDARD',
    'END:VTIMEZONE',
    'BEGIN:VEVENT',
    'UID:f1',
    'DTSTART;TZID=Weird Fixed Zone:20260601T120000',
    'END:VEVENT',
    'END:VCALENDAR'
  ].join('\r\n');
  const p = gs.parseIcs(mini);
  // Etc/GMT signs are inverted: Etc/GMT+7 means UTC-7.
  assert.equal(p.tzMap['Weird Fixed Zone'], 'Etc/GMT+7');
});

test('all-day event uses date form and keeps the exclusive DTEND', () => {
  const res = gs.buildEventResource(groups['all-day-1'].master, parsed.tzMap, 'u', 'h');
  assert.deepEqual(res.start, { date: '2026-11-10' });
  assert.deepEqual(res.end, { date: '2026-11-12' });
});

test('UTC start with DURATION and no DTEND computes the end', () => {
  const res = gs.buildEventResource(groups['utc-duration-1'].master, parsed.tzMap, 'u', 'h');
  assert.deepEqual(res.start, { dateTime: '2026-09-10T15:00:00Z', timeZone: 'Etc/UTC' });
  assert.deepEqual(res.end, { dateTime: '2026-09-10T16:30:00', timeZone: 'Etc/UTC' });
});

test('local-time UNTIL is normalized to UTC (New York, December = UTC-5)', () => {
  const res = gs.buildEventResource(groups['local-until-1'].master, parsed.tzMap, 'u', 'h');
  assert.equal(res.recurrence[0], 'RRULE:FREQ=WEEKLY;UNTIL=20261221T143000Z;BYDAY=MO');
});

test('DST boundary: same wall time maps to different offsets across spring-forward', () => {
  // DST starts 2026-03-08 in the US.
  assert.equal(gs.toRfc3339_('2026-03-07T09:00:00', 'America/Denver'), '2026-03-07T09:00:00-07:00');
  assert.equal(gs.toRfc3339_('2026-03-09T09:00:00', 'America/Denver'), '2026-03-09T09:00:00-06:00');
});

test('content hash ignores DTSTAMP churn but detects real changes', () => {
  const g = groups['all-day-1'];
  const h1 = gs.contentHash_(g);
  const stampChanged = {
    master: { raw: g.master.raw.replace('DTSTAMP:20260807T203614Z', 'DTSTAMP:20990101T000000Z') },
    overrides: []
  };
  assert.equal(gs.contentHash_(stampChanged), h1);
  const titleChanged = {
    master: { raw: g.master.raw.replace('Conference travel', 'Conference travel!') },
    overrides: []
  };
  assert.notEqual(gs.contentHash_(titleChanged), h1);
});

test('STATUS:CANCELLED master is detected', () => {
  assert.equal(gs.isCancelled_(groups['cancelled-1'].master), true);
});

test('Exchange title-prefixed cancellation is detected despite STATUS:CONFIRMED', () => {
  const master = groups['title-cancelled-1'].master;
  assert.equal(gs.getProp(master, 'STATUS').value, 'CONFIRMED');
  assert.equal(gs.isCancelled_(master), true);
});

test('title-prefixed cancellation of one occurrence is detected', () => {
  const cancelledOverride = groups['AAMkAGRlY2Ix-recurring-1'].overrides
    .find(o => /Canceled/.test(gs.getProp(o, 'SUMMARY').value));
  assert.equal(gs.isCancelled_(cancelledOverride), true);
});

test('both English spellings match, and ordinary titles do not', () => {
  const mk = summary => gs.parseIcs(
    'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:x\r\nSUMMARY:' + summary +
    '\r\nEND:VEVENT\r\nEND:VCALENDAR').events[0];
  assert.equal(gs.isCancelled_(mk('Cancelled: Standup')), true);
  assert.equal(gs.isCancelled_(mk('Canceled: Standup')), true);
  // A meeting *about* cancellations is not itself cancelled.
  assert.equal(gs.isCancelled_(mk('Cancelled flights debrief')), false);
  assert.equal(gs.isCancelled_(mk('Re: canceled: policy')), false);
});

test('quoted TZID parameter parses', () => {
  const ev = gs.parseIcs(
    'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:q1\r\n' +
    'DTSTART;TZID="W. Europe Standard Time":20260501T100000\r\n' +
    'END:VEVENT\r\nEND:VCALENDAR'
  );
  const dt = ev.events[0].props['DTSTART'][0];
  assert.equal(gs.icsTimeToGoogle(dt.value, dt.params, {}).timeZone, 'Europe/Berlin');
});

test('UTC offset strings parse with and without colons', () => {
  assert.equal(gs.parseUtcOffsetMinutes_('-0700'), -420);
  assert.equal(gs.parseUtcOffsetMinutes_('+05:45'), 345);
  assert.equal(gs.parseUtcOffsetMinutes_('+0000'), 0);
});

test('unknown TZID falls back to the default zone', () => {
  assert.equal(gs.mapTzid('Totally Unknown Zone', {}), 'America/Denver');
});

test('color fields are omitted by default (calendar default color applies)', () => {
  const res = gs.buildEventResource(groups['all-day-1'].master, parsed.tzMap, 'u', 'h');
  assert.equal('colorId' in res, false);
  assert.equal('eventLabelId' in res, false);
});

test('ownership check keeps masters and rejects recurring exceptions', () => {
  const tagged = { extendedProperties: { private: { icsSyncTag: 'outlook-ics-sync', icsUid: 'u1' } } };
  assert.equal(gs.isOwnedMaster_(tagged), true);
  // A modified occurrence inherits the master's extended properties, so it
  // carries the same tag and UID; updating it with a recurrence-bearing
  // resource fails with "Invalid start time".
  const exception = Object.assign({ recurringEventId: 'abc123' }, tagged);
  assert.equal(gs.isOwnedMaster_(exception), false);
  // An event the user created themselves is never touched.
  assert.equal(gs.isOwnedMaster_({ summary: 'Dinner' }), false);
});

test('transient API errors are retried, permanent ones are not', () => {
  const retried = [
    'API call to calendar.events.update failed with error: Rate Limit Exceeded',
    'User Rate Limit Exceeded',
    'Quota exceeded for quota metric',
    'Backend Error',
    'Service invoked too many times',   // matches "too many times"? see below
    'The service is temporarily unavailable',
    'Internal error encountered (500)'
  ];
  // "Service invoked too many times" is Apps Script's own quota message; it
  // must be retried like Google's rate-limit wording.
  retried.forEach(m => assert.equal(gs.isTransientApiError_(new Error(m)), true, m));
  const notRetried = [
    'Invalid start time',
    'Not Found',
    'Invalid value for: eventLabelId',
    'Required parameter is missing'
  ];
  notRetried.forEach(m => assert.equal(gs.isTransientApiError_(new Error(m)), false, m));
});

test('backoff grows exponentially, stays jittered, and is capped', () => {
  const lo = a => gs.backoffDelayMs_(a, () => 0);
  const hi = a => gs.backoffDelayMs_(a, () => 0.999);
  assert.equal(lo(1), 500);          // half the ceiling is the floor
  assert.ok(hi(1) <= 1000);
  assert.ok(lo(2) >= lo(1) * 2 - 1); // each attempt at least doubles
  assert.ok(lo(4) > lo(3));
  // Capped: attempt 20 must not exceed MAX_BACKOFF_MS (32s).
  assert.ok(hi(20) <= 32000, 'uncapped backoff would exceed the run limit');
  // Jitter spreads retries so parallel calls do not resynchronize.
  assert.notEqual(lo(5), hi(5));
});

test('past-event guard: ended events read as past, ongoing series do not', () => {
  const cutoff = Date.UTC(2026, 7, 7); // 2026-08-07
  assert.equal(gs.eventEndedBefore_(
    { end: { dateTime: '2026-06-01T10:00:00-06:00' } }, cutoff), true);
  assert.equal(gs.eventEndedBefore_(
    { end: { dateTime: '2026-09-01T10:00:00-06:00' } }, cutoff), false);
  assert.equal(gs.eventEndedBefore_(
    { end: { date: '2026-06-02' } }, cutoff), true);
  assert.equal(gs.eventEndedBefore_(
    { recurrence: ['RRULE:FREQ=WEEKLY;UNTIL=20260501T170000Z;BYDAY=TU'], end: { dateTime: '2025-01-06T10:00:00-07:00' } },
    cutoff), true);
  assert.equal(gs.eventEndedBefore_(
    { recurrence: ['RRULE:FREQ=WEEKLY;BYDAY=TU'], end: { dateTime: '2025-01-06T10:00:00-07:00' } },
    cutoff), false);
});
