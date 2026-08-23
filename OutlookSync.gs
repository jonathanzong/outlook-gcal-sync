/**
 * Outlook → Google Calendar one-way sync.
 *
 * Fetches a published Outlook ICS feed on a timer and mirrors its events
 * into a Google calendar as real events (visible to anyone the calendar
 * is shared with, and counted in free/busy).
 *
 * Design notes:
 *  - Timezones are passed through, never hand-converted. Outlook publishes
 *    Windows timezone names ("Mountain Standard Time"); these are mapped to
 *    IANA names ("America/Denver") and handed to Google, which does all math.
 *  - Recurring events are NOT expanded here. The RRULE/EXDATE lines are
 *    passed to Google verbatim (with TZIDs remapped), so Google owns
 *    recurrence expansion. Modified single instances (RECURRENCE-ID) are
 *    patched onto the recurring event after it is created.
 *  - Every event this script creates carries a private tag
 *    (icsSyncTag). The deletion pass only ever touches tagged events, so
 *    events you created yourself are never modified or removed.
 *  - Unchanged events are skipped via a content hash, so steady-state runs
 *    make almost no API calls.
 *
 * SETUP (one time):
 *  1. In the Apps Script editor, click "Services +" and add
 *     "Google Calendar API" (identifier must be "Calendar").
 *  2. Store the feed URL: Project Settings → Script Properties → add
 *     property "ICS_URL" with your published Outlook calendar URL.
 *     (The URL is a secret — anyone holding it can read your calendar —
 *     which is why it lives in Script Properties, not in this file.)
 *  3. Run setup() once and grant permissions.
 * That installs a trigger that runs sync() every SYNC_INTERVAL_MINUTES.
 *
 * To uninstall: run removeTriggersAndSyncedEvents().
 */

// ------------------------------- CONFIG -------------------------------

// Normally '' — the URL is read from Script Properties (key "ICS_URL") so
// this file can live in a public repo. Hardcode it here only if you never
// commit the file anywhere.
var ICS_URL = '';

function getIcsUrl_() {
  if (ICS_URL) return ICS_URL;
  var url = PropertiesService.getScriptProperties().getProperty('ICS_URL');
  if (!url || url.indexOf('http') !== 0) {
    throw new Error('No feed URL. Add Script Property "ICS_URL" ' +
      '(Project Settings → Script Properties) with your published Outlook calendar URL.');
  }
  return url;
}

// 'primary' = your main personal calendar. Or paste another calendar's ID.
var TARGET_CALENDAR_ID = 'primary';

// Tag written into every synced event; the deletion pass only touches
// events carrying this tag. Do not change after the first run.
var SYNC_TAG = 'outlook-ics-sync';

var SYNC_INTERVAL_MINUTES = 15;

// Used when the feed contains a TZID this script cannot map (e.g. Outlook's
// "Customized Time Zone" with no usable offset info).
var DEFAULT_IANA_TZ = 'America/Denver';

// false = synced events make no popup/email reminders on your phone.
var USE_DEFAULT_REMINDERS = false;

// Optional prefix for synced event titles, e.g. '[work] '. '' = none.
var TITLE_PREFIX = '';

// Color for synced events, two systems:
//  - EVENT_LABEL_ID reaches the full current palette (the 24 named colors
//    such as Mango, and custom RGB shades), which the API models as event
//    labels. Label IDs are per-calendar: run listEventLabels() once, find
//    the label whose name/color you want in the execution log, and paste
//    its id here. Takes precedence over EVENT_COLOR_ID when set.
//  - EVENT_COLOR_ID is the classic index palette, '1'–'11': 1 Lavender,
//    2 Sage, 3 Grape, 4 Flamingo, 5 Banana, 6 Tangerine, 7 Peacock,
//    8 Graphite, 9 Blueberry, 10 Basil, 11 Tomato.
// Leave both '' for the calendar's default color.
var EVENT_LABEL_ID = '';
var EVENT_COLOR_ID = '';

// Exchange publishes a cancelled meeting as a still-CONFIRMED event whose
// title it prefixes ("Canceled: Staff sync"), so STATUS:CANCELLED alone
// misses most cancellations. Any title matching this pattern is treated as
// cancelled and kept off Google. The wording is server-language dependent;
// extend the pattern for a non-English tenant, or set it to null to rely on
// STATUS alone.
var CANCELLED_TITLE_PATTERN = /^\s*cancell?ed:\s*/i;

// Retry and pacing. Google returns "Rate Limit Exceeded" for short bursts
// of writes, which a full rewrite (a color change, a first sync) produces;
// the fix is to retry with exponential backoff and to space writes out.
var MAX_ATTEMPTS = 6;          // per API call, including the first try
var BASE_BACKOFF_MS = 1000;    // first retry waits ~1s, then 2s, 4s, ...
var MAX_BACKOFF_MS = 32000;
var WRITE_PAUSE_MS = 120;      // pause after each write; 0 disables pacing

// Apps Script kills a run at 6 minutes (30 on Workspace accounts). The sync
// stops cleanly at this point and leaves the rest for the next run, which
// picks up where this one stopped because unchanged events are skipped.
var MAX_RUN_MS = 4.5 * 60 * 1000;

// Outlook only publishes a rolling window (~6 months back). When a past
// event ages out of that window it disappears from the feed; false keeps
// such events on Google as history, true deletes them along with genuinely
// cancelled ones.
var DELETE_PAST_EVENTS = false;

// ------------------------- WINDOWS → IANA MAP --------------------------
// CLDR windowsZones mapping (territory 001, the canonical zone per name).

var WINDOWS_TZ_MAP = {
  'Dateline Standard Time': 'Etc/GMT+12',
  'UTC-11': 'Etc/GMT+11',
  'Aleutian Standard Time': 'America/Adak',
  'Hawaiian Standard Time': 'Pacific/Honolulu',
  'Marquesas Standard Time': 'Pacific/Marquesas',
  'Alaskan Standard Time': 'America/Anchorage',
  'UTC-09': 'Etc/GMT+9',
  'Pacific Standard Time (Mexico)': 'America/Tijuana',
  'UTC-08': 'Etc/GMT+8',
  'Pacific Standard Time': 'America/Los_Angeles',
  'US Mountain Standard Time': 'America/Phoenix',
  'Mountain Standard Time (Mexico)': 'America/Mazatlan',
  'Mountain Standard Time': 'America/Denver',
  'Yukon Standard Time': 'America/Whitehorse',
  'Central America Standard Time': 'America/Guatemala',
  'Central Standard Time': 'America/Chicago',
  'Easter Island Standard Time': 'Pacific/Easter',
  'Central Standard Time (Mexico)': 'America/Mexico_City',
  'Canada Central Standard Time': 'America/Regina',
  'SA Pacific Standard Time': 'America/Bogota',
  'Eastern Standard Time (Mexico)': 'America/Cancun',
  'Eastern Standard Time': 'America/New_York',
  'Haiti Standard Time': 'America/Port-au-Prince',
  'Cuba Standard Time': 'America/Havana',
  'US Eastern Standard Time': 'America/Indiana/Indianapolis',
  'Turks And Caicos Standard Time': 'America/Grand_Turk',
  'Paraguay Standard Time': 'America/Asuncion',
  'Atlantic Standard Time': 'America/Halifax',
  'Venezuela Standard Time': 'America/Caracas',
  'Central Brazilian Standard Time': 'America/Cuiaba',
  'SA Western Standard Time': 'America/La_Paz',
  'Pacific SA Standard Time': 'America/Santiago',
  'Newfoundland Standard Time': 'America/St_Johns',
  'Tocantins Standard Time': 'America/Araguaina',
  'E. South America Standard Time': 'America/Sao_Paulo',
  'SA Eastern Standard Time': 'America/Cayenne',
  'Argentina Standard Time': 'America/Argentina/Buenos_Aires',
  'Montevideo Standard Time': 'America/Montevideo',
  'Magallanes Standard Time': 'America/Punta_Arenas',
  'Saint Pierre Standard Time': 'America/Miquelon',
  'Bahia Standard Time': 'America/Bahia',
  'UTC-02': 'Etc/GMT+2',
  'Greenland Standard Time': 'America/Nuuk',
  'Azores Standard Time': 'Atlantic/Azores',
  'Cape Verde Standard Time': 'Atlantic/Cape_Verde',
  'UTC': 'Etc/UTC',
  'GMT Standard Time': 'Europe/London',
  'Greenwich Standard Time': 'Atlantic/Reykjavik',
  'Sao Tome Standard Time': 'Africa/Sao_Tome',
  'Morocco Standard Time': 'Africa/Casablanca',
  'W. Europe Standard Time': 'Europe/Berlin',
  'Central Europe Standard Time': 'Europe/Budapest',
  'Romance Standard Time': 'Europe/Paris',
  'Central European Standard Time': 'Europe/Warsaw',
  'W. Central Africa Standard Time': 'Africa/Lagos',
  'Jordan Standard Time': 'Asia/Amman',
  'GTB Standard Time': 'Europe/Bucharest',
  'Middle East Standard Time': 'Asia/Beirut',
  'Egypt Standard Time': 'Africa/Cairo',
  'E. Europe Standard Time': 'Europe/Chisinau',
  'Syria Standard Time': 'Asia/Damascus',
  'West Bank Standard Time': 'Asia/Hebron',
  'South Africa Standard Time': 'Africa/Johannesburg',
  'FLE Standard Time': 'Europe/Kyiv',
  'Israel Standard Time': 'Asia/Jerusalem',
  'South Sudan Standard Time': 'Africa/Juba',
  'Kaliningrad Standard Time': 'Europe/Kaliningrad',
  'Sudan Standard Time': 'Africa/Khartoum',
  'Libya Standard Time': 'Africa/Tripoli',
  'Namibia Standard Time': 'Africa/Windhoek',
  'Arabic Standard Time': 'Asia/Baghdad',
  'Turkey Standard Time': 'Europe/Istanbul',
  'Arab Standard Time': 'Asia/Riyadh',
  'Belarus Standard Time': 'Europe/Minsk',
  'Russian Standard Time': 'Europe/Moscow',
  'E. Africa Standard Time': 'Africa/Nairobi',
  'Iran Standard Time': 'Asia/Tehran',
  'Arabian Standard Time': 'Asia/Dubai',
  'Astrakhan Standard Time': 'Europe/Astrakhan',
  'Azerbaijan Standard Time': 'Asia/Baku',
  'Russia Time Zone 3': 'Europe/Samara',
  'Mauritius Standard Time': 'Indian/Mauritius',
  'Saratov Standard Time': 'Europe/Saratov',
  'Georgian Standard Time': 'Asia/Tbilisi',
  'Volgograd Standard Time': 'Europe/Volgograd',
  'Caucasus Standard Time': 'Asia/Yerevan',
  'Afghanistan Standard Time': 'Asia/Kabul',
  'West Asia Standard Time': 'Asia/Tashkent',
  'Ekaterinburg Standard Time': 'Asia/Yekaterinburg',
  'Pakistan Standard Time': 'Asia/Karachi',
  'Qyzylorda Standard Time': 'Asia/Qyzylorda',
  'India Standard Time': 'Asia/Kolkata',
  'Sri Lanka Standard Time': 'Asia/Colombo',
  'Nepal Standard Time': 'Asia/Kathmandu',
  'Central Asia Standard Time': 'Asia/Almaty',
  'Bangladesh Standard Time': 'Asia/Dhaka',
  'Omsk Standard Time': 'Asia/Omsk',
  'Myanmar Standard Time': 'Asia/Yangon',
  'SE Asia Standard Time': 'Asia/Bangkok',
  'Altai Standard Time': 'Asia/Barnaul',
  'W. Mongolia Standard Time': 'Asia/Hovd',
  'North Asia Standard Time': 'Asia/Krasnoyarsk',
  'N. Central Asia Standard Time': 'Asia/Novosibirsk',
  'Tomsk Standard Time': 'Asia/Tomsk',
  'China Standard Time': 'Asia/Shanghai',
  'North Asia East Standard Time': 'Asia/Irkutsk',
  'Singapore Standard Time': 'Asia/Singapore',
  'W. Australia Standard Time': 'Australia/Perth',
  'Taipei Standard Time': 'Asia/Taipei',
  'Ulaanbaatar Standard Time': 'Asia/Ulaanbaatar',
  'Aus Central W. Standard Time': 'Australia/Eucla',
  'Transbaikal Standard Time': 'Asia/Chita',
  'Tokyo Standard Time': 'Asia/Tokyo',
  'North Korea Standard Time': 'Asia/Pyongyang',
  'Korea Standard Time': 'Asia/Seoul',
  'Yakutsk Standard Time': 'Asia/Yakutsk',
  'Cen. Australia Standard Time': 'Australia/Adelaide',
  'AUS Central Standard Time': 'Australia/Darwin',
  'E. Australia Standard Time': 'Australia/Brisbane',
  'AUS Eastern Standard Time': 'Australia/Sydney',
  'West Pacific Standard Time': 'Pacific/Port_Moresby',
  'Tasmania Standard Time': 'Australia/Hobart',
  'Vladivostok Standard Time': 'Asia/Vladivostok',
  'Lord Howe Standard Time': 'Australia/Lord_Howe',
  'Bougainville Standard Time': 'Pacific/Bougainville',
  'Russia Time Zone 10': 'Asia/Srednekolymsk',
  'Magadan Standard Time': 'Asia/Magadan',
  'Norfolk Standard Time': 'Pacific/Norfolk',
  'Sakhalin Standard Time': 'Asia/Sakhalin',
  'Central Pacific Standard Time': 'Pacific/Guadalcanal',
  'Russia Time Zone 11': 'Asia/Kamchatka',
  'New Zealand Standard Time': 'Pacific/Auckland',
  'UTC+12': 'Etc/GMT-12',
  'Fiji Standard Time': 'Pacific/Fiji',
  'Chatham Islands Standard Time': 'Pacific/Chatham',
  'UTC+13': 'Etc/GMT-13',
  'Tonga Standard Time': 'Pacific/Tongatapu',
  'Samoa Standard Time': 'Pacific/Apia',
  'Line Islands Standard Time': 'Pacific/Kiritimati'
};

// Fallback for custom TZIDs, keyed by the VTIMEZONE's standard UTC offset
// in minutes. Used only when the TZID has a DAYLIGHT rule (observes DST).
var OFFSET_TO_DST_ZONE = {
  '-600': 'Pacific/Honolulu',
  '-540': 'America/Anchorage',
  '-480': 'America/Los_Angeles',
  '-420': 'America/Denver',
  '-360': 'America/Chicago',
  '-300': 'America/New_York',
  '-240': 'America/Halifax',
  '0': 'Europe/London',
  '60': 'Europe/Berlin',
  '120': 'Europe/Helsinki'
};

// --------------------------- API CALL WRAPPER ---------------------------

var RUN_DEADLINE_ = null; // epoch ms; set at the start of each sync run

/**
 * True for errors worth retrying: quota and rate limits, and Google's
 * transient backend failures. A malformed request is not retried.
 */
function isTransientApiError_(error) {
  var text = String((error && error.message) || error);
  // "Service invoked too many times" is Apps Script's own quota wording and
  // must be retried alongside Google's rate-limit wording.
  return /rate limit|ratelimit|quota|user rate|too many (requests|times)|backend error|internal error|transient|try again|temporarily unavailable|\b(429|500|502|503|504)\b/i
    .test(text);
}

/**
 * Exponential backoff with equal jitter: the wait lands between half the
 * attempt's ceiling and the ceiling itself, so retries never exceed
 * MAX_BACKOFF_MS and concurrent calls do not resynchronize. rand() is
 * injectable for testing.
 */
function backoffDelayMs_(attempt, rand) {
  var ceiling = Math.min(MAX_BACKOFF_MS, BASE_BACKOFF_MS * Math.pow(2, attempt - 1));
  var half = ceiling / 2;
  return Math.floor(half + (rand || Math.random)() * half);
}

/** Runs one API call, retrying transient failures. */
function callApi_(label, fn) {
  for (var attempt = 1; ; attempt++) {
    try {
      return fn();
    } catch (e) {
      if (attempt >= MAX_ATTEMPTS || !isTransientApiError_(e)) throw e;
      var waitMs = backoffDelayMs_(attempt);
      if (RUN_DEADLINE_ && Date.now() + waitMs > RUN_DEADLINE_) {
        throw new Error('Out of time while backing off from ' + label + ': ' + e);
      }
      console.warn(label + ' hit a transient error (attempt ' + attempt + '/' +
        MAX_ATTEMPTS + '), retrying in ' + waitMs + ' ms: ' + e);
      Utilities.sleep(waitMs);
    }
  }
}

/** Like callApi_, plus the pacing pause that keeps bursts under the limit. */
function writeApi_(label, fn) {
  var result = callApi_(label, fn);
  if (WRITE_PAUSE_MS) Utilities.sleep(WRITE_PAUSE_MS);
  return result;
}

function outOfTime_() {
  return !!RUN_DEADLINE_ && Date.now() > RUN_DEADLINE_;
}

// ----------------------------- ENTRY POINTS ----------------------------

/** Run once by hand. Installs the timer and does the first sync. */
function setup() {
  getIcsUrl_(); // fail fast if the URL is not configured
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sync') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sync')
    .timeBased()
    .everyMinutes(SYNC_INTERVAL_MINUTES)
    .create();
  sync();
}

/** Removes the timer and every event this script created. */
function removeTriggersAndSyncedEvents() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sync') ScriptApp.deleteTrigger(t);
  });
  var existing = listManagedEvents_();
  Object.keys(existing).forEach(function (uid) {
    try {
      writeApi_('remove ' + uid, function () {
        return Calendar.Events.remove(TARGET_CALENDAR_ID, existing[uid].id);
      });
    } catch (e) {
      console.warn('Could not remove ' + uid + ': ' + e);
    }
  });
  console.log('Removed trigger and ' + Object.keys(existing).length + ' synced events.');
}

/**
 * Run by hand to find the id for EVENT_LABEL_ID. Logs the target calendar's
 * event labels — the entries behind the current color palette (named colors
 * such as Mango, plus any custom shades). If nothing is logged, color one
 * event with the wanted color in the Calendar UI first, then rerun.
 */
function listEventLabels() {
  var cal = callApi_('calendars.get', function () {
    return Calendar.Calendars.get(TARGET_CALENDAR_ID);
  });
  var labels = cal.labelProperties && cal.labelProperties.eventLabels;
  if (!labels || labels.length === 0) {
    console.log('No event labels found on this calendar yet. In the Calendar UI, set one event to the color you want, then run this again.');
    return;
  }
  labels.forEach(function (l) { console.log(JSON.stringify(l)); });
}

/**
 * Run by hand to rebuild every synced event from scratch. Deletes the
 * events this script owns, then syncs. Use after changing settings that
 * only apply at creation time, or to clear out events an older version of
 * the script wrote incorrectly. The trigger is left in place.
 */
function resyncAll() {
  var managed = listManagedEvents_();
  var uids = Object.keys(managed);
  uids.forEach(function (uid) {
    try {
      writeApi_('remove ' + uid, function () {
        return Calendar.Events.remove(TARGET_CALENDAR_ID, managed[uid].id);
      });
    } catch (e) {
      console.warn('Could not remove ' + uid + ': ' + e);
    }
  });
  console.log('Removed ' + uids.length + ' synced events; rebuilding from the feed.');
  sync();
}

/** Main sync. Runs on the timer. */
function sync() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    console.warn('Previous run still going; skipping.');
    return;
  }
  try {
    syncLocked_();
  } finally {
    lock.releaseLock();
  }
}

function syncLocked_() {
  RUN_DEADLINE_ = Date.now() + MAX_RUN_MS;
  var response = UrlFetchApp.fetch(getIcsUrl_(), {
    muteHttpExceptions: true,
    followRedirects: true,
    validateHttpsCertificates: true
  });
  if (response.getResponseCode() !== 200) {
    console.error('Feed fetch failed: HTTP ' + response.getResponseCode() + ' — leaving calendar untouched.');
    return;
  }
  var icsText = response.getContentText('UTF-8');
  if (icsText.indexOf('BEGIN:VCALENDAR') === -1) {
    console.error('Response is not an ICS file — leaving calendar untouched.');
    return;
  }

  var parsed = parseIcs(icsText);
  var groups = groupByUid(parsed.events);
  var groupCount = Object.keys(groups).length;

  var existing = listManagedEvents_(); // icsUid → Google event

  // Safety valve: an empty feed while Google holds synced events means the
  // feed broke (expired link, publishing turned off), not that every
  // meeting vanished. Refuse to mass-delete.
  if (groupCount === 0 && Object.keys(existing).length > 0) {
    console.error('Feed parsed to 0 events but synced events exist — refusing to delete anything.');
    return;
  }

  var stats = { created: 0, updated: 0, unchanged: 0, deleted: 0,
                overridesPatched: 0, errors: 0, deferred: 0 };

  Object.keys(groups).forEach(function (uid) {
    // Retries and pacing cost wall-clock time, so a large rewrite can run
    // past the Apps Script execution limit. Stop before being killed: the
    // remaining events sync on the next run, which skips everything this
    // run already wrote.
    if (outOfTime_()) { stats.deferred++; return; }
    try {
      syncOneGroup_(uid, groups[uid], parsed.tzMap, existing, stats);
    } catch (e) {
      stats.errors++;
      console.error('Event ' + uid + ' failed: ' + e + (e.stack ? '\n' + e.stack : ''));
    }
  });

  // Delete managed events whose UID no longer appears in the feed — except
  // events already over, which vanish from the feed when they age out of
  // Outlook's published window, not because anyone cancelled them.
  var pastCutoff = Date.now() - 24 * 3600 * 1000;
  Object.keys(existing).forEach(function (uid) {
    if (groups[uid] || outOfTime_()) return;
    if (!DELETE_PAST_EVENTS && eventEndedBefore_(existing[uid], pastCutoff)) return;
    try {
      writeApi_('remove ' + uid, function () {
        return Calendar.Events.remove(TARGET_CALENDAR_ID, existing[uid].id);
      });
      stats.deleted++;
    } catch (e) {
      stats.errors++;
      console.error('Delete of ' + uid + ' failed: ' + e);
    }
  });

  console.log(
    'Sync done. feed events: ' + groupCount +
    ', created: ' + stats.created +
    ', updated: ' + stats.updated +
    ', unchanged: ' + stats.unchanged +
    ', deleted: ' + stats.deleted +
    ', instance overrides patched: ' + stats.overridesPatched +
    ', errors: ' + stats.errors +
    ', deferred to next run: ' + stats.deferred
  );
  if (stats.deferred > 0) {
    console.warn(stats.deferred + ' events did not fit in this run and sync on the next one.');
  }
}

// --------------------------- PER-EVENT SYNC ----------------------------

/**
 * True when a VEVENT represents a cancelled meeting, by either signal
 * Exchange uses: an explicit STATUS:CANCELLED, or the title prefix it
 * writes onto cancellations that keep STATUS:CONFIRMED.
 */
function isCancelled_(vevent) {
  var status = getProp(vevent, 'STATUS');
  if (status && status.value.toUpperCase() === 'CANCELLED') return true;
  if (!CANCELLED_TITLE_PATTERN) return false;
  var summary = getProp(vevent, 'SUMMARY');
  return !!summary && CANCELLED_TITLE_PATTERN.test(unescapeIcsText(summary.value));
}

function syncOneGroup_(uid, group, tzMap, existing, stats) {
  var prior = existing[uid];

  // A cancelled master means the whole series is gone.
  if (group.master && isCancelled_(group.master)) {
    if (prior) {
      writeApi_('remove ' + uid, function () {
        return Calendar.Events.remove(TARGET_CALENDAR_ID, prior.id);
      });
      stats.deleted++;
    }
    return;
  }

  // Orphan overrides (master outside the published window): sync each as a
  // standalone event under a synthetic UID.
  if (!group.master) {
    group.overrides.forEach(function (ov) {
      var syntheticUid = uid + '/' + getProp(ov, 'RECURRENCE-ID').value;
      var fakeGroup = { master: ov, overrides: [] };
      syncOneGroup_(syntheticUid, fakeGroup, tzMap, existing, stats);
    });
    return;
  }

  var hash = contentHash_(group);
  if (prior && getPrivateProp_(prior, 'icsHash') === hash) {
    stats.unchanged++;
    return;
  }

  var resource = buildEventResource(group.master, tzMap, uid, hash);

  // The API only processes eventLabelId when the request carries
  // eventLabelVersion=1, and import() does not process it at all — hence
  // the follow-up patch on the create path.
  var labelArgs = EVENT_LABEL_ID ? { eventLabelVersion: 1 } : {};
  var saved;
  if (prior) {
    resource.id = prior.id;
    try {
      saved = writeApi_('update ' + uid, function () {
        return Calendar.Events.update(resource, TARGET_CALENDAR_ID, prior.id, labelArgs);
      });
    } catch (e) {
      // The stored copy cannot accept this resource (a shape Google will not
      // convert in place, or a copy an earlier version of this script wrote
      // wrongly). Replace it: the script owns every event it deletes here.
      console.warn('Update of ' + uid + ' failed (' + e + '); replacing the event.');
      try {
        writeApi_('remove ' + uid, function () {
          return Calendar.Events.remove(TARGET_CALENDAR_ID, prior.id);
        });
      } catch (e2) {}
      delete resource.id;
      saved = createEvent_(resource, labelArgs);
    }
    stats.updated++;
  } else {
    saved = createEvent_(resource, labelArgs);
    stats.created++;
  }

  group.overrides.forEach(function (ov) {
    patchOverride_(saved, ov, tzMap, stats);
  });
}

/** Imports one event, then applies the label (import ignores labels). */
function createEvent_(resource, labelArgs) {
  var saved = writeApi_('import ' + resource.iCalUID, function () {
    return Calendar.Events.import(resource, TARGET_CALENDAR_ID);
  });
  if (EVENT_LABEL_ID) {
    saved = writeApi_('label ' + saved.id, function () {
      return Calendar.Events.patch(
        { eventLabelId: String(EVENT_LABEL_ID) },
        TARGET_CALENDAR_ID, saved.id, labelArgs);
    });
  }
  return saved;
}

/** Applies a RECURRENCE-ID override onto one instance of a recurring event. */
function patchOverride_(masterEvent, override, tzMap, stats) {
  var recId = getProp(override, 'RECURRENCE-ID');
  var origStart = icsTimeToGoogle(recId.value, recId.params, tzMap);

  var listArgs;
  if (origStart.date) {
    listArgs = { originalStart: origStart.date };
  } else {
    listArgs = { originalStart: toRfc3339_(origStart.dateTime, origStart.timeZone) };
  }
  var instances = callApi_('instances of ' + masterEvent.id, function () {
    return Calendar.Events.instances(TARGET_CALENDAR_ID, masterEvent.id, listArgs);
  });
  if (!instances.items || instances.items.length === 0) {
    console.warn('No instance found for override at ' + recId.value + ' of "' +
      (masterEvent.summary || '') + '" — skipping this override.');
    return;
  }
  var instance = instances.items[0];

  if (isCancelled_(override)) {
    instance.status = 'cancelled';
  } else {
    var patch = buildEventResource(override, tzMap, null, null);
    instance.start = patch.start;
    instance.end = patch.end;
    instance.summary = patch.summary;
    instance.description = patch.description;
    instance.location = patch.location;
    instance.transparency = patch.transparency;
    instance.status = 'confirmed';
  }
  if (EVENT_LABEL_ID) instance.eventLabelId = String(EVENT_LABEL_ID);
  writeApi_('override ' + instance.id, function () {
    return Calendar.Events.update(instance, TARGET_CALENDAR_ID, instance.id,
      EVENT_LABEL_ID ? { eventLabelVersion: 1 } : {});
  });
  stats.overridesPatched++;
}

// ------------------------ GOOGLE EVENT BUILDING ------------------------

/** Builds a Calendar API event resource from one VEVENT. */
function buildEventResource(vevent, tzMap, uid, hash) {
  var dtstart = getProp(vevent, 'DTSTART');
  if (!dtstart) throw new Error('VEVENT has no DTSTART');
  var start = icsTimeToGoogle(dtstart.value, dtstart.params, tzMap);

  var end;
  var dtend = getProp(vevent, 'DTEND');
  var duration = getProp(vevent, 'DURATION');
  if (dtend) {
    end = icsTimeToGoogle(dtend.value, dtend.params, tzMap);
  } else if (duration) {
    end = addDuration_(start, duration.value);
  } else {
    // RFC 5545: no DTEND/DURATION → zero length (date-time) or one day (date).
    end = start.date ? addDays_(start) : JSON.parse(JSON.stringify(start));
  }

  var resource = {
    summary: TITLE_PREFIX + textProp(vevent, 'SUMMARY'),
    description: textProp(vevent, 'DESCRIPTION'),
    location: textProp(vevent, 'LOCATION'),
    start: start,
    end: end,
    reminders: USE_DEFAULT_REMINDERS ? { useDefault: true } : { useDefault: false, overrides: [] },
    transparency:
      (getProp(vevent, 'TRANSP') && getProp(vevent, 'TRANSP').value === 'TRANSPARENT')
        ? 'transparent' : 'opaque'
  };

  if (EVENT_LABEL_ID) resource.eventLabelId = String(EVENT_LABEL_ID);
  else if (EVENT_COLOR_ID) resource.colorId = String(EVENT_COLOR_ID);

  var recurrence = buildRecurrence(vevent, tzMap, start.timeZone);
  if (recurrence.length > 0) resource.recurrence = recurrence;

  if (uid !== null) {
    resource.iCalUID = uid;
    resource.extendedProperties = {
      'private': { icsSyncTag: SYNC_TAG, icsUid: uid, icsHash: hash }
    };
  }
  return resource;
}

/**
 * Rebuilds the recurrence property array for Google: RRULE passed through
 * (UNTIL normalized to UTC), EXDATE/RDATE re-emitted with IANA TZIDs.
 */
function buildRecurrence(vevent, tzMap, eventTz) {
  var out = [];

  (vevent.props['RRULE'] || []).forEach(function (p) {
    out.push('RRULE:' + normalizeUntil_(p.value, eventTz));
  });

  ['EXDATE', 'RDATE'].forEach(function (name) {
    var values = [];
    var isDate = false;
    (vevent.props[name] || []).forEach(function (p) {
      var tz = p.params['TZID'] ? mapTzid(p.params['TZID'], tzMap) : null;
      if (p.params['VALUE'] === 'DATE') isDate = true;
      p.value.split(',').forEach(function (v) {
        values.push({ v: v, tz: tz });
      });
    });
    if (values.length === 0) return;
    if (isDate) {
      out.push(name + ';VALUE=DATE:' + values.map(function (x) { return x.v; }).join(','));
      return;
    }
    // Group by (mapped) tz. UTC values (trailing Z) need no TZID.
    var byTz = {};
    values.forEach(function (x) {
      var key = /Z$/.test(x.v) ? '' : (x.tz || eventTz || DEFAULT_IANA_TZ);
      (byTz[key] = byTz[key] || []).push(x.v);
    });
    Object.keys(byTz).forEach(function (tz) {
      var prefix = tz === '' ? name : name + ';TZID=' + tz;
      out.push(prefix + ':' + byTz[tz].join(','));
    });
  });

  return out;
}

/** Google rejects RRULE UNTIL values in local time; convert them to UTC. */
function normalizeUntil_(rrule, eventTz) {
  return rrule.replace(/UNTIL=([0-9]{8}T[0-9]{6})(?![0-9Z])/, function (m, local) {
    var utc = localToUtc_(parseIcsDateTime_(local), eventTz || DEFAULT_IANA_TZ);
    return 'UNTIL=' + Utilities.formatDate(utc, 'Etc/UTC', "yyyyMMdd'T'HHmmss'Z'");
  });
}

/**
 * Converts one ICS date/date-time value into a Google start/end object.
 * Forms handled: VALUE=DATE, UTC ("...Z"), TZID-qualified local, floating.
 */
function icsTimeToGoogle(value, params, tzMap) {
  if (params['VALUE'] === 'DATE' || /^[0-9]{8}$/.test(value)) {
    return { date: value.substring(0, 4) + '-' + value.substring(4, 6) + '-' + value.substring(6, 8) };
  }
  if (/Z$/.test(value)) {
    var w = parseIcsDateTime_(value);
    return {
      dateTime: w.y + '-' + pad2_(w.mo) + '-' + pad2_(w.d) + 'T' + pad2_(w.h) + ':' + pad2_(w.mi) + ':' + pad2_(w.s) + 'Z',
      timeZone: 'Etc/UTC'
    };
  }
  var tz = params['TZID'] ? mapTzid(params['TZID'], tzMap) : DEFAULT_IANA_TZ;
  var t = parseIcsDateTime_(value);
  return {
    dateTime: t.y + '-' + pad2_(t.mo) + '-' + pad2_(t.d) + 'T' + pad2_(t.h) + ':' + pad2_(t.mi) + ':' + pad2_(t.s),
    timeZone: tz
  };
}

/** Maps an Outlook TZID to an IANA name. */
function mapTzid(tzid, tzMap) {
  tzid = tzid.replace(/^"|"$/g, '');
  if (WINDOWS_TZ_MAP[tzid]) return WINDOWS_TZ_MAP[tzid];
  if (tzid.indexOf('/') !== -1) return tzid; // already IANA
  if (tzMap && tzMap[tzid]) return tzMap[tzid]; // resolved from VTIMEZONE
  console.warn('Unknown TZID "' + tzid + '" — using ' + DEFAULT_IANA_TZ);
  return DEFAULT_IANA_TZ;
}

// ------------------------------ ICS PARSER -----------------------------

/**
 * Parses ICS text into { events: [vevent], tzMap: {customTzid: ianaName} }.
 * A vevent is { props: {NAME: [{value, params}]}, raw: '...' }.
 */
function parseIcs(text) {
  var lines = unfoldIcsLines(text);
  var events = [];
  var current = null;
  var rawLines = null;

  var vtimezones = [];
  var currentTz = null;
  var tzSection = null; // 'STANDARD' | 'DAYLIGHT'

  lines.forEach(function (line) {
    if (line === 'BEGIN:VEVENT') {
      current = { props: {}, raw: '' };
      rawLines = [];
      return;
    }
    if (line === 'END:VEVENT') {
      if (current) {
        current.raw = rawLines.join('\n');
        events.push(current);
      }
      current = null;
      return;
    }
    if (line === 'BEGIN:VTIMEZONE') { currentTz = { tzid: null, stdOffset: null, hasDst: false }; return; }
    if (line === 'END:VTIMEZONE') { if (currentTz) vtimezones.push(currentTz); currentTz = null; return; }
    if (currentTz) {
      if (line === 'BEGIN:STANDARD') { tzSection = 'STANDARD'; return; }
      if (line === 'BEGIN:DAYLIGHT') { tzSection = 'DAYLIGHT'; currentTz.hasDst = true; return; }
      if (line === 'END:STANDARD' || line === 'END:DAYLIGHT') { tzSection = null; return; }
      var tzProp = parseIcsLine(line);
      if (!tzProp) return;
      if (tzProp.name === 'TZID') currentTz.tzid = tzProp.value;
      if (tzProp.name === 'TZOFFSETTO' && tzSection === 'STANDARD') {
        currentTz.stdOffset = parseUtcOffsetMinutes_(tzProp.value);
      }
      return;
    }
    if (!current) return;

    rawLines.push(line);
    var prop = parseIcsLine(line);
    if (!prop) return;
    (current.props[prop.name] = current.props[prop.name] || []).push({
      value: prop.value,
      params: prop.params
    });
  });

  // Resolve custom TZIDs (not in the Windows map, not IANA) from their
  // VTIMEZONE offsets.
  var tzMap = {};
  vtimezones.forEach(function (tz) {
    if (!tz.tzid || WINDOWS_TZ_MAP[tz.tzid] || tz.tzid.indexOf('/') !== -1) return;
    if (tz.stdOffset === null) return;
    if (tz.hasDst && OFFSET_TO_DST_ZONE[String(tz.stdOffset)]) {
      tzMap[tz.tzid] = OFFSET_TO_DST_ZONE[String(tz.stdOffset)];
    } else {
      // Fixed offset. Etc/GMT signs are inverted: Etc/GMT+7 means UTC-7.
      var hours = tz.stdOffset / 60;
      if (hours === Math.round(hours)) {
        tzMap[tz.tzid] = 'Etc/GMT' + (hours <= 0 ? '+' + (-hours) : '-' + hours);
      }
    }
  });

  return { events: events, tzMap: tzMap };
}

/** Undoes RFC 5545 line folding (continuation lines start with space/tab). */
function unfoldIcsLines(text) {
  var raw = text.split(/\r?\n/);
  var out = [];
  raw.forEach(function (line) {
    if ((line.charAt(0) === ' ' || line.charAt(0) === '\t') && out.length > 0) {
      out[out.length - 1] += line.substring(1);
    } else if (line !== '') {
      out.push(line);
    }
  });
  return out;
}

/** Parses 'NAME;PARAM="a;b";P2=c:value' → {name, params, value}. */
function parseIcsLine(line) {
  var inQuotes = false;
  var colon = -1;
  for (var i = 0; i < line.length; i++) {
    var ch = line.charAt(i);
    if (ch === '"') inQuotes = !inQuotes;
    else if (ch === ':' && !inQuotes) { colon = i; break; }
  }
  if (colon === -1) return null;

  var head = line.substring(0, colon);
  var value = line.substring(colon + 1);

  var parts = [];
  var buf = '';
  inQuotes = false;
  for (var j = 0; j < head.length; j++) {
    var c = head.charAt(j);
    if (c === '"') { inQuotes = !inQuotes; continue; }
    if (c === ';' && !inQuotes) { parts.push(buf); buf = ''; continue; }
    buf += c;
  }
  parts.push(buf);

  var params = {};
  for (var k = 1; k < parts.length; k++) {
    var eq = parts[k].indexOf('=');
    if (eq === -1) continue;
    params[parts[k].substring(0, eq).toUpperCase()] = parts[k].substring(eq + 1);
  }
  return { name: parts[0].toUpperCase(), params: params, value: value };
}

/** Unescapes RFC 5545 text values: \n, \,, \;, \\. */
function unescapeIcsText(s) {
  return s
    .replace(/\\n/gi, '\n')
    .replace(/\\,/g, ',')
    .replace(/\\;/g, ';')
    .replace(/\\\\/g, '\\');
}

// ------------------------------- HELPERS -------------------------------

function getProp(vevent, name) {
  var list = vevent.props[name];
  return list ? list[0] : null;
}

function textProp(vevent, name) {
  var p = getProp(vevent, name);
  return p ? unescapeIcsText(p.value) : '';
}

function getPrivateProp_(event, key) {
  return event.extendedProperties &&
    event.extendedProperties['private'] &&
    event.extendedProperties['private'][key];
}

function groupByUid(events) {
  var groups = {};
  events.forEach(function (ev) {
    var uidProp = getProp(ev, 'UID');
    if (!uidProp) return;
    var uid = uidProp.value;
    var g = (groups[uid] = groups[uid] || { master: null, overrides: [] });
    if (getProp(ev, 'RECURRENCE-ID')) g.overrides.push(ev);
    else g.master = ev;
  });
  return groups;
}

/**
 * Hash of a UID group's content, used to skip unchanged events. DTSTAMP is
 * excluded because Outlook regenerates it on every publish. The
 * presentation config feeds the hash too, so changing TITLE_PREFIX,
 * EVENT_COLOR_ID, or USE_DEFAULT_REMINDERS re-writes every event on the
 * next run instead of applying only to events that later change.
 */
function contentHash_(group) {
  var texts = [];
  if (group.master) texts.push(group.master.raw);
  group.overrides.forEach(function (ov) { texts.push(ov.raw); });
  var clean = texts.join('\n---\n').replace(/^DTSTAMP:.*$/gm, '') +
    '\ncfg2:' + [TITLE_PREFIX, EVENT_COLOR_ID, EVENT_LABEL_ID, USE_DEFAULT_REMINDERS].join('|');
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, clean, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ((b + 256) % 256).toString(16); }).join('');
}

/** Every event this script has written into the target calendar, by icsUid. */
function isOwnedMaster_(event) {
  if (!getPrivateProp_(event, 'icsUid')) return false;
  // A modified occurrence of a recurring event ("exception") inherits its
  // master's extended properties, so it carries this script's tag and UID
  // too, and Events.list returns it alongside the master. Only the master
  // may receive an update carrying a recurrence field; sending one to an
  // exception fails with "Invalid start time".
  return !event.recurringEventId;
}

function listManagedEvents_() {
  var byUid = {};
  var pageToken = null;
  do {
    var page = callApi_('events.list', function () {
      return Calendar.Events.list(TARGET_CALENDAR_ID, {
        privateExtendedProperty: 'icsSyncTag=' + SYNC_TAG,
        showDeleted: false,
        singleEvents: false,
        maxResults: 2500,
        pageToken: pageToken
      });
    });
    (page.items || []).forEach(function (ev) {
      if (!isOwnedMaster_(ev)) return;
      byUid[getPrivateProp_(ev, 'icsUid')] = ev;
    });
    pageToken = page.nextPageToken;
  } while (pageToken);
  return byUid;
}

/**
 * True when a Google event (single or recurring) has no occurrences at or
 * after cutoffMs. Recurring series without an UNTIL are treated as ongoing.
 * (COUNT-bounded series also read as ongoing — computing their real end
 * would require expansion — so a fully-past COUNT series that leaves the
 * feed is deleted rather than kept; Outlook emits UNTIL, so this is rare.)
 */
function eventEndedBefore_(event, cutoffMs) {
  if (event.recurrence && event.recurrence.length) {
    var until = null;
    event.recurrence.forEach(function (r) {
      var m = r.match(/^RRULE:.*UNTIL=(\d{8}(?:T\d{6}Z?)?)/);
      if (m) until = m[1];
    });
    if (!until) return false;
    var w = parseIcsDateTime_(until);
    var end = /T/.test(until)
      ? Date.UTC(w.y, w.mo - 1, w.d, w.h, w.mi, w.s)
      : Date.UTC(w.y, w.mo - 1, w.d + 1);
    return end < cutoffMs;
  }
  if (event.end) {
    if (event.end.dateTime) return new Date(event.end.dateTime).getTime() < cutoffMs;
    if (event.end.date) return new Date(event.end.date + 'T00:00:00Z').getTime() < cutoffMs;
  }
  return false;
}

/** '20250822T120000' → wall-clock parts. */
function parseIcsDateTime_(v) {
  return {
    y: +v.substring(0, 4), mo: +v.substring(4, 6), d: +v.substring(6, 8),
    h: +v.substring(9, 11) || 0, mi: +v.substring(11, 13) || 0, s: +v.substring(13, 15) || 0
  };
}

/**
 * Wall-clock time in an IANA zone → Date (UTC instant), using the two-pass
 * offset trick with Utilities.formatDate so the tz database does the work.
 */
function localToUtc_(w, tz) {
  var asUtc = Date.UTC(w.y, w.mo - 1, w.d, w.h, w.mi, w.s);
  var guess = asUtc;
  for (var i = 0; i < 3; i++) {
    var offMin = parseUtcOffsetMinutes_(Utilities.formatDate(new Date(guess), tz, 'Z'));
    var candidate = asUtc - offMin * 60000;
    if (candidate === guess) break;
    guess = candidate;
  }
  return new Date(guess);
}

/** '-0630' or '+05:45' → minutes east of UTC. */
function parseUtcOffsetMinutes_(s) {
  var m = s.match(/([+-])(\d{2}):?(\d{2})/);
  if (!m) return 0;
  var min = (+m[2]) * 60 + (+m[3]);
  return m[1] === '-' ? -min : min;
}

/** Google start/end object → RFC3339 with the zone's real offset. */
function toRfc3339_(dateTimeLocal, tz) {
  if (/Z$/.test(dateTimeLocal)) return dateTimeLocal;
  var w = {
    y: +dateTimeLocal.substring(0, 4), mo: +dateTimeLocal.substring(5, 7), d: +dateTimeLocal.substring(8, 10),
    h: +dateTimeLocal.substring(11, 13), mi: +dateTimeLocal.substring(14, 16), s: +dateTimeLocal.substring(17, 19)
  };
  var utc = localToUtc_(w, tz);
  var off = Utilities.formatDate(utc, tz, 'Z'); // e.g. -0600
  return dateTimeLocal + off.substring(0, 3) + ':' + off.substring(3);
}

/** Adds an ISO-8601 duration (e.g. PT1H30M, P1D) to a start object. */
function addDuration_(start, dur) {
  var m = dur.match(/^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/);
  if (!m) return JSON.parse(JSON.stringify(start));
  var sign = m[1] === '-' ? -1 : 1;
  var ms = sign * (
    (+m[2] || 0) * 7 * 86400000 + (+m[3] || 0) * 86400000 +
    (+m[4] || 0) * 3600000 + (+m[5] || 0) * 60000 + (+m[6] || 0) * 1000
  );
  if (start.date) {
    var d = new Date(start.date + 'T00:00:00Z');
    d = new Date(d.getTime() + ms);
    return { date: Utilities.formatDate(d, 'Etc/UTC', 'yyyy-MM-dd') };
  }
  // Duration math must happen on the real instant, then convert back to
  // wall time in the event's zone.
  var w = {
    y: +start.dateTime.substring(0, 4), mo: +start.dateTime.substring(5, 7), d: +start.dateTime.substring(8, 10),
    h: +start.dateTime.substring(11, 13), mi: +start.dateTime.substring(14, 16), s: +start.dateTime.substring(17, 19)
  };
  var tz = start.timeZone === 'Etc/UTC' ? 'Etc/UTC' : start.timeZone;
  var end = new Date(localToUtc_(w, tz).getTime() + ms);
  return {
    dateTime: Utilities.formatDate(end, tz, "yyyy-MM-dd'T'HH:mm:ss"),
    timeZone: tz
  };
}

function addDays_(start) {
  var d = new Date(start.date + 'T00:00:00Z');
  d = new Date(d.getTime() + 86400000);
  return { date: Utilities.formatDate(d, 'Etc/UTC', 'yyyy-MM-dd') };
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}
