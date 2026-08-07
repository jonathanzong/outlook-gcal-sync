# outlook-gcal-sync

A Google Apps Script that mirrors a published Outlook (Exchange) calendar into Google Calendar as real events, one way, every 15 minutes. The events land on your own calendar, so anyone you share that calendar with sees them, and they count toward your free/busy availability. Subscribing to the ICS URL from Google Calendar gives neither property: Google refreshes subscriptions on its own 8–24 hour schedule, and a subscribed calendar cannot be shared onward.

The script reads the ICS feed that Outlook's "Publish a calendar" feature serves (Outlook on the web → Settings → Calendar → Shared calendars). It needs no Microsoft credentials and no admin consent, which matters on tenants that block third-party OAuth apps.

## How it works

Each run executes one pipeline: fetch the feed, parse it, group the VEVENTs by UID, diff each group against what the script previously wrote, and apply the difference through the Google Calendar API.

The parser unfolds RFC 5545 continuation lines, splits each line into name, parameters, and value (respecting quoted parameters such as `TZID="W. Europe Standard Time"`), and unescapes text values (`\n`, `\,`, `\;`, `\\`). Grouping by UID matters because a recurring event arrives as several VEVENTs sharing one UID: one master carrying the RRULE, plus one override per modified occurrence, each marked by a RECURRENCE-ID.

The diff uses a content hash stored on each Google event (in a private extended property), so an unchanged event costs zero write calls. Outlook regenerates the DTSTAMP line on every publish, so DTSTAMP is excluded from the hash; without that exclusion every event would rewrite on every run.

Writes go through the Calendar advanced service. New events are created with `Events.import`, which preserves the Outlook UID as the Google `iCalUID`; changed events are updated in place by Google event id. A deletion pass then removes any script-owned event whose UID no longer appears in the feed.

## The two design rules that prevent timezone bugs

The script never does date arithmetic on event times. Outlook publishes each time as a wall-clock value plus a timezone name (`DTSTART;TZID=Mountain Standard Time:20260814T110000`). The script maps the timezone name to its IANA equivalent (`America/Denver`) and passes the wall-clock value and the IANA name to Google unchanged. Google's own timezone database then resolves every instant, including across DST transitions. Hand-converting these values to UTC at sync time is what breaks naive sync scripts: a timestamp converted with today's UTC offset is wrong for events on the other side of a DST boundary.

The script also never expands recurrences. The RRULE line passes to Google verbatim (EXDATEs with their TZIDs remapped), so Google owns occurrence generation. A script that expands RRULEs itself must reimplement BYDAY/BYMONTH logic, DST-aware interval math, and exception dates, and each of those is a bug source.

## Edge cases handled

**Windows timezone names.** Exchange publishes Windows names ("Mountain Standard Time", "GMT Standard Time"), which the Google API rejects. The script carries the full CLDR windowsZones table (about 140 entries) mapping each name to its canonical IANA zone. Note the traps in this namespace: "Mountain Standard Time" means Denver *including* its daylight time, and "GMT Standard Time" means London, not GMT.

**"Customized Time Zone".** Exchange emits this TZID when an event's zone matches no named zone. The script reads the accompanying VTIMEZONE block: if the block has a DAYLIGHT rule, the standard offset selects a representative DST-observing zone (offset −420 → America/Denver); if not, the offset selects a fixed zone (−420 → Etc/GMT+7 — the Etc/GMT sign convention is inverted). A TZID that resolves no other way falls back to `DEFAULT_IANA_TZ` with a logged warning.

**DST boundaries.** Wall-clock times convert to instants (needed only for RFC3339 override lookups and UNTIL normalization) via a two-pass fixed-point loop over `Utilities.formatDate`, so the offset used is the offset in force at that event's date, not at sync time. The test suite pins this with a pair of events straddling the March 2026 spring-forward.

**Local-time UNTIL.** Google rejects an RRULE whose UNTIL is a local time. The script detects `UNTIL=...T...` without a trailing `Z` and converts it to UTC using the event's own zone.

**Modified occurrences (RECURRENCE-ID).** After the master syncs, each override finds its Google instance via `Events.instances` with `originalStart`, then patches that instance's time, title, location, and description. An override with `STATUS:CANCELLED` cancels the single instance instead.

**Orphan overrides.** Outlook sometimes publishes an override whose master falls outside the published window. Each orphan syncs as a standalone event under a synthetic UID (`uid + '/' + recurrenceId`), so it still appears and still deletes cleanly later.

**Cancelled events.** A master with `STATUS:CANCELLED` deletes the whole series from Google.

**All-day events and durations.** `VALUE=DATE` events map to Google's all-day form; the exclusive DTEND convention is the same on both sides, so dates pass through. An event with a DURATION instead of a DTEND gets its end computed on the real instant (in milliseconds), then converted back to wall time in the event's zone.

**Multi-value and multi-line EXDATEs.** EXDATE values are collected across lines, split on commas, and re-grouped by mapped timezone before re-emission.

## Safety properties

Every event the script creates carries a private tag (`icsSyncTag`), and the deletion pass queries by that tag, so the script structurally cannot touch events you created yourself. If the feed fetch fails, returns non-ICS content, or parses to zero events while synced events exist (an expired link or publishing turned off, not a suddenly empty calendar), the run aborts without deleting anything. A script lock prevents two runs from overlapping. Per-event errors are caught and logged; one malformed event cannot stop the rest of the sync. Every run logs a summary line (created/updated/unchanged/deleted/errors), so a breakage shows up in the execution log rather than as silent drift.

## Setup

1. Create a project at [script.google.com](https://script.google.com) and paste in `OutlookSync.gs`.
2. In the left sidebar, click **+** next to **Services** and add **Google Calendar API** (keep the identifier `Calendar`).
3. In **Project Settings → Script Properties**, add a property `ICS_URL` with your published Outlook calendar URL. The URL lives in Script Properties rather than in the code because it is a capability: anyone holding it can read your calendar, so it must stay out of the repo.
4. Select `setup` in the function dropdown and run it. Grant the permission prompts. This performs the first sync and installs a trigger that runs `sync` every `SYNC_INTERVAL_MINUTES`.

To uninstall, run `removeTriggersAndSyncedEvents`, which deletes the trigger and every event the script created.

## Configuration

| Variable | Default | Effect |
|---|---|---|
| `ICS_URL` | `''` | Feed URL override; normally left empty so the URL comes from the `ICS_URL` Script Property. |
| `TARGET_CALENDAR_ID` | `'primary'` | Calendar that receives events. Any calendar ID works. |
| `SYNC_TAG` | `'outlook-ics-sync'` | Ownership tag on synced events. Changing it after the first run orphans previously synced events. |
| `SYNC_INTERVAL_MINUTES` | `15` | Trigger period. |
| `DEFAULT_IANA_TZ` | `'America/Denver'` | Fallback zone for unmappable TZIDs and floating times. |
| `USE_DEFAULT_REMINDERS` | `false` | `false` suppresses reminders on synced events. |
| `TITLE_PREFIX` | `''` | Optional prefix for synced titles, e.g. `'[work] '`. |
| `EVENT_LABEL_ID` | `''` | Event-label id for synced events — the current palette (24 named colors such as Mango, plus custom RGB shades). Label ids are per-calendar: run `listEventLabels()` once and copy the id from the log. Wins over `EVENT_COLOR_ID`. |
| `EVENT_COLOR_ID` | `''` | Classic colorId `'1'`–`'11'`. `''` on both color settings keeps the calendar's default color. |
| `DELETE_PAST_EVENTS` | `false` | `false` keeps ended events on Google when they age out of Outlook's published window; `true` deletes them. |

The presentation settings (`TITLE_PREFIX`, `EVENT_COLOR_ID`, `EVENT_LABEL_ID`, `USE_DEFAULT_REMINDERS`) feed the change-detection hash, so editing one re-writes every synced event on the next run rather than applying only to events that later change in Outlook.

Google reworked event colors in June 2026: the UI now offers 24 named colors plus a custom RGB picker, and the API models the new palette as per-calendar event labels (`eventLabelId` on the event, superseding the index-based `colorId`). The classic `colorId` values still work but only reach the original 11 colors. Two API quirks make labels easy to lose silently: the API only processes `eventLabelId` when the request URL carries `eventLabelVersion=1`, and `Events.import` ignores the field entirely. The script therefore passes the version parameter on every update and applies the label with a follow-up `patch` after each import.

## Development

`npm test` runs the suite (no dependencies, Node ≥ 18) via the built-in `node --test` runner. `test/shims.js` emulates the two Apps Script APIs the pure layer uses — `Utilities.formatDate` (reimplemented on `Intl.DateTimeFormat`, including offset computation) and `Utilities.computeDigest` — then evaluates `OutlookSync.gs` and returns its functions. `test/fixture.ics` is a synthetic feed modeled on a real Exchange publication: Windows and custom TZIDs, folded and escaped text, EXDATEs, a RECURRENCE-ID override, an all-day event, a DURATION event, a cancelled event, and a local-time UNTIL. The tests cover the parse-and-transform layer; the functions that call Google services (`sync`, the upsert and deletion passes) are exercised only in Apps Script itself, so test a behavior change there by running `sync` by hand and reading the execution log. A GitHub Actions workflow runs the suite on every push.

To edit the deployed script from this repo instead of the web editor, [clasp](https://github.com/google/clasp) pushes and pulls Apps Script projects from the command line; its `.clasp.json` (which contains your script ID) is gitignored.

## Limitations

The sync is one-way; nothing written on the Google side flows back to Outlook. Attendees and RSVP state are not copied — the feed does not reliably publish them, and copying attendees would risk Google emailing them. End-to-end latency is the sync interval plus Outlook's own publication lag, since Exchange regenerates the published feed on its own schedule (typically minutes, occasionally longer). The feed only contains the window Outlook publishes (by default roughly six months back and forward). An event that disappears from the feed is deleted from Google only if it still has occurrences in the future; ended events that merely aged out of the window stay on Google as history (see `DELETE_PAST_EVENTS`). One consequence: cancelling an already-finished meeting in Outlook does not remove its Google copy.
