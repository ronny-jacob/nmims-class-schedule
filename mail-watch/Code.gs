/**
 * Timetable Mail Watcher (dedicated Apps Script)
 * ---------------------------------------------------------------------
 * Runs on an installable every-minute time trigger and kicks off the
 * timetable-check GitHub workflow the moment a new timetable email lands
 * — but ONLY on Saturdays and between 8 AM and 11 PM IST.
 *
 * Why Saturday + this schedule:
 *   - The site's users need a CURRENT timetable; real timetable emails
 *     historically arrive on Saturdays before the week starts.
 *   - Mon–Fri relies on the existing hourly GitHub cron (skips 12 AM–8 AM).
 *   - This script adds the near-real-time ("every minute") check the user
 *     wanted, without needing any paid GCP infra (uses Apps Script + the
 *     built-in GmailApp only).
 *
 * No webapp endpoint is required — this is purely a time-driven trigger.
 * The GitHub workflow it dispatches does all heavy lifting (IMAP fetch,
 * extract.py, commit & push) exactly as it does today when run by cron.
 *
 * REQUIRED: install the trigger once (see installSaturdayTrigger below),
 * and set three script properties:
 *   - GITHUB_PAT   : a Fine-grained PAT with Actions:write on the target repo
 *   - GITHUB_REPO  : "owner/name" of the target repo (e.g. ronny-jacob/nmims-class-schedule)
 *   - (optional) GITHUB_WORKFLOW : workflow file name, default check-timetable.yml
 *   - (optional) GITHUB_REF      : branch, default main
 */

// ─── Config ────────────────────────────────────────────────────────────
var TZ = "Asia/Kolkata";
var DEFAULT_REPO = "ronny-jacob/nmims-class-schedule";
var DEFAULT_WORKFLOW = "check-timetable.yml";
var DEFAULT_REF = "main";
var EMAIL_MINUTE = 1; // every-minute trigger

// Active window (IST). No checks between these hours.
var START_HOUR = 8;  // 8 AM
var END_HOUR = 23;   // up to 11 PM

// Days of week that should get the rapid every-minute check (0 = Sun … 6 = Sat)
// Per user: careful checks only on Saturday.
var RAPID_CHECK_DAYS = [6]; // Saturday

var PROP_MARK = "LAST_SEEN_MSG_ID";

var TIMETABLE_PATTERN = /\d{1,2}\.\d{1,2}\.\d{4}\s*to\s*\d{1,2}\.\d{1,2}\.\d{4}/i;
var SUBJECT_PATTERN = /timetable|schedule|time\s*table/i;

// ─── Trigger lifecycle ─────────────────────────────────────────────────

/**
 * Call this ONCE (e.g. run in the Apps Script editor) to install the
 * every-minute trigger. Re-running is idempotent.
 */
function installSaturdayTrigger() {
  removeAllTriggers_();
  ScriptApp.newTrigger("onMinuteTick")
    .timeBased()
    .everyMinutes(EMAIL_MINUTE)
    .create();
  Logger.log("Installed every-minute trigger for %s.", Session.getActiveUser().getEmail());
}

/** Removes the every-minute trigger (to stop the watcher). */
function uninstallTrigger() {
  removeAllTriggers_();
  Logger.log("Removed all time triggers.");
}

function removeAllTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });
}

/**
 * The every-minute entry point. Decides locally whether today/time matches
 * the rapid-check policy; if it does and a new timetable mail has arrived,
 * it dispatches the GitHub workflow.
 */
function onMinuteTick() {
  if (!shouldRunNow_()) return;

  var newMail;
  try {
    newMail = findNewTimetableMail_();
  } catch (e) {
    Logger.log("Mail scan error (non-fatal): %s", e.message);
    return;
  }

  if (!newMail) {
    Logger.log("Minute tick: no new timetable mail yet.");
    return;
  }

  var ref = PropertiesService.getScriptProperties().getProperty("GITHUB_REF") || DEFAULT_REF;
  if (dispatchWorkflow_(ref)) {
    markLastSeen_(newMail);
    Logger.log("Dispatched workflow for mail %s", newMail.id);
  } else {
    Logger.log("Dispatch failed; will retry next minute.");
  }
}

// ─── Window helpers ────────────────────────────────────────────────────

function isRapidDay_(dt) {
  return RAPID_CHECK_DAYS.indexOf(dt.getDay()) !== -1;
}

function istHour_(dt) {
  return Number(Utilities.formatDate(dt, TZ, "H"));
}

function shouldRunNow_() {
  var now = new Date();
  if (!isRapidDay_(now)) return false;
  var h = istHour_(now);
  return h >= START_HOUR && h <= END_HOUR;
}

// ─── Gmail scan ─────────────────────────────────────────────────────────

function findNewTimetableMail_() {
  var lastSeen = PropertiesService.getScriptProperties().getProperty(PROP_MARK);
  var threads = GmailApp.search("newer_than:7d has:attachment", 0, 25);

  var best = null, bestTs = 0;
  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    for (var j = 0; j < msgs.length; j++) {
      var m = msgs[j];
      if (lastSeen && m.getId() === lastSeen) continue;
      if (!isTimetableMessage_(m)) continue;
      var ts = m.getDate().getTime();
      if (ts > bestTs) {
        bestTs = ts;
        best = { id: m.getId(), ts: ts };
      }
    }
  }
  return best;
}

function isTimetableMessage_(m) {
  var subject = String(m.getSubject() || "");
  var atts = m.getAttachments();
  for (var i = 0; i < atts.length; i++) {
    var name = atts[i].getName() || "";
    if (TIMETABLE_PATTERN.test(name)) return true;
    if (/\.(xlsx?)$/i.test(name) && SUBJECT_PATTERN.test(subject)) return true;
  }
  return SUBJECT_PATTERN.test(subject);
}

function markLastSeen_(mail) {
  PropertiesService.getScriptProperties().setProperty(PROP_MARK, mail.id);
}

// ─── GitHub dispatch ────────────────────────────────────────────────────

function githubRepo_(repo) {
  var r = String(repo || "");
  if (r.indexOf('/') === -1) r = "owner/" + r;
  return r;
}

function dispatchWorkflow_(ref) {
  var props = PropertiesService.getScriptProperties();
  var pat = props.getProperty("GITHUB_PAT");
  if (!pat) {
    Logger.log("GITHUB_PAT not set in script properties; skipping dispatch.");
    return false;
  }
  var fullRepo = githubRepo_(props.getProperty("GITHUB_REPO") || DEFAULT_REPO);
  var workflow = props.getProperty("GITHUB_WORKFLOW") || DEFAULT_WORKFLOW;

  var url = "https://api.github.com/repos/" + fullRepo +
    "/actions/workflows/" + encodeURIComponent(workflow) + "/dispatches";

  var response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ ref: ref }),
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + pat,
      Accept: "application/vnd.github.v3+json",
      "User-Agent": "mail-watch"
    }
  });

  var code = response.getResponseCode();
  if (code >= 200 && code < 300) {
    Logger.log("GitHub dispatch %s OK (%s)", workflow, code);
    return true;
  }
  Logger.log("GitHub dispatch failed: %s %s", code, response.getContentText());
  return false;
}
