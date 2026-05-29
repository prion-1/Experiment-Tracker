// ============================================================================
// Cell Culture Tracker — Apps Script backend
// Lock-based collaboration. One editor at a time; others view read-only.
// All edits during an edit session are batched into a single saveAll() write.
// Data lives in a Google Sheet the script creates on first use.
// ============================================================================

const SHEET_NAME_RUNS  = 'runs';
const SHEET_NAME_STEPS = 'steps';
const SHEET_NAME_LOCK  = 'lock';
const PROP_SS_ID       = 'GANTT_SHEET_ID';
const RUNS_HEADERS  = ['id', 'name', 'color', 'order', 'version'];
const STEPS_HEADERS = ['id', 'runId', 'name', 'start', 'dur', 'mode', 'stoppedAt', 'version'];
const LOCK_HEADERS  = ['holder', 'label', 'acquiredAt'];
const LOCK_TTL_MS   = 20 * 60 * 1000;  // 20 min — abandoned edit sessions expire

// ----------------------------------------------------------------------------
// ACCESS CONTROL — edit this list to restrict who can use the app.
// ----------------------------------------------------------------------------
// Leave EMPTY ([]) to allow anyone the deployment grants access to.
// Add lowercase emails to restrict:
//     const ALLOWED_EMAILS = ['alice@uni.de', 'bob@uni.de'];
//
// Email restriction only works if Apps Script can see who's visiting. That's
// reliable in two configurations — pick (a) or (b) when you deploy:
//
//   (a) Workspace mode: every allowed user is in the same Google Workspace
//       domain as you (the script owner). Deploy with
//         "Execute as: Me"
//         "Who has access: Anyone within <your-domain>"
//
//   (b) Personal-Gmail or mixed mode: deploy with
//         "Execute as: User accessing the web app"
//         "Who has access: Anyone with Google account"
//       AND share the data sheet ("Cell Culture Tracker — data") with each
//       allowed user as Editor.
//
// If ALLOWED_EMAILS is empty, the script does not check identity at all.
// ----------------------------------------------------------------------------
const ALLOWED_EMAILS = [
  // 'you@example.com',
];

function _checkAccess() {
  if (ALLOWED_EMAILS.length === 0) return null;
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) {
    throw new Error(
      'Access denied: could not verify your Google identity. ' +
      'The deployment may need to be set to "Execute as: User accessing the web app".'
    );
  }
  const allowed = ALLOWED_EMAILS.map(function (e) { return String(e).toLowerCase(); });
  if (allowed.indexOf(email) < 0) {
    throw new Error('Access denied: ' + email + ' is not on the allowlist.');
  }
  return email;
}

// ----------------------------------------------------------------------------
// Web entry point
// ----------------------------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Cell Culture Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ----------------------------------------------------------------------------
// Spreadsheet bootstrap + schema migration
// ----------------------------------------------------------------------------
function _getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PROP_SS_ID);
  if (id) {
    try {
      const file = DriveApp.getFileById(id);
      if (file.isTrashed()) {
        props.deleteProperty(PROP_SS_ID);
        id = null;
      } else {
        const ss = SpreadsheetApp.openById(id);
        _ensureSchema(ss);
        return ss;
      }
    } catch (e) {
      props.deleteProperty(PROP_SS_ID);
      id = null;
    }
  }
  const ss = SpreadsheetApp.create('Cell Culture Tracker — data');
  const runs = ss.getActiveSheet();
  runs.setName(SHEET_NAME_RUNS);
  runs.getRange(1, 1, 1, RUNS_HEADERS.length).setValues([RUNS_HEADERS]);
  const steps = ss.insertSheet(SHEET_NAME_STEPS);
  steps.getRange(1, 1, 1, STEPS_HEADERS.length).setValues([STEPS_HEADERS]);
  const lock = ss.insertSheet(SHEET_NAME_LOCK);
  lock.getRange(1, 1, 1, LOCK_HEADERS.length).setValues([LOCK_HEADERS]);
  props.setProperty(PROP_SS_ID, ss.getId());
  return ss;
}

function _ensureSchema(ss) {
  const ensureCols = function (sh, headers) {
    if (!sh) return;
    const lastCol = Math.max(1, sh.getLastColumn());
    const cur = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    for (let i = 0; i < headers.length; i++) {
      if (cur[i] !== headers[i]) sh.getRange(1, i + 1).setValue(headers[i]);
    }
  };
  ensureCols(ss.getSheetByName(SHEET_NAME_RUNS),  RUNS_HEADERS);
  ensureCols(ss.getSheetByName(SHEET_NAME_STEPS), STEPS_HEADERS);
  let lockSh = ss.getSheetByName(SHEET_NAME_LOCK);
  if (!lockSh) {
    lockSh = ss.insertSheet(SHEET_NAME_LOCK);
    lockSh.getRange(1, 1, 1, LOCK_HEADERS.length).setValues([LOCK_HEADERS]);
  } else {
    ensureCols(lockSh, LOCK_HEADERS);
  }
}

function _sheet(name) { return _getSpreadsheet().getSheetByName(name); }
function _uid() { return 'id_' + Utilities.getUuid().replace(/-/g, '').slice(0, 12); }

function _readAll(name) {
  const sh = _sheet(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const cols = sh.getLastColumn();
  const head = sh.getRange(1, 1, 1, cols).getValues()[0];
  const rows = sh.getRange(2, 1, last - 1, cols).getValues();
  return rows.filter(function (r) { return r[0]; }).map(function (r) {
    const o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
}

function _toIso(v) {
  if (v instanceof Date) {
    const tz = _getSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return String(v || '');
}

// ----------------------------------------------------------------------------
// LOCK helpers
// ----------------------------------------------------------------------------
function _readLock() {
  const sh = _sheet(SHEET_NAME_LOCK);
  if (!sh || sh.getLastRow() < 2) return null;
  const row = sh.getRange(2, 1, 1, 3).getValues()[0];
  if (!row[0]) return null;
  return {
    holder: String(row[0]),
    label:  String(row[1] || ''),
    acquiredAt: Number(row[2]) || 0
  };
}

function _isLockExpired(lock) {
  if (!lock) return true;
  return (Date.now() - lock.acquiredAt) > LOCK_TTL_MS;
}

function _writeLock(holder, label) {
  const sh = _sheet(SHEET_NAME_LOCK);
  const values = [[holder, label, Date.now()]];
  if (sh.getLastRow() < 2) sh.getRange(2, 1, 1, 3).setValues(values);
  else sh.getRange(2, 1, 1, 3).setValues(values);
}

function _clearLock() {
  const sh = _sheet(SHEET_NAME_LOCK);
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, 1, 3).setValues([['', '', '']]);
}

// Returns the active lock state (null if none or expired).
function _activeLock() {
  const lock = _readLock();
  if (!lock) return null;
  if (_isLockExpired(lock)) {
    _clearLock();
    return null;
  }
  return lock;
}

// ----------------------------------------------------------------------------
// READ
// ----------------------------------------------------------------------------
function loadAll() {
  _checkAccess();
  const runsRaw  = _readAll(SHEET_NAME_RUNS);
  const stepsRaw = _readAll(SHEET_NAME_STEPS);
  const lock = _activeLock();
  return {
    runs: runsRaw.map(function (r) {
      return {
        id: String(r.id),
        name: String(r.name || ''),
        color: String(r.color || '#44d8e0'),
        order: Number(r.order) || 0
      };
    }),
    steps: stepsRaw.map(function (s) {
      const mode = String(s.mode || '').toLowerCase() === 'open' ? 'open' : 'fixed';
      return {
        id: String(s.id),
        runId: String(s.runId),
        name: String(s.name || ''),
        start: _toIso(s.start),
        dur: Number(s.dur) || 0,
        mode: mode,
        stoppedAt: mode === 'open' ? _toIso(s.stoppedAt) : ''
      };
    }),
    lock: lock,
    serverTime: Date.now(),
    lockTtlMs: LOCK_TTL_MS
  };
}

// ----------------------------------------------------------------------------
// LOCK API
// ----------------------------------------------------------------------------
function acquireLock(sessionId, label) {
  _checkAccess();
  if (!sessionId) return {ok: false, reason: 'bad_session'};
  const scriptLock = LockService.getScriptLock();
  if (!scriptLock.tryLock(5000)) return {ok: false, reason: 'busy'};
  try {
    const cur = _readLock();
    if (cur && cur.holder && !_isLockExpired(cur) && cur.holder !== sessionId) {
      return {ok: false, reason: 'held', lock: cur};
    }
    _writeLock(sessionId, String(label || ''));
    return {ok: true, lock: _readLock(), lockTtlMs: LOCK_TTL_MS};
  } finally {
    scriptLock.releaseLock();
  }
}

function releaseLock(sessionId) {
  _checkAccess();
  if (!sessionId) return {ok: false, reason: 'bad_session'};
  const scriptLock = LockService.getScriptLock();
  if (!scriptLock.tryLock(5000)) return {ok: false, reason: 'busy'};
  try {
    const cur = _readLock();
    if (cur && cur.holder === sessionId) _clearLock();
    return {ok: true};
  } finally {
    scriptLock.releaseLock();
  }
}

// ----------------------------------------------------------------------------
// BATCH WRITE: replaces the entire dataset in one transaction.
// Only the current lock holder may call this; lock is released on success.
// ----------------------------------------------------------------------------
function saveAll(sessionId, runs, steps) {
  _checkAccess();
  if (!sessionId) return {ok: false, reason: 'bad_session'};
  const scriptLock = LockService.getScriptLock();
  if (!scriptLock.tryLock(10000)) return {ok: false, reason: 'busy'};
  try {
    const cur = _readLock();
    if (!cur || cur.holder !== sessionId || _isLockExpired(cur)) {
      return {
        ok: false,
        reason: 'no_lock',
        message: 'Your edit session expired or was taken over. Your local changes are still in this tab — click Unlock again to retry saving.'
      };
    }
    const ss = _getSpreadsheet();
    const rSh = ss.getSheetByName(SHEET_NAME_RUNS);
    const sSh = ss.getSheetByName(SHEET_NAME_STEPS);

    // Clear existing data rows (keep headers)
    if (rSh.getLastRow() > 1) rSh.getRange(2, 1, rSh.getLastRow() - 1, rSh.getLastColumn()).clearContent();
    if (sSh.getLastRow() > 1) sSh.getRange(2, 1, sSh.getLastRow() - 1, sSh.getLastColumn()).clearContent();

    // Write fresh data.
    if (runs && runs.length) {
      const rows = runs.map(function (r) {
        return [
          String(r.id || _uid()),
          String(r.name || ''),
          String(r.color || ''),
          Number(r.order) || 0,
          ''
        ];
      });
      rSh.getRange(2, 1, rows.length, 5).setValues(rows);
    }
    if (steps && steps.length) {
      const rows = steps.map(function (s) {
        const mode = String(s.mode || '').toLowerCase() === 'open' ? 'open' : 'fixed';
        return [
          String(s.id || _uid()),
          String(s.runId || ''),
          String(s.name || ''),
          _toIso(s.start),
          Number(s.dur) || 0,
          mode,
          mode === 'open' ? _toIso(s.stoppedAt) : '',
          ''
        ];
      });
      sSh.getRange(2, 1, rows.length, 8).setValues(rows);
    }
    _clearLock();
    return {ok: true};
  } finally {
    scriptLock.releaseLock();
  }
}

// ----------------------------------------------------------------------------
// Admin utilities (run from the editor if needed)
// ----------------------------------------------------------------------------
function _resetForDev_() {
  const ss = _getSpreadsheet();
  [SHEET_NAME_RUNS, SHEET_NAME_STEPS].forEach(function (name) {
    const sh = ss.getSheetByName(name);
    const last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);
  });
  _clearLock();
}

function _forceClearLock_() {
  _clearLock();
}

function _logDataSheetUrl_() {
  const ss = _getSpreadsheet();
  Logger.log('Data sheet id: ' + ss.getId());
  Logger.log('Data sheet url: ' + ss.getUrl());
}

function _unlinkDataSheetForDev_() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_SS_ID);
}
