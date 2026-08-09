// ============================================================================
// xTra - Experiment Tracker — Apps Script backend
// Lock-based collaboration. One editor at a time; others view read-only.
// All edits during an edit session are batched into a single saveAll() write.
// Data lives in a Google Sheet the script creates on first use.
// ============================================================================

const SHEET_NAME_RUNS  = 'runs';
const SHEET_NAME_STEPS = 'steps';
const SHEET_NAME_EDIT_LOG = 'edit_log';
const SHEET_NAME_TEMPLATES = 'templates';
const PROP_SS_ID       = 'GANTT_SHEET_ID';
const PROP_LOCK        = 'GANTT_EDIT_LOCK';
const PROP_SCHEMA_VERSION = 'GANTT_SCHEMA_VERSION';
const PROP_DATA_VERSION = 'GANTT_DATA_VERSION';
const PROP_LAST_EDIT   = 'GANTT_LAST_EDIT';
const PROP_TIME_ZONE   = 'GANTT_TIME_ZONE';
const SCHEMA_VERSION   = '6';
const RUNS_HEADERS  = ['id', 'name', 'color', 'order', 'version', 'stoppedAt', 'plannedEnd', 'metadata'];
const STEPS_HEADERS = ['id', 'runId', 'name', 'start', 'dur', 'mode', 'stoppedAt', 'version', 'metadata'];
const EDIT_LOG_HEADERS = ['timestamp', 'editor', 'sessionId', 'runCount', 'stepCount'];
const TEMPLATES_HEADERS = ['id', 'name', 'runName', 'color', 'runLength', 'stepsJson', 'createdAt', 'updatedAt', 'version'];
const METADATA_MAX_LENGTH = 50000;
const LOCK_TTL_MS   = 30 * 60 * 1000;  // 30 min — abandoned edit sessions expire

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
//       AND share the data sheet ("xTra - Experiment Tracker — data") with each
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
    .setTitle('xTra - Experiment Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ----------------------------------------------------------------------------
// Spreadsheet bootstrap + schema migration
// ----------------------------------------------------------------------------
function _getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(PROP_SS_ID);
  if (id) {
    try {
      const ss = SpreadsheetApp.openById(id);
      _ensureSchemaIfNeeded(ss, props);
      return ss;
    } catch (e) {
      throw new Error(
        'Could not open the xTra - Experiment Tracker data sheet (' + id + '). ' +
        'Check that it still exists and that this deployment account can access it. ' +
        'Original error: ' + (e && e.message ? e.message : String(e))
      );
    }
  }
  return _createSpreadsheet(props);
}

function _createSpreadsheet(props) {
  const ss = SpreadsheetApp.create('xTra - Experiment Tracker — data');
  const runs = ss.getActiveSheet();
  runs.setName(SHEET_NAME_RUNS);
  runs.getRange(1, 1, 1, RUNS_HEADERS.length).setValues([RUNS_HEADERS]);
  const steps = ss.insertSheet(SHEET_NAME_STEPS);
  steps.getRange(1, 1, 1, STEPS_HEADERS.length).setValues([STEPS_HEADERS]);
  const editLog = ss.insertSheet(SHEET_NAME_EDIT_LOG);
  editLog.getRange(1, 1, 1, EDIT_LOG_HEADERS.length).setValues([EDIT_LOG_HEADERS]);
  const templates = ss.insertSheet(SHEET_NAME_TEMPLATES);
  templates.getRange(1, 1, 1, TEMPLATES_HEADERS.length).setValues([TEMPLATES_HEADERS]);
  props.setProperty(PROP_SS_ID, ss.getId());
  props.setProperty(PROP_SCHEMA_VERSION, SCHEMA_VERSION);
  props.setProperty(PROP_DATA_VERSION, '0');
  props.setProperty(PROP_TIME_ZONE, ss.getSpreadsheetTimeZone());
  return ss;
}

function _ensureSchemaIfNeeded(ss, props) {
  const oldVersion = props.getProperty(PROP_SCHEMA_VERSION);
  if (oldVersion === SCHEMA_VERSION) return;
  _ensureSchema(ss);
  if ((Number(oldVersion) || 0) < 3) {
    _migrateRunStops(ss);
    _bumpDataVersion();
  }
  if ((Number(oldVersion) || 0) < 5) {
    _migratePlannedRunEnds(ss);
    _bumpDataVersion();
  }
  props.setProperty(PROP_SCHEMA_VERSION, SCHEMA_VERSION);
  if (!props.getProperty(PROP_DATA_VERSION)) props.setProperty(PROP_DATA_VERSION, '0');
  props.setProperty(PROP_TIME_ZONE, ss.getSpreadsheetTimeZone());
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
  const ensureSheet = function (name, headers) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    ensureCols(sh, headers);
    return sh;
  };
  ensureSheet(SHEET_NAME_RUNS, RUNS_HEADERS);
  ensureSheet(SHEET_NAME_STEPS, STEPS_HEADERS);
  ensureSheet(SHEET_NAME_EDIT_LOG, EDIT_LOG_HEADERS);
  ensureSheet(SHEET_NAME_TEMPLATES, TEMPLATES_HEADERS);
}

function _uid() { return 'id_' + Utilities.getUuid().replace(/-/g, '').slice(0, 12); }

function _readAll(ss, name) {
  const sh = ss.getSheetByName(name);
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

function _metadataText(value) {
  const text = String(value == null ? '' : value);
  if (text.length > METADATA_MAX_LENGTH) {
    throw new Error('Metadata notes cannot exceed ' + METADATA_MAX_LENGTH + ' characters per entry.');
  }
  return text;
}

function _writeRowsWithMetadata(sh, headers, rows) {
  if (!rows.length) return;
  const metadataIndex = headers.indexOf('metadata');
  const values = rows.map(function (row) {
    const copy = row.slice();
    if (metadataIndex >= 0) copy[metadataIndex] = '';
    return copy;
  });
  sh.getRange(2, 1, values.length, headers.length).setValues(values);
  if (metadataIndex < 0) return;
  const metadataValues = rows.map(function (row) {
    return [SpreadsheetApp.newRichTextValue().setText(_metadataText(row[metadataIndex])).build()];
  });
  sh.getRange(2, metadataIndex + 1, metadataValues.length, 1).setRichTextValues(metadataValues);
}

function _toIso(v, tz) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return String(v || '');
}

function _isoDay(v, tz) {
  const iso = _toIso(v, tz);
  if (!iso) return null;
  const parts = iso.split('-').map(Number);
  if (parts.length !== 3 || parts.some(function (n) { return !Number.isFinite(n); })) return null;
  return Math.round(Date.UTC(parts[0], parts[1] - 1, parts[2]) / 86400000);
}

function _isoFromDay(day) {
  return Utilities.formatDate(new Date(day * 86400000), 'UTC', 'yyyy-MM-dd');
}

function _migrateRunStops(ss) {
  const tz = _timeZone(ss);
  const today = _isoDay(_todayIso(tz), tz);
  const runsSh = ss.getSheetByName(SHEET_NAME_RUNS);
  const stepsSh = ss.getSheetByName(SHEET_NAME_STEPS);
  if (!runsSh || runsSh.getLastRow() < 2) return;

  const runHeaders = runsSh.getRange(1, 1, 1, runsSh.getLastColumn()).getValues()[0];
  const stepHeaders = stepsSh ? stepsSh.getRange(1, 1, 1, stepsSh.getLastColumn()).getValues()[0] : [];
  const runStoppedCol = runHeaders.indexOf('stoppedAt') + 1;
  if (!runStoppedCol) return;

  const stepRows = stepsSh && stepsSh.getLastRow() > 1
    ? stepsSh.getRange(2, 1, stepsSh.getLastRow() - 1, stepsSh.getLastColumn()).getValues()
    : [];
  const idx = function (headers, name) { return headers.indexOf(name); };
  const sId = idx(stepHeaders, 'runId');
  const sStart = idx(stepHeaders, 'start');
  const sDur = idx(stepHeaders, 'dur');
  const sMode = idx(stepHeaders, 'mode');
  const sStopped = idx(stepHeaders, 'stoppedAt');
  const spans = {};

  stepRows.forEach(function (row) {
    const runId = sId >= 0 ? String(row[sId] || '') : '';
    const start = sStart >= 0 ? _isoDay(row[sStart], tz) : null;
    if (!runId || start == null) return;
    const mode = sMode >= 0 && String(row[sMode] || '').toLowerCase() === 'open' ? 'open' : 'fixed';
    const stopped = sStopped >= 0 ? _isoDay(row[sStopped], tz) : null;
    const dur = Math.max(Number(sDur >= 0 ? row[sDur] : 0) || 0, 0);
    let end = start + Math.max(dur || 1, 1) - 1;
    let activeOpen = false;
    if (mode === 'open') {
      activeOpen = stopped == null;
      end = activeOpen ? today : Math.max(start, stopped);
    }
    if (!spans[runId]) spans[runId] = {first: start, last: end, activeOpen: false};
    spans[runId].first = Math.min(spans[runId].first, start);
    spans[runId].last = Math.max(spans[runId].last, end);
    spans[runId].activeOpen = spans[runId].activeOpen || activeOpen;
  });

  const runRows = runsSh.getRange(2, 1, runsSh.getLastRow() - 1, runsSh.getLastColumn()).getValues();
  const stoppedValues = runRows.map(function (row) {
    if (row[runStoppedCol - 1]) return [row[runStoppedCol - 1]];
    const runId = String(row[0] || '');
    const span = spans[runId];
    if (!span) return [_todayIso(tz)];
    if (span.activeOpen || (span.first <= today && span.last >= today) || span.first > today) return [''];
    return [_isoFromDay(span.last)];
  });
  runsSh.getRange(2, runStoppedCol, stoppedValues.length, 1).setValues(stoppedValues);
}

function _migratePlannedRunEnds(ss) {
  const tz = _timeZone(ss);
  const today = _isoDay(_todayIso(tz), tz);
  const runsSh = ss.getSheetByName(SHEET_NAME_RUNS);
  if (!runsSh || runsSh.getLastRow() < 2) return;

  const headers = runsSh.getRange(1, 1, 1, runsSh.getLastColumn()).getValues()[0];
  const stoppedCol = headers.indexOf('stoppedAt') + 1;
  const plannedCol = headers.indexOf('plannedEnd') + 1;
  if (!stoppedCol || !plannedCol) return;

  const rows = runsSh.getRange(2, 1, runsSh.getLastRow() - 1, runsSh.getLastColumn()).getValues();
  const stoppedValues = [];
  const plannedValues = [];
  rows.forEach(function (row) {
    const stoppedRaw = row[stoppedCol - 1];
    const plannedRaw = row[plannedCol - 1];
    const stoppedDay = _isoDay(stoppedRaw, tz);
    if (!plannedRaw && stoppedDay != null && stoppedDay > today) {
      stoppedValues.push(['']);
      plannedValues.push([_toIso(stoppedRaw, tz)]);
    } else {
      stoppedValues.push([stoppedRaw || '']);
      plannedValues.push([plannedRaw || '']);
    }
  });
  runsSh.getRange(2, stoppedCol, stoppedValues.length, 1).setValues(stoppedValues);
  runsSh.getRange(2, plannedCol, plannedValues.length, 1).setValues(plannedValues);
}

function _props() {
  return PropertiesService.getScriptProperties();
}

function _timeZone(ss) {
  const props = _props();
  const tz = ss ? ss.getSpreadsheetTimeZone() : (props.getProperty(PROP_TIME_ZONE) || Session.getScriptTimeZone());
  if (ss) props.setProperty(PROP_TIME_ZONE, tz);
  return tz;
}

function _todayIso(tz) {
  return Utilities.formatDate(new Date(), tz || Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _readDataVersion() {
  return _props().getProperty(PROP_DATA_VERSION) || '0';
}

function _bumpDataVersion() {
  const version = Date.now() + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  _props().setProperty(PROP_DATA_VERSION, version);
  return version;
}

function _readLastEditProperty() {
  const raw = _props().getProperty(PROP_LAST_EDIT);
  if (!raw) return null;
  try {
    const meta = JSON.parse(raw);
    if (!meta || !meta.timestamp) return null;
    return {
      timestamp: Number(meta.timestamp) || 0,
      editor: String(meta.editor || '')
    };
  } catch (e) {
    _props().deleteProperty(PROP_LAST_EDIT);
    return null;
  }
}

function _writeLastEditProperty(meta) {
  if (!meta || !meta.timestamp) {
    _props().deleteProperty(PROP_LAST_EDIT);
    return null;
  }
  const clean = {
    timestamp: Number(meta.timestamp) || 0,
    editor: String(meta.editor || '')
  };
  _props().setProperty(PROP_LAST_EDIT, JSON.stringify(clean));
  return clean;
}

function _safeColor(v, fallback) {
  const color = String(v || '').trim();
  return /^#[0-9a-f]{3}([0-9a-f]{3})?$/i.test(color) ? color : (fallback || '#4169e8');
}

// ----------------------------------------------------------------------------
// LOCK helpers
// ----------------------------------------------------------------------------
function _readLock() {
  const raw = PropertiesService.getScriptProperties().getProperty(PROP_LOCK);
  if (!raw) return null;
  try {
    const lock = JSON.parse(raw);
    if (!lock || !lock.holder) return null;
    return {
      holder: String(lock.holder),
      label:  String(lock.label || ''),
      acquiredAt: Number(lock.acquiredAt) || 0
    };
  } catch (e) {
    _clearLock();
    return null;
  }
}

function _isLockExpired(lock) {
  if (!lock) return true;
  return (Date.now() - lock.acquiredAt) > LOCK_TTL_MS;
}

function _writeLock(holder, label) {
  PropertiesService.getScriptProperties().setProperty(PROP_LOCK, JSON.stringify({
    holder: String(holder || ''),
    label: String(label || ''),
    acquiredAt: Date.now()
  }));
}

function _clearLock() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_LOCK);
}

function _appendEditLog(ss, lock, runs, steps) {
  const sh = ss.getSheetByName(SHEET_NAME_EDIT_LOG);
  const when = new Date();
  const meta = {
    timestamp: when.getTime(),
    editor: String((lock && lock.label) || 'Unknown editor')
  };
  sh.appendRow([
    when,
    meta.editor,
    String((lock && lock.holder) || ''),
    runs && runs.length ? runs.length : 0,
    steps && steps.length ? steps.length : 0
  ]);
  return meta;
}

function _readLastEdit(ss) {
  const cached = _readLastEditProperty();
  if (cached) return cached;
  const sh = ss.getSheetByName(SHEET_NAME_EDIT_LOG);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const row = sh.getRange(last, 1, 1, EDIT_LOG_HEADERS.length).getValues()[0];
  const ts = row[0] instanceof Date ? row[0].getTime() : 0;
  return _writeLastEditProperty({
    timestamp: ts,
    editor: String(row[1] || '')
  });
}

function _normalizeRuns(runs, tz) {
  return (runs || []).map(function (r) {
    return {
      id: String(r.id || _uid()),
      name: String(r.name || ''),
      color: _safeColor(r.color, ''),
      order: Number(r.order) || 0,
      stoppedAt: _toIso(r.stoppedAt, tz),
      plannedEnd: _toIso(r.plannedEnd, tz),
      metadata: _metadataText(r.metadata)
    };
  });
}

function _normalizeSteps(steps, tz) {
  return (steps || []).map(function (s) {
    const mode = String(s.mode || '').toLowerCase() === 'open' ? 'open' : 'fixed';
    return {
      id: String(s.id || _uid()),
      runId: String(s.runId || ''),
      name: String(s.name || ''),
      start: _toIso(s.start, tz),
      dur: Number(s.dur) || 0,
      mode: mode,
      stoppedAt: mode === 'open' ? _toIso(s.stoppedAt, tz) : '',
      metadata: _metadataText(s.metadata)
    };
  });
}

function _stepLastDayForRunState(step, tz, todayDay) {
  const start = _isoDay(step.start, tz);
  if (start == null) return null;
  if (step.mode === 'open') {
    const stopped = _isoDay(step.stoppedAt, tz);
    return Math.max(start, stopped == null ? todayDay : stopped);
  }
  const dur = Math.max(Number(step.dur) || 0, 0);
  return start + Math.max(dur || 1, 1) - 1;
}

function _reconcileStoppedRunsWithSteps(runs, steps, tz) {
  const todayDay = _isoDay(_todayIso(tz), tz);
  runs.forEach(function (run) {
    const stoppedDay = _isoDay(run.stoppedAt, tz);
    if (stoppedDay == null) return;
    const hasLaterStep = steps.some(function (step) {
      if (step.runId !== run.id) return false;
      const last = _stepLastDayForRunState(step, tz, todayDay);
      return last != null && last > stoppedDay;
    });
    if (hasLaterStep) run.stoppedAt = '';
  });
  return runs;
}

function _normalizeTemplateSteps(steps) {
  return (steps || []).map(function (s) {
    const startDay = Math.max(1, Math.round(Number(s.startDay) || Number(s.day) || 1));
    const dur = Math.max(0, Math.round(Number(s.dur) || 0));
    return {
      id: String(s.id || _uid()),
      name: String(s.name || ''),
      startDay: startDay,
      dur: dur
    };
  }).sort(function (a, b) {
    return (a.startDay - b.startDay) || a.name.localeCompare(b.name);
  });
}

function _templateEndDay(step) {
  const dur = Math.max(Number(step.dur) || 0, 0);
  return Number(step.startDay) + Math.max(dur || 1, 1) - 1;
}

function _normalizeTemplateInput(t) {
  t = t || {};
  const steps = _normalizeTemplateSteps(t.steps);
  const maxStepEnd = steps.reduce(function (m, s) { return Math.max(m, _templateEndDay(s)); }, 1);
  const runLength = Math.max(1, Math.round(Number(t.runLength) || maxStepEnd || 1));
  const name = String(t.name || '').trim() || 'Untitled template';
  return {
    id: String(t.id || _uid()),
    name: name,
    runName: String(t.runName || '').trim() || name,
    color: _safeColor(t.color, '#4169e8'),
    runLength: Math.max(runLength, maxStepEnd),
    steps: steps,
    createdAt: Number(t.createdAt) || 0,
    updatedAt: Number(t.updatedAt) || 0
  };
}

function _loadTemplates(ss) {
  const rows = _readAll(ss, SHEET_NAME_TEMPLATES);
  return rows.map(function (row) {
    let steps = [];
    try {
      steps = JSON.parse(String(row.stepsJson || '[]'));
      if (!Array.isArray(steps)) steps = [];
    } catch (e) {
      steps = [];
    }
    const t = _normalizeTemplateInput({
      id: row.id,
      name: row.name,
      runName: row.runName,
      color: row.color,
      runLength: row.runLength,
      steps: steps
    });
    t.createdAt = row.createdAt instanceof Date ? row.createdAt.getTime() : Number(row.createdAt) || 0;
    t.updatedAt = row.updatedAt instanceof Date ? row.updatedAt.getTime() : Number(row.updatedAt) || 0;
    return t;
  }).sort(function (a, b) {
    return a.name.localeCompare(b.name) || String(a.id).localeCompare(String(b.id));
  });
}

function _writeTemplates(ss, templates) {
  const sh = ss.getSheetByName(SHEET_NAME_TEMPLATES);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  const rows = (templates || []).map(function (t) {
    return [
      t.id,
      t.name,
      t.runName,
      t.color,
      t.runLength,
      JSON.stringify(t.steps || []),
      t.createdAt || '',
      t.updatedAt || '',
      ''
    ];
  });
  if (rows.length) sh.getRange(2, 1, rows.length, TEMPLATES_HEADERS.length).setValues(rows);
}

function _metaPayload() {
  const tz = _timeZone(null);
  return {
    lock: _activeLock(),
    lastEdit: _readLastEditProperty(),
    dataVersion: _readDataVersion(),
    serverTime: Date.now(),
    lockTtlMs: LOCK_TTL_MS,
    todayIso: _todayIso(tz)
  };
}

function _loadPayload(ss) {
  const tz = _timeZone(ss);
  const runsRaw  = _readAll(ss, SHEET_NAME_RUNS);
  const stepsRaw = _readAll(ss, SHEET_NAME_STEPS);
  const meta = _metaPayload();
  const normalizedSteps = _normalizeSteps(stepsRaw, tz);
  meta.runs = _reconcileStoppedRunsWithSteps(_normalizeRuns(runsRaw, tz), normalizedSteps, tz).map(function (r) {
    if (!r.color) r.color = '#44d8e0';
    return r;
  });
  meta.steps = normalizedSteps;
  meta.templates = _loadTemplates(ss);
  meta.lastEdit = _readLastEdit(ss);
  meta.todayIso = _todayIso(tz);
  return meta;
}

function _payloadFromNormalized(runs, steps, lock, lastEdit, dataVersion, tz, templates) {
  return {
    ok: true,
    runs: runs,
    steps: steps,
    templates: templates || [],
    lock: lock,
    lastEdit: lastEdit,
    dataVersion: dataVersion,
    serverTime: Date.now(),
    lockTtlMs: LOCK_TTL_MS,
    todayIso: _todayIso(tz)
  };
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
  return _loadPayload(_getSpreadsheet());
}

function loadMeta() {
  _checkAccess();
  return _metaPayload();
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
      const meta = _metaPayload();
      meta.ok = false;
      meta.reason = 'held';
      meta.lock = cur;
      return meta;
    }
    _writeLock(sessionId, String(label || ''));
    const meta = _metaPayload();
    meta.ok = true;
    meta.lock = _readLock();
    return meta;
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
    const meta = _metaPayload();
    meta.ok = true;
    return meta;
  } finally {
    scriptLock.releaseLock();
  }
}

// ----------------------------------------------------------------------------
// BATCH WRITE: replaces the entire dataset in one transaction.
// Only the current lock holder may call this; lock is released on success.
// ----------------------------------------------------------------------------
function saveAll(sessionId, runs, steps, templates) {
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
    const tz = _timeZone(ss);
    const rSh = ss.getSheetByName(SHEET_NAME_RUNS);
    const sSh = ss.getSheetByName(SHEET_NAME_STEPS);
    const normalizedSteps = _normalizeSteps(steps, tz);
    const normalizedRuns = _reconcileStoppedRunsWithSteps(_normalizeRuns(runs, tz), normalizedSteps, tz);
    const runRows = normalizedRuns.map(function (r) {
      return [r.id, r.name, r.color, r.order, '', r.stoppedAt, r.plannedEnd, r.metadata];
    });
    const stepRows = normalizedSteps.map(function (s) {
      return [s.id, s.runId, s.name, s.start, s.dur, s.mode, s.stoppedAt, '', s.metadata];
    });

    // Clear existing data rows (keep headers)
    if (rSh.getLastRow() > 1) rSh.getRange(2, 1, rSh.getLastRow() - 1, rSh.getLastColumn()).clearContent();
    if (sSh.getLastRow() > 1) sSh.getRange(2, 1, sSh.getLastRow() - 1, sSh.getLastColumn()).clearContent();

    // Write fresh data.
    _writeRowsWithMetadata(rSh, RUNS_HEADERS, runRows);
    _writeRowsWithMetadata(sSh, STEPS_HEADERS, stepRows);
    let normalizedTemplates = _loadTemplates(ss);
    if (Array.isArray(templates)) {
      const now = Date.now();
      const existingById = {};
      normalizedTemplates.forEach(function (t) { existingById[t.id] = t; });
      normalizedTemplates = templates.map(function (t) {
        const normalized = _normalizeTemplateInput(t);
        const existing = existingById[normalized.id];
        normalized.createdAt = Number(t.createdAt) || (existing && existing.createdAt) || now;
        normalized.updatedAt = Number(t.updatedAt) || (existing && existing.updatedAt) || now;
        return normalized;
      }).sort(function (a, b) {
        return a.name.localeCompare(b.name) || String(a.id).localeCompare(String(b.id));
      });
      _writeTemplates(ss, normalizedTemplates);
    }
    const lastEdit = _appendEditLog(ss, cur, normalizedRuns, normalizedSteps);
    SpreadsheetApp.flush();
    _writeLastEditProperty(lastEdit);
    const dataVersion = _bumpDataVersion();
    _clearLock();
    return _payloadFromNormalized(normalizedRuns, normalizedSteps, null, lastEdit, dataVersion, tz, normalizedTemplates);
  } finally {
    scriptLock.releaseLock();
  }
}

function saveTemplate(sessionId, template) {
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
        message: 'Your edit session expired or was taken over. Unlock again before saving a template.'
      };
    }
    const ss = _getSpreadsheet();
    const normalized = _normalizeTemplateInput(template);
    const now = Date.now();
    const templates = _loadTemplates(ss);
    const idx = templates.findIndex(function (t) { return t.id === normalized.id; });
    if (idx >= 0) {
      normalized.createdAt = templates[idx].createdAt || now;
      normalized.updatedAt = now;
      templates[idx] = normalized;
    } else {
      normalized.createdAt = now;
      normalized.updatedAt = now;
      templates.push(normalized);
    }
    templates.sort(function (a, b) {
      return a.name.localeCompare(b.name) || String(a.id).localeCompare(String(b.id));
    });
    _writeTemplates(ss, templates);
    const lastEdit = _appendEditLog(ss, cur, [], []);
    SpreadsheetApp.flush();
    _writeLastEditProperty(lastEdit);
    const dataVersion = _bumpDataVersion();
    const meta = _metaPayload();
    meta.ok = true;
    meta.templates = templates;
    meta.dataVersion = dataVersion;
    meta.lock = _readLock();
    meta.lastEdit = lastEdit;
    return meta;
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
  _props().deleteProperty(PROP_LAST_EDIT);
  _bumpDataVersion();
}

function _forceClearLock_() {
  _clearLock();
}

function _logDataSheetUrl_() {
  const ss = _getSpreadsheet();
  Logger.log('Data sheet id: ' + ss.getId());
  Logger.log('Data sheet url: ' + ss.getUrl());
}

function _forceSchemaMigration_() {
  const ss = SpreadsheetApp.openById(_props().getProperty(PROP_SS_ID));
  _ensureSchema(ss);
  _migrateRunStops(ss);
  _migratePlannedRunEnds(ss);
  _bumpDataVersion();
  _props().setProperty(PROP_SCHEMA_VERSION, SCHEMA_VERSION);
}

function _unlinkDataSheetForDev_() {
  const props = _props();
  [
    PROP_SS_ID,
    PROP_SCHEMA_VERSION,
    PROP_DATA_VERSION,
    PROP_LAST_EDIT,
    PROP_TIME_ZONE
  ].forEach(function (key) { props.deleteProperty(key); });
  _clearLock();
}
