const DEFAULT_CONFIG = Object.freeze({
  sheetName: 'Calendar',
  stateSheetName: '_calendar_state',
  tableName: 'Calendar',
  statusCell: 'L1',

  // Lower bound for managed imports: yyyy-mm-dd
  importStartDate: '2024-01-01',

  calendarNames: ['Event', 'dedc', 'EEC', 'CTG'],
  defaultCalendarName: 'Event',

  header: [
    'Calendar',
    'Event',
    'Date',
    'Start',
    'End',
    'Duration',
    'Customer',
    'Project',
    'InvoiceNumber',
    'InvoiceDate',
  ],

  stateHeader: ['EventKey', 'RowKind'],

  rowKind: {
    normal: 'NORMAL',
    changedCopy: 'CHANGED_COPY',
    unmanaged: 'UNMANAGED',
  },

  propertyPrefix: 'CALSYNC_TOKEN_',
  configPropertyKey: 'CALSYNC_CONFIG_JSON',

  colors: {
    normal: '#000000',
    invoiced: '#7A1F1F',
    changed: '#1B5E20',
  },

  menu: {
    title: 'Calendar Import',
    item: 'Update calendar sheet',
    configItem: 'Configuration...',
  },

  toastTitle: 'Calendar import',
});

const CONFIG_SHEET_SPEC = Object.freeze({
  legacyName: 'Config',
  technicalName: '_gcal2gsheet_config',
  ranges: {
    json: 'CFG_JSON',
    lastValidJson: 'CFG_LAST_VALID_JSON',
    importStartDate: 'CFG_IMPORT_START_DATE',
    calendarNames: 'CFG_CALENDAR_NAMES',
    defaultCalendarName: 'CFG_DEFAULT_CALENDAR_NAME',
    validity: 'CFG_VALIDITY',
    debugLog: 'CFG_DEBUG_LOG',
  },
});
const CONFIG_DIALOG_REVISION = '2026-04-29-r4';

let CONFIG = freezeConfigCopy_(DEFAULT_CONFIG);

function refreshConfig_() {
  const state = readConfigStateFromSheet_();
  if (!state.isValid) {
    throw new Error(state.message);
  }

  CONFIG = freezeConfigCopy_(state.config);
  return CONFIG;
}

function getConfigForDialog_() {
  const state = readConfigStateFromSheet_();

  return {
    config: cloneConfig_(state.config),
    defaults: cloneConfig_(DEFAULT_CONFIG),
    revision: CONFIG_DIALOG_REVISION,
    validation: {
      isValid: state.isValid,
      message: state.message,
    },
  };
}

function saveConfigFromDialog_(payload) {
  const saveId = `cfg-save-${new Date().toISOString()}`;
  logStorageDebug_('save-config', `${saveId} begin`);
  logStorageDebug_('save-config.step', '1) validate payload presence');

  if (!payload) {
    throw new Error('Missing configuration payload.');
  }

  logStorageDebug_('save-config.step', '2) merge payload with defaults');
  const basicConfig = withConfigSaveDebug_('save-config.merge', () => mergeConfigWithDefaults_(payload.config || {}));
  logStorageDebug_('save-config.step', '3) validate merged config');
  withConfigSaveDebug_('save-config.validate', () => validateConfig_(basicConfig));
  logStorageDebug_('save-config.step', '4) read persisted config state');
  const persistedState = withConfigSaveDebug_('save-config.read-state', () => readConfigStateFromSheet_());
  logStorageDebug_('save-config.step', '5) compute scope-change decision');
  const previousConfig = cloneConfig_(persistedState.config);
  const scopeChanged = !persistedState.isValid
    || hasScopeAffectingConfigChange_(previousConfig, basicConfig);
  logStorageDebug_('save-config.step', '6) write config values to config sheet');
  withConfigSaveDebug_('save-config.write-sheet', () => writeConfigToSheet_(basicConfig));

  if (scopeChanged) {
    logStorageDebug_('save-config.step', '7) scope changed: open properties store');
    const props = getConfigPropertiesStore_();
    logStorageDebug_('save-config.step', '8) clear sync-token properties');
    withConfigSaveDebug_('save-config.clear-sync', () => clearSyncTokenProperties_(props, [
      DEFAULT_CONFIG.propertyPrefix,
      previousConfig.propertyPrefix,
      basicConfig.propertyPrefix,
    ]));
  }

  logStorageDebug_('save-config.step', '9) refresh in-memory CONFIG snapshot');
  CONFIG = freezeConfigCopy_(basicConfig);
  logStorageDebug_('save-config', `${saveId} done`);
  return { success: true, saveId };
}

function resetConfigToDefault_() {
  writeConfigToSheet_(cloneConfig_(DEFAULT_CONFIG));

  clearSyncTokenProperties_(getConfigPropertiesStore_(), [
    DEFAULT_CONFIG.propertyPrefix,
    CONFIG.propertyPrefix,
  ]);

  CONFIG = freezeConfigCopy_(DEFAULT_CONFIG);
  return { success: true };
}

function readConfigStateFromSheet_() {
  const refs = ensureConfigSheetAndRanges_();
  const jsonRaw = toText_(refs.namedRanges.json.getValue()).trim();
  const lastValidJsonRaw = toText_(refs.namedRanges.lastValidJson.getValue()).trim();
  const importStartDateOverride = normalizeDateCellToIsoOrText_(
    refs.namedRanges.importStartDate.getValue(),
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
  );
  const calendarNamesOverride = toText_(refs.namedRanges.calendarNames.getValue()).trim();
  const defaultCalendarNameOverride = toText_(refs.namedRanges.defaultCalendarName.getValue()).trim();

  let parsedConfig;
  let validationMessage;
  try {
    parsedConfig = jsonRaw ? JSON.parse(jsonRaw) : cloneConfig_(DEFAULT_CONFIG);
  } catch (error) {
    try {
      parsedConfig = lastValidJsonRaw
        ? JSON.parse(lastValidJsonRaw)
        : cloneConfig_(DEFAULT_CONFIG);
      validationMessage = `Invalid JSON in ${CONFIG_SHEET_SPEC.ranges.json}: ${error.message}. Restored the last valid configuration snapshot; fix ${CONFIG_SHEET_SPEC.ranges.json} and save.`;
    } catch (fallbackError) {
      parsedConfig = cloneConfig_(DEFAULT_CONFIG);
      validationMessage = `Invalid JSON in ${CONFIG_SHEET_SPEC.ranges.json}: ${error.message}. No valid snapshot available; defaults loaded.`;
    }
  }

  if (importStartDateOverride) {
    parsedConfig.importStartDate = importStartDateOverride;
  }
  if (calendarNamesOverride) {
    parsedConfig.calendarNames = calendarNamesOverride
      .split(',')
      .map((v) => v.trim())
      .filter((v) => v);
  }
  if (defaultCalendarNameOverride) {
    parsedConfig.defaultCalendarName = defaultCalendarNameOverride;
  }

  const merged = mergeConfigWithDefaults_(parsedConfig || {});
  if (!validationMessage) {
    try {
      validateConfig_(merged);
      validationMessage = 'VALID';
    } catch (error) {
      validationMessage = error.message;
    }
  }

  const isValid = validationMessage === 'VALID';
  refs.namedRanges.validity.setValue(validationMessage);

  return {
    isValid,
    message: validationMessage,
    config: merged,
  };
}

function writeConfigToSheet_(config) {
  const refs = ensureConfigSheetAndRanges_();
  const jsonPayload = JSON.stringify(config, null, 2);
  refs.namedRanges.json.setValue(jsonPayload);
  refs.namedRanges.lastValidJson.setValue(jsonPayload);
  refs.namedRanges.importStartDate.setValue(config.importStartDate || '');
  refs.namedRanges.calendarNames.setValue((config.calendarNames || []).join(', '));
  refs.namedRanges.defaultCalendarName.setValue(config.defaultCalendarName || '');
  refs.namedRanges.validity.setValue('VALID');
}

function ensureConfigSheetAndRanges_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet available.');
  }

  let sheet = resolveManagedConfigSheet_(ss);
  if (!sheet.isSheetHidden()) {
    sheet.hideSheet();
  }

  const layout = [
    ['Key', 'Value'],
    ['ConfigJson', ''],
    ['LastValidConfigJson', ''],
    ['ImportStartDate', ''],
    ['CalendarNames', ''],
    ['DefaultCalendarName', ''],
    ['Validity', ''],
    ['DebugLog', ''],
  ];

  const current = sheet.getRange(1, 1, layout.length, 1).getValues();
  const legacyLayoutKeys = [
    'Key',
    'ConfigJson',
    'ImportStartDate',
    'CalendarNames',
    'DefaultCalendarName',
    'Validity',
  ];
  const currentKeys = current.map((row) => toText_(row[0]));
  const isLegacyLayout = legacyLayoutKeys.every((key, i) => currentKeys[i] === key);
  if (isLegacyLayout) {
    const legacyValues = sheet.getRange(2, 2, legacyLayoutKeys.length - 1, 1).getValues();
    const jsonValue = legacyValues[0][0];
    const importStartDateValue = legacyValues[1][0];
    const calendarNamesValue = legacyValues[2][0];
    const defaultCalendarNameValue = legacyValues[3][0];
    const validityValue = legacyValues[4][0];
    sheet.getRange(2, 2, 6, 1).setValues([
      [jsonValue],
      [jsonValue],
      [importStartDateValue],
      [calendarNamesValue],
      [defaultCalendarNameValue],
      [validityValue],
    ]);
  }

  const keysNeedWrite = layout.some((row, i) => current[i][0] !== row[0]);
  if (keysNeedWrite) {
    sheet.getRange(1, 1, layout.length, 1).setValues(layout.map((row) => [row[0]]));
  }

  const namedRanges = {
    json: ensureManagedNamedRange_(ss, sheet, CONFIG_SHEET_SPEC.ranges.json, 2, 2),
    lastValidJson: ensureManagedNamedRange_(ss, sheet, CONFIG_SHEET_SPEC.ranges.lastValidJson, 3, 2),
    importStartDate: ensureManagedNamedRange_(ss, sheet, CONFIG_SHEET_SPEC.ranges.importStartDate, 4, 2),
    calendarNames: ensureManagedNamedRange_(ss, sheet, CONFIG_SHEET_SPEC.ranges.calendarNames, 5, 2),
    defaultCalendarName: ensureManagedNamedRange_(ss, sheet, CONFIG_SHEET_SPEC.ranges.defaultCalendarName, 6, 2),
    validity: ensureManagedNamedRange_(ss, sheet, CONFIG_SHEET_SPEC.ranges.validity, 7, 2),
    debugLog: ensureManagedNamedRange_(ss, sheet, CONFIG_SHEET_SPEC.ranges.debugLog, 8, 2),
  };

  if (toText_(namedRanges.json.getValue()).trim() === '') {
    const defaultJson = JSON.stringify(DEFAULT_CONFIG, null, 2);
    namedRanges.json.setValue(defaultJson);
    namedRanges.lastValidJson.setValue(defaultJson);
    namedRanges.importStartDate.setValue(DEFAULT_CONFIG.importStartDate);
    namedRanges.calendarNames.setValue(DEFAULT_CONFIG.calendarNames.join(', '));
    namedRanges.defaultCalendarName.setValue(DEFAULT_CONFIG.defaultCalendarName);
    namedRanges.validity.setValue('VALID');
    namedRanges.debugLog.setValue('');
  }

  return { sheet, namedRanges };
}

function resolveManagedConfigSheet_(ss) {
  const preferred = ss.getSheetByName(CONFIG_SHEET_SPEC.legacyName);
  if (preferred) {
    return preferred;
  }

  const technicalSheet = ss.getSheetByName(CONFIG_SHEET_SPEC.technicalName);
  if (technicalSheet && isManagedConfigSheetCandidate_(ss, technicalSheet)) {
    return technicalSheet;
  }

  const legacySheet = ss.getSheetByName(CONFIG_SHEET_SPEC.legacyName);
  if (legacySheet && isManagedConfigSheetCandidate_(ss, legacySheet)) {
    return legacySheet;
  }

  const prefixCandidates = ss
    .getSheets()
    .filter((sheet) => {
      const name = sheet.getName();
      return (
        name === CONFIG_SHEET_SPEC.technicalName
        || name.indexOf(`${CONFIG_SHEET_SPEC.technicalName}_`) === 0
        || name === CONFIG_SHEET_SPEC.legacyName
      );
    });
  for (let i = 0; i < prefixCandidates.length; i += 1) {
    if (isManagedConfigSheetCandidate_(ss, prefixCandidates[i])) {
      return prefixCandidates[i];
    }
  }

  throw new Error(
    'Managed config sheet "Config" is missing. Please create/restore the "Config" sheet; new sheets are intentionally not auto-created.'
  );
}

function insertManagedConfigSheet_(ss) {
  const baseName = CONFIG_SHEET_SPEC.technicalName;
  if (!ss.getSheetByName(baseName)) {
    return ss.insertSheet(baseName);
  }

  for (let i = 1; i <= 99; i += 1) {
    const candidate = `${baseName}_${i}`;
    if (!ss.getSheetByName(candidate)) {
      return ss.insertSheet(candidate);
    }
  }

  throw new Error(`Unable to create managed config sheet for base name ${baseName}.`);
}

function isOwnedConfigSheet_(ss, sheet) {
  const namedRangeNames = Object.keys(CONFIG_SHEET_SPEC.ranges).map(
    (key) => CONFIG_SHEET_SPEC.ranges[key]
  );
  return namedRangeNames.every((name) => {
    return !!findManagedNamedRange_(ss, sheet, name);
  });
}

function isManagedConfigSheetCandidate_(ss, sheet) {
  return isOwnedConfigSheet_(ss, sheet) || hasManagedConfigLayout_(sheet);
}

function hasManagedConfigLayout_(sheet) {
  const expectedKeys = [
    'Key',
    'ConfigJson',
    'LastValidConfigJson',
    'ImportStartDate',
    'CalendarNames',
    'DefaultCalendarName',
    'Validity',
    'DebugLog',
  ];
  const values = sheet.getRange(1, 1, expectedKeys.length, 1).getValues();
  return expectedKeys.every((key, index) => toText_(values[index][0]) === key);
}

function ensureManagedNamedRange_(ss, sheet, baseName, row, col) {
  const desiredRange = sheet.getRange(row, col, 1, 1);
  const candidateNames = [baseName, `${baseName}__GCAL2GSHEET`];

  let suffix = 1;
  while (candidateNames.length < 12) {
    candidateNames.push(`${baseName}__GCAL2GSHEET_${suffix}`);
    suffix += 1;
  }

  for (let i = 0; i < candidateNames.length; i += 1) {
    const name = candidateNames[i];
    const existing = ss.getRangeByName(name);
    if (!existing || existing.getSheet().getSheetId() === sheet.getSheetId()) {
      if (
        !existing
        || existing.getA1Notation() !== desiredRange.getA1Notation()
        || existing.getSheet().getSheetId() !== sheet.getSheetId()
      ) {
        withConfigSaveDebug_(`named-range:${name}`, () => ss.setNamedRange(name, desiredRange));
      }
      return ss.getRangeByName(name);
    }
  }

  throw new Error(`Unable to reserve managed named range for ${baseName}.`);
}

function findManagedNamedRange_(ss, sheet, baseName) {
  const candidateNames = [baseName, `${baseName}__GCAL2GSHEET`];
  for (let i = 1; i <= 10; i += 1) {
    candidateNames.push(`${baseName}__GCAL2GSHEET_${i}`);
  }

  for (let i = 0; i < candidateNames.length; i += 1) {
    const range = ss.getRangeByName(candidateNames[i]);
    if (range && range.getSheet().getSheetId() === sheet.getSheetId()) {
      return range;
    }
  }

  return null;
}


function withConfigSaveDebug_(phase, fn) {
  try {
    return fn();
  } catch (error) {
    const message = error && error.message ? error.message : String(error);
    logStorageDebug_(phase, message);
    throw error;
  }
}

function getConfigPropertiesStore_() {
  try {
    logStorageDebug_('properties-store', 'Using DocumentProperties');
    return PropertiesService.getDocumentProperties();
  } catch (documentError) {
    logStorageDebug_('properties-store', `DocumentProperties unavailable: ${documentError}`);
    try {
      logStorageDebug_('properties-store', 'Using ScriptProperties');
      return PropertiesService.getScriptProperties();
    } catch (scriptError) {
      logStorageDebug_('properties-store', `ScriptProperties unavailable: ${scriptError}`);
      logStorageDebug_('properties-store', 'Falling back to no-op properties store');
      return getNoopPropertiesStore_();
    }
  }
}

function hasScopeAffectingConfigChange_(previousConfig, nextConfig) {
  if (!previousConfig) {
    return true;
  }

  if (previousConfig.importStartDate !== nextConfig.importStartDate) {
    return true;
  }

  const previousCalendars = (previousConfig.calendarNames || []).join('\n');
  const nextCalendars = (nextConfig.calendarNames || []).join('\n');
  if (previousCalendars !== nextCalendars) {
    return true;
  }

  return previousConfig.defaultCalendarName !== nextConfig.defaultCalendarName;
}

function clearSyncTokenProperties_(props, prefixes) {
  const uniquePrefixes = Array.from(
    new Set((prefixes || []).filter((prefix) => typeof prefix === 'string' && prefix.length > 0))
  );
  if (uniquePrefixes.length === 0) {
    return;
  }

  let allProps;
  try {
    allProps = props.getProperties();
  } catch (error) {
    if (isPermissionDeniedError_(error)) {
      logStorageDebug_('clear-sync-token-properties', `Ignored denied read while listing properties: ${error}`);
      return;
    }
    throw error;
  }
  Object.keys(allProps).forEach((key) => {
    if (!uniquePrefixes.some((prefix) => key.startsWith(prefix))) {
      return;
    }
    try {
      props.deleteProperty(key);
    } catch (error) {
      if (!isPermissionDeniedError_(error)) {
        throw error;
      }
      logStorageDebug_('clear-sync-token-properties', `Ignored denied delete for key "${key}": ${error}`);
    }
  });
}

function mergeConfigWithDefaults_(overrideConfig) {
  return mergeKnownShape_(DEFAULT_CONFIG, overrideConfig || {});
}

function mergeKnownShape_(baseValue, overrideValue) {
  if (Array.isArray(baseValue)) {
    if (!Array.isArray(overrideValue)) {
      return baseValue.slice();
    }
    return overrideValue.slice();
  }

  if (baseValue && typeof baseValue === 'object') {
    const merged = {};
    Object.keys(baseValue).forEach((key) => {
      const nextOverride = overrideValue && Object.prototype.hasOwnProperty.call(overrideValue, key)
        ? overrideValue[key]
        : undefined;
      merged[key] = mergeKnownShape_(baseValue[key], nextOverride);
    });
    return merged;
  }

  if (overrideValue === undefined || overrideValue === null) {
    return baseValue;
  }

  return overrideValue;
}

function validateConfig_(config) {
  assertString_(config.sheetName, 'sheetName');
  assertString_(config.stateSheetName, 'stateSheetName');
  assertString_(config.tableName, 'tableName');
  assertString_(config.statusCell, 'statusCell');
  assertString_(config.importStartDate, 'importStartDate');

  assertStrictIsoDate_(config.importStartDate, 'importStartDate');

  assertStringArray_(config.calendarNames, 'calendarNames');
  assertString_(config.defaultCalendarName, 'defaultCalendarName');
  if (!config.calendarNames.includes(config.defaultCalendarName)) {
    throw new Error('defaultCalendarName must exist in calendarNames.');
  }

  assertStringArrayWithLength_(
    config.header,
    'header',
    DEFAULT_CONFIG.header.length
  );
  assertExactArrayMatch_(config.header, DEFAULT_CONFIG.header, 'header');
  assertStringArrayWithLength_(
    config.stateHeader,
    'stateHeader',
    DEFAULT_CONFIG.stateHeader.length
  );
  assertExactArrayMatch_(
    config.stateHeader,
    DEFAULT_CONFIG.stateHeader,
    'stateHeader'
  );

  if (!config.rowKind || typeof config.rowKind !== 'object') {
    throw new Error('Invalid rowKind object.');
  }
  assertString_(config.rowKind.normal, 'rowKind.normal');
  assertString_(config.rowKind.changedCopy, 'rowKind.changedCopy');
  assertString_(config.rowKind.unmanaged, 'rowKind.unmanaged');

  assertString_(config.propertyPrefix, 'propertyPrefix');
  assertString_(config.configPropertyKey, 'configPropertyKey');

  if (!config.colors || typeof config.colors !== 'object') {
    throw new Error('Invalid colors object.');
  }
  assertString_(config.colors.normal, 'colors.normal');
  assertString_(config.colors.invoiced, 'colors.invoiced');
  assertString_(config.colors.changed, 'colors.changed');

  if (!config.menu || typeof config.menu !== 'object') {
    throw new Error('Invalid menu object.');
  }
  assertString_(config.menu.title, 'menu.title');
  assertString_(config.menu.item, 'menu.item');
  assertString_(config.menu.configItem, 'menu.configItem');

  assertString_(config.toastTitle, 'toastTitle');
}

function assertString_(value, fieldName) {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(`Invalid ${fieldName}.`);
  }
}

function assertStrictIsoDate_(value, fieldName) {
  const trimmed = value.trim();
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(trimmed);
  if (!match) {
    throw new Error(
      `Invalid ${fieldName}: "${value}". Use ISO date format YYYY-MM-DD (example: 2024-01-01).`
    );
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const utcDate = new Date(Date.UTC(year, month - 1, day));
  const isRealDate = utcDate.getUTCFullYear() === year
    && utcDate.getUTCMonth() === month - 1
    && utcDate.getUTCDate() === day;
  if (!isRealDate) {
    throw new Error(
      `Invalid ${fieldName}: "${value}" is not a real calendar date. Use format YYYY-MM-DD (example: 2024-01-01).`
    );
  }
}

function assertStringArray_(value, fieldName) {
  if (!Array.isArray(value) || value.length === 0) {
    throw new Error(`Invalid ${fieldName}: expected a non-empty string array.`);
  }

  value.forEach((entry, index) => {
    if (typeof entry !== 'string' || entry.trim() === '') {
      throw new Error(`Invalid ${fieldName}[${index}].`);
    }
  });
}

function assertStringArrayWithLength_(value, fieldName, expectedLength) {
  assertStringArray_(value, fieldName);
  if (value.length !== expectedLength) {
    throw new Error(
      `Invalid ${fieldName}: expected exactly ${expectedLength} item(s), got ${value.length}.`
    );
  }
}

function assertExactArrayMatch_(value, expected, fieldName) {
  for (let i = 0; i < expected.length; i += 1) {
    if (value[i] !== expected[i]) {
      throw new Error(
        `Invalid ${fieldName}: structural reordering is not supported. Expected "${expected[i]}" at position ${i + 1}.`
      );
    }
  }
}

function cloneConfig_(value) {
  return JSON.parse(JSON.stringify(value));
}

function freezeConfigCopy_(config) {
  return Object.freeze(cloneConfig_(config));
}

function normalizeDateCellToIsoOrText_(value, timeZone) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, timeZone || 'Etc/UTC', 'yyyy-MM-dd');
  }
  return toText_(value).trim();
}

function isPermissionDeniedError_(error) {
  const message = error && error.message ? String(error.message) : String(error);
  const upperMessage = message.toUpperCase();
  return (
    upperMessage.includes('PERMISSION_DENIED')
    || upperMessage.includes('ACCESS_DENIED')
    || (
      upperMessage.includes('READING FROM STORAGE')
      && upperMessage.includes('DENIED')
    )
  );
}

function getNoopPropertiesStore_() {
  return {
    getProperty() {
      return '';
    },
    setProperty() {},
    getProperties() {
      return {};
    },
    setProperties() {},
    deleteProperty() {},
  };
}

function logStorageDebug_(phase, message) {
  const line = `${new Date().toISOString()} [storage-debug] ${phase}: ${message}`;
  console.log(line);
  Logger.log(line);
  appendStorageDebugToSheet_(line);
}

function appendStorageDebugToSheet_(line) {
  try {
    const refs = ensureConfigSheetAndRanges_();
    const sheet = refs.sheet;
    const cell = refs.namedRanges.debugLog;
    const existing = toText_(cell.getValue()).trim();
    const existingCount = Number(existing);
    const nextCount = Number.isFinite(existingCount) && existingCount > 0 ? existingCount + 1 : 1;
    cell.setValue(String(nextCount));

    const startRow = 10;
    const maxLines = 200;
    const headerRange = sheet.getRange(startRow - 1, 1, 1, 2);
    if (
      toText_(headerRange.getCell(1, 1).getValue()) !== 'DebugTimestamp'
      || toText_(headerRange.getCell(1, 2).getValue()) !== 'DebugMessage'
    ) {
      headerRange.setValues([['DebugTimestamp', 'DebugMessage']]);
    }

    const timestamp = new Date();
    let nextRow = Math.max(sheet.getLastRow() + 1, startRow);
    if (nextRow >= startRow + maxLines) {
      sheet
        .getRange(startRow + 1, 1, maxLines - 1, 2)
        .moveTo(sheet.getRange(startRow, 1, maxLines - 1, 2));
      nextRow = startRow + maxLines - 1;
    }
    sheet.getRange(nextRow, 1, 1, 2).setValues([[timestamp, line]]);
  } catch (error) {
    const fallback = `[storage-debug] failed to persist debug line: ${error}`;
    console.log(fallback);
    Logger.log(fallback);
  }
}
