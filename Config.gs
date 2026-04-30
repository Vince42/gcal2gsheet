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
  },

  toastTitle: 'Calendar import',
});

const CONFIG_SHEET_SPEC = Object.freeze({
  legacyName: 'Config',
  keyHeader: 'Key',
  valueHeader: 'Value',
  keys: {
    json: 'ConfigJson',
    lastValidJson: 'LastValidConfigJson',
    importStartDate: 'ImportStartDate',
    calendarNames: 'CalendarNames',
    defaultCalendarName: 'DefaultCalendarName',
    validity: 'Validity',
  },
});

let CONFIG = freezeConfigCopy_(DEFAULT_CONFIG);

function refreshConfig_() {
  const state = readConfigStateFromSheet_();
  if (!state.isValid) {
    throw new Error(state.message);
  }

  CONFIG = freezeConfigCopy_(state.config);
  return CONFIG;
}




function readConfigStateFromSheet_() {
  const refs = ensureConfigSheetAndRanges_();
  const jsonRaw = toText_(refs.valuesByKey[CONFIG_SHEET_SPEC.keys.json]).trim();
  const lastValidJsonRaw = toText_(refs.valuesByKey[CONFIG_SHEET_SPEC.keys.lastValidJson]).trim();
  const importStartDateOverride = normalizeDateCellToIsoOrText_(
    refs.valuesByKey[CONFIG_SHEET_SPEC.keys.importStartDate],
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
  );
  const calendarNamesOverride = toText_(refs.valuesByKey[CONFIG_SHEET_SPEC.keys.calendarNames]).trim();
  const defaultCalendarNameOverride = toText_(refs.valuesByKey[CONFIG_SHEET_SPEC.keys.defaultCalendarName]).trim();

  let parsedConfig;
  let validationMessage;
  try {
    parsedConfig = jsonRaw ? JSON.parse(jsonRaw) : cloneConfig_(DEFAULT_CONFIG);
  } catch (error) {
    try {
      parsedConfig = lastValidJsonRaw
        ? JSON.parse(lastValidJsonRaw)
        : cloneConfig_(DEFAULT_CONFIG);
      validationMessage = `Invalid JSON in ${CONFIG_SHEET_SPEC.keys.json}: ${error.message}. Restored the last valid configuration snapshot; fix ${CONFIG_SHEET_SPEC.keys.json} and save.`;
    } catch (fallbackError) {
      parsedConfig = cloneConfig_(DEFAULT_CONFIG);
      validationMessage = `Invalid JSON in ${CONFIG_SHEET_SPEC.keys.json}: ${error.message}. No valid snapshot available; defaults loaded.`;
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
  refs.cellsByKey[CONFIG_SHEET_SPEC.keys.validity].setValue(validationMessage);

  return {
    isValid,
    message: validationMessage,
    config: merged,
  };
}

function writeConfigToSheet_(config) {
  const refs = ensureConfigSheetAndRanges_();
  const jsonPayload = JSON.stringify(config, null, 2);
  const detailRows = [
    {
      key: 'json',
      value: jsonPayload,
      sourceRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.json].getA1Notation(),
      targetRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.json].getA1Notation(),
    },
    {
      key: 'lastValidJson',
      value: jsonPayload,
      sourceRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.lastValidJson].getA1Notation(),
      targetRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.lastValidJson].getA1Notation(),
    },
    {
      key: 'importStartDate',
      value: config.importStartDate || '',
      sourceRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.importStartDate].getA1Notation(),
      targetRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.importStartDate].getA1Notation(),
    },
    {
      key: 'calendarNames',
      value: (config.calendarNames || []).join(', '),
      sourceRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.calendarNames].getA1Notation(),
      targetRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.calendarNames].getA1Notation(),
    },
    {
      key: 'defaultCalendarName',
      value: config.defaultCalendarName || '',
      sourceRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.defaultCalendarName].getA1Notation(),
      targetRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.defaultCalendarName].getA1Notation(),
    },
  ];

  refs.cellsByKey[CONFIG_SHEET_SPEC.keys.json].setValue(jsonPayload);
  refs.cellsByKey[CONFIG_SHEET_SPEC.keys.lastValidJson].setValue(jsonPayload);
  refs.cellsByKey[CONFIG_SHEET_SPEC.keys.importStartDate].setValue(config.importStartDate || '');
  refs.cellsByKey[CONFIG_SHEET_SPEC.keys.calendarNames].setValue((config.calendarNames || []).join(', '));
  refs.cellsByKey[CONFIG_SHEET_SPEC.keys.defaultCalendarName].setValue(config.defaultCalendarName || '');
  refs.cellsByKey[CONFIG_SHEET_SPEC.keys.validity].setValue('VALID');
  return detailRows;
}

function ensureConfigSheetAndRanges_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet available.');
  }
  const sheet = resolveManagedConfigSheet_(ss);

  const layout = [
    [CONFIG_SHEET_SPEC.keyHeader, CONFIG_SHEET_SPEC.valueHeader],
    ['ConfigJson', ''],
    ['LastValidConfigJson', ''],
    ['ImportStartDate', ''],
    ['CalendarNames', ''],
    ['DefaultCalendarName', ''],
    ['Validity', ''],
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

  const cellsByKey = getConfigCellsByKey_(sheet);
  const valuesByKey = getConfigValuesByKey_(sheet);
  if (toText_(valuesByKey[CONFIG_SHEET_SPEC.keys.json]).trim() === '') {
    const defaultJson = JSON.stringify(DEFAULT_CONFIG, null, 2);
    cellsByKey[CONFIG_SHEET_SPEC.keys.json].setValue(defaultJson);
    cellsByKey[CONFIG_SHEET_SPEC.keys.lastValidJson].setValue(defaultJson);
    cellsByKey[CONFIG_SHEET_SPEC.keys.importStartDate].setValue(DEFAULT_CONFIG.importStartDate);
    cellsByKey[CONFIG_SHEET_SPEC.keys.calendarNames].setValue(DEFAULT_CONFIG.calendarNames.join(', '));
    cellsByKey[CONFIG_SHEET_SPEC.keys.defaultCalendarName].setValue(DEFAULT_CONFIG.defaultCalendarName);
    cellsByKey[CONFIG_SHEET_SPEC.keys.validity].setValue('VALID');
  }

  return { sheet, cellsByKey, valuesByKey: getConfigValuesByKey_(sheet) };
}

function resolveManagedConfigSheet_(ss) {
  const sheet = ss.getSheetByName(CONFIG_SHEET_SPEC.legacyName);
  if (!sheet) {
    throw new Error('Sheet "Config" is missing. Please create/restore it.');
  }
  if (!isManagedConfigSheetCandidate_(sheet)) {
    throw new Error(
      'Sheet "Config" exists but is not managed by this script. Rename your existing sheet and create/restore the managed Config layout.'
    );
  }
  return sheet;
}

function isManagedConfigSheetCandidate_(sheet) {
  return hasManagedConfigLayout_(sheet) || hasLegacyManagedConfigLayout_(sheet);
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
  ];
  const values = sheet.getRange(1, 1, expectedKeys.length, 1).getValues();
  return expectedKeys.every((key, index) => toText_(values[index][0]) === key);
}

function hasLegacyManagedConfigLayout_(sheet) {
  const expectedKeys = [
    'Key',
    'ConfigJson',
    'ImportStartDate',
    'CalendarNames',
    'DefaultCalendarName',
    'Validity',
  ];
  const values = sheet.getRange(1, 1, expectedKeys.length, 1).getValues();
  return expectedKeys.every((key, index) => toText_(values[index][0]) === key);
}

function getConfigCellsByKey_(sheet) {
  const values = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
  const cellsByKey = {};
  Object.keys(CONFIG_SHEET_SPEC.keys).forEach((name) => {
    const key = CONFIG_SHEET_SPEC.keys[name];
    for (let i = 0; i < values.length; i += 1) {
      if (toText_(values[i][0]).trim() === key) {
        cellsByKey[key] = sheet.getRange(i + 1, 2);
        break;
      }
    }
    if (!cellsByKey[key]) {
      throw new Error(`Missing config key "${key}" in column A of Config sheet.`);
    }
  });
  return cellsByKey;
}

function getConfigValuesByKey_(sheet) {
  const cellsByKey = getConfigCellsByKey_(sheet);
  const valuesByKey = {};
  Object.keys(cellsByKey).forEach((key) => {
    valuesByKey[key] = cellsByKey[key].getValue();
  });
  return valuesByKey;
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
    const startRow = 9;
    const maxLines = 200;
    const headerRange = sheet.getRange(startRow - 1, 1, 1, 2);
    if (
      toText_(headerRange.getCell(1, 1).getValue()) !== 'Timestamp'
      || toText_(headerRange.getCell(1, 2).getValue()) !== 'Message'
    ) {
      headerRange.setValues([['Timestamp', 'Message']]);
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
