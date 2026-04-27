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
  name: 'Config',
  ranges: {
    json: 'CFG_JSON',
    importStartDate: 'CFG_IMPORT_START_DATE',
    calendarNames: 'CFG_CALENDAR_NAMES',
    defaultCalendarName: 'CFG_DEFAULT_CALENDAR_NAME',
    validity: 'CFG_VALIDITY',
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

function getConfigForDialog_() {
  const state = readConfigStateFromSheet_();

  return {
    config: cloneConfig_(state.config),
    defaults: cloneConfig_(DEFAULT_CONFIG),
    validation: {
      isValid: state.isValid,
      message: state.message,
    },
  };
}

function saveConfigFromDialog_(payload) {
  if (!payload) {
    throw new Error('Missing configuration payload.');
  }

  const basicConfig = mergeConfigWithDefaults_(payload.config || {});
  validateConfig_(basicConfig);
  const previousConfig = cloneConfig_(CONFIG);
  const scopeChanged = hasScopeAffectingConfigChange_(previousConfig, basicConfig);
  writeConfigToSheet_(basicConfig);

  const props = getConfigPropertiesStore_();
  if (scopeChanged) {
    clearSyncTokenProperties_(props, [
      DEFAULT_CONFIG.propertyPrefix,
      previousConfig.propertyPrefix,
      basicConfig.propertyPrefix,
    ]);
  }

  CONFIG = freezeConfigCopy_(basicConfig);
  return { success: true };
}

function resetConfigToDefault_() {
  writeConfigToSheet_(cloneConfig_(DEFAULT_CONFIG));

  const props = getConfigPropertiesStore_();
  clearSyncTokenProperties_(props, [DEFAULT_CONFIG.propertyPrefix, CONFIG.propertyPrefix]);

  CONFIG = freezeConfigCopy_(DEFAULT_CONFIG);
  return { success: true };
}

function readConfigStateFromSheet_() {
  const refs = ensureConfigSheetAndRanges_();
  const jsonRaw = toText_(refs.namedRanges.json.getValue()).trim();
  const importStartDateOverride = toText_(refs.namedRanges.importStartDate.getValue()).trim();
  const calendarNamesOverride = toText_(refs.namedRanges.calendarNames.getValue()).trim();
  const defaultCalendarNameOverride = toText_(refs.namedRanges.defaultCalendarName.getValue()).trim();

  let parsedConfig;
  let validationMessage;
  try {
    parsedConfig = jsonRaw ? JSON.parse(jsonRaw) : cloneConfig_(DEFAULT_CONFIG);
  } catch (error) {
    parsedConfig = cloneConfig_(DEFAULT_CONFIG);
    validationMessage = `Invalid JSON in ${CONFIG_SHEET_SPEC.ranges.json}: ${error.message}`;
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
    config: isValid ? merged : cloneConfig_(DEFAULT_CONFIG),
  };
}

function writeConfigToSheet_(config) {
  const refs = ensureConfigSheetAndRanges_();
  refs.namedRanges.json.setValue(JSON.stringify(config, null, 2));
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

  let sheet = ss.getSheetByName(CONFIG_SHEET_SPEC.name);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_SPEC.name);
  }
  if (!sheet.isSheetHidden()) {
    sheet.hideSheet();
  }

  const layout = [
    ['Key', 'Value'],
    ['ConfigJson', ''],
    ['ImportStartDate', ''],
    ['CalendarNames', ''],
    ['DefaultCalendarName', ''],
    ['Validity', ''],
  ];

  const current = sheet.getRange(1, 1, layout.length, 1).getValues();
  const keysNeedWrite = layout.some((row, i) => current[i][0] !== row[0]);
  if (keysNeedWrite) {
    sheet.getRange(1, 1, layout.length, 1).setValues(layout.map((row) => [row[0]]));
  }

  function ensureNamedRange(name, row, col) {
    const range = sheet.getRange(row, col, 1, 1);
    const existing = ss.getRangeByName(name);
    if (!existing || existing.getA1Notation() !== range.getA1Notation() || existing.getSheet().getSheetId() !== sheet.getSheetId()) {
      ss.setNamedRange(name, range);
    }
    return ss.getRangeByName(name);
  }

  const namedRanges = {
    json: ensureNamedRange(CONFIG_SHEET_SPEC.ranges.json, 2, 2),
    importStartDate: ensureNamedRange(CONFIG_SHEET_SPEC.ranges.importStartDate, 3, 2),
    calendarNames: ensureNamedRange(CONFIG_SHEET_SPEC.ranges.calendarNames, 4, 2),
    defaultCalendarName: ensureNamedRange(CONFIG_SHEET_SPEC.ranges.defaultCalendarName, 5, 2),
    validity: ensureNamedRange(CONFIG_SHEET_SPEC.ranges.validity, 6, 2),
  };

  if (toText_(namedRanges.json.getValue()).trim() === '') {
    namedRanges.json.setValue(JSON.stringify(DEFAULT_CONFIG, null, 2));
    namedRanges.importStartDate.setValue(DEFAULT_CONFIG.importStartDate);
    namedRanges.calendarNames.setValue(DEFAULT_CONFIG.calendarNames.join(', '));
    namedRanges.defaultCalendarName.setValue(DEFAULT_CONFIG.defaultCalendarName);
    namedRanges.validity.setValue('VALID');
  }

  return { sheet, namedRanges };
}

function getConfigPropertiesStore_() {
  try {
    return PropertiesService.getDocumentProperties();
  } catch (error) {
    const message = error && error.message ? String(error.message) : String(error);
    if (message.toUpperCase().includes('PERMISSION_DENIED')) {
      return PropertiesService.getScriptProperties();
    }
    throw error;
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
  return previousCalendars !== nextCalendars;
}

function clearSyncTokenProperties_(props, prefixes) {
  const uniquePrefixes = Array.from(
    new Set((prefixes || []).filter((prefix) => typeof prefix === 'string' && prefix.length > 0))
  );
  if (uniquePrefixes.length === 0) {
    return;
  }

  const allProps = props.getProperties();
  Object.keys(allProps).forEach((key) => {
    if (uniquePrefixes.some((prefix) => key.startsWith(prefix))) {
      props.deleteProperty(key);
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

  if (!/^\d{4}-\d{2}-\d{2}$/.test(config.importStartDate)) {
    throw new Error(`Invalid importStartDate: ${config.importStartDate}`);
  }

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
