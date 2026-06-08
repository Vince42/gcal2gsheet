const LEGACY_CALENDAR_HEADER = Object.freeze([
  'Calendar',
  'Event',
  'Date',
  'Start',
  'End',
  'Duration',
  'Status',
]);

const DEFAULT_CALENDAR_HEADER = Object.freeze([
  'Calendar',
  'Event',
  'Date',
  'Start',
  'End',
  'Duration',
  'State',
]);

const LEGACY_INVOICING_HEADER = Object.freeze([
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
]);

const LEGACY_NON_BILLABLE_HEADER = Object.freeze([
  'Calendar',
  'Event',
  'Date',
  'Start',
  'End',
  'Duration',
  'Reason',
]);

const DEFAULT_HEADER = Object.freeze(['ID'].concat(DEFAULT_CALENDAR_HEADER));

const DEFAULT_INVOICING_HEADER = Object.freeze(['EventID'].concat(LEGACY_INVOICING_HEADER));

const DEFAULT_NON_BILLABLE_HEADER = Object.freeze(['EventID'].concat(LEGACY_NON_BILLABLE_HEADER));

const OBSOLETE_STATE_CONFIG_KEYS = Object.freeze([
  'stateSheetName',
  'invoicingStateSheetName',
  'nonBillableStateSheetName',
  'stateHeader',
  'invoicingStateHeader',
  'nonBillableStateHeader',
]);

const LEGACY_STATE_SHEET_NAME_PROPERTY_KEY = 'CALSYNC_LEGACY_STATE_SHEET_NAMES_JSON';

const DEFAULT_LEGACY_STATE_SHEET_NAMES = Object.freeze({
  calendar: ['_calendar_state'],
  invoicing: ['_invoicing_state'],
  nonBillable: ['_non_billable_state'],
});

const LEGACY_STATE_SHEET_CONFIG_KEY_GROUPS = Object.freeze({
  calendar: 'stateSheetName',
  invoicing: 'invoicingStateSheetName',
  nonBillable: 'nonBillableStateSheetName',
});

const DEFAULT_CONFIG = Object.freeze({
  sheetName: 'Calendar',
  tableName: 'Calendar',
  invoicingSheetName: 'Invoicing',
  invoicingTableName: 'Invoicing',
  nonBillableSheetName: 'Non-Billable',
  nonBillableTableName: 'NonBillable',
  statusCell: buildDefaultStatusCell_(DEFAULT_HEADER),

  // Lower bound for managed imports: yyyy-mm-dd
  importStartDate: '2024-01-01',

  calendarNames: ['Calendar'],
  defaultCalendarName: 'Calendar',

  header: DEFAULT_HEADER.slice(),
  invoicingHeader: DEFAULT_INVOICING_HEADER.slice(),
  nonBillableHeader: DEFAULT_NON_BILLABLE_HEADER.slice(),

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
    nonBillable: '#666666',
  },

  menu: {
    title: 'Invoicing',
    item: 'Update calendar sheet',
  },

  toastTitle: 'Invoicing',
});

const CONFIG_SHEET_SPEC = Object.freeze({
  legacyName: 'Config',
  keyHeader: 'Key',
  valueHeader: 'Value',
  keys: {
    json: 'ConfigJson',
    schemaRegistryJson: 'SchemaRegistryJson',
    validity: 'Validity',
  },
});

let CONFIG = freezeConfigCopy_(DEFAULT_CONFIG);
let LEGACY_STATE_SHEET_NAME_CANDIDATES = {
  calendar: DEFAULT_LEGACY_STATE_SHEET_NAMES.calendar.slice(),
  invoicing: DEFAULT_LEGACY_STATE_SHEET_NAMES.invoicing.slice(),
  nonBillable: DEFAULT_LEGACY_STATE_SHEET_NAMES.nonBillable.slice(),
};

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
  const schemaRaw = toText_(refs.valuesByKey[CONFIG_SHEET_SPEC.keys.schemaRegistryJson]).trim();

  let parsedConfig;
  let validationMessage;
  try {
    parsedConfig = JSON.parse(jsonRaw);
    rememberLegacyStateSheetNameCandidates_(parsedConfig);
  } catch (error) {
    validationMessage = `Invalid JSON in ${CONFIG_SHEET_SPEC.keys.json}: ${error.message}`;
  }

  let schemaRegistry;
  if (!validationMessage) {
    try {
      schemaRegistry = JSON.parse(schemaRaw);
      validateSchemaRegistry_(schemaRegistry);
    } catch (error) {
      validationMessage = `Invalid ${CONFIG_SHEET_SPEC.keys.schemaRegistryJson}: ${error.message}`;
    }
  }

  let merged = null;
  if (!validationMessage) {
    try {
      merged = mergeConfigWithDefaults_(parsedConfig || {});
      validateConfigStrictWithSchema_(merged, parsedConfig || {}, schemaRegistry);
      validationMessage = 'VALID';
    } catch (error) {
      validationMessage = error.message;
    }
  }

  const isValid = validationMessage === 'VALID';
  if (isValid) {
    writeNormalizedConfigState_(refs, merged);
  }
  refs.cellsByKey[CONFIG_SHEET_SPEC.keys.validity].setValue(validationMessage);

  return {
    isValid,
    message: validationMessage,
    config: merged || mergeConfigWithDefaults_({}),
  };
}

function writeNormalizedConfigState_(refs, config) {
  const jsonPayload = JSON.stringify(cloneConfig_(config), null, 2);
  const schemaPayload = JSON.stringify(buildDefaultSchemaRegistry_(), null, 2);
  if (toText_(refs.cellsByKey[CONFIG_SHEET_SPEC.keys.json].getValue()) !== jsonPayload) {
    refs.cellsByKey[CONFIG_SHEET_SPEC.keys.json].setValue(jsonPayload);
  }
  if (toText_(refs.cellsByKey[CONFIG_SHEET_SPEC.keys.schemaRegistryJson].getValue()) !== schemaPayload) {
    refs.cellsByKey[CONFIG_SHEET_SPEC.keys.schemaRegistryJson].setValue(schemaPayload);
  }
}

function writeConfigToSheet_(config) {
  const refs = ensureConfigSheetAndRanges_();
  const jsonPayload = JSON.stringify(cloneConfig_(config), null, 2);
  const schemaPayload = JSON.stringify(buildDefaultSchemaRegistry_(), null, 2);
  const detailRows = [
    {
      key: 'json',
      value: jsonPayload,
      sourceRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.json].getA1Notation(),
      targetRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.json].getA1Notation(),
    },
    {
      key: 'schemaRegistryJson',
      value: schemaPayload,
      sourceRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.schemaRegistryJson].getA1Notation(),
      targetRange: refs.cellsByKey[CONFIG_SHEET_SPEC.keys.schemaRegistryJson].getA1Notation(),
    },
  ];

  refs.cellsByKey[CONFIG_SHEET_SPEC.keys.json].setValue(jsonPayload);
  refs.cellsByKey[CONFIG_SHEET_SPEC.keys.schemaRegistryJson].setValue(schemaPayload);
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
    ['SchemaRegistryJson', ''],
    ['Validity', ''],
  ];

  const current = sheet.getRange(1, 1, layout.length, 1).getValues();
  const currentKeys = current.map((row) => toText_(row[0]));
  const legacyValuesByKey = getColumnBValuesByKey_(sheet);

  const keysNeedWrite = layout.some((row, i) => current[i][0] !== row[0]);
  if (keysNeedWrite) {
    sheet.getRange(1, 1, layout.length, 1).setValues(layout.map((row) => [row[0]]));
    restoreKnownConfigValuesAfterLayoutRewrite_(sheet, legacyValuesByKey);
  }

  const cellsByKey = getConfigCellsByKey_(sheet);
  const valuesByKey = getConfigValuesByKey_(sheet);
  if (toText_(valuesByKey[CONFIG_SHEET_SPEC.keys.json]).trim() === '') {
    const defaultJson = JSON.stringify(DEFAULT_CONFIG, null, 2);
    cellsByKey[CONFIG_SHEET_SPEC.keys.json].setValue(defaultJson);
  }
  ensureSchemaRegistryCellInitialized_(cellsByKey[CONFIG_SHEET_SPEC.keys.schemaRegistryJson]);
  cellsByKey[CONFIG_SHEET_SPEC.keys.validity].setValue('VALID');

  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), layout.length), 2).setFontFamily('Courier New');
  return { sheet, cellsByKey, valuesByKey: getConfigValuesByKey_(sheet) };
}

function getColumnBValuesByKey_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = sheet.getRange(1, 1, lastRow, 2).getValues();
  const byKey = {};
  values.forEach((row) => {
    const key = toText_(row[0]).trim();
    if (key) {
      byKey[key] = row[1];
    }
  });
  return byKey;
}

function restoreKnownConfigValuesAfterLayoutRewrite_(sheet, legacyValuesByKey) {
  const configJsonCell = readConfigSettingCell_(sheet, CONFIG_SHEET_SPEC.keys.json);
  const validityCell = readConfigSettingCell_(sheet, CONFIG_SHEET_SPEC.keys.validity);
  if (configJsonCell && legacyValuesByKey[CONFIG_SHEET_SPEC.keys.json] !== undefined) {
    configJsonCell.setValue(legacyValuesByKey[CONFIG_SHEET_SPEC.keys.json]);
  }
  if (validityCell && legacyValuesByKey[CONFIG_SHEET_SPEC.keys.validity] !== undefined) {
    validityCell.setValue(legacyValuesByKey[CONFIG_SHEET_SPEC.keys.validity]);
  }
}

function readConfigSettingCell_(sheet, settingName) {
  const values = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
  for (let i = 0; i < values.length; i += 1) {
    if (toText_(values[i][0]).trim() === settingName) {
      return sheet.getRange(i + 1, 2);
    }
  }
  return null;
}

function ensureSchemaRegistryCellInitialized_(schemaRegistryCell) {
  const defaultSchemaPayload = JSON.stringify(buildDefaultSchemaRegistry_(), null, 2);
  const currentValue = toText_(schemaRegistryCell.getValue()).trim();
  if (!currentValue) {
    schemaRegistryCell.setValue(defaultSchemaPayload);
    return;
  }
  try {
    const parsed = JSON.parse(currentValue);
    validateSchemaRegistry_(parsed);
  } catch (error) {
    schemaRegistryCell.setValue(defaultSchemaPayload);
  }
}

function resolveManagedConfigSheet_(ss) {
  const sheet = ss.getSheetByName(CONFIG_SHEET_SPEC.legacyName);
  if (!sheet) {
    const candidate = findSingleManagedConfigSheetCandidate_(ss);
    if (!candidate) {
      throw new Error('Sheet "Config" was not found. Expected: a worksheet named exactly "Config". Solution: create a sheet named "Config" (or rename the managed config sheet to "Config") and keep your settings in column B.');
    }
    candidate.setName(CONFIG_SHEET_SPEC.legacyName);
    return candidate;
  }
  // Accept the agreed worksheet name "Config" as the source of truth.
  // The layout is normalized by ensureConfigSheetAndRanges_().
  return sheet;
}

function findSingleManagedConfigSheetCandidate_(ss) {
  const candidates = ss
    .getSheets()
    .filter((sheet) => isManagedConfigSheetCandidate_(sheet) && sheet.getName() !== CONFIG_SHEET_SPEC.legacyName);
  if (candidates.length !== 1) {
    return null;
  }
  return candidates[0];
}

function isManagedConfigSheetCandidate_(sheet) {
  return hasManagedConfigLayout_(sheet) || hasLegacyManagedConfigLayout_(sheet);
}

function hasManagedConfigLayout_(sheet) {
  const compatibleLayouts = [
    [
      'Key',
      'ConfigJson',
      'LastValidConfigJson',
      'StatusCell',
      'ImportStartDate',
      'CalendarNames',
      'DefaultCalendarName',
      'Validity',
    ],
    [
      'Key',
      'ConfigJson',
      'ImportStartDate',
      'CalendarNames',
      'DefaultCalendarName',
      'Validity',
    ],
  ];
  return compatibleLayouts.some((layout) => {
    const values = sheet.getRange(1, 1, layout.length, 1).getValues();
    return layout.every((key, index) => toText_(values[index][0]) === key);
  });
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
      throw new Error(`Config sheet issue: missing key "${key}" in column A. Expected: column A contains the required keys (Key, ConfigJson, LastValidConfigJson, StatusCell, ImportStartDate, CalendarNames, DefaultCalendarName, Validity). Solution: restore the key labels in column A (rows 1-8).`);
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

function readConfigSettingValue_(sheet, settingName) {
  const rows = Math.max(sheet.getLastRow(), 1);
  const keyValues = sheet.getRange(1, 1, rows, 1).getValues();
  for (let i = 0; i < keyValues.length; i += 1) {
    const key = toText_(keyValues[i][0]).trim();
    if (!key) {
      break;
    }
    if (key === settingName) {
      return sheet.getRange(i + 1, 2).getValue();
    }
  }
  return 'n/a';
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


function rememberLegacyStateSheetNameCandidates_(rawConfig) {
  const candidates = collectLegacyStateSheetNameCandidatesFromConfig_(rawConfig || {});
  LEGACY_STATE_SHEET_NAME_CANDIDATES = mergeLegacyStateSheetNameCandidateSets_(
    LEGACY_STATE_SHEET_NAME_CANDIDATES,
    candidates
  );
  persistLegacyStateSheetNameCandidates_(LEGACY_STATE_SHEET_NAME_CANDIDATES);
}

function getLegacyStateSheetNameCandidates_() {
  return mergeLegacyStateSheetNameCandidateSets_(
    DEFAULT_LEGACY_STATE_SHEET_NAMES,
    readPersistedLegacyStateSheetNameCandidates_(),
    LEGACY_STATE_SHEET_NAME_CANDIDATES
  );
}

function collectLegacyStateSheetNameCandidatesFromConfig_(rawConfig) {
  const config = rawConfig && typeof rawConfig === 'object' ? rawConfig : {};
  const result = {};
  Object.keys(DEFAULT_LEGACY_STATE_SHEET_NAMES).forEach((group) => {
    result[group] = uniqueNonEmptyStrings_([
      config[LEGACY_STATE_SHEET_CONFIG_KEY_GROUPS[group]],
    ].concat(DEFAULT_LEGACY_STATE_SHEET_NAMES[group]));
  });
  return result;
}

function mergeLegacyStateSheetNameCandidateSets_() {
  const merged = {};
  Object.keys(DEFAULT_LEGACY_STATE_SHEET_NAMES).forEach((group) => {
    merged[group] = [];
  });

  Array.prototype.slice.call(arguments).forEach((candidateSet) => {
    if (!candidateSet || typeof candidateSet !== 'object') {
      return;
    }
    Object.keys(DEFAULT_LEGACY_STATE_SHEET_NAMES).forEach((group) => {
      merged[group] = uniqueNonEmptyStrings_(merged[group].concat(candidateSet[group] || []));
    });
  });

  return merged;
}

function uniqueNonEmptyStrings_(values) {
  const seen = new Set();
  const result = [];
  (values || []).forEach((value) => {
    const text = toText_(value).trim();
    if (!text || seen.has(text)) {
      return;
    }
    seen.add(text);
    result.push(text);
  });
  return result;
}

function cloneLegacyStateSheetNameCandidates_(value) {
  return mergeLegacyStateSheetNameCandidateSets_(value);
}

function persistLegacyStateSheetNameCandidates_(candidates) {
  if (typeof PropertiesService === 'undefined') {
    return;
  }
  try {
    getConfigPropertiesStore_().setProperty(
      LEGACY_STATE_SHEET_NAME_PROPERTY_KEY,
      JSON.stringify(cloneLegacyStateSheetNameCandidates_(candidates))
    );
  } catch (error) {
    if (!isPermissionDeniedError_(error)) {
      throw error;
    }
    logStorageDebug_('legacy-state-sheet-names', `Ignored denied write while saving legacy sheet names: ${error}`);
  }
}

function readPersistedLegacyStateSheetNameCandidates_() {
  if (typeof PropertiesService === 'undefined') {
    return null;
  }
  try {
    const raw = getConfigPropertiesStore_().getProperty(LEGACY_STATE_SHEET_NAME_PROPERTY_KEY);
    if (!raw) {
      return null;
    }
    return mergeLegacyStateSheetNameCandidateSets_(JSON.parse(raw));
  } catch (error) {
    if (isPermissionDeniedError_(error)) {
      logStorageDebug_('legacy-state-sheet-names', `Ignored denied read while loading legacy sheet names: ${error}`);
      return null;
    }
    logStorageDebug_('legacy-state-sheet-names', `Ignored invalid legacy sheet name cache: ${error}`);
    return null;
  }
}

function mergeConfigWithDefaults_(overrideConfig) {
  return mergeKnownShape_(DEFAULT_CONFIG, normalizeConfigOverrideForCurrentSchema_(overrideConfig || {}));
}

function normalizeConfigOverrideForCurrentSchema_(overrideConfig) {
  const normalized = cloneConfig_(overrideConfig || {});

  normalizeManagedHeaderConfigKeys_(normalized);
  normalizeStatusCellForCurrentManagedLayout_(normalized);
  normalizeLegacyMenuConfigForCurrentProduct_(normalized);

  OBSOLETE_STATE_CONFIG_KEYS.forEach((key) => {
    delete normalized[key];
  });

  if (normalized.nonBillableTableName === 'Non-Billable') {
    normalized.nonBillableTableName = DEFAULT_CONFIG.nonBillableTableName;
  }

  return normalized;
}

function normalizeLegacyMenuConfigForCurrentProduct_(normalizedConfig) {
  if (
    normalizedConfig.menu &&
    typeof normalizedConfig.menu === 'object' &&
    toText_(normalizedConfig.menu.title).trim() === 'Calendar Import'
  ) {
    normalizedConfig.menu.title = DEFAULT_CONFIG.menu.title;
  }

  if (toText_(normalizedConfig.toastTitle).trim() === 'Calendar import') {
    normalizedConfig.toastTitle = DEFAULT_CONFIG.toastTitle;
  }
}

function normalizeStatusCellForCurrentManagedLayout_(normalizedConfig) {
  if (!Object.prototype.hasOwnProperty.call(normalizedConfig, 'statusCell')) {
    return;
  }

  const normalizedStatusCell = toText_(normalizedConfig.statusCell).trim().toUpperCase();
  const defaultStatusCell = buildDefaultStatusCell_(DEFAULT_HEADER);
  const knownGeneratedStatusCells = [
    buildDefaultStatusCell_(LEGACY_CALENDAR_HEADER),
    buildDefaultStatusCell_(LEGACY_INVOICING_HEADER),
    buildDefaultStatusCell_(DEFAULT_INVOICING_HEADER),
  ];

  if (
    !/^[A-Z]+[1-9][0-9]*$/.test(normalizedStatusCell)
    || knownGeneratedStatusCells.includes(normalizedStatusCell)
  ) {
    normalizedConfig.statusCell = defaultStatusCell;
  } else {
    normalizedConfig.statusCell = normalizedStatusCell;
  }
}

function normalizeManagedHeaderConfigKeys_(normalizedConfig) {
  // Header arrays are managed structural constants, not user-editable business
  // configuration. Older generated ConfigJson payloads stored invoice-oriented
  // Calendar headers here, and strict validation must migrate those stale values
  // to the current inline-ID/status contract instead of blocking recovery/import.
  [
    'header',
    'invoicingHeader',
    'nonBillableHeader',
  ].forEach((key) => {
    if (Object.prototype.hasOwnProperty.call(normalizedConfig, key)) {
      delete normalizedConfig[key];
    }
  });
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


function buildDefaultSchemaRegistry_() {
  return {
    allowedKeys: Object.keys(DEFAULT_CONFIG).sort(),
  };
}

function validateSchemaRegistry_(schemaRegistry) {
  if (!schemaRegistry || typeof schemaRegistry !== 'object') {
    throw new Error('schema registry must be an object.');
  }
  if (!Array.isArray(schemaRegistry.allowedKeys) || schemaRegistry.allowedKeys.length === 0) {
    throw new Error('schema registry must contain non-empty allowedKeys array.');
  }
  schemaRegistry.allowedKeys.forEach((key, index) => {
    if (typeof key !== 'string' || !key.trim()) {
      throw new Error(`schema registry allowedKeys[${index}] must be a non-empty string.`);
    }
  });
}

function validateConfigStrictWithSchema_(normalizedConfig, rawConfig, schemaRegistry) {
  const raw = rawConfig || {};
  const allowedKeys = new Set(
    (schemaRegistry.allowedKeys || [])
      .concat(Object.keys(DEFAULT_CONFIG))
      .concat(OBSOLETE_STATE_CONFIG_KEYS)
  );
  Object.keys(raw).forEach((key) => {
    if (!allowedKeys.has(key)) {
      throw new Error(`Invalid config: unknown key "${key}" is not allowed.`);
    }
  });
  validateConfig_(normalizedConfig);
}

function validateConfig_(config) {
  assertString_(config.sheetName, 'sheetName');
  assertString_(config.tableName, 'tableName');
  assertValidTableName_(config.tableName, 'tableName');
  assertString_(config.invoicingSheetName, 'invoicingSheetName');
  assertString_(config.invoicingTableName, 'invoicingTableName');
  assertValidTableName_(config.invoicingTableName, 'invoicingTableName');
  assertString_(config.nonBillableSheetName, 'nonBillableSheetName');
  assertString_(config.nonBillableTableName, 'nonBillableTableName');
  assertValidTableName_(config.nonBillableTableName, 'nonBillableTableName');
  assertString_(config.statusCell, 'StatusCell');
  assertA1CellReference_(config.statusCell, 'StatusCell');
  assertString_(config.importStartDate, 'importStartDate');

  assertStrictIsoDate_(config.importStartDate, 'importStartDate');

  assertStringArray_(config.calendarNames, 'calendarNames');
  assertString_(config.defaultCalendarName, 'defaultCalendarName');
  if (!config.calendarNames.includes(config.defaultCalendarName)) {
    throw new Error('Invalid config: defaultCalendarName is not part of calendarNames. Expected: defaultCalendarName exactly matches one of the comma-separated CalendarNames entries. Solution: either add the default calendar to CalendarNames or change DefaultCalendarName.');
  }

  assertStringArrayWithLength_(
    config.header,
    'header',
    DEFAULT_CONFIG.header.length
  );
  assertExactArrayMatch_(config.header, DEFAULT_CONFIG.header, 'header');
  assertStringArrayWithLength_(
    config.invoicingHeader,
    'invoicingHeader',
    DEFAULT_CONFIG.invoicingHeader.length
  );
  assertExactArrayMatch_(config.invoicingHeader, DEFAULT_CONFIG.invoicingHeader, 'invoicingHeader');
  assertStringArrayWithLength_(
    config.nonBillableHeader,
    'nonBillableHeader',
    DEFAULT_CONFIG.nonBillableHeader.length
  );
  assertExactArrayMatch_(
    config.nonBillableHeader,
    DEFAULT_CONFIG.nonBillableHeader,
    'nonBillableHeader'
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
  assertString_(config.colors.nonBillable, 'colors.nonBillable');

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

function assertValidTableName_(value, fieldName) {
  const normalized = String(value || '').trim();
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(normalized)) {
    throw new Error(
      `${fieldName} must be a Google Sheets table-compatible name using only letters, numbers, and underscores, and it must not start with a number.`
    );
  }
  if (/^[A-Za-z]+[1-9][0-9]*$/.test(normalized)) {
    throw new Error(`${fieldName} must not look like a cell reference.`);
  }
}

function assertA1CellReference_(value, fieldName) {
  const normalized = String(value || '').trim();
  if (!/^[A-Za-z]+[1-9][0-9]*$/.test(normalized)) {
    throw new Error(`${fieldName} must be a valid single-cell A1 reference like "L1".`);
  }
}

function sanitizeStatusCell_(value, header) {
  const normalized = String(value || '').trim();
  if (/^[A-Za-z]+[1-9][0-9]*$/.test(normalized)) {
    return normalized;
  }
  return buildDefaultStatusCell_(header);
}

function buildDefaultStatusCell_(header) {
  const headerWidth = Array.isArray(header) ? header.length : 0;
  const oneBasedColumn = Math.max(1, headerWidth + 2);
  return `${toA1ColumnLabel_(oneBasedColumn)}1`;
}

function toA1ColumnLabel_(oneBasedColumn) {
  let n = Number(oneBasedColumn);
  let label = '';
  while (n > 0) {
    const remainder = (n - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label || 'A';
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
  appendStorageDebugToSheet_(phase, message);
}


function appendStorageDebugToSheet_(phase, message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureLogSheet_(ss);
    const startRow = 2;
    const maxLines = 500;
    const timestamp = new Date();
    let nextRow = Math.max(sheet.getLastRow() + 1, startRow);
    if (nextRow >= startRow + maxLines) {
      sheet.getRange(startRow + 1, 1, maxLines - 1, 5).moveTo(sheet.getRange(startRow, 1, maxLines - 1, 5));
      nextRow = startRow + maxLines - 1;
    }
    sheet.getRange(nextRow, 1, 1, 5).setValues([[timestamp, 'DEBUG', 'storage', String(phase), String(message)]]);
  } catch (error) {
    const fallback = `[storage-debug] failed to persist debug line: ${error}`;
    console.log(fallback);
    Logger.log(fallback);
  }
}

function writeValidityMessage_(message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss && ss.getSheetByName(CONFIG_SHEET_SPEC.legacyName);
    if (!sheet) {
      return;
    }
    const validityCell = readConfigSettingCell_(sheet, CONFIG_SHEET_SPEC.keys.validity);
    if (validityCell) {
      validityCell.setValue(String(message || 'Invalid configuration.'));
    }
  } catch (error) {
    Logger.log(`failed to write Validity message: ${error}`);
  }
}

function ensureLogSheet_(ss) {
  let sheet = ss.getSheetByName('Log');
  if (!sheet) {
    sheet = ss.insertSheet('Log');
  }
  const headers = [['Timestamp', 'Level', 'Component', 'Event', 'Message']];
  const range = sheet.getRange(1, 1, 1, 5);
  range.setValues(headers);
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 1), 5).setFontFamily('Courier New');
  return sheet;
}

function resetConfigAndLogSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet available.');
  }

  const configSheet = resolveManagedConfigSheet_(ss);
  const recoveredConfig = findBestConfigJsonInSheet_(configSheet) || cloneConfig_(DEFAULT_CONFIG);
  const normalizedConfig = mergeConfigWithDefaults_(recoveredConfig);
  validateConfig_(normalizedConfig);

  const layout = [
    [CONFIG_SHEET_SPEC.keyHeader, CONFIG_SHEET_SPEC.valueHeader],
    [CONFIG_SHEET_SPEC.keys.json, ''],
    [CONFIG_SHEET_SPEC.keys.schemaRegistryJson, ''],
    [CONFIG_SHEET_SPEC.keys.validity, ''],
  ];
  configSheet.getRange(1, 1, layout.length, 2).setValues(layout);
  configSheet.getRange(2, 2).setValue(JSON.stringify(normalizedConfig, null, 2));
  configSheet.getRange(3, 2).setValue(JSON.stringify(buildDefaultSchemaRegistry_(), null, 2));
  configSheet.getRange(4, 2).setValue('VALID');
  configSheet.getRange(1, 1, Math.max(configSheet.getMaxRows(), layout.length), 2).setFontFamily('Courier New');

  const logSheet = ensureLogSheet_(ss);
  const lastRow = logSheet.getLastRow();
  if (lastRow > 1) {
    logSheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  }
  logSheet.getRange(1, 1, Math.max(logSheet.getMaxRows(), 1), 5).setFontFamily('Courier New');
}

function findBestConfigJsonInSheet_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastColumn = Math.max(sheet.getLastColumn(), 2);
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  let fallback = null;

  for (let row = 0; row < values.length; row += 1) {
    for (let col = 0; col < values[row].length; col += 1) {
      const text = toText_(values[row][col]).trim();
      if (!text || !/^\s*[\{\[]/.test(text)) {
        continue;
      }
      try {
        const parsed = JSON.parse(text);
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
          const merged = mergeConfigWithDefaults_(parsed);
          validateConfig_(merged);
          return parsed;
        }
      } catch (error) {
        fallback = fallback || null;
      }
    }
  }
  return fallback;
}
