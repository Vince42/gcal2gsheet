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

let CONFIG = freezeConfigCopy_(DEFAULT_CONFIG);

function refreshConfig_() {
  const props = getConfigPropertiesStore_();
  const raw = props.getProperty(DEFAULT_CONFIG.configPropertyKey);

  if (!raw) {
    CONFIG = freezeConfigCopy_(DEFAULT_CONFIG);
    return CONFIG;
  }

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (error) {
    throw new Error(`Invalid saved configuration JSON: ${error.message}`);
  }

  const merged = mergeConfigWithDefaults_(parsed || {});
  validateConfig_(merged);
  CONFIG = freezeConfigCopy_(merged);
  return CONFIG;
}

function getConfigForDialog_() {
  const current = refreshConfig_();

  return {
    config: cloneConfig_(current),
    defaults: cloneConfig_(DEFAULT_CONFIG),
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

  const props = getConfigPropertiesStore_();
  props.setProperty(DEFAULT_CONFIG.configPropertyKey, JSON.stringify(basicConfig));
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
  const props = getConfigPropertiesStore_();
  props.deleteProperty(DEFAULT_CONFIG.configPropertyKey);
  CONFIG = freezeConfigCopy_(DEFAULT_CONFIG);
  return { success: true };
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
