function onOpen() {
  let configError = null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  logStorageDebug_('onOpen.start', new Date().toISOString());
  try {
    refreshConfig_();
  } catch (error) {
    configError = error;
    logStorageDebug_('config.error', String(error));
  } finally {
    logStorageDebug_('onOpen.finish', new Date().toISOString());
  }

  const ui = SpreadsheetApp.getUi();
  ensureMenuVisible_(ui);

  if (configError) {
    const warningMessage = `Configuration issue detected: ${configError.message} Check the "Validity" row in the "Config" sheet for the exact validation result.`;
    showToastMessage_(ss, warningMessage, { severity: 'warning' });
    ui.alert('Configuration validation failed', warningMessage, ui.ButtonSet.OK);
  }
}

function onEdit(e) {
  if (!e || !e.range) {
    return;
  }
  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== 'Config') {
    return;
  }
  try {
    refreshConfig_();
  } catch (error) {
    SpreadsheetApp.getUi().alert('Configuration validation failed', String(error.message || error), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ensureMenuVisible_(ui) {
  if (!ui) {
    return;
  }

  const fallbackTitle = DEFAULT_CONFIG.menu.title;
  const fallbackItem = DEFAULT_CONFIG.menu.item;
  const fallbackFunction = 'updateCalendarSheets';

  try {
    const menuTitle =
      CONFIG && CONFIG.menu && typeof CONFIG.menu.title === 'string' && CONFIG.menu.title.trim()
        ? CONFIG.menu.title
        : fallbackTitle;
    const menuItem =
      CONFIG && CONFIG.menu && typeof CONFIG.menu.item === 'string' && CONFIG.menu.item.trim()
        ? CONFIG.menu.item
        : fallbackItem;

    ui.createMenu(menuTitle).addItem(menuItem, fallbackFunction).addToUi();
  } catch (error) {
    // Quality gate: never fail menu rendering because of config issues.
    ui.createMenu(fallbackTitle).addItem(fallbackItem, fallbackFunction).addToUi();
    logStorageDebug_('menu.fallback', String(error));
  }
}

function updateCalendarSheets() {
  refreshConfig_();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  const timeZone = ss.getSpreadsheetTimeZone();
  const uiState = captureUiState_(ss);
  const scope = buildScope_(timeZone);

  try {
    setProgress_(ss, 'Resolving calendars...');
    const calendars = resolveCalendars_();

    setProgress_(ss, 'Preparing sheets...');
    const sheet = ensureCalendarSheet_(ss);
    const stateSheet = ensureStateSheet_(ss);
    ensureHeader_(sheet);
    ensureStateHeader_(stateSheet);
    ensureSheetFormatting_(sheet);
    ensureTable_(spreadsheetId, sheet);

    setProgress_(ss, 'Reading existing sheet state...');
    const existingState = readExistingState_(sheet, stateSheet, timeZone, scope);
    const syncTokens = loadSyncTokens_(calendars);

    let useIncremental =
      existingState.hasManagedRows &&
      syncTokens.length === calendars.length &&
      syncTokens.every((item) => !!item.syncToken);

    let fetchResult;
    if (useIncremental) {
      try {
        setProgress_(ss, 'Fetching incremental changes...');
        fetchResult = fetchIncrementalChanges_(ss, calendars, timeZone);
      } catch (error) {
        clearSyncTokens_(calendars);
        useIncremental = false;
        setProgress_(ss, 'Sync token invalid. Falling back to full import...');
        fetchResult = fetchFullSnapshot_(ss, calendars, timeZone, scope);
      }
    } else {
      setProgress_(ss, 'Performing initial full import...');
      fetchResult = fetchFullSnapshot_(ss, calendars, timeZone, scope);
    }

    const changedNotifications = [];
    let finalRows;

    setProgress_(ss, 'Rebuilding worksheet data...');
    if (useIncremental) {
      finalRows = rebuildFromIncremental_(
        existingState,
        fetchResult.deltaByKey,
        changedNotifications,
        timeZone,
        scope
      );
    } else {
      finalRows = rebuildFromFullSnapshot_(
        existingState,
        fetchResult.currentByKey,
        changedNotifications,
        timeZone,
        scope
      );
    }

    setProgress_(ss, 'Removing duplicates...');
    finalRows = removeManagedDuplicates_(finalRows, scope);

    setProgress_(ss, `Writing ${finalRows.length} row(s)...`);
    writeVisibleBody_(sheet, finalRows);
    writeStateBody_(stateSheet, finalRows);
    applyNumberFormats_(sheet);
    applyRowColors_(sheet, finalRows);
    ensureTableRange_(spreadsheetId, sheet);

    saveSyncTokens_(fetchResult.nextSyncTokens);

    if (changedNotifications.length > 0) {
      setProgress_(ss, `Done. ${changedNotifications.length} changed invoiced event(s) detected.`);
      SpreadsheetApp.getUi().alert(
        'Changed invoiced events detected',
        buildChangedRowsMessage_(changedNotifications),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      showToastMessage_(ss, `${changedNotifications.length} changed invoiced event(s) detected.`, {
        severity: 'warning',
      });
    } else {
      setProgress_(ss, 'Done.');
      showToastMessage_(ss, 'Calendar import finished.', { severity: 'info' });
    }
  } finally {
    restoreUiState_(ss, uiState);
    lock.releaseLock();
  }
}
