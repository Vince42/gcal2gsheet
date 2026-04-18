function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.menu.title)
    .addItem(CONFIG.menu.item, 'updateCalendarSheets')
    .addToUi();
}

function updateCalendarSheets() {
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
      ss.toast(
        `${changedNotifications.length} changed invoiced event(s) detected.`,
        CONFIG.toastTitle,
        8
      );
    } else {
      setProgress_(ss, 'Done.');
      ss.toast('Calendar import finished.', CONFIG.toastTitle, 5);
    }
  } finally {
    restoreUiState_(ss, uiState);
    lock.releaseLock();
  }
}
