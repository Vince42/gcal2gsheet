function onOpen() {
  let configError = null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  logStorageDebug_('onOpen.start', new Date().toISOString());
  try {
    refreshConfig_();
  } catch (error) {
    configError = error;
    writeValidityMessage_(error && error.message ? error.message : String(error));
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

    ui.createMenu(menuTitle)
      .addItem(menuItem, fallbackFunction)
      .addSeparator()
      .addSubMenu(
        ui.createMenu('Filter for')
          .addItem('Open', 'filterCalendarForOpen')
          .addItem('Invoiced', 'filterCalendarForInvoiced')
          .addItem('Non-Billable', 'filterCalendarForNonBillable')
      )
      .addSubMenu(
        ui.createMenu('Mark as')
          .addItem('Invoiced', 'markSelectedCalendarRowsAsInvoiced')
          .addItem('Non-Billable', 'markSelectedCalendarRowsAsNonBillable')
      )
      .addToUi();
  } catch (error) {
    // Quality gate: never fail menu rendering because of config issues.
    ui.createMenu(fallbackTitle)
      .addItem(fallbackItem, fallbackFunction)
      .addSeparator()
      .addSubMenu(
        ui.createMenu('Filter for')
          .addItem('Open', 'filterCalendarForOpen')
          .addItem('Invoiced', 'filterCalendarForInvoiced')
          .addItem('Non-Billable', 'filterCalendarForNonBillable')
      )
      .addSubMenu(
        ui.createMenu('Mark as')
          .addItem('Invoiced', 'markSelectedCalendarRowsAsInvoiced')
          .addItem('Non-Billable', 'markSelectedCalendarRowsAsNonBillable')
      )
      .addToUi();
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

    setProgress_(ss, 'Checking managed sheets and columns...');
    const managedSheets = ensureManagedWorkbookStructure_(ss, spreadsheetId);
    const sheet = managedSheets.sheet;
    const invoicingSheet = managedSheets.invoicingSheet;
    const nonBillableSheet = managedSheets.nonBillableSheet;

    const migratedInvoiceCount = migrateCalendarInvoicesToInvoicing_(sheet, invoicingSheet);
    const repairedInvoiceStateCount = repairInvoicingStateFromCalendarRows_(sheet, invoicingSheet);

    clearRetiredCalendarInvoiceColumns_(sheet);

    setProgress_(ss, 'Performing full import for self-healing reconciliation...');
    const fetchResult = fetchFullSnapshot_(ss, calendars, timeZone, scope);
    const repairedImportedInvoiceStateCount = repairInvoicingStateFromImportedEvents_(
      fetchResult.currentByKey,
      invoicingSheet
    );
    const repairedNonBillableStateCount = repairNonBillableStateFromImportedEvents_(
      fetchResult.currentByKey,
      nonBillableSheet
    );
    const invoiceStore = readInvoicingState_(invoicingSheet);
    const nonBillableStore = readNonBillableState_(nonBillableSheet);
    applyRegisterStatusesToImportedEvents_(fetchResult.currentByKey, invoiceStore, nonBillableStore);

    setProgress_(ss, 'Reading existing sheet state...');
    const existingState = readExistingState_(sheet, timeZone, scope, invoiceStore, nonBillableStore);

    setProgress_(ss, 'Rebuilding worksheet data...');
    let finalRows = rebuildFromFullSnapshot_(
      existingState,
      fetchResult.currentByKey,
      scope
    );

    setProgress_(ss, 'Removing duplicates...');
    finalRows = removeManagedDuplicates_(finalRows, scope);

    setProgress_(ss, `Writing ${finalRows.length} row(s)...`);
    writeVisibleBody_(sheet, finalRows);
    applyNumberFormats_(sheet);
    applyNumberFormats_(invoicingSheet, CONFIG.invoicingHeader);
    applyNumberFormats_(nonBillableSheet, CONFIG.nonBillableHeader);
    applyRowColors_(sheet, finalRows);
    ensureTableRange_(spreadsheetId, sheet);
    ensureTableRange_(spreadsheetId, invoicingSheet, CONFIG.invoicingTableName, CONFIG.invoicingHeader);
    ensureTableRange_(spreadsheetId, nonBillableSheet, CONFIG.nonBillableTableName, CONFIG.nonBillableHeader);

    saveSyncTokens_(fetchResult.nextSyncTokens);

    const ignoredBeforeImportStartCount = existingState.ignoredBeforeImportStartCount || 0;
    if (ignoredBeforeImportStartCount > 0) {
      showToastMessage_(
        ss,
        `${ignoredBeforeImportStartCount} row(s) before ${CONFIG.importStartDate} were excluded from this update and left unchanged.`,
        { severity: 'info' }
      );
    }

    if (migratedInvoiceCount > 0) {
      showToastMessage_(
        ss,
        `${migratedInvoiceCount} invoiced row(s) were moved to the Invoicing register.`,
        { severity: 'info' }
      );
    }

    const totalRepairedInvoiceStateCount = repairedInvoiceStateCount + repairedImportedInvoiceStateCount;
    if (totalRepairedInvoiceStateCount > 0) {
      showToastMessage_(
        ss,
        `${totalRepairedInvoiceStateCount} Invoicing row(s) were linked to calendar events.`,
        { severity: 'info' }
      );
    }

    if (repairedNonBillableStateCount > 0) {
      showToastMessage_(
        ss,
        `${repairedNonBillableStateCount} Non-Billable row(s) were linked to calendar events.`,
        { severity: 'info' }
      );
    }

    setProgress_(ss, 'Done.');
    showToastMessage_(ss, 'Calendar import finished.', { severity: 'info' });
  } finally {
    restoreUiState_(ss, uiState);
    lock.releaseLock();
  }
}
