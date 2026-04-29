function setProgress_(ss, message) {
  ss.toast(message, CONFIG.toastTitle, 5);

  const sheet = ss.getSheetByName(CONFIG.sheetName);
  if (sheet) {
    sheet.getRange(CONFIG.statusCell).setValue(message);
  }

  SpreadsheetApp.flush();
}

function showConfigDialog_() {
  const template = HtmlService.createTemplateFromFile('ConfigDialog');
  template.modelJson = JSON.stringify(getConfigForDialog_())
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/'/g, '\\u0027');
  const html = template
    .evaluate()
    .setWidth(860)
    .setHeight(700)
    .setTitle('Calendar Import Configuration');

  SpreadsheetApp.getUi().showModalDialog(html, 'Calendar Import Configuration');
}

function saveConfigDialog_(payload) {
  try {
    const result = saveConfigFromDialog_(payload);
    refreshConfig_();
    showConfigSaveDetails_(result);
    return result;
  } catch (error) {
    const message = error && error.message ? error.message : String(error);
    throw new Error(`saveConfigDialog_ failed after revision ${CONFIG_DIALOG_REVISION}: ${message}`);
  }
}

function showConfigSaveDetails_(result) {
  const details = result && Array.isArray(result.writeDetails) ? result.writeDetails : [];
  if (details.length === 0) {
    return;
  }

  const ui = SpreadsheetApp.getUi();
  details.forEach((item) => {
    const message = [
      `Item key: ${item.key}`,
      `Item value: ${item.value}`,
      `Came from range: ${item.sourceRange}`,
      `Written to range: ${item.targetRange}`,
    ].join('\n');
    logStorageDebug_(
      'config-save-item',
      `key=${item.key}; value=${item.value}; from=${item.sourceRange}; to=${item.targetRange}`
    );
    ui.alert('Configuration save detail', message, ui.ButtonSet.OK);
  });
}

function resetConfigDialog_() {
  try {
    const result = resetConfigToDefault_();
    refreshConfig_();
    return result;
  } catch (error) {
    const message = error && error.message ? error.message : String(error);
    throw new Error(`resetConfigDialog_ failed after revision ${CONFIG_DIALOG_REVISION}: ${message}`);
  }
}

function saveConfigDialog(payload) {
  return saveConfigDialog_(payload);
}

function resetConfigDialog() {
  return resetConfigDialog_();
}

function captureUiState_(ss) {
  const activeSheet = ss.getActiveSheet();
  const activeRange = ss.getActiveRange();

  return {
    sheetName: activeSheet ? activeSheet.getName() : null,
    a1Notation: activeRange ? activeRange.getA1Notation() : null,
  };
}

function restoreUiState_(ss, uiState) {
  if (!uiState || !uiState.sheetName) {
    return;
  }

  const sheet = ss.getSheetByName(uiState.sheetName);
  if (!sheet) {
    return;
  }

  ss.setActiveSheet(sheet, false);

  if (uiState.a1Notation) {
    try {
      sheet.setActiveRange(sheet.getRange(uiState.a1Notation));
    } catch (error) {
      // Ignore invalid range restoration.
    }
  }
}

function buildChangedRowsMessage_(items) {
  const lines = items.slice(0, 25).map((item) => {
    return `${item.calendar} | ${item.date} ${item.start}-${item.end} | ${item.title}`;
  });

  if (items.length > 25) {
    lines.push(`... and ${items.length - 25} more.`);
  }

  return lines.join('\n');
}
