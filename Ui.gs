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

function setProgress_(ss, message) {
  if (!ss) {
    return;
  }
  ss.toast(String(message || ''), CONFIG.toastTitle, 3);
}
