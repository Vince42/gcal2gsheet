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
  const lines = [
    `${items.length} registered event(s) changed since the last update.`,
    'Existing register rows were preserved; changed follow-up rows were added for review.',
    '',
  ];

  items.slice(0, 25).forEach((item) => {
    lines.push(`${item.calendar} | ${item.date} ${item.start}-${item.end} | ${item.title}`);
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
  const text = String(message || '');
  showToastMessage_(ss, text, { severity: 'info' });
}

function showToastMessage_(ss, message, options) {
  if (!ss) {
    return;
  }

  const text = String(message || '').trim();
  if (!text) {
    return;
  }

  const severity = normalizeToastSeverity_(options && options.severity);
  const title = buildToastTitleBySeverity_(CONFIG.toastTitle, severity);
  const chunks = splitToastMessage_(text, 90);
  const seconds = estimateToastDurationSeconds_(text, severity);

  chunks.forEach((chunk, index) => {
    const partLabel = chunks.length > 1 ? `[${index + 1}/${chunks.length}] ` : '';
    ss.toast(`${partLabel}${chunk}\n\n`, title, seconds);
  });

  if (severity === 'error' || severity === 'warning') {
    writeStatusCellMessage_(ss, text, text);
    return;
  }

  writeStatusCellMessage_(ss, text);
}

function normalizeToastSeverity_(severity) {
  const value = String(severity || 'info').toLowerCase();
  if (value === 'error' || value === 'warning') {
    return value;
  }
  return 'info';
}

function buildToastTitleBySeverity_(baseTitle, severity) {
  const title = String(baseTitle || 'Notification');
  if (severity === 'error') {
    return `🔴 ${title}`;
  }
  if (severity === 'warning') {
    return `🟠 ${title}`;
  }
  return title;
}

function splitToastMessage_(message, maxChunkLength) {
  const maxLen = Math.max(80, Number(maxChunkLength) || 220);
  const text = String(message || '').replace(/\s+/g, ' ').trim();
  if (text.length <= maxLen) {
    return [text];
  }

  const sentences = text
    .split(/(?<=[.!?])\s+/)
    .map((part) => part.trim())
    .filter((part) => part);
  const sourceParts = sentences.length > 1 ? sentences : [text];

  const chunks = [];
  sourceParts.forEach((part) => {
    const words = part.split(' ');
    let current = '';
    words.forEach((word) => {
      if (!current) {
        current = word;
        return;
      }
      if ((current + ' ' + word).length > maxLen) {
        chunks.push(current);
        current = word;
        return;
      }
      current += ' ' + word;
    });
    if (current) {
      chunks.push(current);
    }
  });

  return chunks;
}

function estimateToastDurationSeconds_(message, severity) {
  const text = String(message || '');
  const base = severity === 'error' ? 30 : severity === 'warning' ? 24 : 15;
  const extraByLength = Math.ceil(text.length / 140) * 4;
  return Math.min(120, Math.max(base, base + extraByLength));
}

function writeStatusCellMessage_(ss, message, comment) {
  if (!ss || !CONFIG || !CONFIG.sheetName || !CONFIG.statusCell) {
    return;
  }
  if (!isA1CellReference_(CONFIG.statusCell)) {
    return;
  }
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    return;
  }
  const cell = sheet.getRange(CONFIG.statusCell);
  cell.setValue(String(message || ''));
  if (comment) {
    cell.setComment(String(comment));
    return;
  }
  cell.setComment('');
}

function isA1CellReference_(value) {
  return /^[A-Za-z]+[1-9][0-9]*$/.test(String(value || '').trim());
}
