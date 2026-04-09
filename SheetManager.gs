function getResearchFields_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FIELDS_TAB);
  if (!sheet) throw new Error(`Sheet "${FIELDS_TAB}" not found. Create a tab named "${FIELDS_TAB}".`);
  const data = sheet.getDataRange().getValues();
  return data
    .filter((row, i) => {
      if (!row[0] || !String(row[0]).trim()) return false;
      if (i === 0 && String(row[0]).trim().toLowerCase() === 'field name') return false;
      return true;
    })
    .map(row => ({
      name: String(row[0]).trim(),
      instructions: String(row[1] || '').trim(),
      format: String(row[2] || 'text').trim()
    }));
}

// Returns the header row array. Notes is always last.
function buildHeaders_(fields) {
  const headers = ['Company'];
  fields.forEach(f => {
    headers.push(f.name);
    headers.push(f.name + ' Confidence');
  });
  headers.push('Notes');
  return headers;
}

// Writes headers every run. Returns the 1-indexed column number of the Notes column.
function getOrWriteHeaders_(sheet, fields) {
  const headers = buildHeaders_(fields);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return headers.length;
}

// Returns rows where column A has a company name and column B is empty.
function getUnprocessedRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const numCols = Math.max(sheet.getLastColumn(), 2);
  const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  return data
    .map((row, i) => ({
      rowIndex: i + 2,
      companyName: String(row[0]).trim(),
      hasData: String(row[1] || '').trim() !== ''
    }))
    .filter(r => r.companyName && !r.hasData);
}

// Writes Claude's results for one company row.
// fields: [{name, instructions, format}]
// results: {"Field Name": {value, confidence, source}}
// notesColIndex: 1-indexed column number of the Notes column
function writeCompanyResults_(sheet, rowIndex, fields, results, notesColIndex) {
  fields.forEach((field, i) => {
    const valueCol = 2 + i * 2;      // B=2, D=4, F=6...
    const confidenceCol = 3 + i * 2; // C=3, E=5, G=7...
    const result = results[field.name] || {};
    sheet.getRange(rowIndex, valueCol).setValue(result.value !== undefined ? result.value : 'N/A');
    sheet.getRange(rowIndex, confidenceCol).setValue(result.confidence !== undefined ? result.confidence : '');
  });

  const lowConfidenceNotes = fields
    .filter(f => {
      const r = results[f.name];
      return r && typeof r.confidence === 'number' && r.confidence < CONFIDENCE_NOTES_THRESHOLD;
    })
    .map(f => `${f.name}: ${results[f.name].source || 'low confidence'}`)
    .join('; ');

  if (lowConfidenceNotes) {
    sheet.getRange(rowIndex, notesColIndex).setValue(lowConfidenceNotes);
  }
}

// Writes ERROR into all value columns and the error message into Notes.
function writeError_(sheet, rowIndex, fields, errorMessage, notesColIndex) {
  fields.forEach((_, i) => {
    sheet.getRange(rowIndex, 2 + i * 2).setValue('ERROR');
  });
  sheet.getRange(rowIndex, notesColIndex).setValue(errorMessage);
}

// Applies green/yellow/red formatting to all Confidence columns. Preserves unrelated rules.
function applyConditionalFormatting_(sheet, fields) {
  const confidenceCols = fields.map((_, i) => 3 + i * 2); // C=3, E=5, G=7...
  const colSet = new Set(confidenceCols);

  const keptRules = sheet.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().some(r => colSet.has(r.getColumn()))
  );

  const newRules = [];
  confidenceCols.forEach(col => {
    const range = sheet.getRange(2, col, 1000, 1);
    newRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(80)
        .setBackground('#b7e1cd')
        .setRanges([range])
        .build()
    );
    newRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(50, 79)
        .setBackground('#fce8b2')
        .setRanges([range])
        .build()
    );
    newRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(50)
        .setBackground('#f4c7c3')
        .setRanges([range])
        .build()
    );
  });

  sheet.setConditionalFormatRules([...keptRules, ...newRules]);
}

// Returns rows where any confidence column is below CONFIDENCE_NOTES_THRESHOLD.
function getLowConfidenceRows_(sheet, fields) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const numCols = Math.max(sheet.getLastColumn(), 2);
  const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  return data
    .map((row, i) => ({
      rowIndex: i + 2,
      companyName: String(row[0]).trim(),
      hasData: String(row[1] || '').trim() !== ''
    }))
    .filter(r => {
      if (!r.companyName || !r.hasData) return false;
      return fields.some((_, i) => {
        const confidence = data[r.rowIndex - 2][2 + i * 2]; // confidence col is 3+i*2, zero-indexed = 2+i*2
        return typeof confidence === 'number' && confidence < CONFIDENCE_NOTES_THRESHOLD;
      });
    });
}

// Clears columns B onward (preserves company names in column A) and removes formatting.
function clearResults_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 2) return;
  sheet.getRange(1, 2, lastRow, lastCol - 1).clearContent();
  sheet.clearConditionalFormatRules();
}
