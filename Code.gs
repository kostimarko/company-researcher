function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Company Research')
    .addItem('Run Research', 'runResearch')
    .addItem('Research Low Confidence with Web Search', 'runWebSearchForLowConfidence')
    .addSeparator()
    .addItem('Clear Results', 'clearResultsWithConfirm')
    .addToUi();
}

function runResearch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let fields, sheet;
  try {
    fields = getResearchFields_();
    sheet = ss.getSheetByName(COMPANIES_TAB);
    if (!sheet) throw new Error(`Sheet "${COMPANIES_TAB}" not found. Create a tab named "${COMPANIES_TAB}".`);
  } catch (e) {
    ui.alert('Setup Error', e.message, ui.ButtonSet.OK);
    return;
  }

  if (fields.length === 0) {
    ui.alert('No Fields Configured', `Add fields to the "${FIELDS_TAB}" tab first.`, ui.ButtonSet.OK);
    return;
  }

  const notesColIndex = getOrWriteHeaders_(sheet, fields);
  applyConditionalFormatting_(sheet, fields);

  const allUnprocessed = getUnprocessedRows_(sheet);
  if (allUnprocessed.length === 0) {
    ui.alert('Nothing to Process', 'All companies already have results. Use "Clear Results" to re-run.', ui.ButtonSet.OK);
    return;
  }

  // Resume from checkpoint if a previous run timed out
  const props = PropertiesService.getScriptProperties();
  const checkpointRow = parseInt(props.getProperty(CHECKPOINT_KEY) || '0', 10);
  const toProcess = checkpointRow > 0
    ? allUnprocessed.filter(r => r.rowIndex > checkpointRow)
    : allUnprocessed;

  const batches = chunkArray_(toProcess, BATCH_SIZE);
  let processed = 0;
  let errors = 0;

  batches.forEach((batch, batchIdx) => {
    const names = batch.map(r => r.companyName);
    ss.toast(
      `Batch ${batchIdx + 1} of ${batches.length} — ${names.length} companies`,
      'Company Research',
      5
    );

    let batchResults = null;
    try {
      batchResults = callClaudeResearch_(names, fields);
    } catch (e) {
      Logger.log(`Batch ${batchIdx + 1} error: ${e.message}`);
      batch.forEach(row => {
        writeError_(sheet, row.rowIndex, fields, e.message, notesColIndex);
        errors++;
      });
    }

    if (batchResults) {
      batch.forEach(row => {
        const result = batchResults[row.companyName];
        if (result) {
          writeCompanyResults_(sheet, row.rowIndex, fields, result, notesColIndex);
        } else {
          writeError_(sheet, row.rowIndex, fields, 'Not returned in batch response', notesColIndex);
          errors++;
        }
      });
    }

    props.setProperty(CHECKPOINT_KEY, String(batch[batch.length - 1].rowIndex));
    processed += batch.length;

    if (batchIdx < batches.length - 1) Utilities.sleep(API_CALL_DELAY_MS);
  });

  props.deleteProperty(CHECKPOINT_KEY);

  const summary = errors > 0
    ? `Processed ${processed} companies. ${errors} error(s) — see the Notes column.`
    : `Done! Processed ${processed} companies.`;
  ss.toast(summary, 'Company Research Complete', 10);
}

function chunkArray_(arr, size) {
  const chunks = [];
  for (let i = 0; i < arr.length; i += size) {
    chunks.push(arr.slice(i, i + size));
  }
  return chunks;
}

function runWebSearchForLowConfidence() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let fields, sheet;
  try {
    fields = getResearchFields_();
    sheet = ss.getSheetByName(COMPANIES_TAB);
    if (!sheet) throw new Error(`Sheet "${COMPANIES_TAB}" not found.`);
  } catch (e) {
    ui.alert('Setup Error', e.message, ui.ButtonSet.OK);
    return;
  }

  const notesColIndex = getOrWriteHeaders_(sheet, fields);
  const lowConfRows = getLowConfidenceRows_(sheet, fields);

  if (lowConfRows.length === 0) {
    ui.alert('Nothing to Search', 'No rows found with confidence below the threshold.', ui.ButtonSet.OK);
    return;
  }

  const confirm = ui.alert(
    'Web Search',
    `Found ${lowConfRows.length} companies with low confidence. Run web search for these?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  let updated = 0;
  let errors = 0;

  lowConfRows.forEach((row, idx) => {
    ss.toast(`Web searching: ${row.companyName} (${idx + 1} of ${lowConfRows.length})`, 'Company Research', 3);
    try {
      const wsResults = callClaudeResearch_([row.companyName], fields, true);
      const companyResult = wsResults[row.companyName];
      if (companyResult) {
        writeCompanyResults_(sheet, row.rowIndex, fields, companyResult, notesColIndex);
        updated++;
      }
    } catch (e) {
      Logger.log(`Web search failed for ${row.companyName}: ${e.message}`);
      errors++;
    }
    if (idx < lowConfRows.length - 1) Utilities.sleep(API_CALL_DELAY_MS);
  });

  const summary = errors > 0
    ? `Updated ${updated} companies. ${errors} error(s) — see the Notes column.`
    : `Done! Updated ${updated} companies with web search results.`;
  ss.toast(summary, 'Web Search Complete', 10);
}

function clearResultsWithConfirm() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Clear Results',
    'This will remove all research results. Company names in column A will be preserved. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (result !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COMPANIES_TAB);
  if (!sheet) {
    ui.alert('Error', `Sheet "${COMPANIES_TAB}" not found.`, ui.ButtonSet.OK);
    return;
  }

  clearResults_(sheet);
  PropertiesService.getScriptProperties().deleteProperty(CHECKPOINT_KEY);
  ui.alert('Done', 'Results cleared. Company names in column A preserved.', ui.ButtonSet.OK);
}
