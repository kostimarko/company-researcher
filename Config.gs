// ─── Sheet names ───────────────────────────────────────────────────────────────
const COMPANIES_TAB = 'Companies';
const FIELDS_TAB = 'Research Fields';

// ─── Claude API ────────────────────────────────────────────────────────────────
const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages';
const ANTHROPIC_MODEL = 'claude-sonnet-4-6';
const ANTHROPIC_VERSION = '2023-06-01';

// ─── Behavior ──────────────────────────────────────────────────────────────────
const API_CALL_DELAY_MS = 1000;
const BATCH_SIZE = 5;
const CONFIDENCE_NOTES_THRESHOLD = 70;
const CHECKPOINT_KEY = 'LAST_PROCESSED_ROW';

// ─── API key ───────────────────────────────────────────────────────────────────
function getApiKey_() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Anthropic API Key',
      'Enter your Anthropic API key (saved to Script Properties):',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return null;
    key = response.getResponseText().trim();
    if (!key) return null;
    props.setProperty('ANTHROPIC_API_KEY', key);
  }
  return key;
}
