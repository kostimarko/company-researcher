function buildPrompt_(companyNames, fields, useWebSearch) {
  const fieldList = fields.map(f => {
    let line = `- ${f.name}`;
    if (f.instructions) line += `: ${f.instructions}`;
    if (f.format) line += ` (format: ${f.format})`;
    return line;
  }).join('\n');

  const companyList = companyNames.map(name => `- "${name}"`).join('\n');

  const instruction = useWebSearch
    ? 'Use web search to find accurate, current information for each company.'
    : 'Use your knowledge to find information for each company. If you are uncertain — for example, a local or obscure business — reflect that with a low confidence score and a brief source note. Use low scores (below 50) when you have little or no reliable information.';

  return `Research the following data points for each company listed below. ${instruction}

Companies:
${companyList}

Fields to research for each company:
${fieldList}

For each field, return your answer and a confidence score (0–100) reflecting how certain you are.

Return ONLY a valid JSON object where each key is EXACTLY the company name as provided above — no markdown, no code fences, no extra text:
{
  "Company Name": {
    "Field Name": {
      "value": "the answer",
      "confidence": 95,
      "source": "brief note on source"
    }
  }
}`;
}

function callClaudeResearch_(companyNames, fields, useWebSearch) {
  const apiKey = getApiKey_();
  if (!apiKey) throw new Error('No API key provided.');

  const payload = {
    model: ANTHROPIC_MODEL,
    max_tokens: 8192,
    messages: [{ role: 'user', content: buildPrompt_(companyNames, fields, useWebSearch) }]
  };

  if (useWebSearch) {
    payload.tools = [{ type: 'web_search_20250305', name: 'web_search' }];
  }

  const requestHeaders = {
    'x-api-key': apiKey,
    'anthropic-version': ANTHROPIC_VERSION,
    'content-type': 'application/json'
  };

  if (useWebSearch) {
    requestHeaders['anthropic-beta'] = 'web-search-2025-03-05';
  }

  const options = {
    method: 'post',
    headers: requestHeaders,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = callWithRetry_(options, 3);
  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code !== 200) throw new Error(`Claude API ${code}: ${body.substring(0, 300)}`);

  const data = JSON.parse(body);
  const textContent = data.content
    .filter(block => block.type === 'text')
    .map(block => block.text)
    .join('');

  const jsonMatch = textContent.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error(`No JSON found in Claude response: ${textContent.substring(0, 200)}`);

  return JSON.parse(jsonMatch[0]);
}

function callWithRetry_(options, maxRetries) {
  const retryableCodes = new Set([429, 500, 502, 503, 504, 529]);
  let delay = 1000;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    const response = UrlFetchApp.fetch(ANTHROPIC_API_URL, options);
    const code = response.getResponseCode();
    if (!retryableCodes.has(code)) return response;
    if (attempt === maxRetries) return response;
    Logger.log(`Status ${code} — retrying in ${delay}ms (attempt ${attempt + 1}/${maxRetries})`);
    Utilities.sleep(delay);
    delay *= 2;
  }
}
