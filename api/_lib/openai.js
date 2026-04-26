const model = process.env.OPENAI_MODEL || 'gpt-5.5';
const OPENAI_TIMEOUT_MS = 59000;
const OPENAI_START_TIMEOUT_MS = 25000;
const PENDING_RESPONSE_STATUSES = new Set(['queued', 'in_progress']);

function buildInstructions(sourceKind, sourceName) {
  return [
    'You are Uncle Kev\'s Distillery principal data discovery analyst.',
    `Source kind: ${sourceKind}. Batch name: ${sourceName}.`,
    'Return one compact valid JSON object only. No markdown.',
    'Analyze retrieved evidence, not just filenames. Mark weak evidence as blocked with the smallest source needed to finish.',
    'Required top-level keys: processName, businessFunction, recommendation, decisionRequired, systemsInScope, criticalOutputs, overallRiskRating, estimatedDollarExposure, executiveBrief, reportSections, items, relationships, artifacts, backlog, evidenceIndex, lineageNodes, lineageEdges, failureRisks, openQuestions.',
    'reportSections: exactly 12 concise sections. Each section MUST use the title property with these exact titles in order: Executive Snapshot; Scope, Coverage, and Confidence; Business Mission of the Process; Current-State Operating Model; System and Artifact Landscape; Data Flow and Process Flow Summary; Transformations and Business Logic; Recursive Lineage and Source-of-Truth Assessment; Controls, Exceptions, and Failure Modes; Financial Impact and Business Exposure; Recommendations and Action Plan; Open Questions and Decisions Needed. Never use generic names like Analysis Section. Body under 75 words. Include confidence and evidenceIds.',
    'items: up to 12 high-signal canonical nodes. Each item requires id, type, name, businessPurpose, owner, evidence, confidence, criticality, upstream, downstream, failureImpact, dollarExposure, recommendedAction, status.',
    'recommendedAction requires mode, summary, owner, priority, acceptanceCriteria. dollarExposure requires low, base, high, assumptions.',
    'relationships up to 20. backlog up to 12. evidenceIndex up to 12. openQuestions up to 8. Use evidence objects that tie findings to retrieved content or explicit blockers.',
    'estimatedDollarExposure must cover revenue at risk, gross margin at risk, cash timing impact, rework labor cost, and compliance/SLA/customer exposure. If pricing evidence is missing, use zero low/base/high and state the exact business inputs needed.',
    'artifacts must list the 9 Discovery_Action_Pack outputs with short purpose text and status final: 01_Executive_Decision_Brief.pdf, 02_Current_State_Architecture_Report.pdf, 03_Technical_Discovery_Workbook.xlsx, 04_Auto_Documentation_Pack, 05_Diagram_Pack, 06_Financial_Impact_Model.xlsx, 07_Action_Backlog.csv, 08_Evidence_Archive, 09_Metadata_Manifest.json.',
    'Every critical output must have current-state narrative, diagram coverage, recursive lineage, business logic extraction, financial exposure, and a clear action recommendation or blocker.'
  ].join('\n');
}

function extractOutputText(result) {
  if (typeof result.output_text === 'string') {
    return result.output_text;
  }

  const chunks = [];
  for (const item of result.output || []) {
    for (const content of item.content || []) {
      if (content.type === 'output_text' && typeof content.text === 'string') {
        chunks.push(content.text);
      }
    }
  }

  return chunks.join('\n').trim();
}

function parseJsonOutput(outputText) {
  if (!outputText) {
    return null;
  }

  try {
    return JSON.parse(outputText);
  } catch {
    const fenced = outputText.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (fenced) {
      try {
        return JSON.parse(fenced[1]);
      } catch {
        return null;
      }
    }

    const firstBrace = outputText.indexOf('{');
    const lastBrace = outputText.lastIndexOf('}');
    if (firstBrace >= 0 && lastBrace > firstBrace) {
      try {
        return JSON.parse(outputText.slice(firstBrace, lastBrace + 1));
      } catch {
        return null;
      }
    }
  }

  return null;
}

async function synthesizeWithOpenAI(payload) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    const error = new Error('OPENAI_API_KEY is not configured.');
    error.statusCode = 500;
    throw error;
  }

  const {
    sourceKind,
    sourceName,
    extractedText,
    knownArtifacts = [],
    targetOutputs = []
  } = payload;

  const abortController = new AbortController();
  const timeout = setTimeout(() => abortController.abort(), OPENAI_TIMEOUT_MS);

  let response;
  try {
    response = await fetch('https://api.openai.com/v1/responses', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`
      },
      signal: abortController.signal,
      body: JSON.stringify({
        model,
        reasoning: { effort: 'low' },
        instructions: buildInstructions(sourceKind, sourceName),
        max_output_tokens: 9000,
        text: {
          format: {
            type: 'json_object'
          }
        },
        input: [
          {
            role: 'user',
            content: [
              {
                type: 'input_text',
                text: JSON.stringify({
                  sourceKind,
                  sourceName,
                  knownArtifacts,
                  targetOutputs,
                  extractedText
                })
              }
            ]
          }
        ]
      })
    });
  } catch (error) {
    if (error?.name === 'AbortError') {
      const timeoutError = new Error('OpenAI synthesis timed out before Vercel could finish the request.');
      timeoutError.statusCode = 504;
      throw timeoutError;
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }

  const responseText = await response.text();
  let result;
  try {
    result = responseText ? JSON.parse(responseText) : {};
  } catch {
    const error = new Error('OpenAI returned a non-JSON response.');
    error.statusCode = 502;
    error.detail = responseText.slice(0, 1000);
    throw error;
  }
  if (!response.ok) {
    const error = new Error('OpenAI request failed.');
    error.statusCode = response.status;
    error.detail = result;
    throw error;
  }

  const outputText = extractOutputText(result);
  const canonicalDelta = parseJsonOutput(outputText);
  if (!canonicalDelta || typeof canonicalDelta !== 'object') {
    const error = new Error('OpenAI did not return a canonical discovery JSON object.');
    error.statusCode = 502;
    error.detail = outputText.slice(0, 1000);
    throw error;
  }

  return {
    model,
    outputText,
    canonicalDelta,
    raw: result
  };
}

function buildRequestBody(payload, background = false) {
  const {
    sourceKind,
    sourceName,
    extractedText,
    knownArtifacts = [],
    targetOutputs = []
  } = payload;

  return {
    model,
    reasoning: { effort: 'low' },
    instructions: buildInstructions(sourceKind, sourceName),
    max_output_tokens: 9000,
    store: true,
    background,
    text: {
      format: {
        type: 'json_object'
      }
    },
    input: [
      {
        role: 'user',
        content: [
          {
            type: 'input_text',
            text: JSON.stringify({
              sourceKind,
              sourceName,
              knownArtifacts,
              targetOutputs,
              extractedText
            })
          }
        ]
      }
    ]
  };
}

async function openAIJson(path, options = {}) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    const error = new Error('OPENAI_API_KEY is not configured.');
    error.statusCode = 500;
    throw error;
  }

  const abortController = new AbortController();
  const timeout = setTimeout(() => abortController.abort(), options.timeoutMs || OPENAI_START_TIMEOUT_MS);

  let response;
  try {
    response = await fetch(`https://api.openai.com/v1${path}`, {
      method: options.method || 'GET',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`
      },
      signal: abortController.signal,
      body: options.body ? JSON.stringify(options.body) : undefined
    });
  } catch (error) {
    if (error?.name === 'AbortError') {
      const timeoutError = new Error('OpenAI did not acknowledge the background analysis before the request window closed.');
      timeoutError.statusCode = 504;
      throw timeoutError;
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }

  const responseText = await response.text();
  let result;
  try {
    result = responseText ? JSON.parse(responseText) : {};
  } catch {
    const error = new Error('OpenAI returned a non-JSON response.');
    error.statusCode = 502;
    error.detail = responseText.slice(0, 1000);
    throw error;
  }

  if (!response.ok) {
    const error = new Error('OpenAI request failed.');
    error.statusCode = response.status;
    error.detail = result;
    throw error;
  }

  return result;
}

function responseToSynthesis(result) {
  if (PENDING_RESPONSE_STATUSES.has(result.status)) {
    return {
      pending: true,
      responseId: result.id,
      responseStatus: result.status,
      model
    };
  }

  if (result.status && result.status !== 'completed') {
    const error = new Error(result.error?.message || `OpenAI response ended with status ${result.status}.`);
    error.statusCode = 502;
    error.detail = result.error || result.incomplete_details || result;
    throw error;
  }

  const outputText = extractOutputText(result);
  const canonicalDelta = parseJsonOutput(outputText);
  if (!canonicalDelta || typeof canonicalDelta !== 'object') {
    const error = new Error('OpenAI did not return a canonical discovery JSON object.');
    error.statusCode = 502;
    error.detail = outputText.slice(0, 1000);
    throw error;
  }

  return {
    pending: false,
    responseId: result.id,
    responseStatus: result.status || 'completed',
    model,
    outputText,
    canonicalDelta,
    raw: result
  };
}

async function startBackgroundSynthesis(payload) {
  const result = await openAIJson('/responses', {
    method: 'POST',
    timeoutMs: OPENAI_START_TIMEOUT_MS,
    body: buildRequestBody(payload, true)
  });

  return responseToSynthesis(result);
}

async function retrieveBackgroundSynthesis(responseId) {
  if (!responseId) {
    const error = new Error('responseId is required.');
    error.statusCode = 400;
    throw error;
  }

  const result = await openAIJson(`/responses/${encodeURIComponent(responseId)}`, {
    method: 'GET',
    timeoutMs: OPENAI_START_TIMEOUT_MS
  });

  return responseToSynthesis(result);
}

module.exports = {
  model,
  parseJsonOutput,
  retrieveBackgroundSynthesis,
  startBackgroundSynthesis,
  synthesizeWithOpenAI
};
