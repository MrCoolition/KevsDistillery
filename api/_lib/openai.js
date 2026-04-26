const model = process.env.OPENAI_MODEL || 'gpt-5.5';
const OPENAI_TIMEOUT_MS = 59000;

function buildInstructions(sourceKind, sourceName) {
  return [
    'You are Uncle Kev\'s Distillery principal data discovery analyst.',
    `Source kind: ${sourceKind}. Batch name: ${sourceName}.`,
    'Return one compact valid JSON object only. No markdown.',
    'Analyze retrieved evidence, not just filenames. Mark weak evidence as blocked with the smallest source needed to finish.',
    'Required top-level keys: processName, businessFunction, recommendation, decisionRequired, systemsInScope, criticalOutputs, overallRiskRating, estimatedDollarExposure, executiveBrief, reportSections, items, relationships, artifacts, backlog, evidenceIndex, lineageNodes, lineageEdges, failureRisks, openQuestions.',
    'reportSections: exactly 6 concise sections named Executive Snapshot, Scope and Evidence, Current-State Operating Model, Lineage and Business Logic, Controls and Failure Modes, Action Plan. Body under 45 words. Include confidence and evidenceIds.',
    'items: max 6 high-signal nodes. Each item requires id, type, name, businessPurpose, owner, evidence, confidence, criticality, upstream, downstream, failureImpact, dollarExposure, recommendedAction, status.',
    'recommendedAction requires mode, summary, owner, priority, acceptanceCriteria. dollarExposure requires low, base, high, assumptions.',
    'relationships max 8. backlog max 5. evidenceIndex max 4. openQuestions max 3. Use one evidence object per item.',
    'artifacts must list the 9 Discovery_Action_Pack outputs with short purpose text. Keep all prose short.'
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
        max_output_tokens: 4800,
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

module.exports = {
  model,
  parseJsonOutput,
  synthesizeWithOpenAI
};
