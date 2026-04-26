const model = process.env.OPENAI_MODEL || 'gpt-5.5';

function buildInstructions(sourceKind, sourceName) {
  return [
    'You are THE DISTILLERY discovery analyst.',
    `Source kind: ${sourceKind}. Source name: ${sourceName}.`,
    'Return valid JSON only, with no markdown fences.',
    'The JSON must be a canonical discovery model delta with items, relationships, evidence, business impact assumptions, artifact recommendations, and backlog actions.',
    'Every discovered item must include id, type, businessPurpose, owner, evidence, confidence, criticality, upstream, downstream, failureImpact, dollarExposure, and recommendedAction. Use dollarExposure only as an optional impact estimate when source evidence supports it.',
    'If a finding lacks evidence, confidence, or a next action, mark it unfinished and name the smallest source needed to finish it.',
    'Trace upstream recursively until a terminal condition is reached or a blocker is documented.',
    'Keep executive narrative concise, but preserve technical depth in structured fields.'
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

  const response = await fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model,
      reasoning: { effort: 'high' },
      instructions: buildInstructions(sourceKind, sourceName),
      max_output_tokens: 10000,
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

  const result = await response.json();
  if (!response.ok) {
    const error = new Error('OpenAI request failed.');
    error.statusCode = response.status;
    error.detail = result;
    throw error;
  }

  const outputText = extractOutputText(result);
  return {
    model,
    outputText,
    canonicalDelta: parseJsonOutput(outputText),
    raw: result
  };
}

module.exports = {
  model,
  parseJsonOutput,
  synthesizeWithOpenAI
};
