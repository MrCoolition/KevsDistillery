import { createServer } from 'node:http';

import { loadDotEnv } from './env.mjs';

loadDotEnv();

const model = 'gpt-5.5';
const port = Number(process.env.DISTILLERY_API_PORT ?? 8787);
const apiKey = process.env.OPENAI_API_KEY;

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': 'http://127.0.0.1:4200',
    'Access-Control-Allow-Headers': 'content-type',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS'
  });
  response.end(JSON.stringify(payload));
}

async function readJson(request) {
  const chunks = [];
  for await (const chunk of request) {
    chunks.push(chunk);
  }

  return JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}');
}

function buildInstructions(sourceKind, sourceName) {
  return [
    'You are Uncle Kev\'s Distillery discovery analyst.',
    `Source kind: ${sourceKind}. Batch name: ${sourceName}.`,
    'Return valid JSON only.',
    'The JSON must describe a canonical discovery model delta with items, relationships, evidence, business impact assumptions, and backlog actions.',
    'Every discovered item must include id, type, businessPurpose, owner, evidence, confidence, criticality, upstream, downstream, failureImpact, dollarExposure, and recommendedAction. Use dollarExposure only as an optional impact estimate when source evidence supports it.',
    'If a finding lacks evidence, confidence, or a next action, mark it unfinished and name the smallest source needed to finish it.',
    'Trace upstream recursively until a terminal condition is reached or a blocker is documented.'
  ].join('\n');
}

async function synthesizeDiscovery(payload) {
  if (!apiKey) {
    return {
      statusCode: 500,
      payload: {
        error: 'OPENAI_API_KEY is not loaded. Add it to .env or the server environment.'
      }
    };
  }

  const { sourceKind, sourceName, extractedText, knownArtifacts = [] } = payload;
  if (!sourceKind || !sourceName || !extractedText) {
    return {
      statusCode: 400,
      payload: {
        error: 'sourceKind, sourceName, and extractedText are required.'
      }
    };
  }

  const openAiResponse = await fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model,
      reasoning: { effort: 'high' },
      instructions: buildInstructions(sourceKind, sourceName),
      input: [
        {
          role: 'user',
          content: [
            {
              type: 'input_text',
              text: JSON.stringify({
                knownArtifacts,
                extractedText
              })
            }
          ]
        }
      ]
    })
  });

  const result = await openAiResponse.json();
  if (!openAiResponse.ok) {
    return {
      statusCode: openAiResponse.status,
      payload: {
        error: 'OpenAI request failed.',
        detail: result
      }
    };
  }

  return {
    statusCode: 200,
    payload: {
      model,
      outputText: result.output_text ?? '',
      raw: result
    }
  };
}

const server = createServer(async (request, response) => {
  if (request.method === 'OPTIONS') {
    sendJson(response, 204, {});
    return;
  }

  if (request.method === 'GET' && request.url === '/health') {
    sendJson(response, 200, {
      ok: true,
      model,
      hasOpenAIKey: Boolean(apiKey)
    });
    return;
  }

  if (request.method === 'POST' && request.url === '/api/discovery/synthesize') {
    try {
      const result = await synthesizeDiscovery(await readJson(request));
      sendJson(response, result.statusCode, result.payload);
    } catch (error) {
      sendJson(response, 500, {
        error: error instanceof Error ? error.message : 'Unknown discovery synthesis error.'
      });
    }
    return;
  }

  sendJson(response, 404, {
    error: 'Not found.'
  });
});

server.listen(port, '127.0.0.1', () => {
  console.log(`Uncle Kev's Distillery discovery API listening on http://127.0.0.1:${port}`);
});
