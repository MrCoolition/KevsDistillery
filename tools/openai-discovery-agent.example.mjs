import { readFileSync } from 'node:fs';

import { loadDotEnv } from './env.mjs';

loadDotEnv();

const apiKey = process.env.OPENAI_API_KEY;

if (!apiKey) {
  console.error('Set OPENAI_API_KEY before running this server-side example.');
  process.exit(1);
}

const sourceFile = process.argv[2];

if (!sourceFile) {
  console.error('Usage: node tools/openai-discovery-agent.example.mjs <extracted-source-text-file>');
  process.exit(1);
}

const sourceText = readFileSync(sourceFile, 'utf8');

const response = await fetch('https://api.openai.com/v1/responses', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    Authorization: `Bearer ${apiKey}`
  },
  body: JSON.stringify({
    model: 'gpt-5.5',
    reasoning: { effort: 'high' },
    instructions: [
      'You are THE DISTILLERY discovery analyst.',
      'Return a concise canonical discovery model delta as JSON.',
      'Every discovered item must include id, type, businessPurpose, owner, evidence, confidence, criticality, upstream, downstream, failureImpact, dollarExposure, and recommendedAction. Use dollarExposure only as an optional impact estimate when source evidence supports it.',
      'If evidence, confidence, or next action is missing, mark the finding unfinished and name the smallest source needed to finish it.'
    ].join('\n'),
    input: sourceText
  })
});

if (!response.ok) {
  console.error(await response.text());
  process.exit(1);
}

const result = await response.json();
console.log(JSON.stringify(result, null, 2));
