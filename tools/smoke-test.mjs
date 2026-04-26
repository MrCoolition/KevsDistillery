import { readFileSync } from 'node:fs';

const requiredFiles = [
  'package.json',
  'angular.json',
  'src/main.ts',
  'src/app/app.ts',
  'src/app/app.html',
  'src/app/app.css',
  'src/app/ai-orchestration.ts',
  'src/app/discovery-model.ts',
  'src/app/sample-discovery.ts',
  'api/health.js',
  'api/admin/migrate.js',
  'api/discovery/status.js',
  'api/discovery/synthesize.js',
  'api/_lib/db.js',
  'database/schema.sql',
  'vercel.json',
  'preview.html',
  'public/vendor/jszip.min.js',
  'tools/serve-distillery.mjs'
];

const canonicalFields = [
  'id',
  'type',
  'businessPurpose',
  'owner',
  'evidence',
  'confidence',
  'criticality',
  'upstream',
  'downstream',
  'failureImpact',
  'dollarExposure',
  'recommendedAction'
];

const source = requiredFiles.map((file) => readFileSync(file, 'utf8')).join('\n');
const missing = canonicalFields.filter((field) => !source.includes(field));

if (missing.length > 0) {
  console.error(`Missing canonical fields: ${missing.join(', ')}`);
  process.exit(1);
}

for (const file of requiredFiles) {
  readFileSync(file, 'utf8');
}

const packageJson = JSON.parse(readFileSync('package.json', 'utf8'));
const angularVersion = packageJson.dependencies['@angular/core'];

if (angularVersion !== '21.2.8') {
  console.error(`Expected Angular 21.2.8, found ${angularVersion}`);
  process.exit(1);
}

if (!source.includes("OPENAI_DISCOVERY_MODEL = 'gpt-5.5'")) {
  console.error('Expected OpenAI discovery model to be gpt-5.5');
  process.exit(1);
}

console.log("Uncle Kev's Distillery smoke test passed.");
