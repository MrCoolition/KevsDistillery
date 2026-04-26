import { readFileSync } from 'node:fs';

import { loadDotEnv } from './env.mjs';

loadDotEnv();

const databaseUrl = process.env.DATABASE_URL || process.env.POSTGRES_URL || process.env.NEON_DATABASE_URL;

if (!databaseUrl) {
  console.error('Set DATABASE_URL, POSTGRES_URL, or NEON_DATABASE_URL before running the migration.');
  process.exit(1);
}

const { neon } = await import('@neondatabase/serverless');
const sql = neon(databaseUrl);
const statements = readFileSync('database/schema.sql', 'utf8')
  .split(/;\s*(?:\r?\n|$)/)
  .map((statement) => statement.trim())
  .filter(Boolean);

for (const statement of statements) {
  await sql.query(`${statement};`);
}

console.log(`Applied ${statements.length} Neon schema statements.`);
