import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { neon } from '@neondatabase/serverless';
import { loadLocalEnv } from './load-local-env.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
loadLocalEnv(root);

const databaseUrl = process.env.DATABASE_URL;
if (!databaseUrl) {
  console.error('DATABASE_URL is missing. Add the Neon pooled or direct connection string, then rerun db:migrate.');
  process.exit(1);
}

const schemaPath = path.join(root, 'db', 'schema.sql');
const schemaSql = await readFile(schemaPath, 'utf8');
const statements = splitSqlStatements(schemaSql);
const sql = neon(databaseUrl);

await sql.transaction((txn) => statements.map((statement) => txn.query(statement)));

const rows = await sql`
  select migration_id, applied_at
  from data_discovery.schema_migrations
  order by applied_at desc
  limit 5
`;

console.log(`Applied ${statements.length} schema statement(s) to Neon.`);
console.log(`Latest migration: ${rows[0]?.migration_id ?? 'none'}`);

function splitSqlStatements(sqlText) {
  const statements = [];
  let current = '';
  let quote = null;
  let dollarTag = null;
  let lineComment = false;
  let blockComment = false;

  for (let index = 0; index < sqlText.length; index += 1) {
    const char = sqlText[index];
    const next = sqlText[index + 1];

    if (lineComment) {
      current += char;
      if (char === '\n') {
        lineComment = false;
      }
      continue;
    }

    if (blockComment) {
      current += char;
      if (char === '*' && next === '/') {
        current += next;
        index += 1;
        blockComment = false;
      }
      continue;
    }

    if (!quote && !dollarTag && char === '-' && next === '-') {
      current += char + next;
      index += 1;
      lineComment = true;
      continue;
    }

    if (!quote && !dollarTag && char === '/' && next === '*') {
      current += char + next;
      index += 1;
      blockComment = true;
      continue;
    }

    if (!quote && char === '$') {
      const match = sqlText.slice(index).match(/^\$[A-Za-z_][A-Za-z0-9_]*\$|^\$\$/);
      if (match) {
        const tag = match[0];
        current += tag;
        index += tag.length - 1;
        dollarTag = dollarTag === tag ? null : tag;
        continue;
      }
    }

    if (!dollarTag && (char === "'" || char === '"')) {
      if (quote === char && next === char) {
        current += char + next;
        index += 1;
        continue;
      }
      quote = quote === char ? null : quote || char;
      current += char;
      continue;
    }

    if (!quote && !dollarTag && char === ';') {
      const statement = current.trim();
      if (statement) {
        statements.push(statement);
      }
      current = '';
      continue;
    }

    current += char;
  }

  const trailing = current.trim();
  if (trailing) {
    statements.push(trailing);
  }
  return statements;
}
