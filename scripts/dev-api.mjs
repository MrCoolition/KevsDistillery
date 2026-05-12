import http from 'node:http';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { loadLocalEnv } from './load-local-env.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const loadedEnvFiles = loadLocalEnv(root);
console.log(
  `Local API env loaded from ${loadedEnvFiles.length ? loadedEnvFiles.join(', ') : 'no .env files'}; OPENAI_API_KEY ${
    process.env.OPENAI_API_KEY ? 'present' : 'missing'
  }; OPENAI_MODEL ${process.env.OPENAI_MODEL ?? 'gpt-5.5'}`,
);
const require = createRequire(import.meta.url);

compileApi();

const { default: handler } = require(path.join(root, '.tmp_api', 'dossiers.js'));

const server = http.createServer((req, res) => {
  if (!req.url?.startsWith('/api/dossiers')) {
    res.statusCode = 404;
    res.end('Not found');
    return;
  }

  const apiRes = decorateResponse(res);
  Promise.resolve(handler(req, apiRes)).catch((error) => {
    if (!res.headersSent) {
      res.statusCode = 500;
      res.setHeader('content-type', 'text/plain; charset=utf-8');
    }
    res.end(error instanceof Error ? error.message : 'Local API error');
  });
});

server.listen(3001, '127.0.0.1', () => {
  console.log('Local dossier API listening at http://127.0.0.1:3001/api/dossiers');
});

function compileApi() {
  const result = spawnSync(
    process.execPath,
    ['./node_modules/typescript/bin/tsc', '-p', 'tsconfig.api.json', '--outDir', '.tmp_api', '--noEmit', 'false'],
    {
      cwd: root,
      stdio: 'inherit',
      env: process.env,
    },
  );

  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

function decorateResponse(res) {
  res.status = (code) => {
    res.statusCode = code;
    return res;
  };

  res.json = (value) => {
    if (!res.headersSent) {
      res.setHeader('content-type', 'application/json; charset=utf-8');
    }
    res.end(JSON.stringify(value));
    return res;
  };

  res.send = (value) => {
    if (Buffer.isBuffer(value) || typeof value === 'string') {
      res.end(value);
      return res;
    }
    res.json(value);
    return res;
  };

  return res;
}
