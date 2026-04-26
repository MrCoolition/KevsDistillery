import { createRequire } from 'node:module';
import { createServer } from 'node:http';
import { extname, join, normalize, resolve } from 'node:path';
import { existsSync, readFileSync } from 'node:fs';

import { loadDotEnv } from './env.mjs';

loadDotEnv();

const require = createRequire(import.meta.url);
const root = resolve('.');
const port = Number(process.env.DISTILLERY_APP_PORT || 4173);

const apiRoutes = new Map([
  ['/api/health', require('../api/health.js')],
  ['/api/admin/migrate', require('../api/admin/migrate.js')],
  ['/api/discovery/synthesize', require('../api/discovery/synthesize.js')],
  ['/api/discovery/runs', require('../api/discovery/runs.js')],
  ['/api/discovery/run', require('../api/discovery/run.js')]
]);

const contentTypes = new Map([
  ['.html', 'text/html; charset=utf-8'],
  ['.js', 'text/javascript; charset=utf-8'],
  ['.css', 'text/css; charset=utf-8'],
  ['.json', 'application/json; charset=utf-8'],
  ['.svg', 'image/svg+xml; charset=utf-8']
]);

function sendFile(response, path) {
  const extension = extname(path);
  response.statusCode = 200;
  response.setHeader('Content-Type', contentTypes.get(extension) || 'application/octet-stream');
  response.end(readFileSync(path));
}

function safePath(pathname) {
  const clean = normalize(pathname).replace(/^(\.\.[/\\])+/, '');
  return resolve(root, clean.replace(/^[/\\]/, ''));
}

const server = createServer(async (request, response) => {
  const url = new URL(request.url || '/', `http://${request.headers.host || '127.0.0.1'}`);
  const route = apiRoutes.get(url.pathname);

  if (route) {
    await route(request, response);
    return;
  }

  if (request.method !== 'GET' && request.method !== 'HEAD') {
    response.statusCode = 405;
    response.end('Method not allowed');
    return;
  }

  if (url.pathname.startsWith('/assets/') || url.pathname.startsWith('/vendor/')) {
    const assetPath = safePath(join('public', url.pathname));
    if (assetPath.startsWith(root) && existsSync(assetPath)) {
      sendFile(response, assetPath);
      return;
    }
  }

  sendFile(response, resolve(root, 'preview.html'));
});

server.listen(port, '127.0.0.1', () => {
  console.log(`Uncle Kev's Distillery is running at http://127.0.0.1:${port}`);
});
