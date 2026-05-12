import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { loadLocalEnv } from './load-local-env.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const loadedEnvFiles = loadLocalEnv(root);
console.log(
  `Local env loaded from ${loadedEnvFiles.length ? loadedEnvFiles.join(', ') : 'no .env files'}; OPENAI_API_KEY ${
    process.env.OPENAI_API_KEY ? 'present' : 'missing'
  }; OPENAI_MODEL ${process.env.OPENAI_MODEL ?? 'gpt-5.5'}`,
);
const esbuildBinary = path.join(root, 'node_modules', '@esbuild', 'win32-x64', 'esbuild.exe');
const env = {
  ...process.env,
  ESBUILD_BINARY_PATH: esbuildBinary,
};

const children = [];

start('api', process.execPath, ['./scripts/dev-api.mjs']);
start('ui', process.execPath, [
  './node_modules/@angular/cli/bin/ng.js',
  'serve',
  '--host',
  '127.0.0.1',
  '--port',
  '4200',
  '--proxy-config',
  'proxy.conf.json',
]);

process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);

function start(name, command, args) {
  const child = spawn(command, args, {
    cwd: root,
    env,
    stdio: 'inherit',
  });

  children.push(child);
  child.on('exit', (code, signal) => {
    if (code && code !== 0) {
      console.error(`${name} exited with code ${code}`);
      shutdown();
    }
    if (signal) {
      console.error(`${name} exited from signal ${signal}`);
      shutdown();
    }
  });
}

function shutdown() {
  for (const child of children) {
    if (!child.killed) {
      child.kill();
    }
  }
  process.exit(0);
}
