import fs from 'node:fs';
import path from 'node:path';

export function loadLocalEnv(root) {
  const loadedFiles = [];

  for (const fileName of ['.env', '.env.local']) {
    const envPath = path.join(root, fileName);
    if (!fs.existsSync(envPath)) {
      continue;
    }

    loadedFiles.push(fileName);
    const lines = fs.readFileSync(envPath, 'utf8').split(/\r?\n/);
    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith('#')) {
        continue;
      }

      const match = trimmed.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
      if (!match) {
        continue;
      }

      const key = match[1];
      let value = match[2] ?? '';
      if (
        (value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'"))
      ) {
        value = value.slice(1, -1);
      }

      process.env[key] = value;
    }
  }

  return loadedFiles;
}
