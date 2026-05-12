import fs from 'node:fs';
import path from 'node:path';

export type RuntimeEnvLoadResult = {
  loadedFiles: string[];
  openAiKeyPresent: boolean;
  openAiKeyPrefix: string;
  openAiKeySuffix: string;
  openAiKeyLength: number;
  openAiModel: string;
};

export function loadRuntimeEnv(): RuntimeEnvLoadResult {
  if (!process.env.VERCEL) {
    loadEnvFiles(process.cwd(), ['.env', '.env.local']);
  }

  const apiKey = process.env.OPENAI_API_KEY?.trim() ?? '';
  if (apiKey) {
    process.env.OPENAI_API_KEY = apiKey;
  }

  return {
    loadedFiles: loadedFiles(process.cwd(), ['.env', '.env.local']),
    openAiKeyPresent: Boolean(apiKey),
    openAiKeyPrefix: apiKey.slice(0, 7),
    openAiKeySuffix: apiKey.slice(-4),
    openAiKeyLength: apiKey.length,
    openAiModel: process.env.OPENAI_MODEL ?? 'gpt-5.5',
  };
}

function loadEnvFiles(root: string, fileNames: string[]): void {
  for (const fileName of fileNames) {
    const envPath = path.join(root, fileName);
    if (!fs.existsSync(envPath)) {
      continue;
    }

    const lines = fs.readFileSync(envPath, 'utf8').split(/\r?\n/);
    for (const line of lines) {
      const parsed = parseEnvLine(line);
      if (!parsed) {
        continue;
      }
      process.env[parsed.key] = parsed.value;
    }
  }
}

function loadedFiles(root: string, fileNames: string[]): string[] {
  if (process.env.VERCEL) {
    return ['vercel-environment'];
  }
  return fileNames.filter((fileName) => fs.existsSync(path.join(root, fileName)));
}

function parseEnvLine(line: string): { key: string; value: string } | null {
  const trimmed = line.replace(/^\uFEFF/, '').trim();
  if (!trimmed || trimmed.startsWith('#')) {
    return null;
  }

  const match = trimmed.match(/^(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
  if (!match) {
    return null;
  }

  const key = match[1];
  let value = match[2] ?? '';
  value = parseEnvValue(value);
  return { key, value };
}

function parseEnvValue(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) {
    return '';
  }

  const quote = trimmed[0];
  if (quote === '"' || quote === "'") {
    const end = findClosingQuote(trimmed, quote);
    return end > 0 ? trimmed.slice(1, end) : trimmed.slice(1);
  }

  return trimmed.replace(/\s+#.*$/, '').trim();
}

function findClosingQuote(value: string, quote: string): number {
  for (let index = value.length - 1; index > 0; index -= 1) {
    if (value[index] === quote && value[index - 1] !== '\\') {
      return index;
    }
  }
  return -1;
}
