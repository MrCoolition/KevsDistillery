import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import type { VercelRequest, VercelResponse } from '@vercel/node';
import formidable, { type File } from 'formidable';
import {
  MASTER_DOSSIER_STANDARD,
  PACKAGE_FOLDERS,
  REQUIRED_DIAGRAMS,
  RUN_PROMPT,
  UPLOAD_INSTRUCTION,
} from './_lib/contract.js';
import { enrichWithOpenAI } from './_lib/ai.js';
import { persistDossierRun } from './_lib/db.js';
import { buildDiscoveryModel } from './_lib/extractors.js';
import { buildDossierPackage } from './_lib/package-builder.js';
import { loadRuntimeEnv } from './_lib/runtime-env.js';
import type { UploadedSource } from './_lib/types.js';

export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(req: VercelRequest, res: VercelResponse): Promise<void> {
  const startedAt = Date.now();
  if (req.method === 'GET') {
    res.status(200).json({
      uploadInstruction: UPLOAD_INSTRUCTION,
      runPrompt: RUN_PROMPT,
      packageFolders: PACKAGE_FOLDERS,
      requiredDiagrams: REQUIRED_DIAGRAMS.map((diagram) => diagram.file),
      standard: MASTER_DOSSIER_STANDARD,
    });
    return;
  }

  if (req.method !== 'POST') {
    res.setHeader('allow', 'GET, POST');
    res.status(405).send('Method not allowed');
    return;
  }

  let sources: UploadedSource[] = [];

  try {
    sources = await parseUpload(req);
    if (!sources.length) {
      res.status(400).send('Upload at least one source file.');
      return;
    }

    logStage(startedAt, 'upload parsed', {
      files: sources.map((source) => ({ name: source.originalName, size: source.size })),
    });
    const model = await buildDiscoveryModel(sources);
    logStage(startedAt, 'deterministic model built', {
      nodes: model.nodes.length,
      evidence: model.evidence.length,
      blockers: model.blockedSources.length,
    });
    const envInfo = loadRuntimeEnv();
    logStage(startedAt, 'runtime env loaded', {
      files: envInfo.loadedFiles,
      openAiKeyPresent: envInfo.openAiKeyPresent,
      openAiKeyPrefix: envInfo.openAiKeyPrefix,
      openAiKeySuffix: envInfo.openAiKeySuffix,
      openAiKeyLength: envInfo.openAiKeyLength,
      openAiModel: envInfo.openAiModel,
    });
    await enrichWithOpenAI(model);
    logStage(startedAt, 'ai enrichment complete', {
      enabled: model.aiNarrative.enabled,
      error: model.aiNarrative.error ?? '',
    });
    const packageResult = await buildDossierPackage(model);
    logStage(startedAt, 'zip package built', {
      bytes: packageResult.buffer.byteLength,
      files: packageResult.summary.packageFileCount,
      qa: packageResult.summary.qaStatus,
    });
    await persistDossierRun(model, packageResult.buffer, packageResult.summary);
    logStage(startedAt, 'persistence complete', { persisted: Boolean(process.env.DATABASE_URL) });

    res.setHeader('content-type', 'application/zip');
    res.setHeader('content-disposition', `attachment; filename="${model.packageName}.zip"`);
    res.setHeader('x-dossier-summary', encodeURIComponent(JSON.stringify(packageResult.summary)));
    res.status(200).send(packageResult.buffer);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unexpected dossier generation failure.';
    logStage(startedAt, 'generation failed', { message });
    res.status(500).send(message);
  } finally {
    await cleanupTempFiles(sources);
  }
}

function logStage(startedAt: number, stage: string, details: Record<string, unknown>): void {
  console.info(`[dossier] ${stage} +${Date.now() - startedAt}ms ${JSON.stringify(details)}`);
}

async function parseUpload(req: VercelRequest): Promise<UploadedSource[]> {
  const uploadDir = await ensureUploadDir();
  const form = formidable({
    multiples: true,
    maxFiles: 25,
    maxFileSize: 1024 * 1024 * 1024,
    maxTotalFileSize: 2 * 1024 * 1024 * 1024,
    uploadDir,
    keepExtensions: true,
  });

  const [, files] = await new Promise<[formidable.Fields, formidable.Files]>((resolve, reject) => {
    form.parse(req, (error, fields, parsedFiles) => {
      if (error) {
        reject(error);
        return;
      }
      resolve([fields, parsedFiles]);
    });
  });

  const uploadedFiles = Object.values(files)
    .flatMap((value) => (Array.isArray(value) ? value : [value]))
    .filter((value): value is File => Boolean(value));

  const sources: UploadedSource[] = [];
  for (const file of uploadedFiles) {
    const buffer = await fs.readFile(file.filepath);
    sources.push({
      originalName: file.originalFilename ?? path.basename(file.filepath),
      mimeType: file.mimetype ?? undefined,
      size: file.size,
      buffer,
      tempPath: file.filepath,
    });
  }

  return sources;
}

async function cleanupTempFiles(sources: UploadedSource[]): Promise<void> {
  await Promise.allSettled(
    sources
      .map((source) => source.tempPath)
      .filter((tempPath): tempPath is string => Boolean(tempPath))
      .map((tempPath) => fs.unlink(tempPath)),
  );
}

async function ensureUploadDir(): Promise<string> {
  if (process.env.VERCEL) {
    return '/tmp';
  }
  const uploadDir =
    process.env.DISCOVERY_UPLOAD_DIR?.trim() ||
    path.join(os.tmpdir(), 'data-source-discovery-distillery', 'uploads');
  await fs.mkdir(uploadDir, { recursive: true });
  return uploadDir;
}
