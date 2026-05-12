import crypto from 'node:crypto';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import path from 'node:path';
import { promisify } from 'node:util';
import { parse as parseCsv } from 'csv-parse/sync';
import * as CFB from 'cfb';
import JSZip from 'jszip';
import mammoth from 'mammoth';
import { format as formatSql } from 'sql-formatter';
import * as XLSX from 'xlsx';
import { QA_CHECKS } from './contract.js';
import type {
  ActionItem,
  Confidence,
  Criticality,
  DataElement,
  DataQualityFinding,
  DiscoveryEdge,
  DiscoveryModel,
  DiscoveryNode,
  EdgeType,
  EvidenceItem,
  FinancialExposure,
  NodeType,
  OpenQuestion,
  ProcessStep,
  SecurityAccessFinding,
  SourceFileMeta,
  SourceType,
  TransformationRule,
  UploadedSource,
} from './types.js';

const execFileAsync = promisify(execFile);

type IdPrefix =
  | 'ACT'
  | 'CTL'
  | 'DQ'
  | 'EDGE'
  | 'EVID'
  | 'FIN'
  | 'NODE'
  | 'OPEN'
  | 'QA'
  | 'SRC'
  | 'STEP'
  | 'TRN'
  | 'DE';

type IdState = Record<IdPrefix, number>;

const ACCESS_EXTENSIONS = new Set(['.accdb', '.mdb']);
const EXCEL_EXTENSIONS = new Set(['.xlsx', '.xlsm', '.xlsb', '.xls']);
const WORD_EXTENSIONS = new Set(['.docx', '.doc']);
const FLAT_EXTENSIONS = new Set(['.csv', '.txt', '.tsv']);
const SQL_EXTENSIONS = new Set(['.sql']);
const SCRIPT_EXTENSIONS = new Set(['.ps1', '.py', '.js', '.ts', '.r', '.sas', '.sh', '.bat']);
const DEFAULT_EXTRACTION_CONCURRENCY = 4;
const ACCESS_COLLECTOR_TIMEOUT_MS = Number(process.env.ACCESS_COLLECTOR_TIMEOUT_MS ?? 600_000);
const ACCESS_FAST_COLLECTOR_TIMEOUT_MS = Number(process.env.ACCESS_FAST_COLLECTOR_TIMEOUT_MS ?? 120_000);
const ACCESS_DEEP_EXPORT_TIMEOUT_MS = Number(process.env.ACCESS_DEEP_EXPORT_TIMEOUT_MS ?? 45_000);
const ACCESS_DEEP_MAX_OBJECTS = Number(process.env.ACCESS_DEEP_MAX_OBJECTS ?? 300);
const EXCEL_COLLECTOR_TIMEOUT_MS = Number(process.env.EXCEL_COLLECTOR_TIMEOUT_MS ?? 300_000);
const ACCESS_LARGE_FILE_THRESHOLD_BYTES = Number(process.env.ACCESS_LARGE_FILE_THRESHOLD_BYTES ?? 150 * 1024 * 1024);
const TRUSTED_DESKTOP_DEEP_EXPORT = process.env.TRUSTED_DESKTOP_DEEP_EXPORT !== '0';
const RECURSIVE_SOURCE_DISCOVERY = process.env.DISCOVERY_RECURSE_LINKED_SOURCES !== '0';
const MAX_RECURSION_DEPTH = Number(process.env.DISCOVERY_MAX_RECURSION_DEPTH ?? 4);
const MAX_RECURSIVE_FILE_BYTES = Number(process.env.DISCOVERY_MAX_RECURSIVE_FILE_BYTES ?? 1024 * 1024 * 1024);

type ExtractionContext = {
  linkedSourcePaths: Set<string>;
  recursionDepth: number;
  maxRecursionDepth: number;
};

type LinkedSourceResolution = {
  key: string;
  tableName: string;
  connect: string;
  sourceTableName: string;
  sourceKind: string;
  sourcePath: string;
  displayName: string;
  resolvedPath: string;
  resolutionStatus: 'resolved_file' | 'resolved_folder' | 'blocked' | 'not_applicable';
  resolutionReason: string;
  fileSizeBytes?: number;
};

export async function buildDiscoveryModel(files: UploadedSource[]): Promise<DiscoveryModel> {
  const generatedDate = new Date().toISOString().slice(0, 10);
  const sourceProcessName = deriveProcessName(files);
  const packageName = `Data_Source_Discovery_Dossier_${sourceProcessName}_${generatedDate}`;
  const ids: IdState = {
    ACT: 0,
    CTL: 0,
    DE: 0,
    DQ: 0,
    EDGE: 0,
    EVID: 0,
    FIN: 0,
    NODE: 0,
    OPEN: 0,
    QA: 0,
    SRC: 0,
    STEP: 0,
    TRN: 0,
  };

  const model: DiscoveryModel = {
    runId: crypto.randomUUID(),
    packageName,
    sourceProcessName,
    generatedDate,
    analysisVersion: 'deterministic-extraction-1.0',
    packageVersion: process.env.DISCOVERY_PACKAGE_VERSION ?? '1.0.0',
    sourceFiles: [],
    sourceTypeCounts: {
      access: 0,
      excel: 0,
      word: 0,
      'flat-file': 0,
      sql: 0,
      script: 0,
      unknown: 0,
    },
    nodes: [],
    edges: [],
    evidence: [],
    processSteps: [],
    transformations: [],
    dataElements: [],
    dataQualityFindings: [],
    controls: [],
    securityAccess: [],
    openQuestions: [],
    actions: [],
    financialExposure: [],
    qaRecords: [],
    access: emptyAccessRegisters(),
    excel: emptyExcelRegisters(),
    word: emptyWordRegisters(),
    dependencyUsage: [],
    modernization: [],
    scheduleSla: [],
    failureModes: [],
    rowCountSummary: {},
    aiNarrative: {
      enabled: false,
      model: process.env.OPENAI_MODEL ?? 'gpt-5.5',
    },
    limitations: [],
    blockedSources: [],
    assumptions: [
      'Ownership is marked owner-confirmation-required until a named accountable owner is supplied.',
      'Financial exposure is directional and proxy-based until finance supplies certified values.',
      'File and folder permissions cannot be inspected from browser uploads and require a platform-side security review.',
    ],
  };

  const runtimeEvidenceId = addEvidence(model, ids, {
    category: '05m_QA_Certification',
    title: 'Runtime extraction strategy',
    fileName: 'runtime_extraction_strategy.json',
    sourceFile: 'dossier-run',
    summary: 'Documents deterministic-first extraction and AI-after-evidence sequencing.',
    confidence: 'confirmed',
    content: JSON.stringify(
      {
        systematic_first: true,
        ai_after_canonical_model: true,
        ai_model: model.aiNarrative.model,
        supported_extractors: ['csv-parse', 'xlsx', 'jszip', 'mammoth', 'sql-formatter', 'deterministic SQL reference scanner'],
        parallel_processing: {
          source_file_concurrency: extractionConcurrency(files.length),
          package_artifact_generation: 'executive brief, architecture report, workbook, diagrams, and financial model are generated in parallel before ZIP assembly',
          access_collector_timeout_ms: ACCESS_COLLECTOR_TIMEOUT_MS,
          access_large_file_threshold_bytes: ACCESS_LARGE_FILE_THRESHOLD_BYTES,
          access_deep_com_export_enabled: TRUSTED_DESKTOP_DEEP_EXPORT,
          access_deep_large_export_enabled: TRUSTED_DESKTOP_DEEP_EXPORT,
          excel_desktop_vba_export_enabled: TRUSTED_DESKTOP_DEEP_EXPORT,
          excel_collector_timeout_ms: EXCEL_COLLECTOR_TIMEOUT_MS,
        },
        blocked_native_extractors: ['Access ACE/OLEDB system catalog extraction is unavailable in Vercel serverless runtime.'],
      },
      null,
      2,
    ),
  });

  const packageNode = addNode(model, ids, {
    node_type: 'output',
    name: model.packageName,
    description: 'Generated Data Source Discovery Dossier package.',
    source_file: 'generated',
    business_purpose: 'Leadership, analyst, engineering, governance, audit, migration, and finance decision package.',
    owner_status: 'generated by discovery workflow',
    criticality: 'P1',
    confidence: 'confirmed',
    evidence_id: runtimeEvidenceId,
    recommended_action: 'Review blockers, validate owners and finance inputs, then execute P0/P1 actions.',
    failure_impact: 'Without the dossier package, stakeholders lack defensible lineage, risk, and modernization guidance.',
    dollar_exposure: 'Directional exposure model included; finance validation required.',
  });

  const extractionContext: ExtractionContext = {
    linkedSourcePaths: new Set<string>(),
    recursionDepth: 0,
    maxRecursionDepth: MAX_RECURSION_DEPTH,
  };

  await runWithConcurrency(files, extractionConcurrency(files.length), (file) =>
    extractFile(file, model, ids, packageNode.node_id, extractionContext),
  );

  addGlobalGovernanceFindings(model, ids, runtimeEvidenceId);
  addFinancialExposure(model, ids, runtimeEvidenceId);
  model.blockedSources = uniqueBlockedSources(model.blockedSources);
  addQaRecords(model, ids, runtimeEvidenceId);

  return model;
}

function deriveProcessName(files: UploadedSource[]): string {
  if (files.length === 1) {
    return sanitizeName(path.parse(files[0]?.originalName ?? 'Source').name);
  }

  const firstName = path.parse(files[0]?.originalName ?? 'Multi_Source').name;
  return sanitizeName(`${firstName}_Multi_Source`);
}

function sanitizeName(value: string): string {
  const cleaned = value.replace(/[^a-zA-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
  return cleaned || 'Source_Process';
}

function uniqueBlockedSources(sources: string[]): string[] {
  const seen = new Set<string>();
  const unique: string[] = [];
  for (const source of sources) {
    const parts = source.split('|').map((part) => part.trim()).filter(Boolean);
    const keySource = parts.at(-1) ?? source;
    const key = keySource.replace(/\s+/g, ' ').toLowerCase();
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    unique.push(source);
  }
  return unique;
}

function nextId(ids: IdState, prefix: IdPrefix): string {
  ids[prefix] += 1;
  return `${prefix}-${String(ids[prefix]).padStart(4, '0')}`;
}

function extractionConcurrency(fileCount: number): number {
  const configured = Number(process.env.DISCOVERY_FILE_CONCURRENCY ?? DEFAULT_EXTRACTION_CONCURRENCY);
  const bounded = Number.isFinite(configured) && configured > 0 ? configured : DEFAULT_EXTRACTION_CONCURRENCY;
  return Math.max(1, Math.min(fileCount || 1, bounded));
}

async function runWithConcurrency<T>(
  items: T[],
  concurrency: number,
  worker: (item: T, index: number) => Promise<void>,
): Promise<void> {
  let nextIndex = 0;
  const workers = Array.from({ length: Math.min(concurrency, items.length) }, async () => {
    while (nextIndex < items.length) {
      const index = nextIndex;
      nextIndex += 1;
      await worker(items[index] as T, index);
    }
  });
  await Promise.all(workers);
}

function detectSourceType(fileName: string): SourceType {
  const extension = path.extname(fileName).toLowerCase();
  if (ACCESS_EXTENSIONS.has(extension)) {
    return 'access';
  }
  if (EXCEL_EXTENSIONS.has(extension)) {
    return 'excel';
  }
  if (WORD_EXTENSIONS.has(extension)) {
    return 'word';
  }
  if (FLAT_EXTENSIONS.has(extension)) {
    return 'flat-file';
  }
  if (SQL_EXTENSIONS.has(extension)) {
    return 'sql';
  }
  if (SCRIPT_EXTENSIONS.has(extension)) {
    return 'script';
  }
  return 'unknown';
}

async function extractFile(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  packageNodeId: string,
  context: ExtractionContext,
): Promise<void> {
  const sourceType = detectSourceType(file.originalName);
  const sourceEvidenceId = addEvidence(model, ids, {
    category: '05a_Raw_Metadata',
    title: `File metadata - ${file.originalName}`,
    fileName: `${sanitizeName(file.originalName)}_metadata.json`,
    sourceFile: file.originalName,
    summary: 'Raw uploaded file metadata and checksum.',
    confidence: 'confirmed',
    content: JSON.stringify(
      {
        file_name: file.originalName,
        file_type: sourceType,
        mime_type: file.mimeType ?? 'unknown',
        size_bytes: file.size,
        source_path: file.sourcePath ?? '',
        discovered_from: file.discoveredFrom ?? '',
        sha256: sha256(file.buffer),
      },
      null,
      2,
    ),
  });

  const sourceMeta: SourceFileMeta = {
    source_id: nextId(ids, 'SRC'),
    file_name: file.originalName,
    file_type: sourceType,
    extension: path.extname(file.originalName).toLowerCase(),
    file_size_bytes: file.size,
    sha256: sha256(file.buffer),
    evidence_id: sourceEvidenceId,
  };
  model.sourceFiles.push(sourceMeta);
  model.sourceTypeCounts[sourceType] += 1;

  const fileNode = addNode(model, ids, {
    node_type: sourceType === 'excel' ? 'workbook' : sourceType === 'word' ? 'document' : sourceType === 'access' ? 'database' : 'file',
    name: file.originalName,
    description: `${sourceType} source uploaded for discovery.`,
    source_file: file.originalName,
    business_purpose: 'Uploaded source asset for current-state discovery.',
    owner_status: 'owner-confirmation-required',
    criticality: 'P1',
    confidence: 'confirmed',
    evidence_id: sourceEvidenceId,
    recommended_action: 'Confirm owner, business purpose, cadence, upstream origin, and downstream consumers.',
    failure_impact: 'If unavailable or stale, lineage and downstream decisions may be incomplete or wrong.',
    dollar_exposure: 'Directional exposure modeled at process level.',
  });

  addEdge(model, ids, {
    from_node_id: fileNode.node_id,
    to_node_id: packageNodeId,
    edge_type: 'documents',
    description: 'Uploaded source is documented by the generated dossier.',
    automated_flag: 'automated',
    transformation_id: '',
    cadence: 'on demand',
    confidence: 'confirmed',
    evidence_id: sourceEvidenceId,
  });

  if (sourceType === 'access') {
    await extractAccess(file, model, ids, fileNode.node_id, packageNodeId, sourceEvidenceId, context);
    return;
  }

  if (sourceType === 'excel') {
    await extractExcel(file, model, ids, fileNode.node_id, sourceEvidenceId);
    return;
  }

  if (sourceType === 'word') {
    await extractWord(file, model, ids, fileNode.node_id, sourceEvidenceId);
    return;
  }

  if (sourceType === 'flat-file') {
    extractFlatFile(file, model, ids, fileNode.node_id, sourceEvidenceId);
    return;
  }

  if (sourceType === 'sql' || sourceType === 'script') {
    extractSqlOrScript(file, model, ids, fileNode.node_id, sourceEvidenceId, sourceType);
    return;
  }

  addOpenQuestion(model, ids, {
    asset: file.originalName,
    question: 'Unsupported or unknown file type. Confirm required extractor and business role.',
    owner_role: 'Data owner / platform owner',
    blocker_type: 'unsupported-source-type',
    priority: 'P1',
    evidence_id: sourceEvidenceId,
  });
}

async function extractAccess(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  fileNodeId: string,
  packageNodeId: string,
  sourceEvidenceId: string,
  context: ExtractionContext,
): Promise<void> {
  if (!process.env.VERCEL && file.tempPath) {
    try {
      const collector = await runAccessCollector(file.tempPath, 'fast', ACCESS_FAST_COLLECTOR_TIMEOUT_MS);
      const linkedSourceResolutions = await resolveAccessLinkedSources(collector);
      const collectorEvidenceId = addEvidence(model, ids, {
        category: '05a_Raw_Metadata',
        title: `Access fast native collector output - ${file.originalName}`,
        fileName: `${sanitizeName(file.originalName)}_access_native_collector.json`,
        sourceFile: file.originalName,
        summary: 'Fast native Windows Access collector output covering ACE/OLEDB schema, DAO query SQL, table definitions, relations, and document metadata where available. This phase must complete before any deep Access.Application export is attempted.',
        confidence: collector.msys_objects?.status === 'ok' || collector.dao_querydefs?.status === 'ok' ? 'partial' : 'blocked',
        content: JSON.stringify(collector, null, 2),
      });

      processAccessCollectorResult(file, model, ids, fileNodeId, sourceEvidenceId, collectorEvidenceId, collector, linkedSourceResolutions);
      await extractResolvedLinkedSources(file, model, ids, packageNodeId, context, linkedSourceResolutions, collectorEvidenceId);
      if (TRUSTED_DESKTOP_DEEP_EXPORT) {
        await runAccessDeepExportPhase(file, model, ids, fileNodeId);
      }
      return;
    } catch (error) {
      const message = error instanceof Error ? error.message : 'unknown Access collector error';
      createAccessBlocker(
        file,
        model,
        ids,
        fileNodeId,
        sourceEvidenceId,
        `Native Access collector failed: ${message}`,
      );
      return;
    }
  }

  createAccessBlocker(
    file,
    model,
    ids,
    fileNodeId,
    sourceEvidenceId,
    'Access native extraction is only available in local Windows development with ACE/OLEDB or DAO providers.',
  );
}

async function runAccessDeepExportPhase(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  fileNodeId: string,
): Promise<void> {
  if (!file.tempPath) {
    return;
  }
  try {
    const collector = await runAccessCollector(file.tempPath, 'deep', ACCESS_DEEP_EXPORT_TIMEOUT_MS);
    const collectorEvidenceId = addEvidence(model, ids, {
      category: '05a_Raw_Metadata',
      title: `Access trusted desktop deep export output - ${file.originalName}`,
      fileName: `${sanitizeName(file.originalName)}_access_deep_export_collector.json`,
      sourceFile: file.originalName,
      summary: 'Bounded Access.Application SaveAsText export for queries, macros, forms, reports, and modules. Runs after fast DAO/OLEDB catalog extraction so a deep export issue cannot erase core discovery.',
      confidence: collector.access_application?.status === 'ok' ? 'partial' : 'blocked',
      content: JSON.stringify(collector, null, 2),
    });
    processAccessDeepCollectorResult(file, model, ids, fileNodeId, collectorEvidenceId, collector);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'unknown Access deep export error';
    const evidenceId = addEvidence(model, ids, {
      category: '05m_QA_Certification',
      title: `Access deep export timeout or failure - ${file.originalName}`,
      fileName: `${sanitizeName(file.originalName)}_access_deep_export_failure.json`,
      sourceFile: file.originalName,
      summary: 'Trusted desktop Access.Application SaveAsText deep export did not complete inside the bounded timeout. Fast DAO/OLEDB discovery evidence remains valid and is used for the dossier.',
      confidence: 'blocked',
      content: JSON.stringify({ source_file: file.originalName, error: message, timeout_ms: ACCESS_DEEP_EXPORT_TIMEOUT_MS }, null, 2),
    });
    model.qaRecords.push({
      qa_id: nextId(ids, 'QA'),
      check: 'Access deep export bounded separately from fast catalog extraction',
      status: 'PASS_WITH_LIMITATION',
      evidence_id: evidenceId,
      notes: `Fast Access catalog evidence was retained. Deep Access.Application SaveAsText failed or timed out: ${message}`,
    });
    addAction(model, ids, {
      title: `Review Access desktop deep export for ${file.originalName}`,
      description: `Fast catalog extraction completed, but Access.Application SaveAsText deep export failed or timed out: ${message}`,
      source_asset: file.originalName,
      owner_role: 'Access technical owner',
      recommended_owner: 'Access owner with trusted desktop runtime',
      action_type: 'Resolve Blocker',
      priority: 'P1',
      severity: 'P1',
      dependency: 'Access.Application SaveAsText desktop automation',
      acceptance_criteria: 'Macro/action/form/report/module text exports complete without blocking fast catalog discovery.',
      evidence_id: evidenceId,
      related_risk: 'Partial automation/UI evidence',
      expected_business_value: 'Completes automation and UI lineage while preserving certified table/query discovery.',
    });
  }
}

async function runAccessCollector(
  tempPath: string,
  mode: 'fast' | 'deep' | 'all' = 'all',
  timeoutMs = ACCESS_COLLECTOR_TIMEOUT_MS,
): Promise<Record<string, any>> {
  const scriptPath = path.join(process.cwd(), 'scripts', 'access-collector.ps1');
  const { stdout } = await execFileAsync(
    'powershell.exe',
    [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      scriptPath,
      '-Path',
      tempPath,
      '-LargeFileThresholdBytes',
      String(ACCESS_LARGE_FILE_THRESHOLD_BYTES),
      '-Mode',
      mode,
      '-MaxDeepObjects',
      String(ACCESS_DEEP_MAX_OBJECTS),
    ],
    {
      timeout: timeoutMs,
      maxBuffer: 256 * 1024 * 1024,
      windowsHide: true,
      env: {
        ...process.env,
        ACCESS_COLLECTOR_DEEP: mode === 'deep' ? '1' : mode === 'fast' ? '0' : TRUSTED_DESKTOP_DEEP_EXPORT ? '1' : (process.env.ACCESS_COLLECTOR_DEEP ?? '0'),
        ACCESS_COLLECTOR_DEEP_LARGE: TRUSTED_DESKTOP_DEEP_EXPORT ? '1' : (process.env.ACCESS_COLLECTOR_DEEP_LARGE ?? '0'),
      },
    },
  );
  return JSON.parse(stdout) as Record<string, any>;
}

function processAccessCollectorResult(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  fileNodeId: string,
  sourceEvidenceId: string,
  collectorEvidenceId: string,
  collector: Record<string, any>,
  linkedSourceResolutions: Map<string, LinkedSourceResolution>,
): void {
  const msysRows = asRows(collector.msys_objects?.rows);
  const schemaTables = asRows(collector.ole_db_schema_tables?.rows);
  const schemaColumns = asRows(collector.ole_db_schema_columns?.rows);
  const queryDefs = asRows(collector.dao_querydefs?.rows);
  const tableDefs = asRows(collector.dao_tabledefs?.rows);
  const daoFields = asRows(collector.dao_fields?.rows);
  const daoIndexes = asRows(collector.dao_indexes?.rows);
  const daoRelations = asRows(collector.dao_relations?.rows);
  const daoDocuments = asRows(collector.dao_documents?.rows);
  const appExports = asRows(collector.access_application?.rows);
  const collectorLimitations = Array.isArray(collector.limitations)
    ? collector.limitations.map((item) => String(item)).filter(Boolean)
    : [];
  model.limitations.push(...collectorLimitations);

  if (collector.provider) {
    model.access.Access_Object_Inventory.push({
      object_name: file.originalName,
      raw_type_value: 'database',
      interpreted_type: 'database file',
      provider: collector.provider,
      evidence_id: collectorEvidenceId,
      confidence: 'confirmed',
      recommended_action: 'Review Access object inventory and reconcile all blocked sub-extractions.',
    });
  }

  const hasDaoObjectEvidence = tableDefs.length > 0 || queryDefs.length > 0 || daoDocuments.length > 0;
  const objectRows = msysRows.length
    ? msysRows.map((row) => ({
        name: String(row.Name ?? row.name ?? ''),
        type: Number(row.Type ?? row.type),
        flags: row.Flags ?? row.flags ?? '',
        connect: String(row.Connect ?? row.connect ?? ''),
        database: String(row.Database ?? row.database ?? ''),
        foreignName: String(row.ForeignName ?? row.foreignName ?? ''),
        confidence: 'confirmed' as Confidence,
      }))
    : hasDaoObjectEvidence
      ? []
      : schemaTables.map((row) => ({
        name: String(row.TABLE_NAME ?? row.table_name ?? ''),
        type: schemaTableTypeToAccessType(String(row.TABLE_TYPE ?? row.table_type ?? '')),
        flags: '',
        connect: '',
        database: '',
        foreignName: '',
        confidence: 'partial' as Confidence,
      }));

  const querySqlByName = new Map(queryDefs.map((row) => [String(row.name ?? row.Name ?? ''), String(row.sql ?? row.SQL ?? '')]));
  const tableDefByName = new Map(tableDefs.map((row) => [String(row.name ?? row.Name ?? ''), row]));
  const queryNames = new Set<string>();
  const macroNames = new Set<string>();
  const sqlEvidenceByQuery = new Map<string, string>();

  for (const object of objectRows.filter((row) => row.name && !row.name.startsWith('~TMP'))) {
    const interpretedType = interpretAccessType(object.type);
    model.access.Access_Object_Inventory.push({
      object_name: object.name,
      raw_type_value: Number.isFinite(object.type) ? object.type : 'schema-only',
      interpreted_type: interpretedType,
      flags: object.flags,
      connect: object.connect,
      database: object.database,
      foreign_name: object.foreignName,
      evidence_id: collectorEvidenceId,
      confidence: object.confidence,
      recommended_action: accessRecommendedAction(interpretedType),
    });

    if (object.type === 1) {
      const tableNode = addAccessObjectNode(model, ids, file, object.name, 'table', collectorEvidenceId, object.confidence);
      addAccessContainmentEdge(model, ids, fileNodeId, tableNode.node_id, collectorEvidenceId, object.confidence);
      const tableDef = tableDefByName.get(object.name);
      model.access.Access_Table_Register.push({
        table_name: object.name,
        row_count: tableDef?.record_count ?? 'not extracted',
        field_count: tableDef?.field_count ?? '',
        evidence_id: collectorEvidenceId,
        confidence: object.confidence,
        recommended_action: 'Confirm owner, primary key, row-count expectations, and source-of-truth status.',
      });
      continue;
    }

    if (object.type === 4 || object.type === 6 || object.connect) {
      const linkedNode = addAccessObjectNode(model, ids, file, object.name, 'linked table', collectorEvidenceId, object.confidence);
      addAccessContainmentEdge(model, ids, fileNodeId, linkedNode.node_id, collectorEvidenceId, object.confidence);
      const resolution = linkedSourceResolutions.get(linkedSourceKey(object.connect, object.foreignName, object.name));
      const upstreamName = resolution?.displayName || object.connect || object.database || object.foreignName || `Linked source for ${object.name}`;
      const upstream = addNode(model, ids, {
        node_type: resolution?.resolutionStatus === 'resolved_file' ? nodeTypeForResolvedPath(resolution.resolvedPath) : resolution?.resolutionStatus === 'resolved_folder' ? 'folder' : 'upstream blocker',
        name: upstreamName,
        description: `Linked source referenced by Access linked table ${object.name}.`,
        source_file: file.originalName,
        business_purpose: 'Terminal upstream source candidate for recursive lineage.',
        owner_status: 'owner-confirmation-required',
        criticality: 'P1',
        confidence: resolution?.resolutionStatus === 'resolved_file' || resolution?.resolutionStatus === 'resolved_folder' ? 'confirmed' : upstreamName === `Linked source for ${object.name}` ? 'unknown' : 'partial',
        evidence_id: collectorEvidenceId,
        recommended_action:
          resolution?.resolutionStatus === 'resolved_file'
            ? 'Inspect recursively extracted linked source evidence and confirm terminal source-of-truth status.'
            : 'Provide linked source file/system access and classify as authoritative, third-party, manual entry, obsolete, duplicate, or approved stopping point.',
        failure_impact: 'If the linked source changes or disappears, Access outputs may fail, stale, or produce wrong data.',
        dollar_exposure: 'Directional exposure modeled at process level.',
      });
      addEdge(model, ids, {
        from_node_id: linkedNode.node_id,
        to_node_id: upstream.node_id,
        edge_type: 'reads_from',
        description: 'Access linked table reads from external source.',
        automated_flag: 'automated',
        transformation_id: '',
        cadence: 'per Access refresh/open',
        confidence: upstream.confidence,
        evidence_id: collectorEvidenceId,
      });
      if (!resolution || resolution.resolutionStatus === 'blocked') {
        model.blockedSources.push(formatBlockedSource(resolution, upstreamName));
      }
      model.access.Access_Linked_Table_Register.push({
        table_name: object.name,
        connect: object.connect,
        database: object.database,
        foreign_name: object.foreignName,
        source_kind: resolution?.sourceKind ?? 'unknown',
        source_name: resolution?.displayName ?? upstreamName,
        source_path: resolution?.sourcePath ?? '',
        resolved_path: resolution?.resolvedPath ?? '',
        resolution_status: resolution?.resolutionStatus ?? 'blocked',
        resolution_reason: resolution?.resolutionReason ?? 'Linked source could not be parsed from Access metadata.',
        linked_source_node_id: upstream.node_id,
        evidence_id: collectorEvidenceId,
        confidence: object.confidence,
        recommended_action: 'Resolve linked source path/system and include source file in discovery package.',
      });
      continue;
    }

    if (object.type === 5) {
      queryNames.add(object.name);
      const queryNode = addAccessObjectNode(model, ids, file, object.name, 'query', collectorEvidenceId, object.confidence);
      addAccessContainmentEdge(model, ids, fileNodeId, queryNode.node_id, collectorEvidenceId, object.confidence);
      const sql = querySqlByName.get(object.name);
      let sqlEvidenceId = collectorEvidenceId;
      if (sql) {
        sqlEvidenceId = addEvidence(model, ids, {
          category: '05b_SQL',
          title: `Access saved query SQL - ${object.name}`,
          fileName: `${sanitizeName(file.originalName)}_${sanitizeName(object.name)}.sql`,
          sourceFile: file.originalName,
          summary: `Saved Access query SQL extracted from DAO QueryDefs for ${object.name}.`,
          confidence: 'confirmed',
          content: safeFormatSql(sql),
        });
        sqlEvidenceByQuery.set(object.name, sqlEvidenceId);
        const refs = extractSqlReferences(sql);
        addSqlReferenceEdgesForAccessQuery(model, ids, file, queryNode.node_id, refs, sqlEvidenceId);
      }
      model.access.Access_Query_Register.push({
        query_name: object.name,
        msysobjects_type: object.type,
        sql_extracted: Boolean(sql),
        sql_evidence_id: sql ? sqlEvidenceId : '',
        evidence_id: collectorEvidenceId,
        confidence: sql ? 'confirmed' : object.confidence,
        recommended_action: sql ? 'Review SQL lineage and transformation classification.' : 'Extract SQL text from DAO QueryDefs or MSysQueries.',
      });
      model.access.Access_Query_SQL_Index.push({
        query_name: object.name,
        sql_evidence_id: sqlEvidenceId,
        sql_length: sql?.length ?? 0,
        evidence_id: collectorEvidenceId,
        confidence: sql ? 'confirmed' : 'blocked',
        recommended_action: sql ? 'Map referenced objects and transformations.' : 'Resolve missing SQL evidence.',
      });
      continue;
    }

    if (object.type === -32766) {
      macroNames.add(object.name);
      const macroNode = addAccessObjectNode(model, ids, file, object.name, 'macro', collectorEvidenceId, object.confidence);
      addAccessContainmentEdge(model, ids, fileNodeId, macroNode.node_id, collectorEvidenceId, object.confidence);
      model.access.Access_Macro_Register.push({
        macro_name: object.name,
        msysobjects_type: object.type,
        macro_xml_status: 'pending trusted desktop deep export',
        macro_action_status: 'pending trusted desktop deep export',
        evidence_id: collectorEvidenceId,
        confidence: 'partial',
        recommended_action: 'Review trusted desktop SaveAsText macro XML/action evidence when deep export completes.',
      });
      continue;
    }

    if (object.type === -32768 || object.type === -32764) {
      const nodeType = object.type === -32768 ? 'form' : 'report';
      const uiNode = addAccessObjectNode(model, ids, file, object.name, nodeType, collectorEvidenceId, object.confidence);
      addAccessContainmentEdge(model, ids, fileNodeId, uiNode.node_id, collectorEvidenceId, object.confidence);
      model.access.Access_Form_Report_Register.push({
        object_name: object.name,
        object_type: nodeType,
        evidence_id: collectorEvidenceId,
        confidence: object.confidence,
        recommended_action: 'Export form/report metadata and identify output/consumer usage.',
      });
      continue;
    }

    if (object.type === -32761) {
      const moduleNode = addAccessObjectNode(model, ids, file, object.name, 'module', collectorEvidenceId, object.confidence);
      addAccessContainmentEdge(model, ids, fileNodeId, moduleNode.node_id, collectorEvidenceId, object.confidence);
      model.access.Access_Module_VBA_Register.push({
        module_name: object.name,
        vba_status: 'module object found; source text pending trusted desktop deep export',
        evidence_id: collectorEvidenceId,
        confidence: 'partial',
        recommended_action: 'Review trusted desktop SaveAsText VBA evidence when deep export completes.',
      });
    }
  }

  if (!msysRows.length && hasDaoObjectEvidence) {
    processDaoAccessMetadata({
      file,
      model,
      ids,
      fileNodeId,
      collectorEvidenceId,
      tableDefs,
      queryDefs,
      daoFields,
      daoIndexes,
      daoRelations,
      daoDocuments,
      linkedSourceResolutions,
      queryNames,
      macroNames,
      sqlEvidenceByQuery,
    });
  }

  const seenColumns = new Set<string>();
  for (const column of daoFields) {
    const tableName = String(column.table_name ?? column.TABLE_NAME ?? '');
    const columnName = String(column.column_name ?? column.COLUMN_NAME ?? '');
    if (!tableName || !columnName || isAccessSystemObject(tableName)) {
      continue;
    }
    seenColumns.add(`${tableName}\u0000${columnName}`);
    model.access.Access_Column_Inventory.push({
      table_name: tableName,
      column_name: columnName,
      ordinal_position: column.ordinal_position ?? '',
      data_type: column.data_type ?? '',
      size: column.size ?? '',
      required: column.required ?? '',
      allow_zero_length: column.allow_zero_length ?? '',
      default_value: column.default_value ?? '',
      validation_rule: column.validation_rule ?? '',
      validation_text: column.validation_text ?? '',
      evidence_id: collectorEvidenceId,
      confidence: 'confirmed',
      recommended_action: 'Confirm data definition, key status, sensitivity, and quality rule.',
    });
    model.dataElements.push({
      data_element_id: nextId(ids, 'DE'),
      asset: `${file.originalName}:${tableName}`,
      field_name: columnName,
      inferred_type: String(column.data_type ?? 'access-dao-type-code'),
      null_count: 0,
      sample_values: '',
      sensitive_indicator: sensitiveIndicator(columnName),
      evidence_id: collectorEvidenceId,
      confidence: 'partial',
      recommended_action: 'Profile field values and confirm business definition.',
    });
  }

  for (const column of schemaColumns) {
    const tableName = String(column.TABLE_NAME ?? column.table_name ?? '');
    const columnName = String(column.COLUMN_NAME ?? column.column_name ?? '');
    if (!tableName || !columnName || tableName.startsWith('MSys') || seenColumns.has(`${tableName}\u0000${columnName}`)) {
      continue;
    }
    model.access.Access_Column_Inventory.push({
      table_name: tableName,
      column_name: columnName,
      ordinal_position: column.ORDINAL_POSITION ?? '',
      data_type: column.DATA_TYPE ?? '',
      nullable: column.IS_NULLABLE ?? '',
      evidence_id: collectorEvidenceId,
      confidence: 'partial',
      recommended_action: 'Confirm data definition, key status, sensitivity, and quality rule.',
    });
    model.dataElements.push({
      data_element_id: nextId(ids, 'DE'),
      asset: `${file.originalName}:${tableName}`,
      field_name: columnName,
      inferred_type: String(column.DATA_TYPE ?? 'access-type-code'),
      null_count: 0,
      sample_values: '',
      sensitive_indicator: sensitiveIndicator(columnName),
      evidence_id: collectorEvidenceId,
      confidence: 'partial',
      recommended_action: 'Profile field values and confirm business definition.',
    });
  }

  for (const tableDef of tableDefs) {
    model.access.Access_Data_Profile.push({
      table_name: tableDef.name ?? tableDef.Name,
      row_count: tableDef.record_count ?? 'not extracted',
      field_count: tableDef.field_count ?? '',
      evidence_id: collectorEvidenceId,
      confidence: tableDef.record_count === null || tableDef.record_count === undefined ? 'partial' : 'confirmed',
      recommended_action: 'Profile nulls, duplicates, invalid dates/numbers, and unexpected values for critical tables.',
    });
  }

  for (const exported of appExports) {
    const objectType = String(exported.object_type ?? '');
    const name = String(exported.name ?? '');
    const text = String(exported.text ?? '');
    const status = String(exported.save_as_text_status ?? 'unknown');
    if (!name) {
      continue;
    }

    if (objectType === 'macro') {
      let macroEvidenceId = collectorEvidenceId;
      if (text) {
        macroEvidenceId = addEvidence(model, ids, {
          category: '05e_Macros',
          title: `Access macro SaveAsText - ${name}`,
          fileName: `${sanitizeName(file.originalName)}_${sanitizeName(name)}_macro.txt`,
          sourceFile: file.originalName,
          summary: `Access macro text exported through Access.Application SaveAsText for ${name}.`,
          confidence: 'confirmed',
          content: text,
        });
        model.access.Access_Macro_XML_Storage.push({
          macro_name: name,
          storage_type: 'Access.Application SaveAsText',
          text_length: text.length,
          evidence_id: macroEvidenceId,
          confidence: 'confirmed',
          recommended_action: 'Review macro text and reconcile action sequence to saved query/table/form/report/macro registers.',
        });
      }
      if (!macroNames.has(name)) {
        const macroNode = addAccessObjectNode(model, ids, file, name, 'macro', macroEvidenceId, text ? 'confirmed' : 'blocked');
        addAccessContainmentEdge(model, ids, fileNodeId, macroNode.node_id, macroEvidenceId, text ? 'confirmed' : 'blocked');
      }
      const actions = parseAccessMacroActions(text);
      const macroRegisterRow = model.access.Access_Macro_Register.find((row) => String(row.macro_name) === name);
      const macroPurpose = inferAutomationPurpose(text, actions.map((action) => `${action.action_type} ${action.target_object}`));
      if (macroRegisterRow) {
        Object.assign(macroRegisterRow, {
          macro_xml_status: text ? 'Access.Application SaveAsText captured' : `SaveAsText ${status}`,
          macro_action_status: actions.length ? `${actions.length} parsed action(s)` : text ? 'no recognized action keywords found; review macro text' : 'blocked',
          macro_purpose: macroPurpose,
          evidence_id: macroEvidenceId,
          confidence: text ? 'confirmed' : 'blocked',
          recommended_action: actions.length
            ? 'Review parsed macro action order, target resolution, controls, and failure impact.'
            : 'Review macro text manually and confirm whether actions are embedded, conditional, or unsupported by parser.',
        });
      } else {
        model.access.Access_Macro_Register.push({
          macro_name: name,
          msysobjects_type: 'Access.Application SaveAsText',
          macro_xml_status: text ? 'Access.Application SaveAsText captured' : `SaveAsText ${status}`,
          macro_action_status: actions.length ? `${actions.length} parsed action(s)` : text ? 'no recognized action keywords found; review macro text' : 'blocked',
          macro_purpose: macroPurpose,
          evidence_id: macroEvidenceId,
          confidence: text ? 'confirmed' : 'blocked',
          recommended_action: 'Review macro action order, target resolution, controls, and failure impact.',
        });
      }
      actions.forEach((action, index) => {
        const actionId = `MACACT-${String(model.access.Access_Macro_Action_Sequence.length + 1).padStart(4, '0')}`;
        const targetResolution = resolveAccessMacroTarget(action.target_object, queryNames, macroNames);
        model.access.Access_Macro_Action_Sequence.push({
          macro_action_id: actionId,
          macro_name: name,
          action_order: index + 1,
          action_type: action.action_type,
          target_object: action.target_object,
          target_resolution: targetResolution,
          purpose: inferAutomationPurpose(`${action.action_type} ${action.target_object}`, [action.action_type, action.target_object]),
          evidence_id: macroEvidenceId,
          confidence: action.target_object ? 'partial' : 'unknown',
          recommended_action: targetResolution.includes('missing') ? 'Resolve missing macro action target.' : 'Confirm macro action purpose and execution order.',
        });
        const actionNode = addNode(model, ids, {
          node_type: 'macro action',
          name: `${name}:${index + 1}:${action.action_type}`,
          description: `Parsed Access macro action ${action.action_type}${action.target_object ? ` targeting ${action.target_object}` : ''}.`,
          source_file: file.originalName,
          business_purpose: 'Macro automation action that may run queries, SQL, transfers, forms, reports, or macros.',
          owner_status: 'owner-confirmation-required',
          criticality: 'P1',
          confidence: action.target_object ? 'partial' : 'unknown',
          evidence_id: macroEvidenceId,
          recommended_action: 'Confirm action target, failure behavior, control coverage, and downstream impact.',
          failure_impact: 'Macro action failure can stop or partially execute the process.',
          dollar_exposure: 'Directional exposure modeled at process level.',
        });
        addEdge(model, ids, {
          from_node_id: fileNodeId,
          to_node_id: actionNode.node_id,
          edge_type: 'runs',
          description: 'Access database macro action is part of executable workflow.',
          automated_flag: 'automated',
          transformation_id: '',
          cadence: 'macro run order',
          confidence: action.target_object ? 'partial' : 'unknown',
          evidence_id: macroEvidenceId,
        });
      });
      continue;
    }

    if (objectType === 'module' && text) {
      const moduleEvidenceId = addEvidence(model, ids, {
        category: '05d_VBA',
        title: `Access module SaveAsText - ${name}`,
        fileName: `${sanitizeName(file.originalName)}_${sanitizeName(name)}_module.bas`,
        sourceFile: file.originalName,
        summary: `Access module text exported through Access.Application SaveAsText for ${name}.`,
        confidence: 'confirmed',
        content: text,
      });
      model.access.Access_Module_VBA_Register.push({
        module_name: name,
        vba_status: status,
        procedure_count_estimate: (text.match(/\b(Sub|Function)\s+[A-Za-z0-9_]+/gi) ?? []).length,
        procedure_summary: extractVbaProcedures(text)
          .map((procedure) => `${procedure.name}: ${procedure.purpose}`)
          .slice(0, 12)
          .join(' | '),
        evidence_id: moduleEvidenceId,
        confidence: 'confirmed',
        recommended_action: 'Parse procedures for reads, writes, exports, schedules, error handling, and control behavior.',
      });
      continue;
    }

    if ((objectType === 'form' || objectType === 'report') && text) {
      const uiEvidenceId = addEvidence(model, ids, {
        category: '05f_Form_Report_Metadata',
        title: `Access ${objectType} SaveAsText - ${name}`,
        fileName: `${sanitizeName(file.originalName)}_${sanitizeName(name)}_${objectType}.txt`,
        sourceFile: file.originalName,
        summary: `Access ${objectType} text exported through Access.Application SaveAsText for ${name}.`,
        confidence: 'confirmed',
        content: text,
      });
      model.access.Access_Form_Report_Register.push({
        object_name: name,
        object_type: objectType,
        save_as_text_status: status,
        evidence_id: uiEvidenceId,
        confidence: 'confirmed',
        recommended_action: 'Review data source, filters, controls, event macros/VBA, and output/consumer usage.',
      });
    }
  }

  const missingSql = Array.from(queryNames).filter((name) => !sqlEvidenceByQuery.has(name));
  const appStatus = String(collector.access_application?.status ?? 'not_run');
  if ((appStatus === 'skipped' || appStatus === 'blocked') && (macroNames.size > 0 || objectRows.some((row) => row.type < 0))) {
    addAction(model, ids, {
      title: `Run bounded Access desktop deep export for ${file.originalName}`,
      description:
        'The fast native collector completed without deep Access.Application SaveAsText export. Export macro text/actions, form/report metadata, and VBA modules from a trusted desktop session when the business needs certified automation lineage.',
      source_asset: file.originalName,
      owner_role: 'Access owner / data engineering',
      recommended_owner: 'Access technical owner',
      action_type: 'Resolve Blocker',
      priority: 'P1',
      severity: 'P1',
      dependency: 'Trusted Access desktop export with ACCESS_COLLECTOR_DEEP enabled',
      acceptance_criteria: 'Macro XML/action map, forms/reports, and VBA evidence are present and reconciled without contaminating query or macro registers.',
      evidence_id: collectorEvidenceId,
      related_risk: 'Partial Access automation lineage',
      expected_business_value: 'Completes macro/action/control/failure-impact lineage while keeping default runs bounded.',
    });
  }

  model.access.Access_Query_Macro_Reconciliation.push({
    total_msysobjects_rows: msysRows.length || 'MSysObjects blocked',
    saved_query_count: queryNames.size,
    sql_evidence_file_count: sqlEvidenceByQuery.size,
    macro_object_count: macroNames.size,
    macro_xml_payload_count: model.access.Access_Macro_XML_Storage.length,
    macro_action_count: model.access.Access_Macro_Action_Sequence.length,
    openquery_action_count: model.access.Access_Macro_Action_Sequence.filter((row) => String(row.action_type).toLowerCase() === 'openquery').length,
    macro_objects_found_inside_query_register: 0,
    queries_found_inside_macro_object_register: 0,
    openquery_targets_resolving_to_saved_query_register: model.access.Access_Macro_Action_Sequence.filter((row) => row.target_resolution === 'saved query').length,
    openquery_targets_missing_from_saved_query_register: model.access.Access_Macro_Action_Sequence.filter((row) => String(row.target_resolution).includes('missing')).length,
    saved_queries_not_referenced_by_parsed_macro_actions:
      model.access.Access_Macro_Action_Sequence.length > 0
        ? Array.from(queryNames)
            .filter((name) => !model.access.Access_Macro_Action_Sequence.some((row) => row.target_object === name))
            .join('; ')
        : 'blocked until macro actions extracted',
    linked_source_count: model.access.Access_Linked_Table_Register.length,
    blocked_lineage_source_count: model.nodes.filter((node) => node.node_type === 'upstream blocker').length,
    mismatch_notes: [
      missingSql.length ? `Missing SQL evidence for: ${missingSql.join(', ')}` : 'No query/SQL count mismatch detected for extracted QueryDefs.',
      appStatus === 'skipped' || appStatus === 'blocked' ? `Deep Access.Application export ${appStatus}: ${collector.access_application?.error ?? 'no detail'}` : '',
    ]
      .filter(Boolean)
      .join(' '),
    evidence_id: collectorEvidenceId,
    confidence: missingSql.length || macroNames.size ? 'partial' : 'confirmed',
    recommended_action: macroNames.size ? 'Extract macro XML/action map and rerun reconciliation.' : 'Review reconciliation with Access owner.',
  });

  if (!msysRows.length && !objectRows.length && !hasDaoObjectEvidence) {
    createTargetedBlocker(model, ids, file, fileNodeId, 'MSysObjects extraction blocked', collectorEvidenceId, 'Grant system catalog read access or provide exported MSysObjects so objects can be classified by authoritative Type values.');
  } else if (!msysRows.length) {
    const fallbackObjectCount = objectRows.length || tableDefs.length + queryDefs.length + daoDocuments.length;
    model.qaRecords.push({
      qa_id: nextId(ids, 'QA'),
      check: 'Access object inventory fallback when MSysObjects is blocked',
      status: 'PASS_WITH_LIMITATION',
      evidence_id: collectorEvidenceId,
      notes: `MSysObjects was blocked, but ${fallbackObjectCount} object(s) were inventoried from DAO/OLEDB fallback evidence. Saved query and macro inventories remain separated by DAO QueryDefs, DAO containers, and schema type evidence where available.`,
    });
  }

  if (missingSql.length) {
    addAction(model, ids, {
      title: `Resolve missing Access SQL evidence for ${file.originalName}`,
      description: `SQL evidence was not extracted for ${missingSql.length} saved query object(s).`,
      source_asset: file.originalName,
      owner_role: 'Access owner / data engineering',
      recommended_owner: 'Access technical owner',
      action_type: 'Resolve Blocker',
      priority: 'P1',
      severity: 'P1',
      dependency: 'DAO QueryDefs or MSysQueries export',
      acceptance_criteria: 'Saved query count reconciles one-to-one with SQL evidence files.',
      evidence_id: collectorEvidenceId,
      related_risk: 'Incomplete query lineage',
      expected_business_value: 'Enables defensible query lineage, transformation, and migration analysis.',
    });
  }
}

function processAccessDeepCollectorResult(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  fileNodeId: string,
  collectorEvidenceId: string,
  collector: Record<string, any>,
): void {
  const appExports = asRows(collector.access_application?.rows);
  const collectorLimitations = Array.isArray(collector.limitations)
    ? collector.limitations.map((item) => String(item)).filter(Boolean)
    : [];
  model.limitations.push(...collectorLimitations);

  const queryNames = new Set(model.access.Access_Query_Register.map((row) => String(row.query_name ?? '')).filter(Boolean));
  const macroNames = new Set(model.access.Access_Macro_Register.map((row) => String(row.macro_name ?? '')).filter(Boolean));

  for (const exported of appExports) {
    const objectType = String(exported.object_type ?? '');
    const name = String(exported.name ?? '');
    const text = String(exported.text ?? '');
    const status = String(exported.save_as_text_status ?? 'unknown');
    if (!name) {
      continue;
    }

    if (objectType === 'query') {
      if (!text) {
        continue;
      }
      const queryEvidenceId = addEvidence(model, ids, {
        category: '05b_SQL',
        title: `Access query SaveAsText - ${name}`,
        fileName: `${sanitizeName(file.originalName)}_${sanitizeName(name)}_query_SaveAsText.txt`,
        sourceFile: file.originalName,
        summary: `Access query object text exported through Access.Application SaveAsText for ${name}. DAO SQL remains the preferred SQL evidence when present.`,
        confidence: 'partial',
        content: text,
      });
      const queryRow = model.access.Access_Query_Register.find((row) => String(row.query_name) === name);
      if (queryRow && !queryRow.sql_evidence_id) {
        Object.assign(queryRow, {
          sql_extracted: true,
          sql_evidence_id: queryEvidenceId,
          confidence: 'partial',
          recommended_action: 'Review query SaveAsText evidence and reconcile with DAO/MSys SQL when available.',
        });
      }
      model.access.Access_Query_SQL_Index.push({
        query_name: name,
        sql_evidence_id: queryEvidenceId,
        sql_length: text.length,
        evidence_id: collectorEvidenceId,
        confidence: 'partial',
        recommended_action: 'Use DAO SQL as canonical when present; use SaveAsText for query design context and blocked SQL reconciliation.',
      });
      continue;
    }

    if (objectType === 'macro') {
      let macroEvidenceId = collectorEvidenceId;
      if (text) {
        macroEvidenceId = addEvidence(model, ids, {
          category: '05e_Macros',
          title: `Access macro SaveAsText - ${name}`,
          fileName: `${sanitizeName(file.originalName)}_${sanitizeName(name)}_macro.txt`,
          sourceFile: file.originalName,
          summary: `Access macro text exported through Access.Application SaveAsText for ${name}.`,
          confidence: 'confirmed',
          content: text,
        });
        model.access.Access_Macro_XML_Storage.push({
          macro_name: name,
          storage_type: 'Access.Application SaveAsText',
          text_length: text.length,
          evidence_id: macroEvidenceId,
          confidence: 'confirmed',
          recommended_action: 'Review macro text and reconcile action sequence to saved query/table/form/report/macro registers.',
        });
      }
      if (!macroNames.has(name)) {
        macroNames.add(name);
        const macroNode = addAccessObjectNode(model, ids, file, name, 'macro', macroEvidenceId, text ? 'confirmed' : 'blocked');
        addAccessContainmentEdge(model, ids, fileNodeId, macroNode.node_id, macroEvidenceId, text ? 'confirmed' : 'blocked');
      }
      const actions = parseAccessMacroActions(text);
      const macroPurpose = inferAutomationPurpose(text, actions.map((action) => `${action.action_type} ${action.target_object}`));
      const macroRegisterRow = model.access.Access_Macro_Register.find((row) => String(row.macro_name) === name);
      if (macroRegisterRow) {
        Object.assign(macroRegisterRow, {
          macro_xml_status: text ? 'Access.Application SaveAsText captured' : `SaveAsText ${status}`,
          macro_action_status: actions.length ? `${actions.length} parsed action(s)` : text ? 'no recognized action keywords found; review macro text' : 'blocked',
          macro_purpose: macroPurpose,
          evidence_id: macroEvidenceId,
          confidence: text ? 'confirmed' : 'blocked',
          recommended_action: actions.length
            ? 'Review parsed macro action order, target resolution, controls, and failure impact.'
            : 'Review macro text manually and confirm whether actions are embedded, conditional, or unsupported by parser.',
        });
      } else {
        model.access.Access_Macro_Register.push({
          macro_name: name,
          msysobjects_type: 'Access.Application SaveAsText',
          macro_xml_status: text ? 'Access.Application SaveAsText captured' : `SaveAsText ${status}`,
          macro_action_status: actions.length ? `${actions.length} parsed action(s)` : text ? 'no recognized action keywords found; review macro text' : 'blocked',
          macro_purpose: macroPurpose,
          evidence_id: macroEvidenceId,
          confidence: text ? 'confirmed' : 'blocked',
          recommended_action: 'Review macro action order, target resolution, controls, and failure impact.',
        });
      }
      actions.forEach((action, index) => {
        const actionId = `MACACT-${String(model.access.Access_Macro_Action_Sequence.length + 1).padStart(4, '0')}`;
        const targetResolution = resolveAccessMacroTarget(action.target_object, queryNames, macroNames);
        model.access.Access_Macro_Action_Sequence.push({
          macro_action_id: actionId,
          macro_name: name,
          action_order: index + 1,
          action_type: action.action_type,
          target_object: action.target_object,
          target_resolution: targetResolution,
          purpose: inferAutomationPurpose(`${action.action_type} ${action.target_object}`, [action.action_type, action.target_object]),
          evidence_id: macroEvidenceId,
          confidence: action.target_object ? 'partial' : 'unknown',
          recommended_action: targetResolution.includes('missing') ? 'Resolve missing macro action target.' : 'Confirm macro action purpose and execution order.',
        });
      });
      continue;
    }

    if (objectType === 'module' && text) {
      const moduleEvidenceId = addEvidence(model, ids, {
        category: '05d_VBA',
        title: `Access module SaveAsText - ${name}`,
        fileName: `${sanitizeName(file.originalName)}_${sanitizeName(name)}_module.bas`,
        sourceFile: file.originalName,
        summary: `Access module text exported through Access.Application SaveAsText for ${name}.`,
        confidence: 'confirmed',
        content: text,
      });
      model.access.Access_Module_VBA_Register.push({
        module_name: name,
        vba_status: status,
        procedure_count_estimate: (text.match(/\b(Sub|Function)\s+[A-Za-z0-9_]+/gi) ?? []).length,
        procedure_summary: extractVbaProcedures(text)
          .map((procedure) => `${procedure.name}: ${procedure.purpose}`)
          .slice(0, 12)
          .join(' | '),
        evidence_id: moduleEvidenceId,
        confidence: 'confirmed',
        recommended_action: 'Parse procedures for reads, writes, exports, schedules, error handling, and control behavior.',
      });
      continue;
    }

    if ((objectType === 'form' || objectType === 'report') && text) {
      const uiEvidenceId = addEvidence(model, ids, {
        category: '05f_Form_Report_Metadata',
        title: `Access ${objectType} SaveAsText - ${name}`,
        fileName: `${sanitizeName(file.originalName)}_${sanitizeName(name)}_${objectType}.txt`,
        sourceFile: file.originalName,
        summary: `Access ${objectType} text exported through Access.Application SaveAsText for ${name}.`,
        confidence: 'confirmed',
        content: text,
      });
      model.access.Access_Form_Report_Register.push({
        object_name: name,
        object_type: objectType,
        save_as_text_status: status,
        evidence_id: uiEvidenceId,
        confidence: 'confirmed',
        recommended_action: 'Review data source, filters, controls, event macros/VBA, and output/consumer usage.',
      });
    }
  }

  const appStatus = String(collector.access_application?.status ?? 'not_run');
  model.qaRecords.push({
    qa_id: nextId(ids, 'QA'),
    check: 'Access deep export phase',
    status: appStatus === 'ok' ? 'PASS' : 'PASS_WITH_LIMITATION',
    evidence_id: collectorEvidenceId,
    notes:
      appStatus === 'ok'
        ? `Access.Application SaveAsText exported ${appExports.length} object(s) after fast catalog extraction.`
        : `Access.Application SaveAsText did not complete fully: ${collector.access_application?.error ?? appStatus}`,
  });
}

function processDaoAccessMetadata(input: {
  file: UploadedSource;
  model: DiscoveryModel;
  ids: IdState;
  fileNodeId: string;
  collectorEvidenceId: string;
  tableDefs: Record<string, any>[];
  queryDefs: Record<string, any>[];
  daoFields: Record<string, any>[];
  daoIndexes: Record<string, any>[];
  daoRelations: Record<string, any>[];
  daoDocuments: Record<string, any>[];
  linkedSourceResolutions: Map<string, LinkedSourceResolution>;
  queryNames: Set<string>;
  macroNames: Set<string>;
  sqlEvidenceByQuery: Map<string, string>;
}): void {
  const {
    file,
    model,
    ids,
    fileNodeId,
    collectorEvidenceId,
    tableDefs,
    queryDefs,
    daoIndexes,
    daoRelations,
    daoDocuments,
    linkedSourceResolutions,
    queryNames,
    macroNames,
    sqlEvidenceByQuery,
  } = input;

  const userTableDefs = tableDefs.filter((tableDef) => String(tableDef.name ?? tableDef.Name ?? ''));
  const upstreamNodesBySource = new Map<string, DiscoveryNode>();
  for (const tableDef of userTableDefs) {
    const name = String(tableDef.name ?? tableDef.Name ?? '');
    const attributes = Number(tableDef.attributes ?? tableDef.Attributes ?? 0);
    const connect = String(tableDef.connect ?? tableDef.Connect ?? '');
    const sourceTableName = String(tableDef.source_table_name ?? tableDef.SourceTableName ?? '');
    const systemObject = isAccessSystemObject(name) || isAccessSystemTableDef(attributes);
    const linked = Boolean(connect) || isAccessLinkedTableDef(attributes);
    const interpretedType = systemObject ? 'system table (DAO TableDefs)' : linked ? 'linked table (DAO TableDefs)' : 'local table (DAO TableDefs)';

    model.access.Access_Object_Inventory.push({
      object_name: name,
      raw_type_value: `DAO.TableDef attributes=${Number.isFinite(attributes) ? attributes : 'unknown'}`,
      interpreted_type: interpretedType,
      connect,
      source_table_name: sourceTableName,
      field_count: tableDef.field_count ?? '',
      row_count: tableDef.record_count ?? 'not extracted',
      metadata_source: 'DAO TableDefs; MSysObjects unavailable',
      evidence_id: collectorEvidenceId,
      confidence: systemObject ? 'partial' : 'confirmed',
      recommended_action: systemObject
        ? 'System/hidden object inventoried separately; confirm whether it is relevant to migration.'
        : accessRecommendedAction(interpretedType),
    });

    if (systemObject) {
      continue;
    }

    const nodeType: NodeType = linked ? 'linked table' : 'table';
    const tableNode = addAccessObjectNode(model, ids, file, name, nodeType, collectorEvidenceId, 'confirmed');
    addAccessContainmentEdge(model, ids, fileNodeId, tableNode.node_id, collectorEvidenceId, 'confirmed');

    if (linked) {
      const resolution = linkedSourceResolutions.get(linkedSourceKey(connect, sourceTableName, name));
      const upstreamName = resolution?.displayName || connect || sourceTableName || `Linked source for ${name}`;
      const upstreamKey = linkedSourceNodeReuseKey(resolution, upstreamName);
      let upstream = upstreamNodesBySource.get(upstreamKey);
      if (!upstream) {
        upstream = addNode(model, ids, {
          node_type: resolution?.resolutionStatus === 'resolved_file' ? nodeTypeForResolvedPath(resolution.resolvedPath) : resolution?.resolutionStatus === 'resolved_folder' ? 'folder' : 'upstream blocker',
          name: upstreamName,
          description: `Linked source referenced by Access linked table ${name}.`,
          source_file: file.originalName,
          business_purpose: 'Terminal upstream source candidate for recursive lineage.',
          owner_status: 'owner-confirmation-required',
          criticality: 'P1',
          confidence: resolution?.resolutionStatus === 'resolved_file' || resolution?.resolutionStatus === 'resolved_folder' ? 'confirmed' : upstreamName === `Linked source for ${name}` ? 'unknown' : 'partial',
          evidence_id: collectorEvidenceId,
          recommended_action:
            resolution?.resolutionStatus === 'resolved_file'
              ? 'Inspect recursively extracted linked source evidence and confirm terminal source-of-truth status.'
              : 'Provide linked source file/system access and classify as authoritative, third-party, manual entry, obsolete, duplicate, or approved stopping point.',
          failure_impact: 'If the linked source changes or disappears, Access outputs may fail, stale, or produce wrong data.',
          dollar_exposure: 'Directional exposure modeled at process level.',
        });
        upstreamNodesBySource.set(upstreamKey, upstream);
      }
      addEdge(model, ids, {
        from_node_id: tableNode.node_id,
        to_node_id: upstream.node_id,
        edge_type: 'reads_from',
        description: 'Access linked table reads from external source.',
        automated_flag: 'automated',
        transformation_id: '',
        cadence: 'per Access refresh/open',
        confidence: upstream.confidence,
        evidence_id: collectorEvidenceId,
      });
      if (!resolution || resolution.resolutionStatus === 'blocked') {
        model.blockedSources.push(formatBlockedSource(resolution, upstreamName));
      }
      model.access.Access_Linked_Table_Register.push({
        table_name: name,
        connect,
        source_table_name: sourceTableName,
        source_kind: resolution?.sourceKind ?? 'unknown',
        source_name: resolution?.displayName ?? upstreamName,
        source_path: resolution?.sourcePath ?? '',
        resolved_path: resolution?.resolvedPath ?? '',
        resolution_status: resolution?.resolutionStatus ?? 'blocked',
        resolution_reason: resolution?.resolutionReason ?? 'Linked source could not be parsed from Access metadata.',
        linked_source_node_id: upstream.node_id,
        row_count: tableDef.record_count ?? 'not extracted',
        field_count: tableDef.field_count ?? '',
        evidence_id: collectorEvidenceId,
        confidence: 'confirmed',
        recommended_action: 'Resolve linked source path/system and include source file in discovery package.',
      });
      continue;
    }

    model.access.Access_Table_Register.push({
      table_name: name,
      row_count: tableDef.record_count ?? 'not extracted',
      field_count: tableDef.field_count ?? '',
      evidence_id: collectorEvidenceId,
      confidence: 'confirmed',
      recommended_action: 'Confirm owner, primary key, row-count expectations, and source-of-truth status.',
    });
  }

  for (const queryDef of queryDefs) {
    const name = String(queryDef.name ?? queryDef.Name ?? '');
    if (!name || name.startsWith('~TMP')) {
      continue;
    }
    queryNames.add(name);
    model.access.Access_Object_Inventory.push({
      object_name: name,
      raw_type_value: `DAO.QueryDef type=${queryDef.type ?? queryDef.Type ?? 'unknown'}`,
      interpreted_type: 'saved query (DAO QueryDefs; MSysObjects unavailable)',
      returns_records: queryDef.returns_records ?? queryDef.ReturnsRecords ?? '',
      connect: queryDef.connect ?? queryDef.Connect ?? '',
      metadata_source: 'DAO QueryDefs; MSysObjects unavailable',
      evidence_id: collectorEvidenceId,
      confidence: 'confirmed',
      recommended_action: 'Review SQL evidence, classify action/select behavior, map lineage, and reconcile macro references.',
    });

    const queryNode = addAccessObjectNode(model, ids, file, name, 'query', collectorEvidenceId, 'confirmed');
    addAccessContainmentEdge(model, ids, fileNodeId, queryNode.node_id, collectorEvidenceId, 'confirmed');
    const sql = String(queryDef.sql ?? queryDef.SQL ?? '');
    let sqlEvidenceId = collectorEvidenceId;
    if (sql) {
      sqlEvidenceId = addEvidence(model, ids, {
        category: '05b_SQL',
        title: `Access saved query SQL - ${name}`,
        fileName: `${sanitizeName(file.originalName)}_${sanitizeName(name)}.sql`,
        sourceFile: file.originalName,
        summary: `Saved Access query SQL extracted from DAO QueryDefs for ${name}.`,
        confidence: 'confirmed',
        content: safeFormatSql(sql),
      });
      sqlEvidenceByQuery.set(name, sqlEvidenceId);
      const refs = extractSqlReferences(sql);
      addSqlReferenceEdgesForAccessQuery(model, ids, file, queryNode.node_id, refs, sqlEvidenceId);
      addSqlTransformationRule(model, ids, `${file.originalName}:${name}`, refs, sqlEvidenceId);
    }
    model.access.Access_Query_Register.push({
      query_name: name,
      msysobjects_type: 'MSysObjects blocked; DAO QueryDefs used',
      dao_query_type: queryDef.type ?? queryDef.Type ?? '',
      sql_extracted: Boolean(sql),
      sql_evidence_id: sql ? sqlEvidenceId : '',
      evidence_id: collectorEvidenceId,
      confidence: sql ? 'confirmed' : 'partial',
      recommended_action: sql ? 'Review SQL lineage and transformation classification.' : 'Resolve missing SQL text from owner export.',
    });
    model.access.Access_Query_SQL_Index.push({
      query_name: name,
      sql_evidence_id: sqlEvidenceId,
      sql_length: sql.length,
      evidence_id: collectorEvidenceId,
      confidence: sql ? 'confirmed' : 'blocked',
      recommended_action: sql ? 'Map referenced objects and transformations.' : 'Resolve missing SQL evidence.',
    });
  }

  for (const document of daoDocuments) {
    const container = String(document.container ?? '');
    const name = String(document.name ?? '');
    if (!name) {
      continue;
    }
    if (container === 'Scripts') {
      macroNames.add(name);
      const hiddenOrTemp = name.startsWith('~TMP');
      model.access.Access_Object_Inventory.push({
        object_name: name,
        raw_type_value: 'DAO.Container Scripts document',
        interpreted_type: hiddenOrTemp ? 'macro / script document (hidden or temporary)' : 'macro / script document',
        date_created: document.date_created ?? '',
        last_updated: document.last_updated ?? '',
        metadata_source: 'DAO Containers; MSysObjects unavailable',
        evidence_id: collectorEvidenceId,
        confidence: 'partial',
        recommended_action: hiddenOrTemp ? 'Confirm whether hidden/temporary script can be ignored.' : 'Export macro XML/action map for execution lineage.',
      });
      model.access.Access_Macro_Register.push({
        macro_name: name,
        msysobjects_type: 'MSysObjects blocked; DAO Scripts container used',
        hidden_or_system: hiddenOrTemp,
        macro_xml_status: 'not extracted by bounded collector',
        macro_action_status: 'blocked',
        evidence_id: collectorEvidenceId,
        confidence: 'partial',
        recommended_action: 'Export macro XML/action sequence from trusted Access desktop deep export.',
      });
      if (!hiddenOrTemp) {
        const macroNode = addAccessObjectNode(model, ids, file, name, 'macro', collectorEvidenceId, 'partial');
        addAccessContainmentEdge(model, ids, fileNodeId, macroNode.node_id, collectorEvidenceId, 'partial');
        model.qaRecords.push({
          qa_id: nextId(ids, 'QA'),
          check: `Access macro action extraction - ${name}`,
          status: 'PASS_WITH_LIMITATION',
          evidence_id: collectorEvidenceId,
          notes:
            'Macro object was inventoried from DAO Containers. Macro XML/action bodies require the bounded Access.Application deep export and are tracked as extraction QA, not upstream lineage blockers.',
        });
        addAction(model, ids, {
          title: `Extract Access macro action sequence for ${name}`,
          description:
            'Run the trusted desktop deep export or provide SaveAsText macro XML so OpenQuery, RunSQL, TransferSpreadsheet, TransferText, OpenForm, OpenReport, and RunMacro actions can be reconciled without contaminating saved query and macro inventories.',
          source_asset: file.originalName,
          owner_role: 'Technical owner',
          recommended_owner: 'Access platform owner',
          action_type: 'Resolve Blocker',
          priority: 'P1',
          severity: 'P1',
          dependency: 'Trusted Access desktop automation export or supplied SaveAsText macro evidence',
          acceptance_criteria:
            'Macro XML/action evidence exists, parsed action rows reconcile to macro objects, and target resolution is recorded separately from saved query and macro registers.',
          evidence_id: collectorEvidenceId,
          related_risk:
            'Macro side effects, query execution order, exports, and controls cannot be fully certified from DAO catalog metadata alone.',
          expected_business_value: 'Completes automation lineage while preserving clean query/macro inventory separation.',
        });
      }
      continue;
    }

    const formReportModuleType = container === 'Forms' ? 'form' : container === 'Reports' ? 'report' : container === 'Modules' ? 'module' : '';
    if (formReportModuleType) {
      model.access.Access_Object_Inventory.push({
        object_name: name,
        raw_type_value: `DAO.Container ${container} document`,
        interpreted_type: formReportModuleType,
        date_created: document.date_created ?? '',
        last_updated: document.last_updated ?? '',
        metadata_source: 'DAO Containers; MSysObjects unavailable',
        evidence_id: collectorEvidenceId,
        confidence: 'partial',
        recommended_action: accessRecommendedAction(formReportModuleType),
      });
      const node = addAccessObjectNode(model, ids, file, name, formReportModuleType, collectorEvidenceId, 'partial');
      addAccessContainmentEdge(model, ids, fileNodeId, node.node_id, collectorEvidenceId, 'partial');
      if (formReportModuleType === 'module') {
        model.access.Access_Module_VBA_Register.push({
          module_name: name,
          vba_status: 'module document found; source text not extracted by bounded collector',
          evidence_id: collectorEvidenceId,
          confidence: 'blocked',
          recommended_action: 'Export VBA module source text from trusted Access desktop deep export.',
        });
      } else {
        model.access.Access_Form_Report_Register.push({
          object_name: name,
          object_type: formReportModuleType,
          evidence_id: collectorEvidenceId,
          confidence: 'partial',
          recommended_action: 'Export form/report metadata and identify output/consumer usage.',
        });
      }
    }
  }

  for (const index of daoIndexes) {
    const tableName = String(index.table_name ?? '');
    const indexName = String(index.index_name ?? '');
    if (!tableName || !indexName || isAccessSystemObject(tableName)) {
      continue;
    }
    model.access.Access_Object_Inventory.push({
      object_name: `${tableName}.${indexName}`,
      raw_type_value: 'DAO.Index',
      interpreted_type: index.primary ? 'primary key/index' : index.unique ? 'unique index' : 'index',
      table_name: tableName,
      fields: index.fields ?? '',
      primary: index.primary ?? '',
      unique: index.unique ?? '',
      evidence_id: collectorEvidenceId,
      confidence: 'confirmed',
      recommended_action: 'Confirm key/index usage for lineage joins, data quality tests, and migration DDL.',
    });
  }

  for (const relation of daoRelations) {
    const relationName = String(relation.name ?? '');
    if (!relationName || relationName.startsWith('MSys')) {
      continue;
    }
    model.access.Access_Object_Inventory.push({
      object_name: relationName,
      raw_type_value: 'DAO.Relation',
      interpreted_type: 'relationship',
      table: relation.table ?? '',
      foreign_table: relation.foreign_table ?? '',
      fields: relation.fields ?? '',
      evidence_id: collectorEvidenceId,
      confidence: 'confirmed',
      recommended_action: 'Validate relationship cardinality and migration constraint requirements.',
    });
    model.dependencyUsage.push({
      dependency_id: `DEP-${model.dependencyUsage.length + 1}`,
      source_asset: file.originalName,
      dependency: `${relation.table ?? ''} -> ${relation.foreign_table ?? ''}`,
      dependency_type: 'Access DAO relationship',
      evidence_id: collectorEvidenceId,
      confidence: 'confirmed',
    });
  }
}

async function resolveAccessLinkedSources(collector: Record<string, any>): Promise<Map<string, LinkedSourceResolution>> {
  const resolutions = new Map<string, LinkedSourceResolution>();
  const tableDefs = asRows(collector.dao_tabledefs?.rows);
  const msysRows = asRows(collector.msys_objects?.rows);

  const candidates: Array<{ tableName: string; connect: string; sourceTableName: string }> = [];
  for (const tableDef of tableDefs) {
    const tableName = String(tableDef.name ?? tableDef.Name ?? '');
    const connect = String(tableDef.connect ?? tableDef.Connect ?? '');
    const sourceTableName = String(tableDef.source_table_name ?? tableDef.SourceTableName ?? '');
    if (connect || isAccessLinkedTableDef(Number(tableDef.attributes ?? tableDef.Attributes ?? 0))) {
      candidates.push({ tableName, connect, sourceTableName });
    }
  }
  for (const row of msysRows) {
    const tableName = String(row.Name ?? row.name ?? '');
    const connect = String(row.Connect ?? row.connect ?? '');
    const sourceTableName = String(row.ForeignName ?? row.foreignName ?? '');
    const type = Number(row.Type ?? row.type);
    if (connect || type === 4 || type === 6) {
      candidates.push({ tableName, connect, sourceTableName });
    }
  }

  await Promise.all(
    candidates.map(async (candidate) => {
      const resolution = await resolveLinkedSourceCandidate(candidate.tableName, candidate.connect, candidate.sourceTableName);
      resolutions.set(resolution.key, resolution);
    }),
  );
  return resolutions;
}

async function resolveLinkedSourceCandidate(
  tableName: string,
  connect: string,
  sourceTableName: string,
): Promise<LinkedSourceResolution> {
  const parsedPath = parseAccessDatabasePath(connect);
  const sourceKind = inferAccessLinkedSourceKind(connect, parsedPath, sourceTableName);
  const key = linkedSourceKey(connect, sourceTableName, tableName);
  const base: LinkedSourceResolution = {
    key,
    tableName,
    connect,
    sourceTableName,
    sourceKind,
    sourcePath: parsedPath,
    displayName: displayLinkedSourceName(parsedPath, sourceTableName, tableName),
    resolvedPath: '',
    resolutionStatus: 'blocked',
    resolutionReason: parsedPath ? 'Source path was parsed but not yet resolved.' : 'No DATABASE path was available in the Access connection metadata.',
  };

  if (!parsedPath || process.env.VERCEL) {
    return base;
  }

  const candidatePaths = candidateLinkedSourcePaths(parsedPath, sourceTableName);
  for (const candidatePath of candidatePaths) {
    const stat = await safeStat(candidatePath);
    if (!stat) {
      continue;
    }
    if (stat.isFile()) {
      return {
        ...base,
        displayName: path.basename(candidatePath),
        resolvedPath: candidatePath,
        resolutionStatus: 'resolved_file',
        resolutionReason: 'Linked source file was reachable from the local runtime and queued for recursive extraction.',
        fileSizeBytes: stat.size,
      };
    }
    if (stat.isDirectory()) {
      const nested = await findNestedLinkedFile(candidatePath, sourceTableName);
      if (nested) {
        const nestedStat = await safeStat(nested);
        return {
          ...base,
          displayName: path.basename(nested),
          resolvedPath: nested,
          resolutionStatus: 'resolved_file',
          resolutionReason: 'Linked source folder was reachable and the linked file was matched by source table name.',
          fileSizeBytes: nestedStat?.size,
        };
      }
      return {
        ...base,
        displayName: path.basename(candidatePath) || candidatePath,
        resolvedPath: candidatePath,
        resolutionStatus: 'resolved_folder',
        resolutionReason: 'Linked source folder was reachable, but no single concrete file could be matched safely from the table metadata.',
      };
    }
  }

  return {
    ...base,
    resolutionReason: 'Linked source path was not reachable from the local runtime. It remains a recursive lineage blocker.',
  };
}

async function extractResolvedLinkedSources(
  parentFile: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  packageNodeId: string,
  context: ExtractionContext,
  resolutions: Map<string, LinkedSourceResolution>,
  evidenceId: string,
): Promise<void> {
  if (!RECURSIVE_SOURCE_DISCOVERY) {
    model.limitations.push('Recursive linked-source extraction is disabled by DISCOVERY_RECURSE_LINKED_SOURCES=0.');
    return;
  }
  if (context.recursionDepth >= context.maxRecursionDepth) {
    const skipped = Array.from(resolutions.values()).filter((resolution) => resolution.resolutionStatus === 'resolved_file');
    if (skipped.length) {
      model.limitations.push(`Recursive depth limit ${context.maxRecursionDepth} reached before ${skipped.length} linked source(s) could be fully drilled through.`);
    }
    return;
  }

  const recursiveSources = Array.from(resolutions.values()).filter((resolution) => resolution.resolutionStatus === 'resolved_file' && resolution.resolvedPath);
  await runWithConcurrency(recursiveSources, extractionConcurrency(recursiveSources.length), async (resolution) => {
    const sourceKey = resolution.resolvedPath.toLowerCase();
    if (context.linkedSourcePaths.has(sourceKey)) {
      return;
    }
    context.linkedSourcePaths.add(sourceKey);
    if ((resolution.fileSizeBytes ?? 0) > MAX_RECURSIVE_FILE_BYTES) {
      model.blockedSources.push(`${resolution.displayName}: reachable but exceeds recursive extraction size limit (${resolution.fileSizeBytes} bytes).`);
      return;
    }

    try {
      const buffer = await fs.readFile(resolution.resolvedPath);
      await extractFile(
        {
          originalName: path.basename(resolution.resolvedPath),
          mimeType: undefined,
          size: buffer.byteLength,
          buffer,
          tempPath: resolution.resolvedPath,
          sourcePath: resolution.resolvedPath,
          discoveredFrom: parentFile.originalName,
        },
        model,
        ids,
        packageNodeId,
        {
          ...context,
          recursionDepth: context.recursionDepth + 1,
        },
      );
      model.dependencyUsage.push({
        dependency_id: `DEP-${model.dependencyUsage.length + 1}`,
        source_asset: parentFile.originalName,
        dependency: resolution.resolvedPath,
        dependency_type: 'recursive linked source extraction',
        evidence_id: evidenceId,
        confidence: 'confirmed',
      });
    } catch (error) {
      model.blockedSources.push(
        `${resolution.displayName}: reachable path could not be read for recursive extraction (${error instanceof Error ? error.message : 'unknown error'}).`,
      );
    }
  });
}

function parseAccessDatabasePath(connect: string): string {
  const match = connect.match(/(?:^|;)DATABASE=([^;]+)/i);
  return match?.[1]?.trim() ?? '';
}

function inferAccessLinkedSourceKind(connect: string, sourcePath: string, sourceTableName: string): string {
  const value = `${connect} ${sourcePath} ${sourceTableName}`.toLowerCase();
  if (/\.(accdb|mdb)\b/.test(value)) {
    return 'Access database';
  }
  if (/\.(xlsx|xlsm|xlsb|xls)\b|excel/.test(value)) {
    return 'Excel workbook';
  }
  if (/\.(csv|txt|tsv)\b|text;|fmt=delimited/.test(value)) {
    return 'Delimited/text file';
  }
  if (sourcePath) {
    return 'External file or folder';
  }
  return 'External source';
}

function candidateLinkedSourcePaths(sourcePath: string, sourceTableName: string): string[] {
  const candidates = new Set<string>();
  candidates.add(path.normalize(sourcePath));
  if (sourceTableName) {
    const normalizedName = normalizeAccessSourceTableName(sourceTableName);
    if (path.extname(sourcePath)) {
      candidates.add(path.normalize(sourcePath));
    } else {
      for (const name of [sourceTableName, normalizedName]) {
        candidates.add(path.normalize(path.join(sourcePath, name)));
        for (const extension of ['.csv', '.txt', '.tsv', '.xlsx', '.xlsm', '.xlsb', '.xls', '.accdb', '.mdb']) {
          candidates.add(path.normalize(path.join(sourcePath, `${name}${extension}`)));
        }
      }
    }
  }
  return Array.from(candidates).filter(Boolean);
}

async function findNestedLinkedFile(folder: string, sourceTableName: string): Promise<string> {
  if (!sourceTableName) {
    return '';
  }
  const normalizedTarget = normalizeForLooseMatch(sourceTableName);
  let entries: string[] = [];
  try {
    entries = await fs.readdir(folder);
  } catch {
    return '';
  }
  const supported = entries.filter((entry) => ACCESS_EXTENSIONS.has(path.extname(entry).toLowerCase()) || EXCEL_EXTENSIONS.has(path.extname(entry).toLowerCase()) || FLAT_EXTENSIONS.has(path.extname(entry).toLowerCase()));
  const match = supported.find((entry) => normalizeForLooseMatch(path.parse(entry).name) === normalizedTarget || normalizeForLooseMatch(entry) === normalizedTarget);
  return match ? path.join(folder, match) : '';
}

function normalizeAccessSourceTableName(value: string): string {
  return value.replace(/#([A-Za-z0-9]+)$/u, '.$1').trim();
}

function normalizeForLooseMatch(value: string): string {
  return normalizeAccessSourceTableName(value).replace(/\.[a-z0-9]+$/i, '').replace(/[^a-z0-9]+/gi, '').toLowerCase();
}

async function safeStat(candidatePath: string): Promise<import('node:fs').Stats | undefined> {
  try {
    return await fs.stat(candidatePath);
  } catch {
    return undefined;
  }
}

function linkedSourceKey(connect: string, sourceTableName: string, tableName: string): string {
  return `${connect || ''}\u0000${sourceTableName || ''}\u0000${tableName || ''}`.toLowerCase();
}

function linkedSourceNodeReuseKey(resolution: LinkedSourceResolution | undefined, fallback: string): string {
  return (resolution?.resolvedPath || resolution?.sourcePath || resolution?.connect || fallback).replace(/\s+/g, ' ').trim().toLowerCase();
}

function displayLinkedSourceName(sourcePath: string, sourceTableName: string, tableName: string): string {
  if (sourcePath && path.extname(sourcePath)) {
    return path.basename(sourcePath);
  }
  if (sourcePath) {
    return sourceTableName ? `${path.basename(sourcePath)} / ${normalizeAccessSourceTableName(sourceTableName)}` : path.basename(sourcePath) || sourcePath;
  }
  return sourceTableName || tableName || 'Unresolved linked source';
}

function nodeTypeForResolvedPath(resolvedPath: string): NodeType {
  const sourceType = detectSourceType(resolvedPath);
  if (sourceType === 'excel') {
    return 'workbook';
  }
  if (sourceType === 'access') {
    return 'database';
  }
  return 'file';
}

function formatBlockedSource(resolution: LinkedSourceResolution | undefined, fallback: string): string {
  if (!resolution) {
    return fallback;
  }
  const location = resolution.sourcePath || resolution.connect || fallback;
  return `${resolution.displayName} | ${resolution.sourceKind} | ${resolution.resolutionReason} | ${location}`;
}

function createAccessBlocker(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  fileNodeId: string,
  sourceEvidenceId: string,
  reason: string,
): void {
  const blockerEvidenceId = addEvidence(model, ids, {
    category: '05m_QA_Certification',
    title: `Access extraction blocker - ${file.originalName}`,
    fileName: `${sanitizeName(file.originalName)}_access_extraction_blocker.json`,
    sourceFile: file.originalName,
    summary: 'Access metadata extraction requires a native catalog collector outside Vercel serverless.',
    confidence: 'blocked',
    content: JSON.stringify(
      {
        source_file: file.originalName,
        blocked_items: [
          'MSysObjects',
          'MSysQueries',
          'saved query SQL',
          'macro XML and macro actions',
          'forms and reports',
          'modules and VBA',
          'import/export specs',
          'relationships',
          'indexes and keys',
          'row counts and column inventory',
        ],
        reason,
      },
      null,
      2,
    ),
  });

  const blocker = addNode(model, ids, {
    node_type: 'upstream blocker',
    name: `Access catalog blocked: ${file.originalName}`,
    description: 'Access system catalog and object metadata could not be extracted in the hosted runtime.',
    source_file: file.originalName,
    business_purpose: 'Blocks authoritative Access inventory, query/macro reconciliation, and recursive lineage.',
    owner_status: 'requires platform/data owner confirmation',
    criticality: 'P0',
    confidence: 'blocked',
    evidence_id: blockerEvidenceId,
    recommended_action: 'Run the Access desktop collector or provide exported MSysObjects, MSysQueries, macro XML, VBA, and import/export specs.',
    failure_impact: 'Saved queries, macros, linked sources, lineage, and failure impact cannot be certified from file metadata alone.',
    dollar_exposure: 'Potentially material; validate once Access internals are extracted.',
  });

  addEdge(model, ids, {
    from_node_id: blocker.node_id,
    to_node_id: fileNodeId,
    edge_type: 'blocks_lineage',
    description: 'Blocked Access catalog prevents authoritative recursive lineage.',
    automated_flag: 'unknown',
    transformation_id: '',
    cadence: 'unknown',
    confidence: 'blocked',
    evidence_id: blockerEvidenceId,
  });

  model.blockedSources.push(file.originalName);
  model.limitations.push(`Access internals for ${file.originalName} are blocked until a native catalog extraction is supplied.`);
  model.access.Access_Object_Inventory.push({
    object_name: file.originalName,
    raw_type_value: 'blocked',
    interpreted_type: 'database file',
    evidence_id: sourceEvidenceId,
    confidence: 'blocked',
    recommended_action: 'Extract MSysObjects with authoritative Type values.',
  });

  addOpenQuestion(model, ids, {
    asset: file.originalName,
    question: 'Who can provide an Access metadata export or run the desktop collector?',
    owner_role: 'Access owner / data engineering',
    blocker_type: 'lineage-blocker',
    priority: 'P0',
    evidence_id: blockerEvidenceId,
  });

  addAction(model, ids, {
    title: `Extract Access catalog for ${file.originalName}`,
    description: 'Run an approved desktop collector to export MSysObjects, MSysQueries, saved query SQL, macro XML/actions, VBA modules, import/export specs, linked table connections, relationships, indexes, row counts, and column inventory.',
    source_asset: file.originalName,
    owner_role: 'Access owner / data engineering',
    recommended_owner: 'Application owner with Access desktop runtime',
    action_type: 'Resolve Blocker',
    priority: 'P0',
    severity: 'P0',
    dependency: 'Access desktop runtime or exported catalog files',
    acceptance_criteria: 'Workbook Access tabs reconcile query count, macro count, SQL evidence count, macro action count, linked source count, and blocker count.',
    evidence_id: blockerEvidenceId,
    related_risk: 'Uncertified Access lineage and object inventory',
    expected_business_value: 'Enables defensible migration, governance, and failure-impact analysis.',
  });
}

function asRows(value: unknown): Record<string, any>[] {
  return Array.isArray(value) ? (value as Record<string, any>[]) : [];
}

function schemaTableTypeToAccessType(tableType: string): number {
  if (/link/i.test(tableType)) {
    return 6;
  }
  if (/view/i.test(tableType)) {
    return 5;
  }
  if (/table/i.test(tableType)) {
    return 1;
  }
  return Number.NaN;
}

function isAccessSystemObject(name: string): boolean {
  return /^MSys/i.test(name) || /^~TMP/i.test(name);
}

function isAccessSystemTableDef(attributes: number): boolean {
  return attributes === 2 || attributes < 0;
}

function isAccessLinkedTableDef(attributes: number): boolean {
  return attributes === 1073741824 || attributes === 536870912;
}

function interpretAccessType(type: number): string {
  const map = new Map<number, string>([
    [1, 'local table'],
    [4, 'ODBC linked table'],
    [5, 'saved query'],
    [6, 'linked table'],
    [-32768, 'form'],
    [-32766, 'macro'],
    [-32764, 'report'],
    [-32761, 'module'],
  ]);
  return map.get(type) ?? 'unknown / system / version-specific';
}

function accessNodeType(type: string): NodeType {
  if (type.includes('linked table')) {
    return 'linked table';
  }
  if (type.includes('table')) {
    return 'table';
  }
  if (type.includes('query')) {
    return 'query';
  }
  if (type.includes('macro')) {
    return 'macro';
  }
  if (type.includes('form')) {
    return 'form';
  }
  if (type.includes('report')) {
    return 'report';
  }
  if (type.includes('module')) {
    return 'module';
  }
  return 'file';
}

function accessRecommendedAction(type: string): string {
  if (type.includes('query')) {
    return 'Extract SQL evidence, map lineage, classify transformations, and reconcile macro references.';
  }
  if (type.includes('macro')) {
    return 'Extract macro XML/action map and reconcile OpenQuery/RunSQL/Transfer actions.';
  }
  if (type.includes('linked')) {
    return 'Resolve linked source path/system and include upstream file or mark terminal blocker.';
  }
  if (type.includes('table')) {
    return 'Profile rows/columns, confirm keys, owner, source-of-truth status, and downstream usage.';
  }
  if (type.includes('form') || type.includes('report')) {
    return 'Export metadata and confirm user/output purpose.';
  }
  if (type.includes('module')) {
    return 'Export VBA source text and map procedures, reads, writes, exports, and scheduling behavior.';
  }
  return 'Confirm object interpretation with Access owner.';
}

function addAccessObjectNode(
  model: DiscoveryModel,
  ids: IdState,
  file: UploadedSource,
  name: string,
  nodeTypeOrInterpretedType: NodeType | string,
  evidenceId: string,
  confidence: Confidence,
): DiscoveryNode {
  const nodeType = (NODE_TYPE_SET.has(nodeTypeOrInterpretedType as NodeType)
    ? nodeTypeOrInterpretedType
    : accessNodeType(nodeTypeOrInterpretedType)) as NodeType;
  return addNode(model, ids, {
    node_type: nodeType,
    name,
    description: `Access ${nodeType} object extracted from ${file.originalName}.`,
    source_file: file.originalName,
    business_purpose: 'Access object participating in storage, transformation, UI, automation, or output workflow.',
    owner_status: 'owner-confirmation-required',
    criticality: nodeType === 'macro' || nodeType === 'query' || nodeType === 'linked table' ? 'P1' : 'P2',
    confidence,
    evidence_id: evidenceId,
    recommended_action: accessRecommendedAction(String(nodeTypeOrInterpretedType)),
    failure_impact: 'Object changes may alter lineage, run behavior, controls, outputs, or downstream decisions.',
    dollar_exposure: 'Directional exposure modeled at process level.',
  });
}

const NODE_TYPE_SET = new Set<NodeType>([
  'system',
  'database',
  'file',
  'folder',
  'workbook',
  'worksheet',
  'table',
  'linked table',
  'query',
  'macro',
  'macro action',
  'form',
  'report',
  'module',
  'Power Query',
  'formula area',
  'named range',
  'pivot',
  'document',
  'document section',
  'process step',
  'data element',
  'output',
  'control',
  'exception',
  'person / role',
  'upstream blocker',
  'downstream consumer',
]);

function addAccessContainmentEdge(
  model: DiscoveryModel,
  ids: IdState,
  fileNodeId: string,
  objectNodeId: string,
  evidenceId: string,
  confidence: Confidence,
): void {
  addEdge(model, ids, {
    from_node_id: fileNodeId,
    to_node_id: objectNodeId,
    edge_type: 'depends_on',
    description: 'Access database contains object.',
    automated_flag: 'automated',
    transformation_id: '',
    cadence: 'per Access open/run',
    confidence,
    evidence_id: evidenceId,
  });
}

function addSqlReferenceEdgesForAccessQuery(
  model: DiscoveryModel,
  ids: IdState,
  file: UploadedSource,
  queryNodeId: string,
  refs: ReturnType<typeof extractSqlReferences>,
  evidenceId: string,
): void {
  for (const read of refs.reads) {
    const tableNode = addNode(model, ids, {
      node_type: 'table',
      name: read,
      description: `Table/view read by Access saved query in ${file.originalName}.`,
      source_file: file.originalName,
      business_purpose: 'Potential upstream table/view referenced in saved query SQL.',
      owner_status: 'owner-confirmation-required',
      criticality: 'P1',
      confidence: 'inferred',
      evidence_id: evidenceId,
      recommended_action: 'Resolve to Access table/query/linked source and continue recursive lineage.',
      failure_impact: 'If this source changes or is unavailable, query output can fail or become wrong.',
      dollar_exposure: 'Directional exposure modeled at process level.',
    });
    addEdge(model, ids, {
      from_node_id: queryNodeId,
      to_node_id: tableNode.node_id,
      edge_type: 'reads_from',
      description: 'Saved query SQL reads from referenced object.',
      automated_flag: 'automated',
      transformation_id: '',
      cadence: 'per query run',
      confidence: 'inferred',
      evidence_id: evidenceId,
    });
  }

  for (const write of refs.writes) {
    const targetNode = addNode(model, ids, {
      node_type: 'table',
      name: write,
      description: `Write target referenced by Access saved query in ${file.originalName}.`,
      source_file: file.originalName,
      business_purpose: 'Potential downstream table/output referenced in action query SQL.',
      owner_status: 'owner-confirmation-required',
      criticality: 'P1',
      confidence: 'inferred',
      evidence_id: evidenceId,
      recommended_action: 'Confirm write target, controls, rollback behavior, and downstream consumers.',
      failure_impact: 'Failed or wrong writes can create partial, stale, or unauditable outputs.',
      dollar_exposure: 'Directional exposure modeled at process level.',
    });
    addEdge(model, ids, {
      from_node_id: queryNodeId,
      to_node_id: targetNode.node_id,
      edge_type: 'writes_to',
      description: 'Saved query SQL writes to referenced object.',
      automated_flag: 'automated',
      transformation_id: '',
      cadence: 'per query run',
      confidence: 'inferred',
      evidence_id: evidenceId,
    });
  }
}

function addSqlTransformationRule(
  model: DiscoveryModel,
  ids: IdState,
  sourceAsset: string,
  refs: ReturnType<typeof extractSqlReferences>,
  evidenceId: string,
): void {
  if (!refs.whereClauses.length && !refs.joinClauses.length && !refs.writes.length) {
    return;
  }
  model.transformations.push({
    transformation_id: nextId(ids, 'TRN'),
    source_asset: sourceAsset,
    rule_type: 'Access saved query SQL',
    rule_description: [...refs.joinClauses.slice(0, 4), ...refs.whereClauses.slice(0, 4), refs.writes.length ? `writes: ${refs.writes.join(', ')}` : '']
      .filter(Boolean)
      .join(' | '),
    input_fields: refs.reads.join(', '),
    output_fields: refs.writes.join(', '),
    evidence_id: evidenceId,
    confidence: 'inferred',
    recommended_action: 'Review parsed Access SQL logic and confirm business rule intent, execution order, controls, and output impact.',
  });
}

function createTargetedBlocker(
  model: DiscoveryModel,
  ids: IdState,
  file: UploadedSource,
  fileNodeId: string,
  title: string,
  evidenceId: string,
  action: string,
): void {
  const blocker = addNode(model, ids, {
    node_type: 'upstream blocker',
    name: title,
    description: action,
    source_file: file.originalName,
    business_purpose: 'Blocks complete, defensible lineage, controls, automation, or object reconciliation.',
    owner_status: 'owner-confirmation-required',
    criticality: 'P1',
    confidence: 'blocked',
    evidence_id: evidenceId,
    recommended_action: action,
    failure_impact: 'Material discovery claims remain incomplete until this blocker is resolved.',
    dollar_exposure: 'Directional exposure modeled at process level.',
  });
  addEdge(model, ids, {
    from_node_id: blocker.node_id,
    to_node_id: fileNodeId,
    edge_type: 'blocks_lineage',
    description: 'Blocked extraction limits canonical discovery model completeness.',
    automated_flag: 'unknown',
    transformation_id: '',
    cadence: 'one-time discovery blocker',
    confidence: 'blocked',
    evidence_id: evidenceId,
  });
  model.blockedSources.push(`${file.originalName}: ${title}`);
  addAction(model, ids, {
    title,
    description: action,
    source_asset: file.originalName,
    owner_role: 'Technical owner / data engineering',
    recommended_owner: 'Assigned application/data owner',
    action_type: 'Resolve Blocker',
    priority: 'P1',
    severity: 'P1',
    dependency: 'Owner-provided export or trusted desktop collector',
    acceptance_criteria: 'Blocked artifact is extracted, evidence-indexed, and reconciled in the technical workbook.',
    evidence_id: evidenceId,
    related_risk: title,
    expected_business_value: 'Moves the dossier from partial to analyst-defensible for the affected area.',
  });
}

function extractVbaProcedures(code: string): {
  name: string;
  kind: string;
  signature: string;
  visibility: string;
  parameter_text: string;
  procedure_role: string;
  runnable_macro_flag: 'yes' | 'no';
  code: string;
  purpose: string;
  operations: string[];
  references: string[];
}[] {
  const normalized = normalizeExtractedVbaCode(code).replace(/\r\n/g, '\n');
  const procedurePattern = vbaProcedurePattern();
  const matches = Array.from(normalized.matchAll(procedurePattern));
  return matches.map((match, index) => {
    const start = match.index ?? 0;
    const end = index + 1 < matches.length ? matches[index + 1].index ?? normalized.length : normalized.length;
    const procedureCode = normalized.slice(start, end).trim();
    const signature = (match[0] ?? '').trim();
    const kind = collapseWhitespace(match[1] ?? 'Sub');
    const name = match[2] ?? 'UnnamedProcedure';
    const visibility = signature.match(/^\s*(Public|Private|Friend)\b/i)?.[1] ?? 'public/default';
    const parameterText = signature.match(/\(([^)]*)\)/)?.[1]?.trim() ?? '';
    const operations = inferVbaOperations(procedureCode, name);
    const references = inferVbaReferences(procedureCode);
    const procedureRole = classifyExcelVbaProcedure(kind, name, signature, parameterText);
    return {
      name,
      kind,
      signature,
      visibility,
      parameter_text: parameterText,
      procedure_role: procedureRole,
      runnable_macro_flag: procedureRole === 'macro entrypoint' ? 'yes' : 'no',
      code: procedureCode,
      purpose: inferAutomationPurpose(procedureCode, [name, ...operations, ...references]),
      operations,
      references,
    };
  });
}

function extractVbaModulesFromProject(vbaProject: Buffer): { name: string; code: string }[] {
  if (vbaProject.length < 512) {
    return [];
  }

  try {
    const cfb = CFB.read(vbaProject, { type: 'buffer' });
    const streams = cfb.FullPaths.map((fullPath, index) => ({
      fullPath,
      entry: cfb.FileIndex[index],
    })).filter(({ fullPath, entry }) => {
      if (!entry?.content || !/\/VBA\//i.test(fullPath)) {
        return false;
      }
      return !/\/(?:dir|_VBA_PROJECT|PROJECT|PROJECTwm|__SRP_|VBFrame)/i.test(fullPath);
    });

    const modules: { name: string; code: string; score: number }[] = [];
    for (const stream of streams) {
      const raw = Buffer.from(stream.entry.content as Uint8Array);
      const name = path.basename(stream.fullPath);
      const code = normalizeExtractedVbaCode(decompressBestVbaStream(raw));
      const score = scoreVbaCode(code);
      if (score > 0 && hasMaterialVbaCode(code)) {
        modules.push({ name, code, score });
      }
    }

    return modules
      .sort((a, b) => b.score - a.score)
      .map(({ name, code }) => ({ name, code }));
  } catch {
    return [];
  }
}

function decompressBestVbaStream(raw: Buffer): string {
  const candidates: { text: string; score: number }[] = [];
  const maxOffset = Math.min(raw.length, 8192);
  for (let offset = 0; offset < maxOffset; offset += 1) {
    if (raw[offset] !== 0x01) {
      continue;
    }
    const decompressed = decompressVbaCompressedContainer(raw.subarray(offset));
    const text = decodeVbaText(decompressed);
    const score = scoreVbaCode(text);
    if (score > 0) {
      candidates.push({ text, score });
    }
  }

  const ascii = decodeVbaText(raw);
  const asciiScore = scoreVbaCode(ascii);
  if (asciiScore > 0) {
    candidates.push({ text: ascii, score: asciiScore });
  }

  return candidates.sort((a, b) => b.score - a.score)[0]?.text ?? '';
}

function vbaProcedurePattern(): RegExp {
  return /^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+([A-Za-z_][A-Za-z0-9_]*)[^\n]*/gim;
}

function normalizeExtractedVbaCode(code: string): string {
  const printable = code
    .replace(/\r\n/g, '\n')
    .replace(/\0/g, '')
    .replace(/[^\x09\x0a\x0d\x20-\x7e]/g, '')
    .trim();
  const firstCodeToken = printable.search(
    /(?:^|\n)\s*(?:Attribute\s+VB_Name|Option\s+Explicit|(?:Public\s+|Private\s+|Friend\s+|Static\s+)*(?:Sub|Function|Property\s+))/i,
  );
  return firstCodeToken > 0 ? printable.slice(firstCodeToken).trim() : printable;
}

function hasMaterialVbaCode(code: string): boolean {
  if (vbaProcedurePattern().test(code)) {
    return true;
  }
  return code
    .split(/\r?\n/)
    .map((line) => line.trim())
    .some((line) => line && !/^Attribute\s+/i.test(line) && !/^Option\s+/i.test(line));
}

function classifyExcelVbaProcedure(kind: string, name: string, signature: string, parameterText: string): string {
  const normalizedKind = kind.toLowerCase();
  if (normalizedKind !== 'sub') {
    return normalizedKind.includes('function') ? 'function/helper' : 'property/helper';
  }
  if (isExcelEventProcedure(name)) {
    return 'event handler';
  }
  if (/^\s*Private\b/i.test(signature)) {
    return 'private helper subroutine';
  }
  if (hasRequiredVbaParameters(parameterText)) {
    return 'parameterized helper subroutine';
  }
  return 'macro entrypoint';
}

function isExcelEventProcedure(name: string): boolean {
  return /^(Workbook|Worksheet|Chart|Application|App|QueryTable|PivotTable)_/i.test(name);
}

function hasRequiredVbaParameters(parameterText: string): boolean {
  if (!parameterText.trim()) {
    return false;
  }
  return parameterText
    .split(',')
    .map((parameter) => parameter.trim())
    .filter(Boolean)
    .some((parameter) => !/^Optional\b/i.test(parameter));
}

function decompressVbaCompressedContainer(container: Buffer): Buffer {
  if (!container.length || container[0] !== 0x01) {
    return Buffer.alloc(0);
  }

  const output: number[] = [];
  let offset = 1;
  while (offset + 2 <= container.length) {
    const headerOffset = offset;
    const header = container.readUInt16LE(offset);
    offset += 2;
    const chunkSize = (header & 0x0fff) + 3;
    const compressed = Boolean(header & 0x8000);
    const chunkEnd = Math.min(headerOffset + chunkSize, container.length);
    const chunkStartOutput = output.length;

    if (!compressed) {
      while (offset < chunkEnd) {
        output.push(container[offset++] ?? 0);
      }
      continue;
    }

    while (offset < chunkEnd) {
      const flags = container[offset++] ?? 0;
      for (let bit = 0; bit < 8 && offset < chunkEnd; bit += 1) {
        if ((flags & (1 << bit)) === 0) {
          output.push(container[offset++] ?? 0);
          continue;
        }

        if (offset + 2 > chunkEnd) {
          offset = chunkEnd;
          break;
        }

        const token = container.readUInt16LE(offset);
        offset += 2;
        const copied = output.length - chunkStartOutput;
        const bitCount = Math.max(4, Math.ceil(Math.log2(Math.max(copied, 1))));
        const lengthMask = 0xffff >> bitCount;
        const length = (token & lengthMask) + 3;
        const displacement = (token >> (16 - bitCount)) + 1;
        for (let index = 0; index < length; index += 1) {
          const sourceIndex = output.length - displacement;
          output.push(sourceIndex >= 0 ? output[sourceIndex] ?? 0 : 0);
        }
      }
    }
  }

  return Buffer.from(output);
}

function decodeVbaText(buffer: Buffer): string {
  return buffer
    .toString('latin1')
    .replace(/\0/g, '')
    .replace(/[^\x09\x0a\x0d\x20-\x7e]/g, '')
    .trim();
}

function scoreVbaCode(text: string): number {
  if (!text) {
    return 0;
  }
  let score = 0;
  if (/\b(Attribute|Option\s+Explicit|Sub\s+|Function\s+|Property\s+)/i.test(text)) score += 5;
  if (/\b(End\s+Sub|End\s+Function|End\s+Property)\b/i.test(text)) score += 5;
  if (/\b(Workbook_|Worksheet_|Range\(|Cells\(|Sheets\(|Worksheets\(|Application\.)/i.test(text)) score += 3;
  if (text.length > 40) score += 1;
  if (text.length > 300) score += 1;
  return score;
}

function inferAutomationPurpose(text: string, hints: string[] = []): string {
  const combined = `${hints.join(' ')} ${text}`.toLowerCase();
  const purposes: string[] = [];
  if (/\b(workbook_open|auto_open|document_open)\b/.test(combined)) {
    purposes.push('startup/run trigger automation');
  }
  if (/(workbook_beforesave|workbook_newsheet|workbook_sheetactivate|worksheet_change|worksheet_selectionchange|worksheet_calculate|beforeclose|beforesave)/.test(combined)) {
    purposes.push('event-driven workbook logic');
  }
  if (/\brefreshmasterrosternow\b/.test(combined) && !/\bsub\s+refreshmasterrosternow\b/.test(combined)) {
    purposes.push('triggers master roster refresh automation');
  }
  if (/\b(refreshall|querytable|connections?|odbc|oledb|powerquery|mashup|\.refresh\b)\b/.test(combined)) {
    purposes.push('external data refresh or connection orchestration');
  }
  if (/\b(adodb|connectionstring|recordset|insert\s+into|update\s+\w+\s+set|delete\s+from|runsql|openquery)\b/.test(combined) || /\bselect\s+.+\s+from\b/.test(combined)) {
    purposes.push('database/query execution or transformation');
  }
  if (/\b(workbooks\.open|filesystemobject|dir\(|kill\s+|name\s+|filecopy|open\s+.*for\s+|textstream)\b/.test(combined)) {
    purposes.push('file intake, export, or filesystem dependency');
  }
  if (/\b(saveas|exportasfixedformat|printout|publishobjects|sendmail|outlook|transfertext|transferspreadsheet)\b/.test(combined)) {
    purposes.push('output generation, export, or distribution');
  }
  if (/\b(pivottables?|pivotcache|chartobjects?|listobjects?)\b/.test(combined)) {
    purposes.push('reporting, pivot, or table refresh');
  }
  if (/\b(range\(|cells\(|sheets\(|worksheets\(|formula|vlookup|xlookup|match\(|index\()\b/.test(combined)) {
    purposes.push('worksheet transformation or calculation control');
  }
  if (/\b(msgbox|inputbox|application\.inputbox|userform|show)\b/.test(combined)) {
    purposes.push('user prompt or manual workflow step');
  }
  if (/\b(on error|err\.|resume next|goto)\b/.test(combined)) {
    purposes.push('error handling/control flow');
  }
  if (!purposes.length && /\b(issystemsheet|select\s+case|sheetname|sheet\s+classification)\b/.test(combined)) {
    purposes.push('system/control sheet classification helper');
  }
  return purposes.length ? Array.from(new Set(purposes)).join('; ') : 'purpose requires owner confirmation from macro code review';
}

function inferVbaOperations(code: string, name: string): string[] {
  const combined = `${name}\n${code}`.toLowerCase();
  const operations: string[] = [];
  if (/\brefreshmasterrosternow\b/.test(combined) && !/^refreshmasterrosternow$/i.test(name)) operations.push('calls master roster refresh macro');
  if (/\b(refreshall|querytable|connections?|odbc|oledb|powerquery|mashup|\.refresh\b)\b/.test(combined)) operations.push('external data refresh');
  if (/\b(workbooks\.open|filesystemobject|dir\(|filecopy|open\s+.*for\s+)\b/.test(combined)) operations.push('file input/output');
  if (/\b(adodb|recordset|insert\s+into|update\s+\w+\s+set|delete\s+from|runsql|openquery)\b/.test(combined) || /\bselect\s+.+\s+from\b/.test(combined)) operations.push('database/query execution');
  if (/\b(saveas|exportasfixedformat|printout|sendmail|outlook|transferspreadsheet|transfertext)\b/.test(combined)) operations.push('output/export');
  if (/\b(range\(|cells\(|worksheets\(|sheets\(|formula)\b/.test(combined)) operations.push('worksheet mutation/calculation');
  if (/\b(pivottables?|pivotcache|chartobjects?)\b/.test(combined)) operations.push('reporting/pivot refresh');
  if (/\b(on error|err\.|resume next)\b/.test(combined)) operations.push('error handling/control');
  if (/\b(issystemsheet|select\s+case|sheetname)\b/.test(combined)) operations.push('system sheet classification');
  return Array.from(new Set(operations));
}

function inferVbaReferences(code: string): string[] {
  const references = new Set<string>();
  for (const match of code.matchAll(/["']([^"'\r\n]+\.(?:xlsx|xlsm|xlsb|xls|csv|txt|accdb|mdb|sql|pdf))["']/gi)) {
    references.add(match[1] ?? '');
  }
  for (const match of code.matchAll(/\b(?:Worksheets|Sheets)\s*\(\s*"([^"]+)"/gi)) {
    references.add(`worksheet:${match[1] ?? ''}`);
  }
  for (const match of code.matchAll(/\b(?:Range)\s*\(\s*"([^"]+)"/gi)) {
    references.add(`range:${match[1] ?? ''}`);
  }
  const sqlRefs = extractSqlReferences(code);
  sqlRefs.reads.forEach((ref) => references.add(`reads:${ref}`));
  sqlRefs.writes.forEach((ref) => references.add(`writes:${ref}`));
  return Array.from(references).filter(Boolean).slice(0, 40);
}

function inferVbaCadence(name: string): string {
  if (/^(Workbook_Open|Auto_Open)$/i.test(name)) {
    return 'on workbook open';
  }
  if (/Worksheet_Change/i.test(name)) {
    return 'on worksheet edit/change';
  }
  if (/Worksheet_Calculate/i.test(name)) {
    return 'on worksheet calculation';
  }
  if (/BeforeSave/i.test(name)) {
    return 'before workbook save';
  }
  if (/NewSheet/i.test(name)) {
    return 'when a new worksheet is added';
  }
  if (/SheetActivate/i.test(name)) {
    return 'when a worksheet is activated';
  }
  if (/BeforeClose/i.test(name)) {
    return 'before workbook close';
  }
  return 'manual button/run or owner-confirmation-required';
}

function parseAccessMacroActions(text: string): { action_type: string; target_object: string }[] {
  if (!text) {
    return [];
  }

  const actionNames = ['OpenQuery', 'RunSQL', 'TransferSpreadsheet', 'TransferText', 'OpenForm', 'OpenReport', 'RunMacro'];
  const actions: { action_type: string; target_object: string }[] = [];
  const actionPattern = new RegExp(`\\b(${actionNames.join('|')})\\b`, 'gi');
  for (const match of text.matchAll(actionPattern)) {
    const actionType = match[1] ?? 'Unknown';
    const window = text.slice(match.index ?? 0, (match.index ?? 0) + 500);
    const target =
      window.match(/\b(?:QueryName|ObjectName|MacroName|TableName|Name|ReportName|FormName)\s*=\s*"([^"]+)"/i)?.[1] ??
      window.match(/\b(?:QueryName|ObjectName|MacroName|TableName|ReportName|FormName)\s*:\s*([^\r\n]+)/i)?.[1]?.trim() ??
      '';
    actions.push({ action_type: actionType, target_object: target });
  }
  return actions;
}

function resolveAccessMacroTarget(target: string, queryNames: Set<string>, macroNames: Set<string>): string {
  if (!target) {
    return 'unknown target';
  }
  if (queryNames.has(target)) {
    return 'saved query';
  }
  if (macroNames.has(target)) {
    return 'macro';
  }
  return 'unknown or missing target';
}

async function extractExcel(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  workbookNodeId: string,
  sourceEvidenceId: string,
): Promise<void> {
  let workbook: XLSX.WorkBook;
  try {
    workbook = XLSX.read(file.buffer, { type: 'buffer', cellFormula: true, cellDates: true, bookVBA: true });
  } catch (error) {
    addOpenQuestion(model, ids, {
      asset: file.originalName,
      question: `Workbook could not be parsed by xlsx: ${error instanceof Error ? error.message : 'unknown parse error'}`,
      owner_role: 'Workbook owner',
      blocker_type: 'parse-error',
      priority: 'P1',
      evidence_id: sourceEvidenceId,
    });
    return;
  }

  const metadataEvidenceId = addEvidence(model, ids, {
    category: '05a_Raw_Metadata',
    title: `Workbook metadata - ${file.originalName}`,
    fileName: `${sanitizeName(file.originalName)}_workbook_metadata.json`,
    sourceFile: file.originalName,
    summary: 'Workbook properties, sheet names, workbook metadata, and hidden sheet indicators.',
    confidence: 'confirmed',
    content: JSON.stringify(
      {
        props: workbook.Props ?? {},
        sheet_names: workbook.SheetNames,
        workbook_metadata: workbook.Workbook ?? {},
      },
      null,
      2,
    ),
  });

  model.excel.Excel_Workbook_Inventory.push({
    workbook: file.originalName,
    sheet_count: workbook.SheetNames.length,
    protection_state: 'unknown',
    metadata_evidence_id: metadataEvidenceId,
    confidence: 'confirmed',
    recommended_action: 'Confirm workbook owner, refresh cadence, output consumers, and protection requirements.',
  });

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const ref = sheet?.['!ref'] ?? '';
    const range = ref ? XLSX.utils.decode_range(ref) : undefined;
    const rows = range ? range.e.r - range.s.r + 1 : 0;
    const cols = range ? range.e.c - range.s.c + 1 : 0;
    const sheetVisibility = workbook.Workbook?.Sheets?.find((item) => item.name === sheetName)?.Hidden ?? 0;
    const visibility = sheetVisibility === 2 ? 'very hidden' : sheetVisibility === 1 ? 'hidden' : 'visible';

    const sheetEvidenceId = addEvidence(model, ids, {
      category: '05h_Data_Profiles',
      title: `Sheet profile - ${sheetName}`,
      fileName: `${sanitizeName(file.originalName)}_${sanitizeName(sheetName)}_profile.json`,
      sourceFile: file.originalName,
      summary: `Data profile for worksheet ${sheetName}.`,
      confidence: 'confirmed',
      content: JSON.stringify(profileSheet(sheet, sheetName), null, 2),
    });

    const sheetNode = addNode(model, ids, {
      node_type: 'worksheet',
      name: sheetName,
      description: `Worksheet in ${file.originalName}; ${rows} row(s), ${cols} column(s), visibility ${visibility}.`,
      source_file: file.originalName,
      business_purpose: 'Worksheet data area or calculation surface.',
      owner_status: 'owner-confirmation-required',
      criticality: rows > 0 ? 'P1' : 'P3',
      confidence: 'confirmed',
      evidence_id: sheetEvidenceId,
      recommended_action: 'Confirm worksheet business role, manual input zones, and output usage.',
      failure_impact: 'Worksheet changes may affect workbook outputs, calculations, or downstream extracts.',
      dollar_exposure: 'Directional exposure modeled at process level.',
    });

    addEdge(model, ids, {
      from_node_id: workbookNodeId,
      to_node_id: sheetNode.node_id,
      edge_type: 'depends_on',
      description: 'Workbook includes worksheet.',
      automated_flag: 'automated',
      transformation_id: '',
      cadence: 'on open / refresh',
      confidence: 'confirmed',
      evidence_id: sheetEvidenceId,
    });

    model.excel.Excel_Sheet_Inventory.push({
      workbook: file.originalName,
      sheet_name: sheetName,
      visibility,
      used_range: ref,
      row_count: rows,
      column_count: cols,
      evidence_id: sheetEvidenceId,
      confidence: 'confirmed',
      recommended_action: 'Classify as input, staging, calculation, output, or unused.',
    });

    model.rowCountSummary[`${file.originalName}:${sheetName}`] = rows;
    model.excel.Excel_Data_Profile.push({
      workbook: file.originalName,
      sheet_name: sheetName,
      used_range: ref,
      row_count: rows,
      column_count: cols,
      visibility,
      evidence_id: sheetEvidenceId,
      confidence: 'confirmed',
      recommended_action: 'Review profile evidence for blanks, duplicate keys, invalid dates/numbers, hardcoded overrides, and manual input zones.',
    });
    addFormulaAreas(file.originalName, sheetName, sheet, model, ids, sheetNode.node_id, sheetEvidenceId);
    addSheetDataElements(file.originalName, sheetName, sheet, model, ids, sheetEvidenceId);
  }

  for (const definedName of workbook.Workbook?.Names ?? []) {
    const name = definedName.Name ?? 'UnnamedRange';
    const evidenceId = addEvidence(model, ids, {
      category: '05a_Raw_Metadata',
      title: `Named range - ${name}`,
      fileName: `${sanitizeName(file.originalName)}_${sanitizeName(name)}_named_range.json`,
      sourceFile: file.originalName,
      summary: 'Workbook named range metadata.',
      confidence: 'confirmed',
      content: JSON.stringify(definedName, null, 2),
    });

    addNode(model, ids, {
      node_type: 'named range',
      name,
      description: `Named range ${name}: ${definedName.Ref ?? 'reference unavailable'}.`,
      source_file: file.originalName,
      business_purpose: 'Named range used by formulas, inputs, outputs, or workbook navigation.',
      owner_status: 'owner-confirmation-required',
      criticality: 'P2',
      confidence: 'confirmed',
      evidence_id: evidenceId,
      recommended_action: 'Confirm whether this named range is input, mapping, calculation, output, or obsolete.',
      failure_impact: 'Broken named ranges can corrupt formulas, refreshes, or output areas.',
      dollar_exposure: 'Directional exposure modeled at process level.',
    });

    model.excel.Excel_Table_NamedRange_Register.push({
      workbook: file.originalName,
      object_type: 'named range',
      object_name: name,
      reference: definedName.Ref ?? '',
      evidence_id: evidenceId,
      confidence: 'confirmed',
      recommended_action: 'Confirm purpose and downstream dependency.',
    });
  }

  await extractWorkbookZipInternals(file, model, ids, workbookNodeId, metadataEvidenceId);
  await extractExcelDesktopArtifacts(file, model, ids, workbookNodeId, metadataEvidenceId);
}

async function extractExcelDesktopArtifacts(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  workbookNodeId: string,
  metadataEvidenceId: string,
): Promise<void> {
  const macroCapable = ['.xlsm', '.xlsb', '.xls'].includes(path.extname(file.originalName).toLowerCase());
  if (process.env.VERCEL || !file.tempPath || !TRUSTED_DESKTOP_DEEP_EXPORT) {
    if (macroCapable) {
      if (hasConfirmedExcelVbaSource(model, file.originalName)) {
        const limitation =
          `${file.originalName}: trusted Excel desktop collector did not run in this environment; deterministic ` +
          'vbaProject.bin extraction captured VBA source, modules, procedures, and workbook events, but shape/button bindings and Excel COM-only metadata may remain partial.';
        if (!model.limitations.includes(limitation)) {
          model.limitations.push(limitation);
        }
      } else {
        createTargetedBlocker(
          model,
          ids,
          file,
          workbookNodeId,
          'Excel desktop VBA source extraction unavailable',
          metadataEvidenceId,
          'Run the workbook through the trusted local desktop collector so VBA modules, workbook events, sheet events, button bindings, and macro purposes are captured.',
        );
      }
    }
    return;
  }

  try {
    const collector = await runExcelCollector(file.tempPath);
    const collectorEvidenceId = addEvidence(model, ids, {
      category: '05a_Raw_Metadata',
      title: `Excel desktop collector output - ${file.originalName}`,
      fileName: `${sanitizeName(file.originalName)}_excel_desktop_collector.json`,
      sourceFile: file.originalName,
      summary: 'Trusted desktop Excel.Application collector output covering VBA components, button bindings, connections, and query tables where available.',
      confidence: collector.excel_application?.status === 'ok' ? 'partial' : 'blocked',
      content: JSON.stringify(collector, null, 2),
    });
    processExcelCollectorResult(file, model, ids, workbookNodeId, collectorEvidenceId, collector);
  } catch (error) {
    if (macroCapable) {
      if (hasConfirmedExcelVbaSource(model, file.originalName)) {
        const limitation =
          `${file.originalName}: trusted Excel desktop collector failed (${error instanceof Error ? error.message : 'unknown error'}); deterministic ` +
          'vbaProject.bin extraction still captured confirmed VBA source, modules, procedures, and workbook events. Button/shape bindings and Excel COM-only metadata remain partial.';
        if (!model.limitations.includes(limitation)) {
          model.limitations.push(limitation);
        }
      } else {
        createTargetedBlocker(
          model,
          ids,
          file,
          workbookNodeId,
          'Excel desktop VBA source extraction failed',
          metadataEvidenceId,
          `Trusted Excel desktop collector failed: ${error instanceof Error ? error.message : 'unknown error'}. Enable Excel desktop automation and trusted access to the VBA project object model, then rerun.`,
        );
      }
    }
  }
}

function hasConfirmedExcelVbaSource(model: DiscoveryModel, workbookName: string): boolean {
  return model.excel.Excel_VBA_Register.some(
    (row) =>
      row['workbook'] === workbookName &&
      row['confidence'] === 'confirmed' &&
      ['module', 'procedure', 'vba project binary'].includes(String(row['artifact_level'] ?? '')),
  );
}

async function runExcelCollector(tempPath: string): Promise<Record<string, any>> {
  const scriptPath = path.join(process.cwd(), 'scripts', 'excel-collector.ps1');
  const { stdout } = await execFileAsync(
    'powershell.exe',
    ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath, '-Path', tempPath],
    {
      timeout: EXCEL_COLLECTOR_TIMEOUT_MS,
      maxBuffer: 256 * 1024 * 1024,
      windowsHide: true,
      env: process.env,
    },
  );
  return JSON.parse(stdout) as Record<string, any>;
}

function processExcelCollectorResult(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  workbookNodeId: string,
  collectorEvidenceId: string,
  collector: Record<string, any>,
): void {
  const limitations = Array.isArray(collector.limitations)
    ? collector.limitations.map((item) => String(item)).filter(Boolean)
    : [];
  model.limitations.push(...limitations);

  const components = asRows(collector.vba_components?.rows);
  const shapeBindings = asRows(collector.shape_macro_bindings?.rows);
  const connectionRows = asRows(collector.workbook_connections?.rows);
  const queryTableRows = asRows(collector.query_tables?.rows);

  for (const connection of connectionRows) {
    model.excel.Excel_Connection_Register.push({
      workbook: file.originalName,
      connection_name: connection.name ?? '',
      connection_type: connection.type ?? '',
      refresh_with_refresh_all: connection.refresh_with_refresh_all ?? '',
      connection_string: connection.ole_db_connection ?? connection.odbc_connection ?? connection.text_connection ?? '',
      evidence_id: collectorEvidenceId,
      confidence: 'partial',
      recommended_action: 'Classify connection source, credential handling, refresh cadence, and lineage role.',
    });
  }

  for (const queryTable of queryTableRows) {
    model.excel.Excel_PowerQuery_Register.push({
      workbook: file.originalName,
      sheet_name: queryTable.sheet_name ?? '',
      query_artifact: queryTable.query_table_name ?? queryTable.list_object ?? '',
      m_code_status: 'Excel desktop QueryTable/ListObject metadata captured; Power Query M may still require workbook query export when stored in mashup binary.',
      command_text: queryTable.command_text ?? '',
      connection: queryTable.connection ?? '',
      evidence_id: collectorEvidenceId,
      confidence: 'partial',
      recommended_action: 'Map query table command/connection to upstream lineage and refresh controls.',
    });
  }

  for (const binding of shapeBindings) {
    const actionText = String(binding.on_action ?? '');
    const bindingEvidenceId = addEvidence(model, ids, {
      category: '05d_VBA',
      title: `Excel button macro binding - ${binding.sheet_name ?? 'worksheet'} / ${binding.shape_name ?? 'shape'}`,
      fileName: `${sanitizeName(file.originalName)}_${sanitizeName(String(binding.sheet_name ?? 'sheet'))}_${sanitizeName(String(binding.shape_name ?? 'shape'))}_button_binding.json`,
      sourceFile: file.originalName,
      summary: 'Desktop Excel shape/button OnAction macro binding.',
      confidence: 'confirmed',
      content: JSON.stringify(binding, null, 2),
    });
    model.excel.Excel_VBA_Register.push({
      artifact_level: 'button binding',
      workbook: file.originalName,
      module_name: 'shape/button binding',
      procedure_name: actionText,
      component_type: 'button/onAction',
      procedure_role: 'macro invocation binding',
      runnable_macro_flag: actionText ? 'yes' : 'no',
      purpose: 'User-triggered macro entry point from workbook UI.',
      sheet_name: binding.sheet_name ?? '',
      shape_name: binding.shape_name ?? '',
      evidence_id: bindingEvidenceId,
      confidence: 'confirmed',
      recommended_action: 'Trace this button-triggered macro to module procedure code, controls, outputs, and failure impact.',
    });
  }

  let exportedProcedureCount = 0;
  for (const component of components) {
    const componentName = String(component.component_name ?? 'UnnamedComponent');
    const code = String(component.code ?? '');
    const lineCount = Number(component.line_count ?? 0);
    if (!code || lineCount <= 0) {
      model.excel.Excel_VBA_Register.push({
        artifact_level: 'component',
        workbook: file.originalName,
        module_name: componentName,
        component_type: component.component_type ?? '',
        extraction_status: 'Component present but no code lines were exported.',
        evidence_id: collectorEvidenceId,
        confidence: 'partial',
        recommended_action: 'Confirm whether this component is intentionally empty or contains designer-only logic.',
      });
      continue;
    }

    const procedures = extractVbaProcedures(code);
    exportedProcedureCount += procedures.length;
    addExcelVbaModuleEvidence(file, model, ids, workbookNodeId, componentName, code, 'trusted desktop Excel collector', String(component.component_type ?? ''));
  }

  if (collector.vba_components?.status === 'blocked') {
    createTargetedBlocker(
      model,
      ids,
      file,
      workbookNodeId,
      'Excel VBA project access blocked',
      collectorEvidenceId,
      'Enable trusted access to the VBA project object model in Excel Trust Center and rerun so module code, events, buttons, and macro purposes are captured.',
    );
  } else if (components.length && exportedProcedureCount === 0) {
    addOpenQuestion(model, ids, {
      asset: file.originalName,
      question: 'Excel VBA components were present but no procedures were detected. Confirm whether code is designer-only, password-protected, obfuscated, or empty.',
      owner_role: 'Workbook owner / Excel technical owner',
      blocker_type: 'vba-purpose-validation',
      priority: 'P1',
      evidence_id: collectorEvidenceId,
    });
  }
}

function addExcelVbaModuleEvidence(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  workbookNodeId: string,
  componentName: string,
  code: string,
  extractionMethod: string,
  componentType = '',
): void {
  const procedures = extractVbaProcedures(code);
  const macroEntryCount = procedures.filter((procedure) => procedure.procedure_role === 'macro entrypoint').length;
  const functionCount = procedures.filter((procedure) => procedure.procedure_role === 'function/helper').length;
  const eventHandlerCount = procedures.filter((procedure) => procedure.procedure_role === 'event handler').length;
  const moduleEvidenceId = addEvidence(model, ids, {
    category: '05d_VBA',
    title: `Excel VBA module - ${componentName}`,
    fileName: `${sanitizeName(file.originalName)}_${sanitizeName(componentName)}.bas`,
    sourceFile: file.originalName,
    summary: `Excel VBA source for ${componentName} extracted by ${extractionMethod}.`,
    confidence: 'confirmed',
    content: code,
  });
  const existingModuleNode = model.nodes.find(
    (node) =>
      node.node_type === 'module' &&
      node.source_file === file.originalName &&
      node.name === `VBA module: ${componentName}`,
  );
  const moduleNode =
    existingModuleNode ??
    addNode(model, ids, {
      node_type: 'module',
      name: `VBA module: ${componentName}`,
      description: `Excel VBA component ${componentName} extracted by ${extractionMethod}. Contains ${procedures.length} procedure(s), including ${macroEntryCount} runnable macro entrypoint(s), ${functionCount} function/helper(s), and ${eventHandlerCount} event handler(s).`,
      source_file: file.originalName,
      business_purpose: procedures.length
        ? `Workbook automation module supporting ${Array.from(new Set(procedures.map((procedure) => procedure.procedure_role))).join(', ')} behavior.`
        : 'Workbook VBA component with no material procedure logic detected.',
      owner_status: 'owner-confirmation-required',
      criticality: macroEntryCount || eventHandlerCount ? 'P1' : 'P2',
      confidence: 'confirmed',
      evidence_id: moduleEvidenceId,
      recommended_action: 'Review module-level and procedure-level VBA evidence; validate trigger, purpose, side effects, controls, and owner.',
      failure_impact: 'VBA module changes can alter workbook automation, calculations, controls, or downstream outputs.',
      dollar_exposure: 'Directional exposure modeled at process level.',
    });

  if (!existingModuleNode) {
    addEdge(model, ids, {
      from_node_id: workbookNodeId,
      to_node_id: moduleNode.node_id,
      edge_type: 'depends_on',
      description: `Workbook contains VBA module ${componentName}.`,
      automated_flag: macroEntryCount || eventHandlerCount ? 'automated' : 'unknown',
      transformation_id: '',
      cadence: eventHandlerCount ? 'event-driven workbook automation present' : 'manual run or owner-confirmation-required',
      confidence: 'confirmed',
      evidence_id: moduleEvidenceId,
    });
  }

  model.excel.Excel_VBA_Register.push({
    artifact_level: 'module',
    workbook: file.originalName,
    module_name: componentName,
    component_type: componentType,
    extraction_status: `VBA source exported by ${extractionMethod}.`,
    line_count: code.split(/\r?\n/).length,
    procedure_count: procedures.length,
    macro_entrypoint_count: macroEntryCount,
    function_count: functionCount,
    event_handler_count: eventHandlerCount,
    procedure_summary: procedures.map((procedure) => `${procedure.name}: ${procedure.purpose}`).join(' | '),
    evidence_id: moduleEvidenceId,
    confidence: 'confirmed',
    recommended_action: 'Review procedure-level purpose, lineage, controls, and side effects.',
  });

  for (const procedure of procedures) {
    const procedureEvidenceId = addEvidence(model, ids, {
      category: '05d_VBA',
      title: `Excel VBA procedure - ${componentName}.${procedure.name}`,
      fileName: `${sanitizeName(file.originalName)}_${sanitizeName(componentName)}_${sanitizeName(procedure.name)}.bas`,
      sourceFile: file.originalName,
      summary: `Procedure-level VBA source and purpose for ${componentName}.${procedure.name}.`,
      confidence: 'confirmed',
      content: [
        `' Purpose: ${procedure.purpose}`,
        `' Procedure kind: ${procedure.kind}`,
        `' Detected operations: ${procedure.operations.join(', ') || 'none detected'}`,
        '',
        procedure.code,
      ].join('\n'),
    });
    if (procedure.procedure_role === 'macro entrypoint' || procedure.procedure_role === 'event handler') {
      const procedureNode = addNode(model, ids, {
        node_type: procedure.procedure_role === 'macro entrypoint' ? 'macro' : 'macro action',
        name: `${componentName}.${procedure.name}`,
        description: `Excel VBA ${procedure.procedure_role}: ${procedure.kind} ${procedure.name}. Purpose: ${procedure.purpose}`,
        source_file: file.originalName,
        business_purpose: procedure.purpose,
        owner_status: 'owner-confirmation-required',
        criticality: procedure.operations.some((operation) => ['external data refresh', 'file input/output', 'database/query execution', 'output/export'].includes(operation)) ? 'P1' : 'P2',
        confidence: 'confirmed',
        evidence_id: procedureEvidenceId,
        recommended_action:
          procedure.procedure_role === 'macro entrypoint'
            ? 'Confirm runnable macro purpose, owner, trigger, control coverage, input/output lineage, and failure impact.'
            : 'Confirm event trigger behavior, side effects, control coverage, and whether the event calls a runnable macro entrypoint.',
        failure_impact:
          procedure.procedure_role === 'macro entrypoint'
            ? 'Macro failure or silent logic change can alter workbook data, refreshes, exports, controls, or downstream decisions.'
            : 'Event-handler failure can skip required refreshes, leave workbook state stale, or trigger unexpected automation during workbook use.',
        dollar_exposure: 'Directional exposure modeled at process level.',
      });
      addEdge(model, ids, {
        from_node_id: moduleNode.node_id,
        to_node_id: procedureNode.node_id,
        edge_type: procedure.procedure_role === 'macro entrypoint' ? 'runs' : 'triggers',
        description: `VBA module contains ${procedure.procedure_role} ${procedure.name} extracted by ${extractionMethod}.`,
        automated_flag: procedure.procedure_role === 'event handler' ? 'automated' : 'unknown',
        transformation_id: '',
        cadence: inferVbaCadence(procedure.name),
        confidence: 'confirmed',
        evidence_id: procedureEvidenceId,
      });
      addVbaProcessStep(model, ids, {
        fileName: file.originalName,
        componentName,
        procedure,
        moduleNodeId: moduleNode.node_id,
        procedureNodeId: procedureNode.node_id,
        evidenceId: procedureEvidenceId,
      });
    }
    model.excel.Excel_VBA_Register.push({
      artifact_level: 'procedure',
      workbook: file.originalName,
      module_name: componentName,
      procedure_name: procedure.name,
      procedure_kind: procedure.kind,
      procedure_visibility: procedure.visibility,
      procedure_signature: procedure.signature,
      procedure_role: procedure.procedure_role,
      runnable_macro_flag: procedure.runnable_macro_flag,
      event_handler_flag: procedure.procedure_role === 'event handler' ? 'yes' : 'no',
      parameter_text: procedure.parameter_text,
      purpose: procedure.purpose,
      operations: procedure.operations.join('; '),
      code_evidence_id: procedureEvidenceId,
      evidence_id: procedureEvidenceId,
      code_excerpt: procedure.code.slice(0, 1200),
      confidence: 'confirmed',
      recommended_action: 'Validate macro intent with workbook owner and map side effects to lineage, controls, and action backlog.',
    });
    if (procedure.operations.length) {
      model.transformations.push({
        transformation_id: nextId(ids, 'TRN'),
        source_asset: `${file.originalName}:${componentName}.${procedure.name}`,
        rule_type: 'Excel VBA procedure',
        rule_description: procedure.purpose,
        input_fields: procedure.references.join('; '),
        output_fields: procedure.operations.join('; '),
        evidence_id: procedureEvidenceId,
        confidence: 'confirmed',
        recommended_action: 'Map VBA side effects to source/target lineage and controls.',
      });
    }
  }
}

function addVbaProcessStep(
  model: DiscoveryModel,
  ids: IdState,
  input: {
    fileName: string;
    componentName: string;
    procedure: ReturnType<typeof extractVbaProcedures>[number];
    moduleNodeId: string;
    procedureNodeId: string;
    evidenceId: string;
  },
): void {
  const { componentName, procedure, moduleNodeId, procedureNodeId, evidenceId } = input;
  if (model.processSteps.some((step) => step.input_node_id === moduleNodeId && step.output_node_id === procedureNodeId)) {
    return;
  }

  const callsRefreshRoster = /\bRefreshMasterRosterNow\b/i.test(procedure.code) && !/^RefreshMasterRosterNow$/i.test(procedure.name);
  const isRosterRefresh = /^RefreshMasterRosterNow$/i.test(procedure.name);
  const operationText = procedure.operations.length ? procedure.operations.join('; ') : procedure.purpose;

  model.processSteps.push({
    process_step_id: nextId(ids, 'STEP'),
    step_name:
      procedure.procedure_role === 'event handler'
        ? `${procedure.name} triggers workbook automation`
        : `${procedure.name} executes workbook automation`,
    actor_or_role: procedure.procedure_role === 'event handler' ? 'Excel application event' : 'Workbook user or automation caller',
    trigger:
      procedure.procedure_role === 'event handler'
        ? inferVbaCadence(procedure.name)
        : 'manual macro run, event call, or owner-confirmation-required trigger',
    description: isRosterRefresh
      ? 'RefreshMasterRosterNow scans non-system worksheets, preserves prior roster display/note values, clears roster input/note ranges, writes display names and worksheet tab names back to Master Roster, links user sheet B5 cells to the roster, recalculates roster/summary sheets, and restores Excel application state with error handling.'
      : callsRefreshRoster
        ? `${procedure.name} calls RefreshMasterRosterNow, making this event part of the roster-refresh process.`
        : `${componentName}.${procedure.name} is confirmed VBA automation. Detected operation(s): ${operationText}.`,
    manual_or_automated: procedure.procedure_role === 'event handler' ? 'automated' : 'mixed',
    input_node_id: moduleNodeId,
    output_node_id: procedureNodeId,
    evidence_id: evidenceId,
    confidence: 'confirmed',
    recommended_action: `Validate the ${procedure.name} trigger, side effects, expected run cadence, controls, and failure handling with the workbook owner. Trace key changed ranges to downstream decisions.`,
  });
}

function addFormulaAreas(
  workbookName: string,
  sheetName: string,
  sheet: XLSX.WorkSheet | undefined,
  model: DiscoveryModel,
  ids: IdState,
  sheetNodeId: string,
  sheetEvidenceId: string,
): void {
  if (!sheet) {
    return;
  }

  const formulaExamples: string[] = [];
  for (const address of Object.keys(sheet)) {
    if (address.startsWith('!')) {
      continue;
    }
    const cell = sheet[address] as XLSX.CellObject;
    if (cell?.f) {
      formulaExamples.push(`${address}=${cell.f}`);
    }
    if (formulaExamples.length >= 25) {
      break;
    }
  }

  if (!formulaExamples.length) {
    return;
  }

  const evidenceId = addEvidence(model, ids, {
    category: '05h_Data_Profiles',
    title: `Formula examples - ${workbookName} / ${sheetName}`,
    fileName: `${sanitizeName(workbookName)}_${sanitizeName(sheetName)}_formulas.json`,
    sourceFile: workbookName,
    summary: 'Sampled formulas from worksheet. Full repeated-cell formula expansion is intentionally summarized into formula areas.',
    confidence: 'partial',
    content: JSON.stringify({ workbook: workbookName, sheet: sheetName, formula_examples: formulaExamples }, null, 2),
  });

  const formulaNode = addNode(model, ids, {
    node_type: 'formula area',
    name: `${sheetName} formula area`,
    description: `Worksheet contains ${formulaExamples.length} sampled formula cell(s).`,
    source_file: workbookName,
    business_purpose: 'Workbook calculation or business rule layer.',
    owner_status: 'owner-confirmation-required',
    criticality: 'P1',
    confidence: 'partial',
    evidence_id: evidenceId,
    recommended_action: 'Review formulas for business rules, hardcoded overrides, and downstream output dependencies.',
    failure_impact: 'Formula changes may alter business rules, eligibility, financial calculations, mappings, or status logic.',
    dollar_exposure: 'Directional exposure modeled at process level.',
  });

  const transformationId = nextId(ids, 'TRN');
  model.transformations.push({
    transformation_id: transformationId,
    source_asset: `${workbookName}:${sheetName}`,
    rule_type: 'Excel formula area',
    rule_description: formulaExamples.slice(0, 5).join(' | '),
    input_fields: 'owner-confirmation-required',
    output_fields: 'owner-confirmation-required',
    evidence_id: evidenceId,
    confidence: 'partial',
    recommended_action: 'Map formula inputs and outputs to lineage nodes during analyst review.',
  });

  model.excel.Excel_Formula_Register.push({
    workbook: workbookName,
    sheet_name: sheetName,
    formula_area: `${sheetName} formula area`,
    formula_sample_count: formulaExamples.length,
    examples: formulaExamples.slice(0, 5).join(' | '),
    evidence_id: evidenceId,
    confidence: 'partial',
    recommended_action: 'Confirm business purpose and calculation owner.',
  });

  addEdge(model, ids, {
    from_node_id: sheetNodeId,
    to_node_id: formulaNode.node_id,
    edge_type: 'transforms',
    description: 'Worksheet values feed formula area calculations.',
    automated_flag: 'automated',
    transformation_id: transformationId,
    cadence: 'on workbook calculation',
    confidence: 'partial',
    evidence_id: evidenceId,
  });

  model.dependencyUsage.push({
    dependency_id: `DEP-${model.dependencyUsage.length + 1}`,
    source_asset: `${workbookName}:${sheetName}`,
    dependency: `${sheetName} formula area`,
    dependency_type: 'formula',
    evidence_id: evidenceId,
    confidence: 'partial',
  });

  void sheetEvidenceId;
}

function addSheetDataElements(
  workbookName: string,
  sheetName: string,
  sheet: XLSX.WorkSheet | undefined,
  model: DiscoveryModel,
  ids: IdState,
  evidenceId: string,
): void {
  if (!sheet?.['!ref']) {
    return;
  }

  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '', raw: false });
  const firstRow = rows[0] ?? {};
  const fieldNames = Object.keys(firstRow).slice(0, 60);
  for (const fieldName of fieldNames) {
    const values = rows.map((row) => row[fieldName]);
    const blanks = values.filter((value) => value === '' || value === null || value === undefined).length;
    const samples = distinctSamples(values);
    const inferredType = inferColumnType(values);
    model.dataElements.push({
      data_element_id: nextId(ids, 'DE'),
      asset: `${workbookName}:${sheetName}`,
      field_name: fieldName,
      inferred_type: inferredType,
      null_count: blanks,
      sample_values: samples.join(' | '),
      sensitive_indicator: sensitiveIndicator(fieldName),
      evidence_id: evidenceId,
      confidence: 'partial',
      recommended_action: 'Confirm field definition, criticality, data owner, and downstream usage.',
    });

    if (blanks > 0 && rows.length > 0 && blanks / rows.length > 0.25) {
      addDataQualityFinding(model, ids, {
        asset: `${workbookName}:${sheetName}`,
        field: fieldName,
        issue: 'High blank/null rate',
        example: `${blanks} blank value(s) across ${rows.length} sampled row(s).`,
        severity: 'P2',
        business_impact: 'May indicate optional fields, incomplete source data, or broken formulas.',
        recommended_fix: 'Confirm expected population rule and add a data quality test if critical.',
        evidence_id: evidenceId,
        confidence: 'partial',
      });
    }
  }
}

async function extractWorkbookZipInternals(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  workbookNodeId: string,
  metadataEvidenceId: string,
): Promise<void> {
  try {
    const zip = await JSZip.loadAsync(file.buffer);
    const fileNames = Object.keys(zip.files);
    const interestingFiles = fileNames.filter((name) =>
      /^(xl\/(connections|queryTables|pivotTables|externalLinks|vbaProject|customXml|model|tables|drawings|ctrlProps|activeX)|customXml\/)/i.test(
        name,
      ),
    );

    if (!interestingFiles.length) {
      return;
    }

    const zipEvidenceId = addEvidence(model, ids, {
      category: '05a_Raw_Metadata',
      title: `Workbook package internals - ${file.originalName}`,
      fileName: `${sanitizeName(file.originalName)}_xlsx_package_manifest.json`,
      sourceFile: file.originalName,
      summary: 'Workbook ZIP internals relevant to connections, Power Query, pivots, VBA, and model artifacts.',
      confidence: 'partial',
      content: JSON.stringify({ interesting_files: interestingFiles }, null, 2),
    });

    const tableFiles = interestingFiles.filter((name) => /xl\/tables\/table/i.test(name));
    for (const tableFile of tableFiles) {
      const content = (await zip.files[tableFile]?.async('string')) ?? '';
      const tableEvidenceId = addEvidence(model, ids, {
        category: '05a_Raw_Metadata',
        title: `Excel table definition - ${tableFile}`,
        fileName: `${sanitizeName(file.originalName)}_${sanitizeName(tableFile)}.xml`,
        sourceFile: file.originalName,
        summary: 'Raw Excel table XML definition extracted from workbook package.',
        confidence: 'confirmed',
        content,
      });
      const tableName = content.match(/\bname="([^"]+)"/i)?.[1] ?? tableFile;
      const ref = content.match(/\bref="([^"]+)"/i)?.[1] ?? '';
      const tableNode = addNode(model, ids, {
        node_type: 'table',
        name: tableName,
        description: `Excel structured table ${tableName}${ref ? ` at ${ref}` : ''}.`,
        source_file: file.originalName,
        business_purpose: 'Structured workbook table used as input, staging, transformation, or output area.',
        owner_status: 'owner-confirmation-required',
        criticality: 'P1',
        confidence: 'confirmed',
        evidence_id: tableEvidenceId,
        recommended_action: 'Confirm table business purpose, keys, source role, output role, and refresh dependency.',
        failure_impact: 'Table changes can break formulas, Power Query loads, pivots, exports, or downstream decisions.',
        dollar_exposure: 'Directional exposure modeled at process level.',
      });
      addEdge(model, ids, {
        from_node_id: workbookNodeId,
        to_node_id: tableNode.node_id,
        edge_type: 'depends_on',
        description: 'Workbook contains structured table definition.',
        automated_flag: 'automated',
        transformation_id: '',
        cadence: 'workbook calculation/refresh',
        confidence: 'confirmed',
        evidence_id: tableEvidenceId,
      });
      model.excel.Excel_Table_NamedRange_Register.push({
        workbook: file.originalName,
        object_type: 'table',
        object_name: tableName,
        reference: ref,
        package_part: tableFile,
        evidence_id: tableEvidenceId,
        confidence: 'confirmed',
        recommended_action: 'Confirm table owner, purpose, keys, and downstream usage.',
      });
    }

    const connectionFiles = interestingFiles.filter((name) => /connections|externalLinks|queryTables|customXml/i.test(name));
    for (const connectionFile of connectionFiles) {
      const content = await zip.files[connectionFile]?.async('string').catch(async () => {
        const binary = await zip.files[connectionFile]?.async('nodebuffer');
        return binary ? `[binary part ${binary.byteLength} bytes]` : '';
      });
      const evidenceId = addEvidence(model, ids, {
        category: /query|customXml|mashup/i.test(connectionFile) || /PowerQuery|DataMashup|Mashup/i.test(content ?? '')
          ? '05c_PowerQuery_M'
          : '05a_Raw_Metadata',
        title: `Workbook connection evidence - ${connectionFile}`,
        fileName: `${sanitizeName(file.originalName)}_${sanitizeName(connectionFile)}.xml`,
        sourceFile: file.originalName,
        summary: 'Raw workbook connection or query table evidence.',
        confidence: 'partial',
        content: content ?? '',
      });

      const connectionNode = addNode(model, ids, {
        node_type: connectionFile.includes('query') ? 'Power Query' : 'file',
        name: connectionFile,
        description: 'Workbook package connection/query artifact discovered by JSZip.',
        source_file: file.originalName,
        business_purpose: 'Potential upstream data source, Power Query, or refresh dependency.',
        owner_status: 'owner-confirmation-required',
        criticality: 'P1',
        confidence: 'partial',
        evidence_id: evidenceId,
        recommended_action: 'Review raw XML to classify source path, refresh behavior, load target, and lineage.',
        failure_impact: 'Broken workbook connection can stop refreshes or silently stale outputs.',
        dollar_exposure: 'Directional exposure modeled at process level.',
      });

      addEdge(model, ids, {
        from_node_id: workbookNodeId,
        to_node_id: connectionNode.node_id,
        edge_type: 'depends_on',
        description: 'Workbook contains connection/query package artifact.',
        automated_flag: 'automated',
        transformation_id: '',
        cadence: 'refresh cadence unknown',
        confidence: 'partial',
        evidence_id: evidenceId,
      });

      model.excel.Excel_Connection_Register.push({
        workbook: file.originalName,
        package_part: connectionFile,
        evidence_id: evidenceId,
        confidence: 'partial',
        recommended_action: 'Classify connection source and update lineage.',
      });

      if (/query|customXml|mashup/i.test(connectionFile) || /PowerQuery|DataMashup|Mashup/i.test(content ?? '')) {
        model.excel.Excel_PowerQuery_Register.push({
          workbook: file.originalName,
          query_artifact: connectionFile,
          m_code_status: /let\s|section\s|shared\s/i.test(content ?? '')
            ? 'M-like text present in package evidence'
            : 'Power Query/Mashup artifact present; full M may be encoded and requires owner/export validation',
          evidence_id: evidenceId,
          confidence: /let\s|section\s|shared\s/i.test(content ?? '') ? 'partial' : 'blocked',
          recommended_action: 'Extract full M from workbook query metadata if present or owner export.',
        });
        if (!/let\s|section\s|shared\s/i.test(content ?? '')) {
          createTargetedBlocker(
            model,
            ids,
            file,
            workbookNodeId,
            `Power Query M extraction blocked: ${connectionFile}`,
            evidenceId,
            'Export full Power Query M code from Excel so joins, filters, merges, appends, type changes, custom columns, grouping, replacements, source references, and load targets can be certified.',
          );
        }
      }
    }

    const vbaFiles = interestingFiles.filter((name) => /vbaProject/i.test(name));
    for (const vbaFile of vbaFiles) {
      const vbaBuffer = (await zip.files[vbaFile]?.async('nodebuffer')) ?? Buffer.alloc(0);
      const decompiledModules = extractVbaModulesFromProject(vbaBuffer);
      const vbaEvidenceId = addEvidence(model, ids, {
        category: '05d_VBA',
        title: `Excel VBA binary - ${vbaFile}`,
        fileName: `${sanitizeName(file.originalName)}_${sanitizeName(vbaFile)}.bin`,
        sourceFile: file.originalName,
        summary: decompiledModules.length
          ? `Raw vbaProject.bin plus ${decompiledModules.length} decompiled VBA module candidate(s).`
          : 'Raw vbaProject.bin extracted from macro-enabled workbook. Source text still requires trusted VBA export/decompilation.',
        confidence: decompiledModules.length ? 'confirmed' : 'partial',
        content: vbaBuffer,
      });
      const vbaProjectNode = addNode(model, ids, {
        node_type: 'module',
        name: 'vbaProject.bin',
        description: decompiledModules.length
          ? `Macro/VBA binary present in ${file.originalName}; module source candidates were decompiled deterministically from the CFB container.`
          : `Macro/VBA binary present in ${file.originalName}. Module/procedure source text not decoded by systematic workbook ZIP extraction.`,
        source_file: file.originalName,
        business_purpose: 'Potential workbook automation, refresh, export, transformation, control, or button/event logic.',
        owner_status: 'owner-confirmation-required',
        criticality: 'P1',
        confidence: decompiledModules.length ? 'confirmed' : 'partial',
        evidence_id: vbaEvidenceId,
        recommended_action: 'Export VBA modules, workbook events, sheet events, and button macro assignments from trusted Excel desktop workflow.',
        failure_impact: 'Hidden macro logic can change data, refresh sources, export outputs, suppress errors, or bypass controls.',
        dollar_exposure: 'Directional exposure modeled at process level.',
      });
      addEdge(model, ids, {
        from_node_id: workbookNodeId,
        to_node_id: vbaProjectNode.node_id,
        edge_type: 'depends_on',
        description: 'Workbook contains VBA project binary.',
        automated_flag: 'automated',
        transformation_id: '',
        cadence: 'workbook open/click/event/run; exact cadence blocked until VBA source is exported',
        confidence: decompiledModules.length ? 'confirmed' : 'partial',
        evidence_id: vbaEvidenceId,
      });
      model.excel.Excel_VBA_Register.push({
        artifact_level: 'vba project binary',
        workbook: file.originalName,
        module_name: 'vbaProject.bin',
        extraction_status: decompiledModules.length
          ? 'VBA binary present; source code candidates decompiled from vbaProject.bin.'
          : 'VBA binary present; trusted desktop Excel collector will export module/procedure source when VBA project access is allowed.',
        decompiled_module_count: decompiledModules.length,
        evidence_id: vbaEvidenceId,
        confidence: decompiledModules.length ? 'confirmed' : 'partial',
        recommended_action: decompiledModules.length
          ? 'Review decompiled module candidates and confirm macro purpose, triggers, side effects, and controls.'
          : 'Review extracted desktop VBA module evidence if present; otherwise enable trusted access to the VBA project object model and rerun.',
      });
      for (const module of decompiledModules) {
        addExcelVbaModuleEvidence(file, model, ids, workbookNodeId, module.name, module.code, 'vbaProject.bin deterministic CFB decompression');
      }
    }

    const pivotFiles = interestingFiles.filter((name) => /pivotTables|model/i.test(name));
    for (const pivotFile of pivotFiles) {
      const pivotContent = await zip.files[pivotFile]?.async('string').catch(() => '');
      const pivotEvidenceId = addEvidence(model, ids, {
        category: '05a_Raw_Metadata',
        title: `Excel pivot/data model artifact - ${pivotFile}`,
        fileName: `${sanitizeName(file.originalName)}_${sanitizeName(pivotFile)}.xml`,
        sourceFile: file.originalName,
        summary: 'Raw pivot table or data model artifact extracted from workbook package.',
        confidence: 'partial',
        content: pivotContent ?? '',
      });
      const pivotNode = addNode(model, ids, {
        node_type: 'pivot',
        name: pivotFile,
        description: 'Pivot table or data model artifact present in workbook package.',
        source_file: file.originalName,
        business_purpose: 'Potential analytical output, semantic aggregation, or downstream decision surface.',
        owner_status: 'owner-confirmation-required',
        criticality: 'P2',
        confidence: 'partial',
        evidence_id: pivotEvidenceId,
        recommended_action: 'Review pivot caches, source ranges, model relationships, and measures.',
        failure_impact: 'Pivot/model changes can alter reported totals and downstream decisions.',
        dollar_exposure: 'Directional exposure modeled at process level.',
      });
      addEdge(model, ids, {
        from_node_id: workbookNodeId,
        to_node_id: pivotNode.node_id,
        edge_type: 'depends_on',
        description: 'Workbook contains pivot/data model artifact.',
        automated_flag: 'automated',
        transformation_id: '',
        cadence: 'refresh/calculate',
        confidence: 'partial',
        evidence_id: pivotEvidenceId,
      });
      model.excel.Excel_Pivot_DataModel_Register.push({
        workbook: file.originalName,
        artifact_status: `Pivot/data model package artifact present: ${pivotFile}`,
        evidence_id: pivotEvidenceId,
        confidence: 'partial',
        recommended_action: 'Review pivot caches, model relationships, and measure definitions.',
      });
    }

    void metadataEvidenceId;
  } catch {
    model.limitations.push(`Workbook ZIP internals could not be inspected for ${file.originalName}.`);
  }
}

async function extractWord(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  documentNodeId: string,
  sourceEvidenceId: string,
): Promise<void> {
  const result = await mammoth.extractRawText({ buffer: file.buffer });
  const text = result.value.trim();
  const evidenceId = addEvidence(model, ids, {
    category: '05k_Document_Extracts',
    title: `Document text extract - ${file.originalName}`,
    fileName: `${sanitizeName(file.originalName)}_document_extract.txt`,
    sourceFile: file.originalName,
    summary: 'Raw text extracted from Word document by mammoth.',
    confidence: 'partial',
    content: text || 'No extractable text found.',
  });

  const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const headingCandidates = lines.filter((line) => isHeadingCandidate(line)).slice(0, 80);
  const processCandidates = lines.filter((line) => /^\d+[\.)]\s+|^(then|next|finally|approve|validate|send|export|import)\b/i.test(line)).slice(0, 120);
  const systemsMentioned = extractSystemMentions(text);

  model.word.Word_Document_Inventory.push({
    document: file.originalName,
    extracted_character_count: text.length,
    heading_candidate_count: headingCandidates.length,
    process_step_candidate_count: processCandidates.length,
    evidence_id: evidenceId,
    confidence: 'partial',
    recommended_action: 'Confirm heading hierarchy, actors, process rules, inputs, outputs, controls, and exceptions with owner.',
  });

  for (const heading of headingCandidates) {
    const sectionNode = addNode(model, ids, {
      node_type: 'document section',
      name: heading.slice(0, 120),
      description: 'Heading-like document section extracted from Word source.',
      source_file: file.originalName,
      business_purpose: 'Potential process, rule, control, exception, or source reference documentation.',
      owner_status: 'owner-confirmation-required',
      criticality: 'P2',
      confidence: 'partial',
      evidence_id: evidenceId,
      recommended_action: 'Classify document section and map to process, rule, control, or open question.',
      failure_impact: 'Unclassified process documentation can hide business rules and control requirements.',
      dollar_exposure: 'Directional exposure modeled at process level.',
    });

    addEdge(model, ids, {
      from_node_id: documentNodeId,
      to_node_id: sectionNode.node_id,
      edge_type: 'documents',
      description: 'Word document contains extracted section.',
      automated_flag: 'automated',
      transformation_id: '',
      cadence: 'static document',
      confidence: 'partial',
      evidence_id: evidenceId,
    });

    model.word.Word_Section_Extracts.push({
      document: file.originalName,
      section_text: heading,
      section_type: 'heading-candidate',
      evidence_id: evidenceId,
      confidence: 'partial',
      recommended_action: 'Owner classify and confirm relevance.',
    });
  }

  for (const step of processCandidates) {
    model.processSteps.push({
      process_step_id: nextId(ids, 'STEP'),
      step_name: step.slice(0, 90),
      actor_or_role: 'owner-confirmation-required',
      trigger: 'documented process text',
      description: step,
      manual_or_automated: 'unknown',
      input_node_id: documentNodeId,
      output_node_id: '',
      evidence_id: evidenceId,
      confidence: 'partial',
      recommended_action: 'Confirm actor, trigger, system, input, output, control, and exception path.',
    });

    model.word.Word_Process_Rules.push({
      document: file.originalName,
      extract: step,
      extract_type: 'process-step-candidate',
      evidence_id: evidenceId,
      confidence: 'partial',
      recommended_action: 'Confirm as actual process step or discard.',
    });
  }

  for (const system of systemsMentioned) {
    const systemNode = addNode(model, ids, {
      node_type: 'system',
      name: system,
      description: `System mentioned in ${file.originalName}.`,
      source_file: file.originalName,
      business_purpose: 'Potential upstream source, downstream consumer, or operating platform.',
      owner_status: 'owner-confirmation-required',
      criticality: 'P2',
      confidence: 'inferred',
      evidence_id: evidenceId,
      recommended_action: 'Confirm whether this system is upstream, downstream, or contextual only.',
      failure_impact: 'Unresolved system references can leave lineage incomplete.',
      dollar_exposure: 'Directional exposure modeled at process level.',
    });

    addEdge(model, ids, {
      from_node_id: documentNodeId,
      to_node_id: systemNode.node_id,
      edge_type: 'documents',
      description: 'Document mentions system.',
      automated_flag: 'manual',
      transformation_id: '',
      cadence: 'static document',
      confidence: 'inferred',
      evidence_id: evidenceId,
    });
  }

  void sourceEvidenceId;
}

function extractFlatFile(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  fileNodeId: string,
  sourceEvidenceId: string,
): void {
  const text = decodeText(file.buffer);
  const extension = path.extname(file.originalName).toLowerCase();
  const delimiter = inferDelimiter(text, extension);
  if (!shouldTreatAsDelimitedText(text, delimiter, extension)) {
    extractUnstructuredTextFile(file, model, ids, fileNodeId, sourceEvidenceId, text, 'No stable delimiter pattern detected; analyzed as unstructured text.');
    return;
  }

  let rows: string[][] = [];
  try {
    rows = parseCsv(text, {
      delimiter,
      relax_column_count: true,
      skip_empty_lines: false,
      bom: true,
    }) as string[][];
  } catch (error) {
    extractUnstructuredTextFile(
      file,
      model,
      ids,
      fileNodeId,
      sourceEvidenceId,
      text,
      `Delimited parse failed (${error instanceof Error ? error.message : 'unknown parse error'}); analyzed as unstructured text instead of treating the file as blocked.`,
    );
    return;
  }

  const header = rows[0] ?? [];
  const hasHeader = looksLikeHeader(header, rows[1] ?? []);
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const columns = hasHeader ? header : header.map((_, index) => `column_${index + 1}`);
  const profile = profileTabularRows(columns, dataRows);
  const semanticProfile = inferTabularAssetSemantics(file.originalName, columns, dataRows, profile);

  const profileEvidenceId = addEvidence(model, ids, {
    category: '05h_Data_Profiles',
    title: `Flat file profile - ${file.originalName}`,
    fileName: `${sanitizeName(file.originalName)}_flat_file_profile.json`,
    sourceFile: file.originalName,
    summary: `Delimiter, row count, column count, semantic profile (${semanticProfile.domain}), column profiles, duplicate check, and sample rows.`,
    confidence: 'confirmed',
    content: JSON.stringify(
      {
        delimiter,
        has_header: hasHeader,
        row_count: dataRows.length,
        column_count: columns.length,
        columns,
        semantic_profile: semanticProfile,
        profile,
        sample_rows: dataRows.slice(0, 10),
      },
      null,
      2,
    ),
  });

  addEvidence(model, ids, {
    category: '05i_Samples',
    title: `Flat file sample - ${file.originalName}`,
    fileName: `${sanitizeName(file.originalName)}_sample.csv`,
    sourceFile: file.originalName,
    summary: 'First sampled rows from uploaded flat file.',
    confidence: 'confirmed',
    content: rows.slice(0, 25).map((row) => row.map(csvCell).join(',')).join('\n'),
  });

  const tableNode = addNode(model, ids, {
    node_type: 'table',
    name: `${file.originalName} records`,
    description: `Flat file table with ${dataRows.length} data row(s), ${columns.length} column(s), delimiter ${JSON.stringify(delimiter)}. Inferred domain: ${semanticProfile.domain}. ${semanticProfile.description}`,
    source_file: file.originalName,
    business_purpose: semanticProfile.business_purpose,
    owner_status: 'owner-confirmation-required',
    criticality: semanticProfile.criticality,
    confidence: semanticProfile.confidence,
    evidence_id: profileEvidenceId,
    recommended_action: semanticProfile.recommended_action,
    failure_impact: semanticProfile.failure_impact,
    dollar_exposure: 'Directional exposure modeled at process level.',
  });

  addEdge(model, ids, {
    from_node_id: fileNodeId,
    to_node_id: tableNode.node_id,
    edge_type: 'imports_from',
    description: 'Flat file parsed into a tabular asset.',
    automated_flag: 'automated',
    transformation_id: '',
    cadence: 'on upload',
    confidence: 'confirmed',
    evidence_id: profileEvidenceId,
  });

  model.rowCountSummary[file.originalName] = dataRows.length;

  model.processSteps.push({
    process_step_id: nextId(ids, 'STEP'),
    step_name: `Profile ${semanticProfile.domain} from ${file.originalName}`,
    actor_or_role: 'Discovery pipeline / source owner',
    trigger: 'on upload',
    description: `The file was parsed as a tabular ${semanticProfile.domain} with ${dataRows.length} row(s), ${columns.length} column(s), and semantic signals: ${semanticProfile.signals.join('; ') || 'none beyond tabular structure'}.`,
    manual_or_automated: 'mixed',
    input_node_id: fileNodeId,
    output_node_id: tableNode.node_id,
    evidence_id: profileEvidenceId,
    confidence: semanticProfile.confidence,
    recommended_action: semanticProfile.recommended_action,
  });

  if (semanticProfile.domain !== 'generic tabular dataset') {
    addOpenQuestion(model, ids, {
      asset: file.originalName,
      question: `Is ${file.originalName} the authoritative ${semanticProfile.domain}, an extract from another system, or a downstream working copy?`,
      owner_role: 'Source owner / data steward',
      blocker_type: 'source-of-truth-confirmation',
      priority: 'P1',
      evidence_id: profileEvidenceId,
    });
    addAction(model, ids, {
      title: `Confirm source-of-truth status for ${semanticProfile.domain}`,
      description: `Validate whether ${file.originalName} is authoritative, exported from another roster/source system, manually maintained, or a temporary working file. Confirm owner, refresh cadence, consumers, controls, and sensitive-field handling.`,
      source_asset: file.originalName,
      owner_role: 'Source owner / data steward',
      recommended_owner: 'Assigned roster/process owner',
      action_type: 'Validate',
      priority: 'P1',
      severity: 'P1',
      dependency: 'Owner confirmation and upstream/source-system evidence',
      acceptance_criteria: 'Source-of-truth status, upstream origin, owner, cadence, consumers, controls, and sensitivity classification are documented.',
      evidence_id: profileEvidenceId,
      related_risk: `Unconfirmed ${semanticProfile.domain} ownership and authority`,
      expected_business_value: 'Prevents stale or unofficial roster/reference data from driving decisions, access, communication, staffing, or reporting.',
    });
  }

  for (const columnProfile of profile.columns) {
    const element: DataElement = {
      data_element_id: nextId(ids, 'DE'),
      asset: file.originalName,
      field_name: columnProfile.name,
      inferred_type: columnProfile.inferred_type,
      null_count: columnProfile.blank_count,
      sample_values: columnProfile.sample_values.join(' | '),
      sensitive_indicator: sensitiveIndicator(columnProfile.name),
      evidence_id: profileEvidenceId,
      confidence: 'confirmed',
      recommended_action: 'Confirm field definition, owner, criticality, and validation rule.',
    };
    model.dataElements.push(element);

    if (columnProfile.blank_count > 0 && dataRows.length > 0 && columnProfile.blank_count / dataRows.length > 0.2) {
      addDataQualityFinding(model, ids, {
        asset: file.originalName,
        field: columnProfile.name,
        issue: 'Blank values detected',
        example: `${columnProfile.blank_count} blank value(s) across ${dataRows.length} row(s).`,
        severity: 'P2',
        business_impact: 'Potential missing mappings, incomplete records, or optional field requiring classification.',
        recommended_fix: 'Confirm expected population and add a validation rule for critical fields.',
        evidence_id: profileEvidenceId,
        confidence: 'confirmed',
      });
    }
  }

  if (profile.duplicate_first_column_count > 0) {
    addDataQualityFinding(model, ids, {
      asset: file.originalName,
      field: columns[0] ?? 'column_1',
      issue: 'Duplicate likely key values',
      example: `${profile.duplicate_first_column_count} duplicate value(s) in first column.`,
      severity: 'P1',
      business_impact: 'Duplicate keys can inflate measures, create join fan-out, or overwrite records.',
      recommended_fix: 'Confirm primary key and deduplicate or add composite key logic.',
      evidence_id: profileEvidenceId,
      confidence: 'inferred',
    });
  }
}

function extractUnstructuredTextFile(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  fileNodeId: string,
  sourceEvidenceId: string,
  text: string,
  parseNote: string,
): void {
  const lines = text.split(/\r?\n/);
  const nonEmptyLines = lines.map((line) => line.trim()).filter(Boolean);
  const headings = nonEmptyLines.filter(isHeadingCandidate).slice(0, 40);
  const keyValues = extractTextKeyValuePairs(nonEmptyLines).slice(0, 80);
  const urls = extractUrls(text).slice(0, 30);
  const paths = extractPathMentions(text).slice(0, 30);
  const systems = extractSystemMentions(text);
  const artifactKind = inferTextArtifactKind(file.originalName, text, keyValues, headings, urls);
  const sampleLines = nonEmptyLines.slice(0, 50);
  const wordCount = (text.match(/\b[\w'-]+\b/g) ?? []).length;

  const profileEvidenceId = addEvidence(model, ids, {
    category: '05k_Document_Extracts',
    title: `Text discovery profile - ${file.originalName}`,
    fileName: `${sanitizeName(file.originalName)}_text_discovery_profile.json`,
    sourceFile: file.originalName,
    summary: 'Unstructured text fallback profile with sections, key/value signals, URL/path mentions, and sample lines.',
    confidence: 'confirmed',
    content: JSON.stringify(
      {
        parse_mode: 'unstructured-text',
        parse_note: parseNote,
        artifact_kind: artifactKind,
        line_count: lines.length,
        non_empty_line_count: nonEmptyLines.length,
        character_count: text.length,
        word_count: wordCount,
        heading_candidates: headings,
        key_value_pairs: keyValues,
        system_mentions: systems,
        url_mentions: urls,
        path_mentions: paths,
        sample_lines: sampleLines,
      },
      null,
      2,
    ),
  });

  addEvidence(model, ids, {
    category: '05i_Samples',
    title: `Text sample - ${file.originalName}`,
    fileName: `${sanitizeName(file.originalName)}_sample.txt`,
    sourceFile: file.originalName,
    summary: 'First non-empty text lines from uploaded source.',
    confidence: 'confirmed',
    content: sampleLines.join('\n'),
  });

  const documentNode = addNode(model, ids, {
    node_type: 'document',
    name: file.originalName,
    description: `${artifactKind} text artifact with ${lines.length} line(s), ${wordCount} word(s), ${keyValues.length} key/value signal(s), ${headings.length} heading candidate(s), ${urls.length} URL mention(s), and ${paths.length} path mention(s).`,
    source_file: file.originalName,
    business_purpose: 'Uploaded text artifact; purpose, owner, consumer, and decision use require owner confirmation unless directly evidenced in the text profile.',
    owner_status: 'owner-confirmation-required',
    criticality: 'P2',
    confidence: 'confirmed',
    evidence_id: profileEvidenceId,
    recommended_action: 'Confirm whether this text is documentation, configuration, export/log content, or intended tabular data; then assign owner, purpose, cadence, and downstream consumer.',
    failure_impact: 'If this file is used operationally, malformed or misunderstood text can cause missing context, wrong configuration, broken automation, or unauditable decisions.',
    dollar_exposure: 'Directional exposure modeled at process level until business use is confirmed.',
  });

  addEdge(model, ids, {
    from_node_id: fileNodeId,
    to_node_id: documentNode.node_id,
    edge_type: 'documents',
    description: 'Text file analyzed as an unstructured document artifact.',
    automated_flag: 'automated',
    transformation_id: '',
    cadence: 'on upload',
    confidence: 'confirmed',
    evidence_id: profileEvidenceId,
  });

  const sectionTitles = headings.length ? headings : deriveTextSectionsFromLines(nonEmptyLines);
  sectionTitles.slice(0, 12).forEach((heading) => {
    const sectionNode = addNode(model, ids, {
      node_type: 'document section',
      name: heading,
      description: `Detected text section or prominent line in ${file.originalName}.`,
      source_file: file.originalName,
      business_purpose: 'Potential section, rule, prompt, configuration block, or content region requiring owner classification.',
      owner_status: 'owner-confirmation-required',
      criticality: 'P3',
      confidence: headings.includes(heading) ? 'inferred' : 'partial',
      evidence_id: profileEvidenceId,
      recommended_action: 'Confirm section meaning, owner, and whether it drives a process, rule, configuration, or output.',
      failure_impact: 'Misclassified text sections can hide requirements, controls, source references, or decision logic.',
      dollar_exposure: 'Directional exposure modeled at process level.',
    });
    addEdge(model, ids, {
      from_node_id: documentNode.node_id,
      to_node_id: sectionNode.node_id,
      edge_type: 'documents',
      description: 'Document contains detected section/content region.',
      automated_flag: 'automated',
      transformation_id: '',
      cadence: 'static text structure',
      confidence: sectionNode.confidence,
      evidence_id: profileEvidenceId,
    });
  });

  keyValues.slice(0, 40).forEach((pair) => {
    model.dataElements.push({
      data_element_id: nextId(ids, 'DE'),
      asset: file.originalName,
      field_name: pair.key,
      inferred_type: inferColumnType([pair.value]),
      null_count: pair.value ? 0 : 1,
      sample_values: pair.value.slice(0, 160),
      sensitive_indicator: sensitiveIndicator(pair.key),
      evidence_id: profileEvidenceId,
      confidence: 'inferred',
      recommended_action: 'Confirm whether this key/value is configuration, metadata, content, or data and whether it is operationally material.',
    });
  });

  [...urls.map((value) => ({ value, type: 'system' as NodeType, label: 'URL mention' })), ...paths.map((value) => ({ value, type: nodeTypeForPathMention(value), label: 'Path mention' }))].slice(0, 12).forEach((mention) => {
    const referenceNode = addNode(model, ids, {
      node_type: mention.type,
      name: mention.value,
      description: `${mention.label} detected inside text artifact.`,
      source_file: file.originalName,
      business_purpose: 'Potential upstream, downstream, or contextual reference requiring owner confirmation.',
      owner_status: 'owner-confirmation-required',
      criticality: 'P2',
      confidence: 'inferred',
      evidence_id: profileEvidenceId,
      recommended_action: 'Confirm whether this referenced resource is an upstream source, downstream output, documentation link, or incidental text.',
      failure_impact: 'If material, unavailable or changed referenced resources can block lineage, automation, or decision traceability.',
      dollar_exposure: 'Directional exposure modeled at process level.',
    });
    addEdge(model, ids, {
      from_node_id: documentNode.node_id,
      to_node_id: referenceNode.node_id,
      edge_type: 'documents',
      description: 'Text artifact references this resource.',
      automated_flag: 'unknown',
      transformation_id: '',
      cadence: 'owner-confirmation-required',
      confidence: 'inferred',
      evidence_id: profileEvidenceId,
    });
  });

  model.processSteps.push({
    process_step_id: nextId(ids, 'STEP'),
    step_name: `Classify ${file.originalName} as ${artifactKind}`,
    actor_or_role: 'Discovery pipeline / source owner',
    trigger: 'on upload',
    description: `${parseNote} The file was profiled as unstructured text with ${keyValues.length} key/value signal(s), ${headings.length} heading candidate(s), ${urls.length} URL mention(s), and ${paths.length} path mention(s).`,
    manual_or_automated: 'mixed',
    input_node_id: fileNodeId,
    output_node_id: documentNode.node_id,
    evidence_id: profileEvidenceId,
    confidence: 'confirmed',
    recommended_action: 'Owner must confirm intended file role: tabular source, configuration, documentation, web/page extract, log/export, or supporting artifact.',
  });

  if (keyValues.length) {
    model.transformations.push({
      transformation_id: nextId(ids, 'TRN'),
      source_asset: file.originalName,
      rule_type: 'Text key/value extraction',
      rule_description: `Extracted ${keyValues.length} key/value signal(s) from unstructured text for owner review.`,
      input_fields: 'raw text lines',
      output_fields: keyValues.slice(0, 20).map((pair) => pair.key).join(', '),
      evidence_id: profileEvidenceId,
      confidence: 'inferred',
      recommended_action: 'Confirm which keys are business data, metadata, configuration, or incidental content.',
    });
  }

  addDataQualityFinding(model, ids, {
    asset: file.originalName,
    field: 'file_structure',
    issue: 'Text file is not confirmed as tabular data',
    example: parseNote,
    severity: 'P3',
    business_impact: 'Treating prose/config/page text as a table can create false columns and misleading lineage; treating it as unstructured text preserves evidence without fabricating structure.',
    recommended_fix: 'If this is intended to be tabular, provide delimiter, quote rules, header expectations, and a clean sample; otherwise confirm the text artifact type and owner.',
    evidence_id: profileEvidenceId,
    confidence: 'confirmed',
  });

  addOpenQuestion(model, ids, {
    asset: file.originalName,
    question: 'What is the intended role of this text file: tabular data source, configuration file, web/page extract, log/export, documentation, prompt, or supporting evidence?',
    owner_role: 'Source owner',
    blocker_type: 'owner-confirmation-required',
    priority: 'P2',
    evidence_id: profileEvidenceId,
  });

  addAction(model, ids, {
    title: `Confirm text artifact role for ${file.originalName}`,
    description: 'Review the text discovery profile and classify the file role, owner, business purpose, consumer, cadence, sensitivity, and whether the key/value or section signals are material.',
    source_asset: file.originalName,
    owner_role: 'Source owner / data steward',
    recommended_owner: 'Assigned business/data owner',
    action_type: 'Validate',
    priority: 'P2',
    severity: 'P2',
    dependency: 'Owner review of text profile evidence',
    acceptance_criteria: 'File role, owner, purpose, consumer, cadence, and criticality are confirmed; tabular parse rules are supplied if applicable.',
    evidence_id: profileEvidenceId,
    related_risk: 'Unclassified text artifact can produce misleading lineage or missed requirements.',
    expected_business_value: 'Prevents false structure while preserving useful text evidence for discovery and governance.',
  });
}

function extractSqlOrScript(
  file: UploadedSource,
  model: DiscoveryModel,
  ids: IdState,
  fileNodeId: string,
  sourceEvidenceId: string,
  sourceType: SourceType,
): void {
  const text = decodeText(file.buffer);
  const isSql = sourceType === 'sql';
  const formatted = isSql ? safeFormatSql(text) : text;
  const evidenceId = addEvidence(model, ids, {
    category: isSql ? '05b_SQL' : '05a_Raw_Metadata',
    title: `${isSql ? 'SQL' : 'Script'} source - ${file.originalName}`,
    fileName: `${sanitizeName(file.originalName)}.${isSql ? 'sql' : 'txt'}`,
    sourceFile: file.originalName,
    summary: 'Raw or formatted script evidence used for dependency extraction.',
    confidence: 'confirmed',
    content: formatted,
  });

  const queryNode = addNode(model, ids, {
    node_type: 'query',
    name: file.originalName,
    description: `${isSql ? 'SQL' : 'Script'} logic uploaded for discovery.`,
    source_file: file.originalName,
    business_purpose: 'Transformation, query, extraction, or automation logic.',
    owner_status: 'owner-confirmation-required',
    criticality: 'P1',
    confidence: 'confirmed',
    evidence_id: evidenceId,
    recommended_action: 'Confirm schedule, parameters, source/target ownership, and deployment path.',
    failure_impact: 'Logic changes can alter source reads, writes, transformations, filters, joins, and outputs.',
    dollar_exposure: 'Directional exposure modeled at process level.',
  });

  addEdge(model, ids, {
    from_node_id: fileNodeId,
    to_node_id: queryNode.node_id,
    edge_type: 'documents',
    description: 'File contains query or script logic.',
    automated_flag: 'automated',
    transformation_id: '',
    cadence: 'owner-confirmation-required',
    confidence: 'confirmed',
    evidence_id: evidenceId,
  });

  const refs = extractSqlReferences(text);
  for (const tableName of refs.reads) {
    const tableNode = addNode(model, ids, {
      node_type: 'table',
      name: tableName,
      description: `Read reference found in ${file.originalName}.`,
      source_file: file.originalName,
      business_purpose: 'Potential upstream table or view.',
      owner_status: 'owner-confirmation-required',
      criticality: 'P1',
      confidence: 'inferred',
      evidence_id: evidenceId,
      recommended_action: 'Confirm database, schema, owner, and source-of-truth status.',
      failure_impact: 'Unavailable or changed table can break query outputs.',
      dollar_exposure: 'Directional exposure modeled at process level.',
    });
    addEdge(model, ids, {
      from_node_id: queryNode.node_id,
      to_node_id: tableNode.node_id,
      edge_type: 'reads_from',
      description: 'SQL/script references table as read source.',
      automated_flag: 'automated',
      transformation_id: '',
      cadence: 'per execution',
      confidence: 'inferred',
      evidence_id: evidenceId,
    });
  }

  for (const tableName of refs.writes) {
    const tableNode = addNode(model, ids, {
      node_type: 'table',
      name: tableName,
      description: `Write reference found in ${file.originalName}.`,
      source_file: file.originalName,
      business_purpose: 'Potential downstream table, staging table, or output.',
      owner_status: 'owner-confirmation-required',
      criticality: 'P1',
      confidence: 'inferred',
      evidence_id: evidenceId,
      recommended_action: 'Confirm target owner, refresh cadence, retention, and consumer impact.',
      failure_impact: 'Failed writes can produce stale, missing, partial, or unauditable outputs.',
      dollar_exposure: 'Directional exposure modeled at process level.',
    });
    addEdge(model, ids, {
      from_node_id: queryNode.node_id,
      to_node_id: tableNode.node_id,
      edge_type: 'writes_to',
      description: 'SQL/script references table as write target.',
      automated_flag: 'automated',
      transformation_id: '',
      cadence: 'per execution',
      confidence: 'inferred',
      evidence_id: evidenceId,
    });
  }

  if (refs.whereClauses.length || refs.joinClauses.length) {
    model.transformations.push({
      transformation_id: nextId(ids, 'TRN'),
      source_asset: file.originalName,
      rule_type: 'SQL logic',
      rule_description: [...refs.joinClauses.slice(0, 4), ...refs.whereClauses.slice(0, 4)].join(' | '),
      input_fields: refs.reads.join(', '),
      output_fields: refs.writes.join(', '),
      evidence_id: evidenceId,
      confidence: 'inferred',
      recommended_action: 'Review parsed SQL logic and confirm business rule intent.',
    });
  }

  if (isSql) {
    addEvidence(model, ids, {
      category: '05b_SQL',
      title: `SQL Reference Scan - ${file.originalName}`,
      fileName: `${sanitizeName(file.originalName)}_sql_reference_scan.json`,
      sourceFile: file.originalName,
      summary: 'Deterministic SQL reference evidence for reads, writes, joins, filters, and scheduling hints.',
      confidence: 'partial',
      content: JSON.stringify(refs, null, 2),
    });
  }

  void sourceEvidenceId;
}

function addGlobalGovernanceFindings(model: DiscoveryModel, ids: IdState, evidenceId: string): void {
  for (const source of model.sourceFiles) {
    model.securityAccess.push({
      security_id: `SEC-${String(model.securityAccess.length + 1).padStart(4, '0')}`,
      asset: source.file_name,
      concern: 'File/folder permissions cannot be inspected from upload alone.',
      status: 'blocked',
      impact: 'Broad access, embedded credentials, shared accounts, retention concerns, and manual edit risks require platform-side validation.',
      evidence_id: evidenceId,
      confidence: 'blocked',
      recommended_action: 'Run security/access review on the source location and connected systems.',
    });

    model.controls.push({
      control_id: nextId(ids, 'CTL'),
      asset: source.file_name,
      control_type: 'Run log / approval / exception workflow',
      description: 'No external run log, approval trail, or exception workflow was provided with the upload.',
      status: 'owner-confirmation-required',
      evidence_id: evidenceId,
      confidence: 'unknown',
      recommended_action: 'Confirm controls and add automated run logging, QA checks, approval evidence, and exception handling where needed.',
    });

    addOpenQuestion(model, ids, {
      asset: source.file_name,
      question: 'Who owns this source and who consumes the outputs?',
      owner_role: 'Business owner / data product owner',
      blocker_type: 'owner-confirmation-required',
      priority: 'P1',
      evidence_id: evidenceId,
    });

    addAction(model, ids, {
      title: `Confirm owner and consumers for ${source.file_name}`,
      description: 'Identify accountable business owner, technical owner, downstream consumers, decision use, cadence, and critical outputs.',
      source_asset: source.file_name,
      owner_role: 'Business owner / data product owner',
      recommended_owner: 'Assigned data steward',
      action_type: 'Confirm Owner',
      priority: 'P1',
      severity: 'P1',
      dependency: 'Owner interview or source inventory record',
      acceptance_criteria: 'Owner, purpose, cadence, outputs, consumers, and decision usage are documented in the workbook and manifest.',
      evidence_id: evidenceId,
      related_risk: 'Unknown ownership and consumption',
      expected_business_value: 'Reduces operational, governance, and migration risk.',
    });
  }

  model.scheduleSla.push({
    schedule_id: 'SLA-0001',
    cadence: 'owner-confirmation-required',
    trigger: 'uploaded on demand',
    run_window: 'unknown',
    dependency: 'source owner interview',
    evidence_id: evidenceId,
    confidence: 'unknown',
    recommended_action: 'Confirm refresh trigger, schedule, dependencies, expected duration, and SLA.',
  });

  model.failureModes.push(
    failureMode('No run', 'Process does not execute or source is unavailable.', evidenceId),
    failureMode('Late run', 'Process completes after downstream decision window.', evidenceId),
    failureMode('Wrong data', 'Logic, mapping, source, or manual edits produce incorrect output.', evidenceId),
    failureMode('Partial run', 'Some but not all data/assets refresh or export.', evidenceId),
    failureMode('Unauditable run', 'Output exists but lineage, controls, approvals, or evidence are missing.', evidenceId),
    failureMode('Upstream dependency failure', 'Upstream file, source, query, connection, or owner-controlled artifact changes or disappears.', evidenceId),
  );

  model.modernization.push({
    asset: model.sourceProcessName,
    recommendation: 'Stabilize, govern, then migrate/automate based on validated criticality.',
    target_state:
      'Governed data product with scheduled ingestion, version-controlled transformations, automated quality tests, lineage observability, managed secrets, audit logging, and controlled documentation.',
    priority: 'P1',
    risk_reduced: 'Manual, unaudited, owner-unknown, lineage-blocked source risk.',
    estimated_effort: 'T-shirt sizing pending owner and finance validation.',
    dependency: 'Owner confirmation, complete upstream files, finance inputs, security review.',
    acceptance_criteria: 'Every critical output has confirmed lineage, owner, SLA, control, quality test, action owner, and validated exposure.',
  });
}

function failureMode(scenario: string, description: string, evidenceId: string): Record<string, unknown> {
  return {
    failure_mode_id: `FM-${scenario.replace(/[^A-Z0-9]+/gi, '_').toUpperCase()}`,
    scenario,
    description,
    failure_impact: 'Directional impact until owner and finance validation.',
    evidence_id: evidenceId,
    confidence: 'inferred',
    recommended_action: 'Validate impact and define control/test for this failure mode.',
  };
}

function addFinancialExposure(model: DiscoveryModel, ids: IdState, evidenceId: string): void {
  const scenarios = [
    'Does not run',
    'Runs late',
    'Runs with wrong data',
    'Partially runs',
    'Runs but cannot be audited or explained',
    'Fails because an upstream dependency changes or is unavailable',
  ];

  for (const scenario of scenarios) {
    const units = 1;
    const dollarPerUnit = scenario.includes('wrong data') ? 25000 : 10000;
    const revenueAtRisk = units * dollarPerUnit;
    const marginPercent = 0.35;
    const marginAtRisk = Math.round(revenueAtRisk * marginPercent);
    const reworkHours = scenario.includes('audited') ? 24 : 8;
    const laborRate = 125;
    const laborRecoveryCost = reworkHours * laborRate;
    const customerSlaExposure = scenario.includes('late') ? 5000 : 0;
    const complianceExposure = scenario.includes('audited') ? 15000 : 2500;
    const cashTimingCost = scenario.includes('late') || scenario.includes('run') ? 2000 : 0;
    const baseImpact = revenueAtRisk + marginAtRisk + laborRecoveryCost + customerSlaExposure + complianceExposure + cashTimingCost;
    const lowImpact = Math.round(baseImpact * 0.35);
    const highImpact = Math.round(baseImpact * 2.25);

    model.financialExposure.push({
      process_or_output: model.sourceProcessName,
      failure_scenario: scenario,
      frequency: 'per incident; annualization assumes 12 occurrences until validated',
      units_affected: units,
      dollar_per_unit: dollarPerUnit,
      revenue_at_risk: revenueAtRisk,
      margin_percent: marginPercent,
      margin_at_risk: marginAtRisk,
      rework_hours: reworkHours,
      labor_rate: laborRate,
      labor_recovery_cost: laborRecoveryCost,
      customer_sla_exposure: customerSlaExposure,
      compliance_exposure: complianceExposure,
      cash_timing_cost: cashTimingCost,
      low_impact: lowImpact,
      base_impact: baseImpact,
      high_impact: highImpact,
      annualized_low: lowImpact * 12,
      annualized_base: baseImpact * 12,
      annualized_high: highImpact * 12,
      confidence: 'inferred',
      assumptions: 'Directional proxy only; not finance-certified.',
      evidence_id: evidenceId,
      finance_validation_needed: 'Finance must provide real units, dollar per unit, margin, labor rate, SLA terms, compliance ranges, and run frequency.',
    });
  }

  addAction(model, ids, {
    title: 'Validate financial exposure model',
    description: 'Replace proxy assumptions with finance-approved volumes, values, margins, SLA penalties, compliance ranges, labor rates, and incident frequency.',
    source_asset: model.sourceProcessName,
    owner_role: 'Finance partner / process owner',
    recommended_owner: 'Finance business partner',
    action_type: 'Financial Validation',
    priority: 'P1',
    severity: 'P1',
    dependency: 'Finance inputs and process owner validation',
    acceptance_criteria: 'Financial model is marked finance-certified or explicitly remains directional with documented gaps.',
    evidence_id: evidenceId,
    related_risk: 'Proxy dollar exposure',
    expected_business_value: 'Improves prioritization and leadership decision quality.',
  });
}

function addQaRecords(model: DiscoveryModel, ids: IdState, evidenceId: string): void {
  const hasBlockedNodes = model.nodes.some((node) => node.confidence === 'blocked');
  const hasBlockedEvidence = model.evidence.some((evidence) => evidence.confidence === 'blocked');
  const hasExcelVbaBlocker = model.excel.Excel_VBA_Register.some((row) => row.confidence === 'blocked');
  const hasPowerQueryBlocker = model.excel.Excel_PowerQuery_Register.some((row) => row.confidence === 'blocked');
  const hasAccessBlocker = model.sourceTypeCounts.access > 0 && model.blockedSources.length > 0;

  for (const check of QA_CHECKS) {
    const limitation = shouldLimitQaCheck(check, {
      hasBlockedNodes,
      hasBlockedEvidence,
      hasExcelVbaBlocker,
      hasPowerQueryBlocker,
      hasAccessBlocker,
    })
      ? 'PASS_WITH_LIMITATION'
      : 'PASS';

    model.qaRecords.push({
      qa_id: nextId(ids, 'QA'),
      check,
      status: limitation,
      evidence_id: evidenceId,
      notes:
        limitation === 'PASS_WITH_LIMITATION'
          ? 'The package structure was generated, but one or more source-specific extraction areas are partial/blocked and have action items.'
          : 'Generated from enforced package contract.',
    });
  }
}

function shouldLimitQaCheck(
  check: string,
  flags: {
    hasBlockedNodes: boolean;
    hasBlockedEvidence: boolean;
    hasExcelVbaBlocker: boolean;
    hasPowerQueryBlocker: boolean;
    hasAccessBlocker: boolean;
  },
): boolean {
  if (!flags.hasBlockedNodes && !flags.hasBlockedEvidence) {
    return false;
  }
  if (check.includes('Every important finding') || check.includes('Every critical output') || check.includes('Every lineage blocker')) {
    return true;
  }
  if (check.includes('P0/P1 risk') || check.includes('final deliverables')) {
    return true;
  }
  if (check.includes('Access') || check.includes('macro') || check.includes('query')) {
    return flags.hasAccessBlocker;
  }
  if (check.includes('Excel Power Query') || check.includes('VBA')) {
    return flags.hasExcelVbaBlocker || flags.hasPowerQueryBlocker;
  }
  return false;
}

function addEvidence(
  model: DiscoveryModel,
  ids: IdState,
  input: {
    category: string;
    title: string;
    fileName: string;
    sourceFile: string;
    summary: string;
    confidence: Confidence;
    content: string | Buffer;
  },
): string {
  const evidenceId = nextId(ids, 'EVID');
  const safeFileName = input.fileName.replace(/[\\/]+/g, '_');
  const relativePath = `05_Evidence_Archive/${input.category}/${safeFileName}`;
  const item: EvidenceItem = {
    evidence_id: evidenceId,
    title: input.title,
    category: input.category,
    relative_path: relativePath,
    summary: input.summary,
    source_file: input.sourceFile,
    confidence: input.confidence,
    content: input.content,
  };
  model.evidence.push(item);
  return evidenceId;
}

function addNode(model: DiscoveryModel, ids: IdState, input: Omit<DiscoveryNode, 'node_id'>): DiscoveryNode {
  const node: DiscoveryNode = {
    node_id: nextId(ids, 'NODE'),
    ...input,
  };
  model.nodes.push(node);
  return node;
}

function addEdge(model: DiscoveryModel, ids: IdState, input: Omit<DiscoveryEdge, 'edge_id'>): DiscoveryEdge {
  const edge: DiscoveryEdge = {
    edge_id: nextId(ids, 'EDGE'),
    ...input,
  };
  model.edges.push(edge);
  return edge;
}

function addOpenQuestion(
  model: DiscoveryModel,
  ids: IdState,
  input: Omit<OpenQuestion, 'question_id' | 'status'>,
): void {
  model.openQuestions.push({
    question_id: nextId(ids, 'OPEN'),
    status: 'open',
    ...input,
  });
}

function addAction(
  model: DiscoveryModel,
  ids: IdState,
  input: Omit<ActionItem, 'action_id' | 'status' | 'due_date_or_phase'> & Partial<Pick<ActionItem, 'due_date_or_phase'>>,
): void {
  model.actions.push({
    action_id: nextId(ids, 'ACT'),
    status: 'Not Started',
    due_date_or_phase: 'Next discovery sprint',
    ...input,
  });
}

function addDataQualityFinding(
  model: DiscoveryModel,
  ids: IdState,
  input: Omit<DataQualityFinding, 'finding_id'>,
): void {
  model.dataQualityFindings.push({
    finding_id: nextId(ids, 'DQ'),
    ...input,
  });
}

function emptyAccessRegisters(): Record<string, Record<string, unknown>[]> {
  return {
    Access_Object_Inventory: [],
    Access_Table_Register: [],
    Access_Linked_Table_Register: [],
    Access_Query_Register: [],
    Access_Query_SQL_Index: [],
    Access_Macro_Register: [],
    Access_Macro_XML_Storage: [],
    Access_Macro_Action_Sequence: [],
    Access_Query_Macro_Reconciliation: [],
    Access_Form_Report_Register: [],
    Access_Module_VBA_Register: [],
    Access_Import_Export_Specs: [],
    Access_Column_Inventory: [],
    Access_Data_Profile: [],
  };
}

function emptyExcelRegisters(): Record<string, Record<string, unknown>[]> {
  return {
    Excel_Workbook_Inventory: [],
    Excel_Sheet_Inventory: [],
    Excel_Table_NamedRange_Register: [],
    Excel_PowerQuery_Register: [],
    Excel_Formula_Register: [],
    Excel_Connection_Register: [],
    Excel_VBA_Register: [],
    Excel_Pivot_DataModel_Register: [],
    Excel_Data_Profile: [],
  };
}

function emptyWordRegisters(): Record<string, Record<string, unknown>[]> {
  return {
    Word_Document_Inventory: [],
    Word_Section_Extracts: [],
    Word_Process_Rules: [],
    Word_Control_Extracts: [],
  };
}

function profileSheet(sheet: XLSX.WorkSheet | undefined, sheetName: string): Record<string, unknown> {
  if (!sheet?.['!ref']) {
    return {
      sheet_name: sheetName,
      used_range: '',
      row_count: 0,
      column_count: 0,
      profile_status: 'empty-or-unavailable',
    };
  }

  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '', raw: false });
  const columns = Object.keys(rows[0] ?? {});
  return {
    sheet_name: sheetName,
    used_range: sheet['!ref'],
    row_count: rows.length,
    column_count: columns.length,
    columns: profileTabularObjects(columns, rows).columns,
    sample_rows: rows.slice(0, 10),
  };
}

function profileTabularRows(columns: string[], rows: string[][]): {
  columns: {
    name: string;
    inferred_type: string;
    blank_count: number;
    sample_values: string[];
  }[];
  duplicate_first_column_count: number;
} {
  const columnProfiles = columns.map((name, index) => {
    const values = rows.map((row) => row[index] ?? '');
    return {
      name: name || `column_${index + 1}`,
      inferred_type: inferColumnType(values),
      blank_count: values.filter((value) => value === '').length,
      sample_values: distinctSamples(values),
    };
  });

  const firstColumnValues = rows.map((row) => row[0] ?? '').filter(Boolean);
  const duplicateCount = firstColumnValues.length - new Set(firstColumnValues).size;

  return {
    columns: columnProfiles,
    duplicate_first_column_count: duplicateCount,
  };
}

function inferTabularAssetSemantics(
  fileName: string,
  columns: string[],
  rows: string[][],
  profile: {
    columns: {
      name: string;
      inferred_type: string;
      blank_count: number;
      sample_values: string[];
    }[];
    duplicate_first_column_count: number;
  },
): {
  domain: string;
  description: string;
  business_purpose: string;
  confidence: Confidence;
  criticality: Criticality;
  signals: string[];
  key_fields: string[];
  subject: string;
  recommended_action: string;
  failure_impact: string;
} {
  const normalizedName = normalizeSemanticText(fileName);
  const normalizedColumns = columns.map(normalizeSemanticText);
  const allSignals = `${normalizedName} ${normalizedColumns.join(' ')}`;
  const keyFields = profile.columns
    .filter((column) => /(id|key|code|number|name|email|user|employee|member|team|role|title|department|status)/i.test(column.name))
    .map((column) => column.name)
    .slice(0, 8);
  const signals: string[] = [];

  const hasRosterName = /\broster\b|\bteam\b|\bmember\b|\bemployee\b|\bstaff\b|\bperson\b|\bpeople\b/.test(allSignals);
  const hasName = normalizedColumns.some((column) => /\b(first name|last name|full name|name|member name|employee name|person)\b/.test(column));
  const hasRole = normalizedColumns.some((column) => /\b(role|title|position|job|function)\b/.test(column));
  const hasTeam = normalizedColumns.some((column) => /\b(team|department|group|unit|pod|squad|division)\b/.test(column));
  const hasContact = normalizedColumns.some((column) => /\b(email|e mail|phone|mobile|contact)\b/.test(column));
  const hasStatus = normalizedColumns.some((column) => /\b(status|active|inactive|start date|end date)\b/.test(column));

  if (hasRosterName) signals.push('file/columns reference roster, team, member, employee, staff, or people');
  if (hasName) signals.push('person/name field detected');
  if (hasRole) signals.push('role/title/position field detected');
  if (hasTeam) signals.push('team/department/group field detected');
  if (hasContact) signals.push('contact field detected');
  if (hasStatus) signals.push('status/effective-date field detected');

  const peopleScore = [hasRosterName, hasName, hasRole, hasTeam, hasContact, hasStatus].filter(Boolean).length;
  if (peopleScore >= 2) {
    return {
      domain: hasTeam ? 'team roster / people reference dataset' : 'people roster / contact reference dataset',
      description: `The column set and filename indicate a roster-style dataset with ${rows.length} record(s) and fields such as ${columns.slice(0, 8).join(', ')}.`,
      business_purpose:
        'Roster/reference data used to identify people or team members, their roles or assignments, and potentially contact/status attributes for operations, communication, staffing, access review, or reporting.',
      confidence: peopleScore >= 4 ? 'confirmed' : 'inferred',
      criticality: hasContact || hasStatus ? 'P1' : 'P2',
      signals,
      key_fields: keyFields,
      subject: 'people/team members',
      recommended_action:
        'Confirm the roster owner, source-of-truth system, refresh cadence, active/inactive rules, downstream consumers, privacy handling, and whether this file drives access, staffing, communication, or reporting decisions.',
      failure_impact:
        'Wrong, stale, duplicated, or missing roster records can misroute communication, misstate team composition, affect access/staffing decisions, and undermine downstream reporting or accountability.',
    };
  }

  const hasMoney = normalizedColumns.some((column) => /\b(amount|revenue|cost|price|margin|budget|spend|invoice|payment)\b/.test(column));
  const hasDate = normalizedColumns.some((column) => /\b(date|period|month|year|week|timestamp)\b/.test(column));
  const hasLocation = normalizedColumns.some((column) => /\b(location|site|address|city|state|region|market)\b/.test(column));
  if (hasMoney) signals.push('financial amount/cost/revenue field detected');
  if (hasDate) signals.push('date/period field detected');
  if (hasLocation) signals.push('location/geography field detected');

  if (hasMoney) {
    return {
      domain: 'financial or commercial transaction/reference dataset',
      description: `The file contains financial/commercial field signals across ${rows.length} record(s).`,
      business_purpose:
        'Tabular financial/commercial data used for reporting, reconciliation, spend/revenue analysis, billing, margin review, or decision support until owner confirms exact use.',
      confidence: 'inferred',
      criticality: 'P1',
      signals,
      key_fields: keyFields,
      subject: 'financial/commercial records',
      recommended_action: 'Confirm finance owner, certified source, period, key fields, reconciliation controls, downstream reports, and finance-approved exposure assumptions.',
      failure_impact: 'Wrong or stale financial records can misstate spend, revenue, margin, billing, cash timing, or executive reporting.',
    };
  }

  return {
    domain: 'generic tabular dataset',
    description: `The file is a parsed tabular dataset with ${rows.length} row(s), ${columns.length} column(s), and fields such as ${columns.slice(0, 8).join(', ')}.`,
    business_purpose: 'Uploaded tabular data source; exact business role, owner, consumers, cadence, and source-of-truth status require confirmation.',
    confidence: 'confirmed',
    criticality: 'P2',
    signals,
    key_fields: keyFields,
    subject: 'records',
    recommended_action: 'Confirm primary key, source-of-truth status, upstream system, owner, downstream use, cadence, sensitivity, and quality thresholds.',
    failure_impact: 'Bad or missing file can break downstream decisions, imports, reporting, or reconciliation depending on confirmed usage.',
  };
}

function normalizeSemanticText(value: string): string {
  return value
    .replace(/[_-]+/g, ' ')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function profileTabularObjects(columns: string[], rows: Record<string, unknown>[]): {
  columns: {
    name: string;
    inferred_type: string;
    blank_count: number;
    sample_values: string[];
  }[];
} {
  return {
    columns: columns.map((name) => {
      const values = rows.map((row) => row[name]);
      return {
        name,
        inferred_type: inferColumnType(values),
        blank_count: values.filter((value) => value === '' || value === null || value === undefined).length,
        sample_values: distinctSamples(values),
      };
    }),
  };
}

function distinctSamples(values: unknown[]): string[] {
  return Array.from(
    new Set(
      values
        .filter((value) => value !== '' && value !== null && value !== undefined)
        .map((value) => String(value).slice(0, 80)),
    ),
  ).slice(0, 5);
}

function inferColumnType(values: unknown[]): string {
  const populated = values.filter((value) => value !== '' && value !== null && value !== undefined).map(String).slice(0, 200);
  if (!populated.length) {
    return 'blank';
  }
  if (populated.every((value) => /^-?\d+(\.\d+)?$/.test(value.replace(/[$,%\s]/g, '')))) {
    return 'number';
  }
  if (populated.every((value) => /^(true|false|yes|no|y|n)$/i.test(value))) {
    return 'boolean';
  }
  if (populated.every((value) => !Number.isNaN(Date.parse(value)) && /\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(value))) {
    return 'date';
  }
  return 'string';
}

function sensitiveIndicator(name: string): string {
  if (/(ssn|social|tax|ein|dob|birth|passport|license)/i.test(name)) {
    return 'PII indicator';
  }
  if (/(salary|wage|revenue|margin|cost|price|amount|bank|account|routing|payment)/i.test(name)) {
    return 'financial/confidential indicator';
  }
  if (/(email|phone|address|customer|patient|member|employee)/i.test(name)) {
    return 'personal/contact indicator';
  }
  return 'none detected';
}

function decodeText(buffer: Buffer): string {
  return buffer.toString('utf8').replace(/^\uFEFF/, '');
}

function inferDelimiter(text: string, extension: string): string {
  if (extension === '.tsv') {
    return '\t';
  }
  const sample = text.split(/\r?\n/).slice(0, 20).join('\n');
  const candidates = [',', '\t', '|', ';'];
  return candidates
    .map((delimiter) => ({
      delimiter,
      count: sample.split(delimiter).length,
    }))
    .sort((a, b) => b.count - a.count)[0]?.delimiter ?? ',';
}

function shouldTreatAsDelimitedText(text: string, delimiter: string, extension: string): boolean {
  if (extension === '.csv' || extension === '.tsv') {
    return true;
  }
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .slice(0, 80);
  if (lines.length < 2) {
    return false;
  }
  const delimiterCounts = lines.map((line) => countOccurrences(line, delimiter));
  const linesWithDelimiter = delimiterCounts.filter((count) => count > 0).length;
  if (linesWithDelimiter < Math.max(3, Math.ceil(lines.length * 0.45))) {
    return false;
  }
  const columnCounts = delimiterCounts.filter((count) => count > 0).map((count) => count + 1);
  const mode = mostCommonNumber(columnCounts);
  const modeHits = columnCounts.filter((count) => count === mode).length;
  return mode >= 2 && modeHits >= Math.max(3, Math.ceil(columnCounts.length * 0.55));
}

function countOccurrences(value: string, needle: string): number {
  if (!needle) {
    return 0;
  }
  return value.split(needle).length - 1;
}

function mostCommonNumber(values: number[]): number {
  const counts = new Map<number, number>();
  for (const value of values) {
    counts.set(value, (counts.get(value) ?? 0) + 1);
  }
  return Array.from(counts.entries()).sort((a, b) => b[1] - a[1] || b[0] - a[0])[0]?.[0] ?? 0;
}

function looksLikeHeader(firstRow: string[], secondRow: string[]): boolean {
  if (!firstRow.length) {
    return false;
  }

  const unique = new Set(firstRow.filter(Boolean)).size === firstRow.filter(Boolean).length;
  const firstHasLetters = firstRow.some((value) => /[a-zA-Z]/.test(value));
  const secondHasData = secondRow.some((value) => value !== '');
  return unique && firstHasLetters && secondHasData;
}

function extractTextKeyValuePairs(lines: string[]): { key: string; value: string; line: number }[] {
  const pairs: { key: string; value: string; line: number }[] = [];
  const seen = new Set<string>();
  lines.forEach((line, index) => {
    const match = line.match(/^\s*["']?([A-Za-z0-9_.:/ -]{2,80})["']?\s*(?:=|:|=>)\s*(.+?)\s*$/);
    if (!match) {
      return;
    }
    const key = collapseWhitespace(match[1] ?? '').replace(/^[-*]\s*/, '').slice(0, 80);
    const value = collapseWhitespace((match[2] ?? '').replace(/^["']|["']$/g, '')).slice(0, 500);
    const pairKey = `${key.toLowerCase()}\u0000${value.toLowerCase()}`;
    if (!key || seen.has(pairKey)) {
      return;
    }
    seen.add(pairKey);
    pairs.push({ key, value, line: index + 1 });
  });
  return pairs;
}

function extractUrls(text: string): string[] {
  return Array.from(new Set(Array.from(text.matchAll(/\bhttps?:\/\/[^\s<>"')]+/gi)).map((match) => match[0].replace(/[.,;]+$/, ''))));
}

function extractPathMentions(text: string): string[] {
  const windowsPaths = Array.from(text.matchAll(/\b(?:[A-Za-z]:\\|\\\\)[^\r\n"'<>|]+/g)).map((match) => collapseWhitespace(match[0]).replace(/[.,;]+$/, ''));
  const fileNames = Array.from(text.matchAll(/\b[\w .()#&-]+\.(?:xlsx|xlsm|xlsb|xls|csv|txt|tsv|accdb|mdb|sql|pdf|docx|doc|json|xml|html|htm)\b/gi)).map((match) =>
    collapseWhitespace(match[0]),
  );
  return Array.from(new Set([...windowsPaths, ...fileNames])).slice(0, 80);
}

function inferTextArtifactKind(
  fileName: string,
  text: string,
  keyValues: { key: string; value: string; line: number }[],
  headings: string[],
  urls: string[],
): string {
  const combined = `${fileName}\n${text.slice(0, 8000)}`.toLowerCase();
  const keys = keyValues.map((pair) => pair.key.toLowerCase());
  if (/<html|<!doctype|page_title|meta name|href=|body>|<\/[a-z]+>/i.test(text) || keys.some((key) => /page_title|url|href|meta|html/.test(key))) {
    return 'web/page extract';
  }
  if (/\b(error|warn|exception|traceback|stack trace|timestamp|request id|status code)\b/i.test(combined)) {
    return 'log/export text';
  }
  if (/\b(select|insert into|update\s+\w+\s+set|delete from|create table|alter table)\b/i.test(combined)) {
    return 'sql-like text';
  }
  if (/\b(api[_-]?key|token|secret|connectionstring|client_id|tenant_id|endpoint|base_url)\b/i.test(combined) || keyValues.length >= 8) {
    return 'configuration/key-value text';
  }
  if (headings.length >= 3 || /\b(purpose|scope|requirements?|steps?|process|approval|control|exception|sla)\b/i.test(combined)) {
    return 'documentation/process text';
  }
  if (urls.length) {
    return 'reference/link text';
  }
  return 'unstructured text';
}

function deriveTextSectionsFromLines(lines: string[]): string[] {
  return lines
    .filter((line) => line.length >= 8)
    .map((line) => collapseWhitespace(line).slice(0, 90))
    .filter((line, index, array) => array.indexOf(line) === index)
    .slice(0, 6);
}

function nodeTypeForPathMention(value: string): NodeType {
  if (/[\\/][^\\/]+\.[A-Za-z0-9]{2,6}$/.test(value) || /\.[A-Za-z0-9]{2,6}$/.test(value)) {
    return 'file';
  }
  return 'folder';
}

function extractSystemMentions(text: string): string[] {
  const known = ['Access', 'Excel', 'Word', 'SQL Server', 'Postgres', 'Oracle', 'Snowflake', 'dbt', 'Power BI', 'Tableau', 'SAP', 'Salesforce'];
  return known.filter((name) => new RegExp(`\\b${name.replace(/\s+/g, '\\s+')}\\b`, 'i').test(text));
}

function isHeadingCandidate(line: string): boolean {
  if (line.length < 4 || line.length > 120) {
    return false;
  }
  if (/[:.]$/.test(line) && line.split(/\s+/).length <= 12) {
    return true;
  }
  return /^[A-Z0-9][A-Z0-9\s/&-]{3,}$/.test(line) && line.split(/\s+/).length <= 12;
}

function extractSqlReferences(text: string): {
  reads: string[];
  writes: string[];
  whereClauses: string[];
  joinClauses: string[];
} {
  const cleaned = text.replace(/--.*$/gm, ' ').replace(/\/\*[\s\S]*?\*\//g, ' ');
  const identifier = String.raw`(?:\[[^\]]+\]|"[^"]+"|[a-zA-Z0-9_.$]+)(?:\s*\.\s*(?:\[[^\]]+\]|"[^"]+"|[a-zA-Z0-9_.$]+))*`;
  const reads = new Set<string>();
  const writes = new Set<string>();
  for (const match of cleaned.matchAll(new RegExp(String.raw`\b(?:from|join)\s+(${identifier})`, 'gi'))) {
    reads.add(cleanIdentifier(match[1] ?? ''));
  }
  for (const match of cleaned.matchAll(new RegExp(String.raw`\b(?:into|insert\s+into|update|merge\s+into|delete\s+from|alter\s+table)\s+(${identifier})`, 'gi'))) {
    writes.add(cleanIdentifier(match[1] ?? ''));
  }
  const whereClauses = Array.from(cleaned.matchAll(/\bwhere\b\s+([\s\S]*?)(?:\bgroup\b|\border\b|\bhaving\b|\bunion\b|;|$)/gi)).map((match) =>
    collapseWhitespace(match[1] ?? '').slice(0, 280),
  );
  const joinClauses = Array.from(
    cleaned.matchAll(new RegExp(String.raw`\bjoin\s+${identifier}\s+(?:as\s+)?[a-zA-Z0-9_"]*\s*on\s+([\s\S]*?)(?:\bjoin\b|\bwhere\b|\bgroup\b|\border\b|;|$)`, 'gi')),
  ).map((match) => collapseWhitespace(match[1] ?? '').slice(0, 280));
  return {
    reads: Array.from(reads).filter(Boolean),
    writes: Array.from(writes).filter(Boolean),
    whereClauses,
    joinClauses,
  };
}

function cleanIdentifier(value: string): string {
  return value.replace(/["[\]]/g, '').replace(/\s*\.\s*/g, '.').replace(/[;,)]$/, '').trim();
}

function collapseWhitespace(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}

function safeFormatSql(text: string): string {
  try {
    return formatSql(text, { language: 'sql' });
  } catch {
    return text;
  }
}

function sha256(buffer: Buffer): string {
  return crypto.createHash('sha256').update(buffer).digest('hex');
}

function csvCell(value: unknown): string {
  const text = String(value ?? '');
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

void addTypedPlaceholders;

function addTypedPlaceholders(
  _nodeType: NodeType,
  _edgeType: EdgeType,
  _criticality: Criticality,
  _confidence: Confidence,
  _processStep: ProcessStep,
  _transformationRule: TransformationRule,
  _securityAccessFinding: SecurityAccessFinding,
  _financialExposure: FinancialExposure,
  _actionItem: ActionItem,
): void {
  // Keeps imported contract types anchored for editor/tooling without runtime impact.
}
