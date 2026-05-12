import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import PDFDocument from 'pdfkit';
import { createHash } from 'node:crypto';
import path from 'node:path';
import {
  ACCESS_AUTO_CSVS,
  ACCESS_WORKBOOK_TABS,
  BASE_AUTO_CSVS,
  BASE_WORKBOOK_TABS,
  EDGE_TYPES,
  EVIDENCE_SUBFOLDERS,
  EXCEL_AUTO_CSVS,
  EXCEL_WORKBOOK_TABS,
  MASTER_DOSSIER_STANDARD,
  MASTER_DOSSIER_STANDARD_VERSION,
  NODE_TYPES,
  PACKAGE_FOLDERS,
  REQUIRED_DIAGRAMS,
  WORD_AUTO_CSVS,
  WORD_WORKBOOK_TABS,
} from './contract.js';
import type { AiCostSummary, DiscoveryModel } from './types.js';

type BuildResult = {
  buffer: Buffer;
  summary: DossierSummary;
  majorArtifacts: Record<string, string>;
};

export type DossierSummary = {
  packageName: string;
  sourceAnalyzed: string;
  fileCount: number;
  packageFileCount: number;
  objectCount: number;
  queryCount: number;
  macroCount: number;
  vbaModuleCount: number;
  vbaProcedureCount: number;
  vbaFunctionCount: number;
  vbaEventHandlerCount: number;
  linkedSourceCount: number;
  lineageBlockers: number;
  qaStatus: string;
  aiEnabled: boolean;
  aiModel: string;
  aiCost: AiCostSummary;
  filePurpose: string;
  executiveSummary: string;
  architectureSummary: string;
  topRisks: string[];
  recommendedPath: string;
  limitations: string[];
  limitationBlocks: Array<{ title: string; detail: string; severity: string }>;
  upstreamCoverage: {
    total: number;
    resolvedFiles: number;
    resolvedFolders: number;
    blocked: number;
  };
  upstreamSources: Array<{
    name: string;
    kind: string;
    status: string;
    location: string;
    table: string;
    action: string;
  }>;
};

type ZipWriter = {
  addText: (relativePath: string, content: string) => void;
  addBinary: (relativePath: string, content: Buffer) => void;
  fileCount: () => number;
};

type DiagramCard = {
  title: string;
  subtitle?: string;
  meta?: string;
  nodeId?: string;
  nodeType?: string;
  criticality?: string;
  confidence?: string;
  accent?: string;
};

type DiagramStage = {
  title: string;
  subtitle?: string;
  cards: DiagramCard[];
  accent?: string;
};

export async function buildDossierPackage(model: DiscoveryModel): Promise<BuildResult> {
  const zip = new JSZip();
  const root = `${model.packageName}/`;
  const filePaths = new Set<string>();
  const writer: ZipWriter = {
    addText(relativePath, content) {
      const zipPath = `${root}${relativePath}`;
      zip.file(zipPath, content);
      filePaths.add(zipPath);
    },
    addBinary(relativePath, content) {
      const zipPath = `${root}${relativePath}`;
      zip.file(zipPath, content);
      filePaths.add(zipPath);
    },
    fileCount() {
      return filePaths.size;
    },
  };

  for (const folder of PACKAGE_FOLDERS) {
    zip.folder(`${root}${folder}`);
  }
  for (const subfolder of EVIDENCE_SUBFOLDERS) {
    zip.folder(`${root}05_Evidence_Archive/${subfolder}`);
  }

  const [executiveBrief, architectureReport, technicalWorkbook, diagramPdfs, financialWorkbook] = await Promise.all([
    createExecutiveBrief(model),
    createArchitectureReport(model),
    createTechnicalWorkbook(model),
    Promise.all(REQUIRED_DIAGRAMS.map(async (diagram) => [diagram.file, await createDiagramPdf(model, diagram)] as const)),
    createFinancialWorkbook(model),
  ]);

  writer.addText('README.md', createReadme(model));
  writer.addBinary('01_Executive_Decision_Brief/Executive_Decision_Brief.pdf', executiveBrief);
  writer.addBinary('02_Current_State_Architecture_Report/Current_State_Architecture_Report.pdf', architectureReport);
  writer.addBinary('03_Technical_Discovery_Workbook/Technical_Discovery_Workbook.xlsx', technicalWorkbook);

  for (const [fileName, pdf] of diagramPdfs) {
    writer.addBinary(`04_Diagram_Pack/${fileName}`, pdf);
  }

  for (const evidence of model.evidence) {
    const content = Buffer.isBuffer(evidence.content) ? evidence.content : Buffer.from(evidence.content, 'utf8');
    writer.addBinary(evidence.relative_path, content);
  }

  for (const subfolder of EVIDENCE_SUBFOLDERS) {
    const hasEvidence = model.evidence.some((evidence) => evidence.relative_path.includes(`/${subfolder}/`));
    if (!hasEvidence) {
      writer.addText(
        `05_Evidence_Archive/${subfolder}/README.md`,
        `No source-specific evidence was available for ${subfolder}. If this concept applies, it is documented as a blocker or open question.\n`,
      );
    }
  }

  const csvFiles = createAutoDocumentationCsvs(model);
  for (const [fileName, csv] of Object.entries(csvFiles)) {
    writer.addText(`06_Auto_Documentation_Pack/${fileName}`, csv);
  }

  writer.addText('08_Action_Backlog/Action_Backlog.csv', rowsToCsv(model.actions));
  writer.addBinary('09_Financial_Impact_Model/Financial_Impact_Model.xlsx', financialWorkbook);
  writer.addText(
    '05_Evidence_Archive/05m_QA_Certification/Master_Dossier_Standard.md',
    `# Master Dossier Standard\n\nVersion: ${MASTER_DOSSIER_STANDARD_VERSION}\nSHA-256: ${masterPromptHash()}\n\n${MASTER_DOSSIER_STANDARD.trim()}\n`,
  );

  const manifest = createManifest(model, writer.fileCount() + 1);
  writer.addText('07_Metadata_Manifest/Metadata_Manifest.json', JSON.stringify(manifest, null, 2));

  const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });

  return {
    buffer,
    summary: createSummary(model, writer.fileCount()),
    majorArtifacts: {
      zip: model.packageName,
      executiveBrief: `${model.packageName}/01_Executive_Decision_Brief/Executive_Decision_Brief.pdf`,
      architectureReport: `${model.packageName}/02_Current_State_Architecture_Report/Current_State_Architecture_Report.pdf`,
      technicalWorkbook: `${model.packageName}/03_Technical_Discovery_Workbook/Technical_Discovery_Workbook.xlsx`,
      diagramPack: `${model.packageName}/04_Diagram_Pack/`,
      evidenceArchive: `${model.packageName}/05_Evidence_Archive/`,
      autoDocumentationPack: `${model.packageName}/06_Auto_Documentation_Pack/`,
      metadataManifest: `${model.packageName}/07_Metadata_Manifest/Metadata_Manifest.json`,
      actionBacklog: `${model.packageName}/08_Action_Backlog/Action_Backlog.csv`,
      financialModel: `${model.packageName}/09_Financial_Impact_Model/Financial_Impact_Model.xlsx`,
    },
  };
}

function createSummary(model: DiscoveryModel, packageFileCount = 0): DossierSummary {
  const upstreamSources = createUpstreamSourceSummaries(model);
  const limitationBlocks = createLimitationBlocks(model);
  const automationStats = createAutomationStats(model);
  const lineageBlockerCount = upstreamSources.filter((source) => source.status === 'Blocked').length;
  return {
    packageName: model.packageName,
    sourceAnalyzed: model.sourceFiles.map((source) => source.file_name).join(', '),
    fileCount: model.sourceFiles.length,
    packageFileCount,
    objectCount: model.nodes.length,
    queryCount: model.nodes.filter((node) => node.node_type === 'query').length,
    macroCount: automationStats.macroCount,
    vbaModuleCount: automationStats.vbaModuleCount,
    vbaProcedureCount: automationStats.vbaProcedureCount,
    vbaFunctionCount: automationStats.vbaFunctionCount,
    vbaEventHandlerCount: automationStats.vbaEventHandlerCount,
    linkedSourceCount: model.nodes.filter((node) => node.node_type === 'linked table').length,
    lineageBlockers: lineageBlockerCount,
    qaStatus: model.qaRecords.some((record) => record.status === 'FAIL')
      ? 'FAIL'
      : model.qaRecords.some((record) => record.status === 'PASS_WITH_LIMITATION')
        ? 'PASS_WITH_LIMITATIONS'
        : 'PASS',
    aiEnabled: Boolean(model.aiNarrative.enabled),
    aiModel: model.aiNarrative.model ?? '',
    aiCost: model.aiNarrative.cost ?? createFallbackAiCostSummary(model.aiNarrative.model ?? ''),
    filePurpose: truncateSummary(model.aiNarrative.filePurpose ?? createDeterministicFilePurpose(model), 520),
    executiveSummary: truncateSummary(
      model.aiNarrative.executiveSummary ??
        'AI narrative synthesis did not return an executive summary; deterministic package artifacts were still generated from evidence.',
    ),
    architectureSummary: truncateSummary(
      model.aiNarrative.architectureSummary ??
        'AI narrative synthesis did not return an architecture summary; use the technical workbook and evidence archive as the canonical record.',
    ),
    topRisks: [
      'No run',
      'Late run',
      'Wrong data',
      'Partial run',
      'Unauditable run',
      'Source dependency failure',
    ],
    recommendedPath:
      model.aiNarrative.recommendedPath ??
      createDeterministicRecommendedPath(model),
    limitations: limitationBlocks.map((block) => `${block.title}: ${block.detail}`).slice(0, 6),
    limitationBlocks,
    upstreamCoverage: {
      total: upstreamSources.length,
      resolvedFiles: upstreamSources.filter((source) => source.status === 'Resolved file').length,
      resolvedFolders: upstreamSources.filter((source) => source.status === 'Resolved folder').length,
      blocked: upstreamSources.filter((source) => source.status === 'Blocked').length,
    },
    upstreamSources: upstreamSources.slice(0, 12),
  };
}

function createDeterministicFilePurpose(model: DiscoveryModel): string {
  const sourceNames = model.sourceFiles.map((source) => source.file_name).join(', ') || model.sourceProcessName;
  const sourceTypes = Object.entries(model.sourceTypeCounts)
    .filter(([, count]) => count > 0)
    .map(([type, count]) => `${count} ${type}`)
    .join(', ');
  const nameSignal = humanizeProcessName(model.sourceProcessName);
  const nodeCounts = model.nodes.reduce<Record<string, number>>((counts, node) => {
    counts[node.node_type] = (counts[node.node_type] ?? 0) + 1;
    return counts;
  }, {});
  const automationStats = createAutomationStats(model);

  if (model.sourceTypeCounts.excel > 0) {
    const automationPhrase =
      automationStats.macroCount || automationStats.vbaProcedureCount
        ? ` with ${automationStats.macroCount} runnable macro entrypoint(s), ${automationStats.vbaModuleCount} VBA module(s), and ${automationStats.vbaProcedureCount} parsed VBA procedure(s)`
        : '';
    return `At a 30,000-foot level, ${sourceNames} is an Excel-based operational data workbook${automationPhrase}. It appears to support the ${nameSignal} process by storing workbook structures, applying formulas or automation logic where present, and producing business-ready data artifacts for review or downstream use. Exact owner, decision use, and cadence remain owner-confirmation-required unless directly evidenced.`;
  }

  if (model.sourceTypeCounts.access > 0) {
    return `At a 30,000-foot level, ${sourceNames} is a Microsoft Access database used as an operational data process container. It appears to support the ${nameSignal} process by organizing tables or linked sources, saved queries, forms/reports, macros, modules, and outputs where extractable. Exact owner, decision use, cadence, and source-of-truth status remain owner-confirmation-required unless directly evidenced.`;
  }

  if (model.sourceTypeCounts.word > 0) {
    return `At a 30,000-foot level, ${sourceNames} is a process/documentation source. It appears to describe the ${nameSignal} process, including business rules, actors, controls, inputs, outputs, or decisions where the document text provides evidence.`;
  }

  if (model.sourceTypeCounts['flat-file'] > 0) {
    if (nodeCounts.document || nodeCounts['document section']) {
      return `At a 30,000-foot level, ${sourceNames} is a text artifact submitted for discovery, not a confirmed tabular file. The distillery profiled it as unstructured text with ${nodeCounts.document ?? 0} document node(s), ${nodeCounts['document section'] ?? 0} detected section(s), and ${model.dataElements.length} key/value or data signal(s). It appears to support the ${nameSignal} process as documentation, configuration, web/page extract, log/export, or supporting evidence until the source owner confirms its intended role.`;
    }
    const primaryTable = model.nodes.find((node) => node.node_type === 'table' && model.sourceFiles.some((source) => source.file_type === 'flat-file' && source.file_name === node.source_file));
    if (primaryTable?.business_purpose && primaryTable.business_purpose !== 'Uploaded tabular data source.') {
      return `At a 30,000-foot level, ${sourceNames} is a parsed flat-file data asset with an inferred business role. ${primaryTable.business_purpose} The parsed file has ${model.rowCountSummary[primaryTable.source_file] ?? 'unknown'} record(s) and ${model.dataElements.length} discovered data element(s). ${primaryTable.description} Ownership, authoritative source status, refresh cadence, downstream consumers, and decision use still require confirmation.`;
    }
    return `At a 30,000-foot level, ${sourceNames} is a structured flat-file data source with ${nodeCounts['data element'] ?? 0} discovered data element(s). It appears to provide tabular input or output data for the ${nameSignal} process, with lineage and ownership requiring confirmation from upstream/downstream owners.`;
  }

  if (model.sourceTypeCounts.sql > 0 || model.sourceTypeCounts.script > 0) {
    return `At a 30,000-foot level, ${sourceNames} is a code or SQL logic source for the ${nameSignal} process. It appears to define reads, writes, joins, filters, parameters, or output behavior that should be governed as transformation logic.`;
  }

  return `At a 30,000-foot level, ${sourceNames} is a ${sourceTypes || 'source artifact'} submitted for discovery. The dossier establishes what can be evidenced from the uploaded file, then marks missing business purpose, owner, consumers, cadence, and lineage as confirmation items.`;
}

function createDeterministicRecommendedPath(model: DiscoveryModel): string {
  const primaryFlatTable = model.nodes.find((node) => node.node_type === 'table' && model.sourceFiles.some((source) => source.file_type === 'flat-file' && source.file_name === node.source_file));
  if (primaryFlatTable?.business_purpose && /roster|people|team member|contact|staffing|access review/i.test(primaryFlatTable.business_purpose)) {
    return 'Treat this as a roster/reference data asset. Confirm whether it is authoritative or an extract, assign the roster owner/steward, document refresh cadence and downstream consumers, validate active/inactive and role/team rules, add privacy controls for contact/person fields, and add basic completeness/duplicate checks before relying on it for staffing, communication, access, or reporting decisions.';
  }
  if (primaryFlatTable) {
    return 'Confirm the flat-file owner, source-of-truth status, upstream origin, refresh cadence, downstream consumers, key fields, sensitivity, and quality thresholds before using it as an operational or reporting input.';
  }
  return 'Stabilize ownership and blockers, govern lineage and controls, then migrate or automate validated critical paths.';
}

function humanizeProcessName(value: string): string {
  const cleaned = path
    .basename(value, path.extname(value))
    .replace(/[_-]+/g, ' ')
    .replace(/\b(v?\d+(\.\d+)*)\b/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
  return cleaned || 'uploaded source';
}


function truncateSummary(value: string, maxLength = 900): string {
  return value.length > maxLength ? `${value.slice(0, maxLength - 1)}...` : value;
}

function createFallbackAiCostSummary(modelName: string): AiCostSummary {
  return {
    pricingSource: 'No OpenAI usage record was attached to this run.',
    totalRequests: 0,
    totalInputTokens: 0,
    totalCachedInputTokens: 0,
    totalBillableInputTokens: 0,
    totalOutputTokens: 0,
    totalReasoningTokens: 0,
    totalTokens: 0,
    estimatedTotalCostUsd: 0,
    estimatedCacheSavingsUsd: 0,
    cacheHitRate: 0,
    optimizationNote: 'No OpenAI usage was recorded for this run.',
    models: [
      {
        model: modelName || 'unknown',
        requests: 0,
        inputTokens: 0,
        cachedInputTokens: 0,
        billableInputTokens: 0,
        outputTokens: 0,
        reasoningTokens: 0,
        totalTokens: 0,
        inputCostUsd: 0,
        cachedInputCostUsd: 0,
        outputCostUsd: 0,
        totalCostUsd: 0,
        cacheSavingsUsd: 0,
        cacheHitRate: 0,
        pricing: {
          inputPerMillion: 0,
          cachedInputPerMillion: 0,
          outputPerMillion: 0,
          pricingAvailable: false,
        },
      },
    ],
  };
}

function createAutomationStats(model: DiscoveryModel): {
  macroCount: number;
  vbaModuleCount: number;
  vbaProcedureCount: number;
  vbaFunctionCount: number;
  vbaEventHandlerCount: number;
} {
  const excelRows = model.excel.Excel_VBA_Register;
  const moduleRows = excelRows.filter((row) => row.artifact_level === 'module');
  const procedureRows = excelRows.filter((row) => row.artifact_level === 'procedure');
  const runnableExcelMacros = new Set(
    procedureRows
      .filter((row) => row.runnable_macro_flag === 'yes' || row.procedure_role === 'macro entrypoint')
      .map((row) => `${row.workbook ?? ''}:${row.module_name ?? ''}.${row.procedure_name ?? ''}`),
  );
  if (!runnableExcelMacros.size) {
    excelRows
      .filter((row) => row.artifact_level === 'button binding' && row.runnable_macro_flag === 'yes')
      .forEach((row) => runnableExcelMacros.add(`${row.workbook ?? ''}:button:${row.procedure_name ?? ''}`));
  }

  return {
    macroCount: model.access.Access_Macro_Register.length + runnableExcelMacros.size,
    vbaModuleCount: moduleRows.length,
    vbaProcedureCount: procedureRows.length,
    vbaFunctionCount: procedureRows.filter((row) => row.procedure_role === 'function/helper').length,
    vbaEventHandlerCount: procedureRows.filter((row) => row.procedure_role === 'event handler').length,
  };
}

function createUpstreamSourceSummaries(model: DiscoveryModel): DossierSummary['upstreamSources'] {
  const rows = model.access.Access_Linked_Table_Register ?? [];
  if (rows.length) {
    const summaries = new Map<string, DossierSummary['upstreamSources'][number]>();
    for (const row of rows) {
      const rawPath = String(row.resolved_path || row.source_path || row.database || row.connect || '');
      const status = String(row.resolution_status ?? '').toLowerCase();
      const location = compactLocation(rawPath);
      const name = String(row.source_name || row.foreign_name || row.source_table_name || row.table_name || 'Linked source');
      const table = String(row.table_name ?? '');
      const key = `${status || 'blocked'}|${location || name}`.toLowerCase();
      const existing = summaries.get(key);
      if (existing) {
        const tableSet = new Set(
          existing.table
            .split(', ')
            .map((item) => item.trim())
            .filter(Boolean),
        );
        if (table) {
          tableSet.add(table);
        }
        existing.table = Array.from(tableSet).join(', ');
        continue;
      }
      summaries.set(key, {
        name,
        kind: String(row.source_kind || inferSourceKindForSummary(rawPath)),
        status: status === 'resolved_file' ? 'Resolved file' : status === 'resolved_folder' ? 'Resolved folder' : 'Blocked',
        location,
        table,
        action:
          status === 'resolved_file'
            ? 'Recursed into linked file evidence.'
            : status === 'resolved_folder'
              ? 'Folder reached; identify exact file set or include folder export.'
              : 'Provide reachable upstream file or mount path, then rerun discovery.',
      });
    }
    return Array.from(summaries.values());
  }

  return model.blockedSources.slice(0, 12).map((source) => ({
    name: source.split('|')[0]?.trim() || 'Blocked source',
    kind: source.split('|')[1]?.trim() || 'External source',
    status: 'Blocked',
    location: compactLocation(source.split('|').at(-1)?.trim() || source),
    table: '',
    action: 'Provide reachable upstream file or owner confirmation.',
  }));
}

function createLimitationBlocks(model: DiscoveryModel): DossierSummary['limitationBlocks'] {
  const blocks = new Map<string, { title: string; detail: string; severity: string }>();
  const add = (title: string, detail: string, severity = 'Action required') => {
    blocks.set(title, { title, detail, severity });
  };

  for (const limitation of model.limitations) {
    const lower = limitation.toLowerCase();
    if (lower.includes('access.application') || lower.includes('saveastext')) {
      add('Deep Access automation export', 'Macro action bodies, form/report text, and VBA source require a trusted desktop deep export; macro objects are still inventoried separately.', 'Partial');
      continue;
    }
    if (lower.includes('openai')) {
      if (lower.includes('not configured')) {
        add('AI synthesis', 'OpenAI narrative enrichment did not run because OPENAI_API_KEY is not configured; deterministic evidence, workbooks, diagrams, and registers were still generated.', 'Optional');
      } else if (lower.includes('incorrect api key') || lower.includes('401')) {
        add('AI synthesis', 'OPENAI_API_KEY was loaded from local configuration, but OpenAI rejected it with an authentication error. Replace the key or confirm the project/account, then rerun for AI narrative enrichment.', 'Action required');
      } else {
        add('AI synthesis', `OpenAI narrative enrichment did not complete: ${truncateText(limitation, 240)} Deterministic evidence, workbooks, diagrams, and registers were still generated.`, 'Action required');
      }
      continue;
    }
    if (lower.includes('large file mode')) {
      add('Large-source guardrails', 'Expensive full row counts are deferred in bounded mode; table, query, field, index, relationship, and linked-source metadata still drive the dossier.', 'Partial');
      continue;
    }
    if (lower.includes('depth limit')) {
      add('Recursive depth limit', limitation, 'Action required');
      continue;
    }
    add('Discovery note', truncateText(limitation, 220), 'Review');
  }

  if (!blocks.size && model.blockedSources.length) {
    add('Upstream reachability', `${model.blockedSources.length} upstream source(s) need a reachable file, mounted path, or owner-approved terminal classification.`, 'Action required');
  }

  return Array.from(blocks.values()).slice(0, 6);
}

function inferSourceKindForSummary(location: string): string {
  const extension = path.extname(location).toLowerCase();
  if (['.xlsx', '.xlsm', '.xlsb', '.xls'].includes(extension)) {
    return 'Excel workbook';
  }
  if (['.accdb', '.mdb'].includes(extension)) {
    return 'Access database';
  }
  if (['.csv', '.txt', '.tsv'].includes(extension)) {
    return 'Delimited/text file';
  }
  return location ? 'External source' : 'Unknown';
}

function compactLocation(location: string): string {
  if (!location) {
    return 'Not resolved';
  }
  const normalized = location.replace(/;?(HDR|IMEX|ACCDB|FMT|CharacterSet|DSN)=[^;]*/gi, '').replace(/;+$/, '').trim();
  const database = normalized.match(/DATABASE=([^;]+)/i)?.[1] ?? normalized;
  if (database.length <= 120) {
    return database;
  }
  const parts = database.split(/[\\/]+/).filter(Boolean);
  return parts.length > 3 ? `.../${parts.slice(-3).join('/')}` : truncateText(database, 120);
}

function truncateText(value: string, maxLength: number): string {
  return value.length <= maxLength ? value : `${value.slice(0, maxLength - 1)}…`;
}

function createReadme(model: DiscoveryModel): string {
  const summary = createSummary(model);
  return `# ${model.packageName}

## Source File(s) Analyzed
${model.sourceFiles.map((source) => `- ${source.file_name} (${source.file_type}, ${source.file_size_bytes} bytes)`).join('\n')}

## Analysis Date
${model.generatedDate}

## Dossier Contract
- Contract version: ${MASTER_DOSSIER_STANDARD_VERSION}
- Contract SHA-256: ${masterPromptHash()}
- Contract evidence: 05_Evidence_Archive/05m_QA_Certification/Master_Dossier_Standard.md

## Folder Guide
- 01_Executive_Decision_Brief: Leadership-ready risk, recommendation, and decision summary.
- 02_Current_State_Architecture_Report: Narrative explanation of how the current process, systems, data flow, controls, and risks work.
- 03_Technical_Discovery_Workbook: Detailed structured inventory of objects, lineage, logic, controls, quality findings, and actions.
- 04_Diagram_Pack: Visual diagrams of value stream, process flow, data flow, lineage, dependencies, controls, failures, and schedule.
- 05_Evidence_Archive: Raw extracted evidence supporting dossier findings.
- 06_Auto_Documentation_Pack: Machine-readable current-state documentation generated from the canonical discovery model.
- 07_Metadata_Manifest: Package-level manifest with file inventory, counts, source metadata, and QA status.
- 08_Action_Backlog: Execution-ready remediation, governance, and modernization backlog.
- 09_Financial_Impact_Model: Low/base/high business exposure model for failure, delay, wrong data, partial run, or unauditable run.

## Primary Deliverables
- Executive_Decision_Brief.pdf: Decision, risk, dollars, urgency, and recommendation.
- Current_State_Architecture_Report.pdf: Current-state operating model, lineage summary, controls, risks, and modernization path.
- Technical_Discovery_Workbook.xlsx: Canonical structured inventory and evidence index.
- Diagram PDFs D01-D09: Standalone graph-backed diagrams with legends and context.
- Metadata_Manifest.json: Valid JSON inventory, counts, limitations, checksums, and QA status.
- Action_Backlog.csv: Jira/ADO-ready action register.
- Financial_Impact_Model.xlsx: Directional low/base/high exposure model.

## Known Blockers or Inaccessible Upstream Sources
${summary.upstreamSources.length ? summary.upstreamSources.map((source) => `- ${source.status}: ${source.name} (${source.kind}) - ${source.location}`).join('\n') : '- None identified from uploaded files.'}

## How To Use The Package
Start with the executive brief for decisions, then use the architecture report for current-state context. Analysts and engineers should work from the technical workbook, evidence archive, and auto-documentation CSVs. Governance and migration teams should focus on lineage blockers, controls, security findings, and the action backlog.

## QA Status Summary
${summary.qaStatus}. ${summary.limitationBlocks.length ? summary.limitationBlocks.map((block) => `${block.title}: ${block.detail}`).join(' ') : 'No package-generation QA failures were detected.'}
`;
}

async function createExecutiveBrief(model: DiscoveryModel): Promise<Buffer> {
  const summary = createSummary(model);
  const baseExposure = model.financialExposure.reduce((sum, row) => sum + row.base_impact, 0);
  return createPdf('Executive Decision Brief', [
    {
      heading: 'Executive Snapshot',
      lines: [
        `Source/process name: ${model.sourceProcessName}`,
        `Source type(s): ${Object.entries(model.sourceTypeCounts)
          .filter(([, count]) => count > 0)
          .map(([type, count]) => `${type} (${count})`)
          .join(', ')}`,
        `Critical outputs: ${model.packageName}`,
        `Current usage: owner-confirmation-required`,
        `Criticality: P1 until owner validation`,
        `Risk level: ${summary.lineageBlockers ? 'High due to lineage blockers' : 'Moderate pending owner validation'}`,
        `Modernization recommendation: Stabilize, govern, then migrate/automate validated critical paths.`,
        `Action priority: ${model.actions.some((action) => action.priority === 'P0') ? 'P0/P1' : 'P1'}`,
        `Estimated directional base exposure: ${currency(baseExposure)}`,
        'Decisions required: confirm owner, finance assumptions, security review, and modernization path.',
      ],
    },
    {
      heading: 'What This Process Does',
      lines: [
        summary.filePurpose,
        model.aiNarrative.executiveSummary ??
          'The uploaded source files were inspected from scratch to create a canonical discovery model covering assets, lineage, evidence, risks, open questions, actions, and financial exposure.',
        'Who depends on it: downstream consumers are owner-confirmation-required unless directly present in uploaded evidence.',
        'Cadence: not supplied in the upload and recorded as an open question.',
      ],
    },
    {
      heading: 'Top Findings',
      lines: topFindings(model).slice(0, 5),
    },
    {
      heading: 'Top Risks',
      lines: [
        'No run: source unavailable or process not executed.',
        'Late run: outputs miss the decision window.',
        'Wrong data: formula, query, mapping, source, or manual edit changes produce incorrect results.',
        'Partial run: some assets refresh while others remain stale.',
        'Unauditable run: lineage, controls, approvals, or evidence are missing.',
        'Source dependency failure: upstream files, connections, or owner-controlled artifacts change or disappear.',
      ],
    },
    {
      heading: 'Financial Exposure Summary',
      lines: [
        `Directional low/base/high exposure: ${currency(model.financialExposure.reduce((sum, row) => sum + row.low_impact, 0))} / ${currency(baseExposure)} / ${currency(model.financialExposure.reduce((sum, row) => sum + row.high_impact, 0))}.`,
        'Buckets: revenue, margin, cash timing, rework labor, customer/SLA, compliance/audit, and decision delay.',
        'Confidence: inferred. Finance validation is mandatory before treating values as certified.',
      ],
    },
    {
      heading: 'Recommended Path',
      lines: [
        model.aiNarrative.recommendedPath ??
          'Stabilize source ownership and controls, govern lineage and security, then rebuild/migrate/automate the validated critical path.',
        `P0/P1 action count: ${model.actions.filter((action) => action.priority === 'P0' || action.priority === 'P1').length}.`,
      ],
    },
    {
      heading: 'Decisions Needed',
      lines: [
        'Assign accountable owner and steward.',
        'Approve blocker resolution for missing upstream/internal metadata.',
        'Validate finance assumptions.',
        'Decide whether target state is rebuild, migrate, automate, retire, or leave-as-is temporarily.',
      ],
    },
  ]);
}

async function createArchitectureReport(model: DiscoveryModel): Promise<Buffer> {
  const summary = createSummary(model);
  return createPdf('Current-State Architecture Report', [
    section('Scope, Coverage, and Confidence', [
      `This report covers ${model.sourceFiles.length} uploaded source file(s) and ${model.nodes.length} graph node(s).`,
      `Confidence is evidence-specific. Blockers and owner-confirmation-required items are explicitly marked.`,
    ]),
    section('Business Mission of the Process', [
      summary.filePurpose,
      model.aiNarrative.architectureSummary ??
        'The process mission is inferred from uploaded artifacts until business owners confirm purpose, decision usage, outputs, and consumers.',
    ]),
    section('Current-State Operating Model', [
      'The uploaded files are treated as current-state artifacts. Manual triggers, run windows, approvals, and exception workflows require owner validation.',
      `Process step candidates captured: ${model.processSteps.length}.`,
      ...model.processSteps.slice(0, 6).map((step) => `${step.process_step_id}: ${step.step_name} - ${step.description} Evidence: ${step.evidence_id}; confidence: ${step.confidence}.`),
    ]),
    section('System and Artifact Landscape', [
      `Source types: ${Object.entries(model.sourceTypeCounts)
        .filter(([, count]) => count > 0)
        .map(([type, count]) => `${type}: ${count}`)
        .join(', ')}.`,
      `Object inventory is in workbook tab 02_Object_Inventory_All. Source inventory is in tab 01_Source_Inventory.`,
      ...excelArchitectureLandscape(model),
    ]),
    section('Process Flow Summary', [
      'Process steps are extracted from documents, scripts, and inferred upload flow. Full human workflow requires interviews.',
      'See workbook tabs 03_Business_Process_Steps and 13_Schedule_SLA.',
    ]),
    section('Data Flow Summary', [
      `Lineage nodes: ${model.nodes.length}. Lineage edges: ${model.edges.length}.`,
      'See D04_Detailed_Data_Flow and workbook tabs 05_Lineage_Nodes and 06_Lineage_Edges.',
    ]),
    section('Transformation and Business Logic Summary', [
      `${model.transformations.length} transformation/rule candidate(s) were captured from formulas, SQL, scripts, or structured metadata.`,
      'Detailed logic belongs in workbook tab 07_Transformations_Rules, not duplicated here.',
      ...vbaArchitectureLines(model),
    ]),
    section('Recursive Lineage and Source-of-Truth Assessment', [
      `${model.blockedSources.length} lineage blocker(s) are documented.`,
      'Terminal lineage conditions are marked as confirmed, inferred, blocked, partial, or owner-confirmation-required.',
      model.blockedSources.length
        ? `Blocked sources: ${model.blockedSources.slice(0, 5).join(' | ')}.`
        : 'No recursive lineage blockers were reported for the uploaded source set; terminal lineage currently stops at the uploaded workbook/source file pending owner-confirmed source-of-truth status.',
    ]),
    section('Controls, Exceptions, and Failure Modes', [
      `${model.controls.length} control/exception record(s) and ${model.failureModes.length} failure mode(s) are included.`,
      'Run logging, approvals, and exception workflow require owner/platform validation unless evidence was uploaded.',
    ]),
    section('Security, Access, and Compliance Summary', [
      `${model.securityAccess.length} security/access finding(s) are listed.`,
      'Permissions, embedded credentials, retention rules, and broad access cannot be certified from browser upload alone.',
    ]),
    section('Data Quality Summary', [
      `${model.dataQualityFindings.length} data quality finding(s) were captured from profileable files.`,
      'Critical field rules require data owner confirmation and automated tests.',
      ...model.dataQualityFindings.slice(0, 8).map((finding) => `${finding.finding_id}: ${finding.asset} / ${finding.field} - ${finding.issue}. ${finding.example} Evidence: ${finding.evidence_id}.`),
    ]),
    section('Financial Impact and Business Exposure Summary', [
      'The financial model is directional and proxy-based.',
      'Finance must provide certified units, dollar values, margins, SLA terms, compliance ranges, labor rates, and run frequency.',
    ]),
    section('Modernization Recommendation', [
      'Recommended path: stabilize ownership and blockers, govern lineage and controls, then migrate/automate critical paths into a governed data product.',
      'Target-state candidates include scheduled ingestion, version-controlled transformations, automated quality tests, lineage observability, managed secrets, audit logging, and controlled documentation.',
    ]),
    section('Open Questions and Decisions Needed', model.openQuestions.slice(0, 10).map((question) => `${question.question_id}: ${question.question}`)),
  ]);
}

function section(heading: string, lines: string[]): { heading: string; lines: string[] } {
  return { heading, lines };
}

function excelArchitectureLandscape(model: DiscoveryModel): string[] {
  if (!model.sourceTypeCounts.excel) {
    return [];
  }
  const sheets = model.excel.Excel_Sheet_Inventory.length;
  const formulaAreas = model.excel.Excel_Formula_Register.length;
  const namedRanges = model.excel.Excel_Table_NamedRange_Register.filter((row) => row['object_type'] === 'named range').length;
  const vbaRows = model.excel.Excel_VBA_Register.length;
  return [
    `Excel landscape: ${sheets} worksheet/register row(s), ${formulaAreas} formula area(s), ${namedRanges} named range(s), and ${vbaRows} VBA register row(s).`,
    `Workbook visibility/profile evidence is in tabs 22_Excel_Sheet_Inventory, 25_Excel_Formula_Register, 27_Excel_VBA_Register, and 29_Excel_Data_Profile.`,
  ];
}

function vbaArchitectureLines(model: DiscoveryModel): string[] {
  const rows = model.excel.Excel_VBA_Register;
  if (!rows.length) {
    return [];
  }
  const modules = rows.filter((row) => row['artifact_level'] === 'module').length;
  const macroEntrypoints = rows.filter((row) => row['procedure_role'] === 'macro entrypoint').length;
  const events = rows.filter((row) => row['procedure_role'] === 'event handler').length;
  const functions = rows.filter((row) => row['procedure_role'] === 'function/helper').length;
  const procedureLines = rows
    .filter((row) => row['artifact_level'] === 'procedure')
    .slice(0, 8)
    .map(
      (row) =>
        `${row['module_name']}.${row['procedure_name']}: ${row['procedure_role']} - ${row['purpose']}. Operations: ${row['operations'] || 'none detected'}. Evidence: ${row['evidence_id']}.`,
    );
  return [
    `VBA automation: ${modules} module(s), ${macroEntrypoints} runnable macro entrypoint(s), ${events} workbook/event handler(s), and ${functions} helper function(s) were extracted from evidence.`,
    ...procedureLines,
  ];
}

async function createDiagramPdf(
  model: DiscoveryModel,
  diagram: (typeof REQUIRED_DIAGRAMS)[number],
): Promise<Buffer> {
  return createVisualDiagramPdf(model, diagram);
}

async function createVisualDiagramPdf(
  model: DiscoveryModel,
  diagram: (typeof REQUIRED_DIAGRAMS)[number],
): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const document = new PDFDocument({ size: 'LETTER', layout: 'landscape', margin: 28, bufferPages: true });
    const chunks: Buffer[] = [];
    document.on('data', (chunk: Buffer) => chunks.push(chunk));
    document.on('error', reject);
    document.on('end', () => resolve(Buffer.concat(chunks)));

    drawDiagramHeader(document, model, diagram);

    if (diagram.file.startsWith('D01')) {
      drawExecutiveValueStream(document, model);
    } else if (diagram.file.startsWith('D02')) {
      drawSystemContext(document, model);
    } else if (diagram.file.startsWith('D03')) {
      drawBusinessSwimlane(document, model);
    } else if (diagram.file.startsWith('D04')) {
      drawDetailedDataFlow(document, model);
    } else if (diagram.file.startsWith('D05')) {
      drawRecursiveLineage(document, model);
    } else if (diagram.file.startsWith('D06')) {
      drawObjectDependencyMap(document, model);
    } else if (diagram.file.startsWith('D07')) {
      drawControlExceptionMap(document, model);
    } else if (diagram.file.startsWith('D08')) {
      drawFailureImpactMap(document, model);
    } else {
      drawScheduleTimeline(document, model);
    }

    drawDiagramLegend(document, model, diagram);
    drawDiagramFooter(document, model);

    document.end();
  });
}

function drawDiagramHeader(
  document: PDFKit.PDFDocument,
  model: DiscoveryModel,
  diagram: (typeof REQUIRED_DIAGRAMS)[number],
): void {
  document.rect(0, 0, 792, 78).fill('#111614');
  document.font('Helvetica-Bold').fontSize(17).fillColor('#ffffff').text(diagram.title, 34, 22, { width: 520 });
  document.font('Helvetica').fontSize(8.8).fillColor('#dbe2df').text(`Purpose: ${diagram.purpose}`, 34, 45, { width: 560 });
  drawPill(document, 604, 22, 132, 'GRAPH-BACKED', '#ff7a1a', '#111614');
  document.font('Helvetica').fontSize(7.2).fillColor('#cbd5d0').text(`Generated ${model.generatedDate}`, 604, 48, { width: 132, align: 'center' });
}

function drawExecutiveValueStream(document: PDFKit.PDFDocument, model: DiscoveryModel): void {
  const stages: DiagramStage[] = [
    { title: 'Source Intake', subtitle: `${model.sourceFiles.length} uploaded source(s)`, cards: sourceCards(model, 2), accent: '#2d7d67' },
    { title: 'Assets', subtitle: 'Detected structures', cards: nodeCards(model, ['workbook', 'worksheet', 'table', 'named range'], 3), accent: '#4b6f9f' },
    { title: 'Logic Distillery', subtitle: 'What changes the data', cards: nodeCards(model, ['macro', 'macro action', 'Power Query', 'formula area', 'module'], 3), accent: '#ff7a1a' },
    { title: 'Outputs', subtitle: 'Decision package', cards: nodeCards(model, ['output'], 2), accent: '#2d7d67' },
    { title: 'Risk and Dollars', subtitle: 'Leadership lens', cards: financialCards(model, 3), accent: '#8a2e1d' },
  ];
  drawStageColumns(document, stages, 34, 104, 724, 312, true);
}

function drawSystemContext(document: PDFKit.PDFDocument, model: DiscoveryModel): void {
  const center = firstNodeCard(model, ['workbook', 'database', 'file'], 'Uploaded source system');
  const source = sourceCards(model, 1)[0] ?? center;
  const around = [
    ...nodeCards(model, ['worksheet'], 4),
    ...nodeCards(model, ['macro', 'macro action', 'formula area', 'Power Query'], 4),
    ...nodeCards(model, ['data element'], 3),
    ...nodeCards(model, ['output'], 1),
  ].slice(0, 10);
  drawCard(document, 316, 222, 160, 66, center, { fill: '#f7fbf8', accent: '#2d7d67' });
  drawCard(document, 60, 218, 150, 64, source, { fill: '#ffffff', accent: '#4b6f9f' });
  drawArrow(document, 210, 250, 316, 250, 'documents');
  const positions = [
    [78, 118],
    [248, 118],
    [486, 118],
    [656, 118],
    [78, 342],
    [248, 342],
    [486, 342],
    [656, 342],
    [316, 120],
    [316, 344],
  ];
  around.forEach((card, index) => {
    const [x, y] = positions[index] ?? [60 + index * 68, 350];
    drawCard(document, x, y, 118, 56, card, { fill: '#ffffff', accent: card.accent });
    drawArrow(document, x + 59, y + 56, 396, 222, card.nodeType === 'output' ? 'exports_to' : 'depends_on', true);
  });
}

function drawBusinessSwimlane(document: PDFKit.PDFDocument, model: DiscoveryModel): void {
  const automationSteps = processStepCards(model, 4);
  const automationObjects = nodeCards(model, ['macro', 'macro action', 'module', 'formula area'], 4);
  const lanes = [
    {
      title: 'Business / Analyst',
      cards: [
        textCard('Upload source file(s)', 'Manual trigger', 'confidence: confirmed', '#4b6f9f'),
        textCard('Review dossier outputs', 'Decision and validation', 'owner confirmation', '#4b6f9f'),
      ],
    },
    {
      title: 'Deterministic Extractors',
      cards: [
        textCard('Workbook metadata', `${model.excel.Excel_Sheet_Inventory?.length ?? 0} sheets`, 'automated', '#2d7d67'),
        textCard('Profiles and registers', `${model.dataQualityFindings.length} DQ finding(s)`, 'automated', '#2d7d67'),
      ],
    },
    {
      title: 'Automation Logic',
      cards: automationSteps.length
        ? automationSteps
        : automationObjects.length
          ? automationObjects
          : [textCard('No automation object found', 'registers empty', 'owner confirmation', '#8a8f89')],
    },
    {
      title: 'Package QA',
      cards: [
        textCard('Build required artifacts', `${model.evidence.length} evidence item(s)`, 'automated', '#ff7a1a'),
        textCard('Run 30 QA gates', createSummary(model).qaStatus, 'contract enforced', '#ff7a1a'),
      ],
    },
  ];
  drawSwimlanes(document, lanes, 34, 106, 724, 318);
}

function drawDetailedDataFlow(document: PDFKit.PDFDocument, model: DiscoveryModel): void {
  const stages: DiagramStage[] = [
    { title: 'Input Files', subtitle: 'Uploaded or linked', cards: sourceCards(model, 3), accent: '#4b6f9f' },
    { title: 'Workbook Structures', subtitle: 'Sheets, tables, ranges', cards: nodeCards(model, ['worksheet', 'table', 'named range', 'pivot'], 4), accent: '#2d7d67' },
    { title: 'Transformation Logic', subtitle: 'Macro, event, formula, query, M', cards: nodeCards(model, ['macro', 'macro action', 'module', 'formula area', 'Power Query', 'query'], 4), accent: '#ff7a1a' },
    { title: 'Outputs and Evidence', subtitle: 'Package and workbook tabs', cards: nodeCards(model, ['output', 'document'], 3), accent: '#2d7d67' },
  ];
  drawStageColumns(document, stages, 44, 104, 704, 320, true);
}

function drawRecursiveLineage(document: PDFKit.PDFDocument, model: DiscoveryModel): void {
  const outputCards = nodeCards(model, ['output'], 2);
  const logicCards = nodeCards(model, ['macro', 'macro action', 'module', 'formula area', 'Power Query', 'query'], 4);
  const assetCards = nodeCards(model, ['worksheet', 'table', 'named range', 'data element'], 4);
  const terminalCards = model.blockedSources.length
    ? model.blockedSources.slice(0, 4).map((source) => textCard(source.split('|')[0]?.trim() || 'Blocked upstream', 'terminal: blocked', 'action required', '#8a2e1d'))
    : sourceCards(model, 2).map((card) => ({ ...card, meta: 'terminal: uploaded source' }));
  drawStageColumns(
    document,
    [
      { title: 'Critical Output', subtitle: 'Where decisions land', cards: outputCards, accent: '#2d7d67' },
      { title: 'Transformation Points', subtitle: 'What changed it', cards: logicCards, accent: '#ff7a1a' },
      { title: 'Data Assets', subtitle: 'Where values live', cards: assetCards, accent: '#4b6f9f' },
      { title: 'Terminal Source', subtitle: 'Confirmed or blocked stop', cards: terminalCards, accent: model.blockedSources.length ? '#8a2e1d' : '#2d7d67' },
    ],
    56,
    104,
    680,
    320,
    true,
  );
}

function drawObjectDependencyMap(document: PDFKit.PDFDocument, model: DiscoveryModel): void {
  const clusters: DiagramStage[] = [
    { title: 'Workbook / File', cards: nodeCards(model, ['workbook', 'file', 'database'], 3), accent: '#4b6f9f' },
    { title: 'Sheets and Ranges', cards: nodeCards(model, ['worksheet', 'table', 'named range', 'pivot'], 5), accent: '#2d7d67' },
    { title: 'Code and Logic', cards: nodeCards(model, ['module', 'macro', 'macro action', 'formula area', 'Power Query', 'query'], 5), accent: '#ff7a1a' },
    { title: 'Outputs / Consumers', cards: nodeCards(model, ['output', 'downstream consumer', 'person / role'], 3), accent: '#2d7d67' },
  ];
  drawStageColumns(document, clusters, 40, 102, 712, 322, false);
  drawMetricRibbon(document, [
    ['Nodes', String(model.nodes.length)],
    ['Edges', String(model.edges.length)],
    ['Dependency edges', String(model.edges.filter((edge) => edge.edge_type === 'depends_on').length)],
    ['Transforms', String(model.transformations.length)],
  ]);
}

function drawControlExceptionMap(document: PDFKit.PDFDocument, model: DiscoveryModel): void {
  const stages: DiagramStage[] = [
    { title: 'Controls', subtitle: 'Validations and checks', cards: recordCards(model.controls, 'control_id', 'control_name', 'control_type', 4, '#2d7d67'), accent: '#2d7d67' },
    { title: 'Exceptions', subtitle: 'Known quality issues', cards: dataQualityCards(model, 4), accent: '#8a2e1d' },
    { title: 'Security / Access', subtitle: 'Compliance posture', cards: recordCards(model.securityAccess, 'security_id', 'asset', 'finding', 4, '#4b6f9f'), accent: '#4b6f9f' },
    { title: 'Actions', subtitle: 'What closes the gap', cards: recordCards(model.actions, 'action_id', 'title', 'priority', 4, '#ff7a1a'), accent: '#ff7a1a' },
  ];
  drawStageColumns(document, stages, 38, 104, 716, 320, false);
}

function drawFailureImpactMap(document: PDFKit.PDFDocument, model: DiscoveryModel): void {
  const exposures = model.financialExposure.slice(0, 6);
  const rows = exposures.length
    ? exposures.map((exposure) => ({
        scenario: exposure.failure_scenario,
        impact: currency(exposure.annualized_base),
        note: `${exposure.process_or_output} | ${exposure.confidence}`,
      }))
    : ['No run', 'Late run', 'Wrong data', 'Partial run', 'Unauditable run', 'Source dependency failure'].map((scenario) => ({
        scenario,
        impact: 'directional',
        note: 'Finance validation required',
      }));
  const y0 = 112;
  rows.forEach((row, index) => {
    const y = y0 + index * 48;
    document.roundedRect(48, y, 696, 34, 5).fillAndStroke(index % 2 ? '#fbfcfa' : '#ffffff', '#d5ddd8');
    document.font('Helvetica-Bold').fontSize(9.2).fillColor('#111614').text(row.scenario, 62, y + 8, { width: 180, ellipsis: true });
    document.font('Helvetica').fontSize(7.6).fillColor('#4f5b55').text(row.note, 254, y + 8, { width: 280, ellipsis: true });
    drawPill(document, 620, y + 7, 92, row.impact, '#8a2e1d', '#fff4ef');
  });
  drawMetricRibbon(document, [
    ['Failure modes', String(model.failureModes.length)],
    ['DQ findings', String(model.dataQualityFindings.length)],
    ['P0/P1 actions', String(model.actions.filter((action) => action.priority === 'P0' || action.priority === 'P1').length)],
    ['Finance status', 'Directional'],
  ]);
}

function drawScheduleTimeline(document: PDFKit.PDFDocument, model: DiscoveryModel): void {
  const sourceSchedule = processStepCards(model, 6);
  const steps = sourceSchedule.length
    ? sourceSchedule
    : [
        textCard('Upload', `${model.sourceFiles.length} source(s)`, 'manual trigger', '#4b6f9f'),
        textCard('Extract', `${model.nodes.length} nodes`, 'deterministic', '#2d7d67'),
        textCard('Profile', `${model.dataQualityFindings.length} DQ findings`, 'automated', '#2d7d67'),
        textCard('AI synth', model.aiNarrative.enabled ? model.aiNarrative.model : 'not run', `${model.aiNarrative.cost?.totalTokens ?? 0} tokens`, '#ff7a1a'),
        textCard('Package', `${model.evidence.length} evidence files`, 'PDF/XLSX/CSV/JSON', '#2d7d67'),
        textCard('QA', createSummary(model).qaStatus, '30 gates', '#4b6f9f'),
      ];
  const y = 220;
  steps.forEach((step, index) => {
    const x = 42 + index * 119;
    drawCard(document, x, y, 100, 64, step, { fill: '#ffffff', accent: step.accent });
    if (index < steps.length - 1) {
      drawArrow(document, x + 100, y + 32, x + 119, y + 32, `${index + 1}`);
    }
  });
  document.font('Helvetica-Bold').fontSize(11).fillColor('#111614').text('Refresh cadence and run windows', 48, 112);
  document.font('Helvetica').fontSize(9).fillColor('#3e4843').text(
    sourceSchedule.length
      ? `Confirmed automation trigger candidates are shown below from workbook tab 03_Business_Process_Steps. Owner validation is still required for production cadence, run window, SLA, and control expectations.`
      : model.scheduleSla.length
        ? `${model.scheduleSla.length} schedule/SLA record(s) captured. See workbook tab 13_Schedule_SLA.`
        : 'No native schedule metadata was present. Cadence requires owner confirmation; the generation sequence shown below is the dossier run timeline.',
    48,
    132,
    { width: 680 },
  );
}

function drawStageColumns(
  document: PDFKit.PDFDocument,
  stages: DiagramStage[],
  x: number,
  y: number,
  width: number,
  height: number,
  connect: boolean,
): void {
  const gap = 12;
  const stageWidth = (width - gap * (stages.length - 1)) / stages.length;
  stages.forEach((stage, index) => {
    const sx = x + index * (stageWidth + gap);
    document.roundedRect(sx, y, stageWidth, height, 6).fillAndStroke('#fbfcfa', '#d5ddd8');
    document.rect(sx, y, stageWidth, 5).fill(stage.accent ?? '#2d7d67');
    document.font('Helvetica-Bold').fontSize(10).fillColor('#111614').text(stage.title, sx + 10, y + 14, { width: stageWidth - 20, ellipsis: true });
    document.font('Helvetica').fontSize(7.5).fillColor('#62706a').text(stage.subtitle ?? '', sx + 10, y + 30, { width: stageWidth - 20, ellipsis: true });
    const cards = stage.cards.length ? stage.cards : [textCard('No evidence captured', 'owner confirmation required', 'confidence: unknown', '#8a8f89')];
    cards.slice(0, 5).forEach((card, cardIndex) => {
      drawCard(document, sx + 10, y + 54 + cardIndex * 46, stageWidth - 20, 38, card, { fill: '#ffffff', accent: card.accent ?? stage.accent });
    });
    if (connect && index < stages.length - 1) {
      drawArrow(document, sx + stageWidth, y + height / 2, sx + stageWidth + gap, y + height / 2, index === 0 ? 'reads_from' : index === stages.length - 2 ? 'exports_to' : 'transforms');
    }
  });
}

function drawSwimlanes(
  document: PDFKit.PDFDocument,
  lanes: { title: string; cards: DiagramCard[] }[],
  x: number,
  y: number,
  width: number,
  height: number,
): void {
  const laneHeight = height / lanes.length;
  lanes.forEach((lane, laneIndex) => {
    const ly = y + laneIndex * laneHeight;
    document.roundedRect(x, ly, width, laneHeight - 8, 5).fillAndStroke(laneIndex % 2 ? '#fbfcfa' : '#ffffff', '#d5ddd8');
    document.font('Helvetica-Bold').fontSize(9).fillColor('#111614').text(lane.title, x + 10, ly + 12, { width: 94 });
    const cards = lane.cards.length ? lane.cards : [textCard('No evidence', 'owner confirmation', 'unknown', '#8a8f89')];
    cards.slice(0, 4).forEach((card, cardIndex) => {
      const cx = x + 126 + cardIndex * 142;
      drawCard(document, cx, ly + 12, 122, 48, card, { fill: '#ffffff', accent: card.accent });
      if (cardIndex < Math.min(cards.length, 4) - 1) {
        drawArrow(document, cx + 122, ly + 36, cx + 142, ly + 36, 'handoff');
      }
    });
  });
}

function drawCard(
  document: PDFKit.PDFDocument,
  x: number,
  y: number,
  width: number,
  height: number,
  card: DiagramCard,
  options: { fill?: string; accent?: string } = {},
): void {
  document.roundedRect(x, y, width, height, 5).fillAndStroke(options.fill ?? '#ffffff', '#ccd6d0');
  document.rect(x, y, 4, height).fill(options.accent ?? card.accent ?? '#2d7d67');
  if (card.nodeId) {
    document.font('Helvetica-Bold').fontSize(6.2).fillColor(options.accent ?? card.accent ?? '#2d7d67').text(card.nodeId, x + 8, y + 5, { width: width - 16, ellipsis: true });
  }
  document.font('Helvetica-Bold').fontSize(7.4).fillColor('#111614').text(card.title, x + 8, y + (card.nodeId ? 14 : 8), {
    width: width - 16,
    height: 14,
    ellipsis: true,
  });
  document.font('Helvetica').fontSize(6.6).fillColor('#53605a').text(card.subtitle ?? card.nodeType ?? '', x + 8, y + (card.nodeId ? 28 : 22), {
    width: width - 16,
    height: 10,
    ellipsis: true,
  });
  if (height > 44) {
    document.font('Helvetica').fontSize(6.2).fillColor('#78827d').text(card.meta ?? confidenceLine(card), x + 8, y + height - 13, {
      width: width - 16,
      height: 8,
      ellipsis: true,
    });
  }
}

function drawArrow(document: PDFKit.PDFDocument, x1: number, y1: number, x2: number, y2: number, label: string, muted = false): void {
  const endX = x2 - 4;
  document.strokeColor(muted ? '#9aa6a0' : '#2d7d67').lineWidth(muted ? 0.5 : 0.85).moveTo(x1 + 4, y1).lineTo(endX, y2).stroke();
  document
    .fillColor(muted ? '#9aa6a0' : '#2d7d67')
    .polygon([endX, y2], [endX - 5, y2 - 3], [endX - 5, y2 + 3])
    .fill();
  const midX = (x1 + endX) / 2;
  const midY = (y1 + y2) / 2;
  document.font('Helvetica').fontSize(5.6).fillColor('#2f3a35').text(label, midX - 24, midY - 8, { width: 60, align: 'center' });
}

function drawPill(document: PDFKit.PDFDocument, x: number, y: number, width: number, text: string, color: string, fill: string): void {
  document.roundedRect(x, y, width, 18, 9).fill(fill);
  document.font('Helvetica-Bold').fontSize(7).fillColor(color).text(text, x + 6, y + 5, { width: width - 12, align: 'center', ellipsis: true });
}

function drawMetricRibbon(document: PDFKit.PDFDocument, metrics: [string, string][]): void {
  const x = 48;
  const y = 438;
  const width = 684;
  const itemWidth = width / metrics.length;
  document.roundedRect(x, y, width, 42, 5).fillAndStroke('#111614', '#111614');
  metrics.forEach(([label, value], index) => {
    const ix = x + index * itemWidth;
    if (index) {
      document.moveTo(ix, y + 8).lineTo(ix, y + 34).strokeColor('#303a35').lineWidth(0.5).stroke();
    }
    document.font('Helvetica').fontSize(6.8).fillColor('#b8c5bf').text(label, ix + 10, y + 9, { width: itemWidth - 20, align: 'center' });
    document.font('Helvetica-Bold').fontSize(12).fillColor('#ffffff').text(value, ix + 10, y + 22, { width: itemWidth - 20, align: 'center', ellipsis: true });
  });
}

function drawDiagramLegend(
  document: PDFKit.PDFDocument,
  model: DiscoveryModel,
  diagram: (typeof REQUIRED_DIAGRAMS)[number],
): void {
  const y = 500;
  document.roundedRect(34, y, 724, 62, 5).strokeColor('#c7d0cb').lineWidth(0.75).stroke();
  document.font('Helvetica-Bold').fontSize(8).fillColor('#111614').text('Legend and contract controls', 46, y + 10);
  const nodeTypes = Array.from(new Set(model.nodes.map((node) => node.node_type))).slice(0, 12).join(', ') || NODE_TYPES.slice(0, 8).join(', ');
  const edgeTypes = Array.from(new Set(model.edges.map((edge) => edge.edge_type))).slice(0, 12).join(', ') || EDGE_TYPES.slice(0, 8).join(', ');
  document.font('Helvetica').fontSize(6.8).fillColor('#303a35').text(`Node types in model: ${nodeTypes}.`, 46, y + 24, { width: 690, ellipsis: true });
  document.font('Helvetica').fontSize(6.8).fillColor('#303a35').text(`Edge types in model: ${edgeTypes}. Automated/manual flags, confidence, evidence IDs, and full node/edge rows live in workbook tabs ${diagram.workbookTabs}.`, 46, y + 36, { width: 690, ellipsis: true });
  document.font('Helvetica').fontSize(6.8).fillColor('#303a35').text('Color: green=confirmed/output, blue=asset/context, orange=logic/transformation, red=blocked/risk. Criticality: P0-P3. Confidence: confirmed, inferred, blocked, unknown, partial, owner-confirmation-required.', 46, y + 48, { width: 690, ellipsis: true });
}

function drawDiagramFooter(document: PDFKit.PDFDocument, model: DiscoveryModel): void {
  document.font('Helvetica').fontSize(6.8).fillColor('#6b736f').text(`Package: ${model.packageName} | Evidence-backed generated view`, 34, 580, { width: 724, align: 'right' });
}

function sourceCards(model: DiscoveryModel, limit: number): DiagramCard[] {
  return model.sourceFiles.slice(0, limit).map((source) => ({
    title: source.file_name,
    subtitle: `${source.file_type} | ${formatBytes(source.file_size_bytes)}`,
    meta: `Evidence ${source.evidence_id}`,
    nodeType: source.file_type,
    confidence: 'confirmed',
    criticality: 'P1',
    accent: '#4b6f9f',
  }));
}

function nodeCards(model: DiscoveryModel, types: string[], limit: number): DiagramCard[] {
  return model.nodes
    .filter((node) => types.includes(node.node_type))
    .sort((left, right) => {
      const typeDelta = types.indexOf(left.node_type) - types.indexOf(right.node_type);
      if (typeDelta) return typeDelta;
      const criticalityDelta = criticalityRank(left.criticality) - criticalityRank(right.criticality);
      if (criticalityDelta) return criticalityDelta;
      return confidenceRank(left.confidence) - confidenceRank(right.confidence);
    })
    .slice(0, limit)
    .map((node) => ({
      title: node.name,
      subtitle: node.node_type,
      meta: `${node.criticality} | ${node.confidence}`,
      nodeId: node.node_id,
      nodeType: node.node_type,
      criticality: node.criticality,
      confidence: node.confidence,
      accent: colorForNodeType(node.node_type, node.confidence),
    }));
}

function processStepCards(model: DiscoveryModel, limit: number): DiagramCard[] {
  return model.processSteps.slice(0, limit).map((step) => ({
    title: step.step_name,
    subtitle: step.trigger,
    meta: `${step.manual_or_automated} | ${step.confidence} | ${step.evidence_id}`,
    criticality: 'P1',
    confidence: step.confidence,
    accent: step.manual_or_automated === 'automated' ? '#2d7d67' : '#ff7a1a',
  }));
}

function criticalityRank(value: string): number {
  return ({ P0: 0, P1: 1, P2: 2, P3: 3 } as Record<string, number>)[value] ?? 4;
}

function confidenceRank(value: string): number {
  return (
    {
      confirmed: 0,
      partial: 1,
      inferred: 2,
      'owner-confirmation-required': 3,
      unknown: 4,
      blocked: 5,
    } as Record<string, number>
  )[value] ?? 6;
}

function firstNodeCard(model: DiscoveryModel, types: string[], fallback: string): DiagramCard {
  return nodeCards(model, types, 1)[0] ?? textCard(fallback, model.sourceFiles[0]?.file_name ?? 'source', 'confidence: confirmed', '#2d7d67');
}

function financialCards(model: DiscoveryModel, limit: number): DiagramCard[] {
  const exposures = model.financialExposure.slice(0, limit);
  if (!exposures.length) {
    return [textCard('Directional model', 'Finance validation required', 'no certified dollars', '#8a2e1d')];
  }
  return exposures.map((exposure) => ({
    title: exposure.failure_scenario,
    subtitle: `Base ${currency(exposure.annualized_base)}`,
    meta: `${exposure.confidence} | finance validation`,
    accent: '#8a2e1d',
  }));
}

function dataQualityCards(model: DiscoveryModel, limit: number): DiagramCard[] {
  if (!model.dataQualityFindings.length) {
    return [textCard('No DQ findings detected', 'Not a clean-data certification', 'rules require owner validation', '#8a8f89')];
  }
  return model.dataQualityFindings.slice(0, limit).map((finding) => ({
    title: finding.issue,
    subtitle: `${finding.asset}.${finding.field}`,
    meta: `${finding.severity} | ${finding.confidence}`,
    accent: '#8a2e1d',
  }));
}

function recordCards(
  rows: Record<string, unknown>[],
  idKey: string,
  titleKey: string,
  subtitleKey: string,
  limit: number,
  accent: string,
): DiagramCard[] {
  if (!rows.length) {
    return [textCard('No evidence captured', 'owner confirmation required', 'confidence: unknown', '#8a8f89')];
  }
  return rows.slice(0, limit).map((row) => ({
    title: String(row[titleKey] ?? row[idKey] ?? 'Record'),
    subtitle: String(row[subtitleKey] ?? ''),
    meta: String(row[idKey] ?? ''),
    accent,
  }));
}

function textCard(title: string, subtitle: string, meta: string, accent: string): DiagramCard {
  return { title, subtitle, meta, accent };
}

function confidenceLine(card: DiagramCard): string {
  return [card.criticality, card.confidence].filter(Boolean).join(' | ');
}

function colorForNodeType(nodeType: string, confidence?: string): string {
  if (confidence === 'blocked') {
    return '#8a2e1d';
  }
  if (['macro', 'macro action', 'module', 'Power Query', 'formula area', 'query'].includes(nodeType)) {
    return '#ff7a1a';
  }
  if (['output', 'control'].includes(nodeType)) {
    return '#2d7d67';
  }
  if (['upstream blocker', 'exception'].includes(nodeType)) {
    return '#8a2e1d';
  }
  return '#4b6f9f';
}

function formatBytes(bytes: number): string {
  if (!bytes) {
    return '0 B';
  }
  const units = ['B', 'KB', 'MB', 'GB'];
  const exponent = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** exponent).toFixed(exponent === 0 ? 0 : 1)} ${units[exponent]}`;
}

async function createTechnicalWorkbook(model: DiscoveryModel): Promise<Buffer> {
  const workbook = createWorkbook('Technical Discovery Workbook', model);
  const tabs = [
    ...BASE_WORKBOOK_TABS,
    ...(model.sourceTypeCounts.access ? ACCESS_WORKBOOK_TABS : []),
    ...(model.sourceTypeCounts.excel ? EXCEL_WORKBOOK_TABS : []),
    ...(model.sourceTypeCounts.word ? WORD_WORKBOOK_TABS : []),
  ];

  for (const tab of tabs) {
    addWorksheet(workbook, tab, workbookRowsForTab(tab, model));
  }

  return workbookToBuffer(workbook);
}

async function createFinancialWorkbook(model: DiscoveryModel): Promise<Buffer> {
  const workbook = createWorkbook('Financial Impact Model', model);
  addWorksheet(workbook, 'Financial_Impact_Model', model.financialExposure);
  addWorksheet(workbook, 'Assumptions', [
    ...model.assumptions.map((assumption, index) => ({ assumption_id: `ASM-${index + 1}`, assumption })),
    {
      assumption_id: 'ASM-FINANCE',
      assumption: 'This model is directional and not finance-certified until finance provides validated inputs.',
    },
  ]);
  return workbookToBuffer(workbook);
}

function createWorkbook(title: string, model: DiscoveryModel): ExcelJS.Workbook {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Data Source Discovery Distillery';
  workbook.created = new Date(model.generatedDate);
  workbook.modified = new Date();
  workbook.subject = model.packageName;
  workbook.title = title;
  workbook.company = 'Generated by Vercel + Neon + deterministic discovery pipeline';
  return workbook;
}

const DATASET_COLUMNS: Record<string, string[]> = {
  '03_Business_Process_Steps': [
    'process_step_id',
    'step_name',
    'actor_or_role',
    'trigger',
    'description',
    'manual_or_automated',
    'input_node_id',
    'output_node_id',
    'evidence_id',
    'confidence',
    'recommended_action',
  ],
  '11_Controls_Exceptions': ['control_id', 'asset', 'control_type', 'description', 'status', 'evidence_id', 'confidence', 'recommended_action'],
  '12_Security_Access': ['security_id', 'asset', 'concern', 'status', 'impact', 'evidence_id', 'confidence', 'recommended_action'],
  '18_Open_Questions': ['question_id', 'status', 'asset', 'question', 'owner_role', 'blocker_type', 'priority', 'evidence_id'],
  '24_Excel_PowerQuery_Register': [
    'workbook',
    'sheet_name',
    'query_artifact',
    'm_code_status',
    'command_text',
    'connection',
    'evidence_id',
    'confidence',
    'recommended_action',
  ],
  '26_Excel_Connection_Register': [
    'workbook',
    'connection_name',
    'connection_type',
    'refresh_with_refresh_all',
    'connection_string',
    'evidence_id',
    'confidence',
    'recommended_action',
  ],
  '28_Excel_Pivot_DataModel_Register': ['workbook', 'artifact_type', 'name', 'location', 'evidence_id', 'confidence', 'recommended_action'],
  '06c_Process_Steps.csv': [
    'process_step_id',
    'step_name',
    'actor_or_role',
    'trigger',
    'description',
    'manual_or_automated',
    'input_node_id',
    'output_node_id',
    'evidence_id',
    'confidence',
    'recommended_action',
  ],
  '06g_Controls_Exceptions.csv': ['control_id', 'asset', 'control_type', 'description', 'status', 'evidence_id', 'confidence', 'recommended_action'],
  '06h_Security_Access.csv': ['security_id', 'asset', 'concern', 'status', 'impact', 'evidence_id', 'confidence', 'recommended_action'],
  '06k_Open_Questions.csv': ['question_id', 'status', 'asset', 'question', 'owner_role', 'blocker_type', 'priority', 'evidence_id'],
  '06p_Excel_PowerQuery_Register.csv': [
    'workbook',
    'sheet_name',
    'query_artifact',
    'm_code_status',
    'command_text',
    'connection',
    'evidence_id',
    'confidence',
    'recommended_action',
  ],
};

function addWorksheet(workbook: ExcelJS.Workbook, name: string, rows: Record<string, unknown>[]): void {
  const worksheet = workbook.addWorksheet(safeSheetName(name));
  const columns = columnsForDataset(name, rows);
  worksheet.columns = columns.map((key) => ({
    header: key,
    key,
    width: Math.min(Math.max(key.length + 4, 16), 42),
  }));
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];
  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: Math.max(rows.length + 1, 1), column: Math.max(columns.length, 1) },
  };
  worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  worksheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF111614' },
  };
  worksheet.getRow(1).alignment = { vertical: 'middle', wrapText: true };

  const bodyRows = rows.length ? rows : [emptyDatasetNotice(columns, 'No records captured for this tab; absence is evidence-reviewed, not a missing worksheet.')];
  for (const row of bodyRows) {
    worksheet.addRow(row);
  }

  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      cell.alignment = { vertical: 'top', wrapText: true };
    });
  });
}

function workbookRowsForTab(tab: string, model: DiscoveryModel): Record<string, unknown>[] {
  const lookup: Record<string, Record<string, unknown>[]> = {
    '00_Package_Control': [
      {
        package_name: model.packageName,
        generated_date: model.generatedDate,
        analysis_version: model.analysisVersion,
        package_version: model.packageVersion,
        master_dossier_standard_version: MASTER_DOSSIER_STANDARD_VERSION,
        master_dossier_standard_sha256: masterPromptHash(),
        ai_enabled: model.aiNarrative.enabled,
        ai_model: model.aiNarrative.model,
        qa_status: createSummary(model).qaStatus,
      },
    ],
    '01_Source_Inventory': model.sourceFiles,
    '02_Object_Inventory_All': model.nodes,
    '03_Business_Process_Steps': model.processSteps,
    '04_Data_Asset_Catalog': model.nodes.filter((node) => ['database', 'file', 'workbook', 'worksheet', 'table', 'document'].includes(node.node_type)),
    '05_Lineage_Nodes': model.nodes,
    '06_Lineage_Edges': model.edges,
    '07_Transformations_Rules': model.transformations,
    '08_Data_Elements': model.dataElements,
    '09_Data_Quality_Findings': model.dataQualityFindings,
    '10_Dependency_Usage_Map': model.dependencyUsage,
    '11_Controls_Exceptions': model.controls,
    '12_Security_Access': model.securityAccess,
    '13_Schedule_SLA': model.scheduleSla,
    '14_Failure_Modes': model.failureModes,
    '15_Financial_Exposure': model.financialExposure,
    '16_Modernization_Recommendations': model.modernization,
    '17_Action_Backlog': model.actions,
    '18_Open_Questions': model.openQuestions,
    '19_Evidence_Index': model.evidence.map(({ content, ...evidence }) => evidence),
    '20_QA_Reconciliation': model.qaRecords,
    '21_Access_Object_Inventory': model.access.Access_Object_Inventory,
    '22_Access_Table_Register': model.access.Access_Table_Register,
    '23_Access_Linked_Table_Register': model.access.Access_Linked_Table_Register,
    '24_Access_Query_Register': model.access.Access_Query_Register,
    '25_Access_Query_SQL_Index': model.access.Access_Query_SQL_Index,
    '26_Access_Macro_Register': model.access.Access_Macro_Register,
    '27_Access_Macro_XML_Storage': model.access.Access_Macro_XML_Storage,
    '28_Access_Macro_Action_Sequence': model.access.Access_Macro_Action_Sequence,
    '29_Access_Query_Macro_Reconciliation': model.access.Access_Query_Macro_Reconciliation,
    '30_Access_Form_Report_Register': model.access.Access_Form_Report_Register,
    '31_Access_Module_VBA_Register': model.access.Access_Module_VBA_Register,
    '32_Access_Import_Export_Specs': model.access.Access_Import_Export_Specs,
    '33_Access_Column_Inventory': model.access.Access_Column_Inventory,
    '34_Access_Data_Profile': model.access.Access_Data_Profile,
    '21_Excel_Workbook_Inventory': model.excel.Excel_Workbook_Inventory,
    '22_Excel_Sheet_Inventory': model.excel.Excel_Sheet_Inventory,
    '23_Excel_Table_NamedRange_Register': model.excel.Excel_Table_NamedRange_Register,
    '24_Excel_PowerQuery_Register': model.excel.Excel_PowerQuery_Register,
    '25_Excel_Formula_Register': model.excel.Excel_Formula_Register,
    '26_Excel_Connection_Register': model.excel.Excel_Connection_Register,
    '27_Excel_VBA_Register': model.excel.Excel_VBA_Register,
    '28_Excel_Pivot_DataModel_Register': model.excel.Excel_Pivot_DataModel_Register,
    '29_Excel_Data_Profile': model.excel.Excel_Data_Profile,
    '21_Word_Document_Inventory': model.word.Word_Document_Inventory,
    '22_Word_Section_Extracts': model.word.Word_Section_Extracts,
    '23_Word_Process_Rules': model.word.Word_Process_Rules,
    '24_Word_Control_Extracts': model.word.Word_Control_Extracts,
  };

  return lookup[tab] ?? [];
}

function createAutoDocumentationCsvs(model: DiscoveryModel): Record<string, string> {
  const files: Record<string, string> = {
    '06a_System_Inventory.csv': rowsToCsvForFile('06a_System_Inventory.csv', model.nodes.filter((node) => ['system', 'database', 'file', 'folder', 'workbook', 'document'].includes(node.node_type))),
    '06b_Object_Inventory.csv': rowsToCsvForFile('06b_Object_Inventory.csv', model.nodes),
    '06c_Process_Steps.csv': rowsToCsvForFile('06c_Process_Steps.csv', model.processSteps),
    '06d_Lineage_Nodes.csv': rowsToCsvForFile('06d_Lineage_Nodes.csv', model.nodes),
    '06e_Lineage_Edges.csv': rowsToCsvForFile('06e_Lineage_Edges.csv', model.edges),
    '06f_Transformations_Rules.csv': rowsToCsvForFile('06f_Transformations_Rules.csv', model.transformations),
    '06g_Controls_Exceptions.csv': rowsToCsvForFile('06g_Controls_Exceptions.csv', model.controls),
    '06h_Security_Access.csv': rowsToCsvForFile('06h_Security_Access.csv', model.securityAccess),
    '06i_Data_Quality_Findings.csv': rowsToCsvForFile('06i_Data_Quality_Findings.csv', model.dataQualityFindings),
    '06j_Dependency_Usage_Map.csv': rowsToCsvForFile('06j_Dependency_Usage_Map.csv', model.dependencyUsage),
    '06k_Open_Questions.csv': rowsToCsvForFile('06k_Open_Questions.csv', model.openQuestions),
  };

  for (const fileName of BASE_AUTO_CSVS) {
    files[fileName] ??= rowsToCsvForFile(fileName, []);
  }

  if (model.sourceTypeCounts.access) {
    files['06l_Access_Query_Register.csv'] = rowsToCsvForFile('06l_Access_Query_Register.csv', model.access.Access_Query_Register);
    files['06m_Access_Macro_Register.csv'] = rowsToCsvForFile('06m_Access_Macro_Register.csv', model.access.Access_Macro_Register);
    files['06n_Access_Macro_Action_Map.csv'] = rowsToCsvForFile('06n_Access_Macro_Action_Map.csv', model.access.Access_Macro_Action_Sequence);
    files['06o_Access_Query_Macro_Reconciliation.csv'] = rowsToCsvForFile('06o_Access_Query_Macro_Reconciliation.csv', model.access.Access_Query_Macro_Reconciliation);
    for (const fileName of ACCESS_AUTO_CSVS) {
      files[fileName] ??= rowsToCsvForFile(fileName, []);
    }
  }

  if (model.sourceTypeCounts.excel) {
    files['06p_Excel_PowerQuery_Register.csv'] = rowsToCsvForFile('06p_Excel_PowerQuery_Register.csv', model.excel.Excel_PowerQuery_Register);
    files['06q_Excel_Formula_Register.csv'] = rowsToCsvForFile('06q_Excel_Formula_Register.csv', model.excel.Excel_Formula_Register);
    files['06r_Excel_VBA_Register.csv'] = rowsToCsvForFile('06r_Excel_VBA_Register.csv', model.excel.Excel_VBA_Register);
    for (const fileName of EXCEL_AUTO_CSVS) {
      files[fileName] ??= rowsToCsvForFile(fileName, []);
    }
  }

  if (model.sourceTypeCounts.word) {
    files['06s_Word_Process_Extracts.csv'] = rowsToCsvForFile('06s_Word_Process_Extracts.csv', model.word.Word_Process_Rules);
    for (const fileName of WORD_AUTO_CSVS) {
      files[fileName] ??= rowsToCsvForFile(fileName, []);
    }
  }

  return files;
}

function createManifest(model: DiscoveryModel, fileCount: number): Record<string, unknown> {
  const automationStats = createAutomationStats(model);
  return {
    package_name: model.packageName,
    generated_date: model.generatedDate,
    source_files: model.sourceFiles.map((source) => source.file_name),
    source_file_sizes: Object.fromEntries(model.sourceFiles.map((source) => [source.file_name, source.file_size_bytes])),
    source_file_types: Object.fromEntries(model.sourceFiles.map((source) => [source.file_name, source.file_type])),
    source_modified_dates: 'not available from browser upload',
    package_version: model.packageVersion,
    analysis_version: model.analysisVersion,
    master_dossier_standard_version: MASTER_DOSSIER_STANDARD_VERSION,
    master_dossier_standard_sha256: masterPromptHash(),
    ai_enabled: model.aiNarrative.enabled,
    ai_model: model.aiNarrative.model,
    ai_cost_summary: model.aiNarrative.cost ?? createFallbackAiCostSummary(model.aiNarrative.model),
    folder_inventory: PACKAGE_FOLDERS,
    deliverable_inventory: [
      'README.md',
      'Executive_Decision_Brief.pdf',
      'Current_State_Architecture_Report.pdf',
      'Technical_Discovery_Workbook.xlsx',
      ...REQUIRED_DIAGRAMS.map((diagram) => diagram.file),
      ...Object.keys(createAutoDocumentationCsvs(model)),
      'Metadata_Manifest.json',
      'Action_Backlog.csv',
      'Financial_Impact_Model.xlsx',
    ],
    file_count: fileCount,
    object_counts: {
      nodes: model.nodes.length,
      edges: model.edges.length,
      evidence: model.evidence.length,
      actions: model.actions.length,
      data_quality_findings: model.dataQualityFindings.length,
    },
    source_type_counts: model.sourceTypeCounts,
    row_count_summary: model.rowCountSummary,
    query_count_summary: {
      query_nodes: model.nodes.filter((node) => node.node_type === 'query').length,
      access_saved_queries: model.access.Access_Query_Register.length,
    },
    macro_count_summary: {
      macro_entrypoints: automationStats.macroCount,
      macro_nodes: model.nodes.filter((node) => node.node_type === 'macro').length,
      macro_action_nodes: model.nodes.filter((node) => node.node_type === 'macro action').length,
      access_macros: model.access.Access_Macro_Register.length,
      excel_macro_entrypoints: automationStats.macroCount - model.access.Access_Macro_Register.length,
      excel_vba_modules: automationStats.vbaModuleCount,
      excel_vba_procedures: automationStats.vbaProcedureCount,
      excel_vba_functions: automationStats.vbaFunctionCount,
      excel_vba_event_handlers: automationStats.vbaEventHandlerCount,
      excel_vba_register_rows: model.excel.Excel_VBA_Register.length,
    },
    linked_source_count: model.nodes.filter((node) => node.node_type === 'linked table').length,
    lineage_blocker_count: model.nodes.filter((node) => node.node_type === 'upstream blocker').length,
    evidence_count: model.evidence.length,
    diagram_count: REQUIRED_DIAGRAMS.length,
    QA_status: createSummary(model).qaStatus,
    known_limitations: model.limitations,
    assumptions: model.assumptions,
    blocked_sources: model.blockedSources,
    checksum_hash_values: Object.fromEntries(model.sourceFiles.map((source) => [source.file_name, source.sha256])),
  };
}

function masterPromptHash(): string {
  return createHash('sha256').update(MASTER_DOSSIER_STANDARD).digest('hex');
}

function rowsToCsvForFile(fileName: string, rows: Record<string, unknown>[]): string {
  return rowsToCsv(rows, DATASET_COLUMNS[fileName]);
}

function rowsToCsv(rows: Record<string, unknown>[], preferredColumns: string[] = []): string {
  const columns = inferColumns(rows, preferredColumns);
  const header = columns.map(csvCell).join(',');
  const body = rows.map((row) => columns.map((column) => csvCell(row[column])).join(',')).join('\n');
  return body ? `${header}\n${body}\n` : `${header}\n`;
}

function columnsForDataset(name: string, rows: Record<string, unknown>[]): string[] {
  return inferColumns(rows, DATASET_COLUMNS[name] ?? []);
}

function inferColumns(rows: Record<string, unknown>[], preferredColumns: string[] = []): string[] {
  const columns = Array.from(new Set(rows.flatMap((row) => Object.keys(row))));
  return [...preferredColumns, ...columns.filter((column) => !preferredColumns.includes(column))].length
    ? [...preferredColumns, ...columns.filter((column) => !preferredColumns.includes(column))]
    : ['status'];
}

function emptyDatasetNotice(columns: string[], message: string): Record<string, unknown> {
  return Object.fromEntries(columns.map((column, index) => [column, index === 0 ? message : '']));
}

function csvCell(value: unknown): string {
  const text = value === undefined || value === null ? '' : typeof value === 'object' ? JSON.stringify(value) : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

async function workbookToBuffer(workbook: ExcelJS.Workbook): Promise<Buffer> {
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
}

async function createPdf(title: string, sections: { heading: string; lines: string[] }[]): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const document = new PDFDocument({ size: 'LETTER', margin: 48, bufferPages: true });
    const chunks: Buffer[] = [];
    document.on('data', (chunk: Buffer) => chunks.push(chunk));
    document.on('error', reject);
    document.on('end', () => resolve(Buffer.concat(chunks)));

    document.font('Helvetica-Bold').fontSize(20).fillColor('#111614').text(title, { lineGap: 4 });
    document.moveDown(0.4);
    document.font('Helvetica').fontSize(9).fillColor('#59615d').text(`Generated ${new Date().toISOString().slice(0, 10)}`);
    document.moveDown(1);

    for (const item of sections) {
      ensureRoom(document, 90);
      document.font('Helvetica-Bold').fontSize(13).fillColor('#111614').text(item.heading);
      document.moveDown(0.3);
      for (const line of item.lines.length ? item.lines : ['No records available.']) {
        ensureRoom(document, 48);
        document.font('Helvetica').fontSize(9.5).fillColor('#262d2a').text(`- ${line}`, {
          lineGap: 2,
          width: 500,
        });
      }
      document.moveDown(0.8);
    }

    const range = document.bufferedPageRange();
    for (let i = range.start; i < range.start + range.count; i += 1) {
      document.switchToPage(i);
      document.font('Helvetica').fontSize(8).fillColor('#6b736f').text(`Page ${i + 1} of ${range.count}`, 48, 742, {
        align: 'right',
        width: 500,
      });
    }

    document.end();
  });
}

function ensureRoom(document: PDFKit.PDFDocument, height: number): void {
  if (document.y + height > document.page.height - document.page.margins.bottom) {
    document.addPage();
  }
}

function topFindings(model: DiscoveryModel): string[] {
  const findings = [
    ...model.nodes
      .filter((node) => node.confidence === 'blocked')
      .slice(0, 3)
      .map((node) => `${node.node_id}: ${node.name}. Risk: ${node.failure_impact}. Evidence: ${node.evidence_id}. Confidence: ${node.confidence}. Action: ${node.recommended_action}`),
    ...model.dataQualityFindings
      .slice(0, 3)
      .map((finding) => `${finding.finding_id}: ${finding.issue} on ${finding.asset}.${finding.field}. Evidence: ${finding.evidence_id}. Confidence: ${finding.confidence}. Action: ${finding.recommended_fix}`),
    ...model.openQuestions
      .slice(0, 3)
      .map((question) => `${question.question_id}: ${question.question}. Evidence: ${question.evidence_id}. Action: owner confirmation required.`),
  ];

  return findings.length ? findings.slice(0, 5) : ['No high-impact findings beyond standard owner, finance, security, and cadence validation were detected.'];
}

function graphSnapshotLines(model: DiscoveryModel): string[] {
  const nodeLines = model.nodes
    .slice(0, 20)
    .map((node) => `${node.node_id} [${node.node_type}; ${node.criticality}; ${node.confidence}]: ${node.name}`);
  const edgeLines = model.edges
    .slice(0, 20)
    .map((edge) => `${edge.edge_id} ${edge.from_node_id} -${edge.edge_type}-> ${edge.to_node_id} [${edge.automated_flag}; ${edge.confidence}]`);
  return [
    `Node count: ${model.nodes.length}. Edge count: ${model.edges.length}.`,
    ...nodeLines,
    ...edgeLines,
    model.nodes.length > 20 || model.edges.length > 20
      ? 'Graph is summarized here for readability. See workbook tabs 05_Lineage_Nodes and 06_Lineage_Edges for complete detail.'
      : 'Complete graph snapshot shown.',
  ];
}

function currency(value: number): string {
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(value);
}

function safeSheetName(name: string): string {
  return name.replace(/[\\/?*:[\]]/g, '_').slice(0, 31);
}
