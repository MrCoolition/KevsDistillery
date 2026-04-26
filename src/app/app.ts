import { ChangeDetectionStrategy, Component, computed, signal } from '@angular/core';

import { aiReadiness, discoveryModel } from './sample-discovery';
import { BacklogAction, Criticality, DiscoveryItem } from './discovery-model';

type ViewId = 'command' | 'imports' | 'model' | 'lineage' | 'reports' | 'impact' | 'backlog';
type SourceKind = 'access' | 'excel' | 'word' | 'database' | 'interview' | 'mixed';

interface NavItem {
  id: ViewId;
  label: string;
}

interface ImportSource {
  kind: SourceKind;
  label: string;
  description: string;
}

interface ApiHealth {
  ok: boolean;
  model: string;
  openAIConfigured: boolean;
  database: {
    configured: boolean;
    ready: boolean;
    schema?: string;
    driver?: string;
    tableCount?: number;
    requiredTableCount?: number;
    missingTables?: string[];
    error?: string;
  };
}

interface SynthesisResponse {
  ok: boolean;
  runId?: string;
  stored?: boolean;
  persistenceError?: string | null;
  counts?: {
    items: number;
    relationships: number;
    artifacts: number;
    backlog: number;
  } | null;
  model?: string;
  fallbackReason?: string | null;
  outputText?: string;
  canonicalDelta?: {
    processName?: string;
    businessFunction?: string;
    recommendation?: string;
    decisionRequired?: string;
    overallRiskRating?: string;
    estimatedDollarExposure?: Record<string, unknown>;
    executiveBrief?: Record<string, unknown>;
    reportSections?: ReportSection[];
    items?: unknown[];
    relationships?: unknown[];
    artifacts?: Array<{
      id?: string;
      name?: string;
      type?: string;
      status?: string;
      audience?: string;
      purpose?: string;
    }>;
    backlog?: unknown[];
  } | null;
  error?: string;
}

interface ReportSection {
  title?: string;
  body?: string;
  confidence?: number;
  evidenceIds?: string[];
}

interface HistoricalRun {
  id: string;
  created_at: string;
  source_kind: string;
  source_name: string;
  model: string;
  status: string;
  confidence: number | null;
  evidence_chars: number;
  item_count: number;
  backlog_count: number;
}

interface StagedSource {
  name: string;
  path: string;
  size: number;
  extension: string;
  status: string;
  text: string;
}

@Component({
  selector: 'td-root',
  standalone: true,
  imports: [],
  templateUrl: './app.html',
  styleUrl: './app.css',
  changeDetection: ChangeDetectionStrategy.OnPush
})
export class App {
  readonly model = discoveryModel;
  readonly ai = aiReadiness;
  readonly activeView = signal<ViewId>('imports');
  readonly selectedItemId = signal<string>('SRC-001');
  readonly searchTerm = signal('');
  readonly importSourceKind = signal<SourceKind>('excel');
  readonly importSourceName = signal('Uncle Kev source batch');
  readonly knownArtifactsText = signal('');
  readonly targetOutputsText = signal('01_Executive_Decision_Brief.pdf\n02_Current_State_Architecture_Report.pdf\n03_Technical_Discovery_Workbook.xlsx\n04_Auto_Documentation_Pack\n05_Diagram_Pack\n06_Financial_Impact_Model.xlsx\n07_Action_Backlog.csv\n08_Evidence_Archive\n09_Metadata_Manifest.json');
  readonly extractedText = signal('Choose files or a folder above. Extracted evidence will collect here before the still runs.');
  readonly apiHealth = signal<ApiHealth | null>(null);
  readonly apiError = signal('');
  readonly isSynthesizing = signal(false);
  readonly synthesisResult = signal<SynthesisResponse | null>(null);
  readonly synthesisError = signal('');
  readonly historicalRuns = signal<HistoricalRun[]>([]);
  readonly stagedSources = signal<StagedSource[]>([]);
  readonly isExtractingSources = signal(false);

  readonly navItems: NavItem[] = [
    { id: 'command', label: 'Command' },
    { id: 'imports', label: 'Imports' },
    { id: 'model', label: 'Canonical Model' },
    { id: 'lineage', label: 'Lineage' },
    { id: 'reports', label: 'Reports' },
    { id: 'impact', label: 'Impact' },
    { id: 'backlog', label: 'Backlog' }
  ];

  readonly importSources: ImportSource[] = [
    {
      kind: 'access',
      label: 'Access',
      description: 'ACCDB metadata, saved SQL, linked tables, macros, VBA, forms, reports.'
    },
    {
      kind: 'excel',
      label: 'Excel',
      description: 'Formulas, hidden sheets, Power Query M, pivots, external links, VBA.'
    },
    {
      kind: 'word',
      label: 'Word',
      description: 'Process steps, controls, approvals, rules, exceptions, SLA language.'
    },
    {
      kind: 'database',
      label: 'Database',
      description: 'Schemas, tables, keys, indexes, row counts, dependency and profile evidence.'
    },
    {
      kind: 'interview',
      label: 'Interview',
      description: 'Tribal rules, owners, recovery paths, manual workarounds, business impact.'
    },
    {
      kind: 'mixed',
      label: 'Mixed Pack',
      description: 'Multi-source evidence batches for a full still run.'
    }
  ];

  readonly selectedItem = computed(() => {
    return this.model.items.find((item) => item.id === this.selectedItemId()) ?? this.model.items[0];
  });

  readonly filteredItems = computed(() => {
    const term = this.searchTerm().trim().toLowerCase();
    if (!term) {
      return this.model.items;
    }

    return this.model.items.filter((item) => {
      return [
        item.id,
        item.name,
        item.type,
        item.owner,
        item.businessPurpose,
        item.recommendedAction.summary,
        ...item.tags
      ].some((value) => value.toLowerCase().includes(term));
    });
  });

  readonly completeItems = computed(() => {
    return this.model.items.filter((item) => this.isFinished(item)).length;
  });

  readonly packageProgress = computed(() => {
    const total = this.model.artifacts.reduce((sum, artifact) => sum + artifact.progress, 0);
    return Math.round(total / this.model.artifacts.length);
  });

  readonly averageConfidence = computed(() => {
    const total = this.model.items.reduce((sum, item) => sum + item.confidence, 0);
    return Math.round(total / this.model.items.length);
  });

  readonly impactBasis = computed(() => this.model.estimatedDollarExposure.assumptions);

  readonly selectedUpstream = computed(() => {
    const item = this.selectedItem();
    return item.upstream.map((id) => this.itemById(id)).filter((node): node is DiscoveryItem => Boolean(node));
  });

  readonly selectedDownstream = computed(() => {
    const item = this.selectedItem();
    return item.downstream.map((id) => this.itemById(id)).filter((node): node is DiscoveryItem => Boolean(node));
  });

  readonly lineageTrace = computed(() => this.traceUpstream(this.selectedItemId()));

  readonly criticalBacklog = computed(() => {
    const priorityOrder: Record<BacklogAction['priority'], number> = { P0: 0, P1: 1, P2: 2, P3: 3 };
    return [...this.model.backlog].sort((a, b) => priorityOrder[a.priority] - priorityOrder[b.priority]);
  });

  readonly knownArtifacts = computed(() => this.toLines(this.knownArtifactsText()));
  readonly targetOutputs = computed(() => this.toLines(this.targetOutputsText()));

  readonly liveArtifacts = computed(() => {
    const generatedArtifacts = this.synthesisResult()?.canonicalDelta?.artifacts;
    if (generatedArtifacts?.length) {
      return generatedArtifacts.map((artifact, index) => ({
        id: artifact.id || String(index + 1).padStart(2, '0'),
        name: artifact.name || 'Generated artifact',
        audience: artifact.audience || 'Distillery crew',
        purpose: artifact.purpose || artifact.type || 'Generated from canonical discovery model.',
        progress: artifact.status === 'final' ? 100 : 76,
        sourceModel: 'canonical graph' as const
      }));
    }

    return this.model.artifacts;
  });

  readonly generatedCounts = computed(() => {
    const result = this.synthesisResult();
    return {
      items: result?.counts?.items ?? result?.canonicalDelta?.items?.length ?? this.model.items.length,
      relationships: result?.counts?.relationships ?? result?.canonicalDelta?.relationships?.length ?? this.model.relationships.length,
      artifacts: result?.counts?.artifacts ?? result?.canonicalDelta?.artifacts?.length ?? this.model.artifacts.length,
      backlog: result?.counts?.backlog ?? result?.canonicalDelta?.backlog?.length ?? this.model.backlog.length
    };
  });

  readonly sourceStats = computed(() => {
    const sources = this.stagedSources();
    const bytes = sources.reduce((sum, source) => sum + source.size, 0);
    const readable = sources.filter((source) => source.status.includes('extracted')).length;
    const evidenceChars = sources.reduce((sum, source) => sum + source.text.length, 0);
    return {
      files: sources.length,
      bytes: this.formatBytes(bytes),
      readable,
      evidenceChars: evidenceChars.toLocaleString()
    };
  });

  readonly limitedExtraction = computed(() => {
    return this.stagedSources().some((source) => {
      return source.status.includes('metadata') || source.status.includes('binary') || source.status.includes('failed');
    });
  });

  readonly runStatus = computed(() => {
    if (this.isSynthesizing()) {
      return 'Running gpt-5.5 against the staged evidence now. The still is working.';
    }

    const result = this.synthesisResult();
    if (result) {
      const counts = result.counts;
      const countText = counts
        ? `${counts.items} items, ${counts.relationships} relationships, ${counts.artifacts} artifacts, ${counts.backlog} actions`
        : 'canonical output returned';
      const persistence = result.stored
        ? 'Persisted to Neon.'
        : `Not persisted to Neon: ${result.persistenceError || 'database write failed.'}`;
      const fallback = result.fallbackReason ? ` ${result.fallbackReason}.` : '';
      const analysis = result.model === 'gpt-5.5'
        ? 'OpenAI analyzed the staged evidence'
        : 'Fallback analysis generated a blocker-backed action pack';
      return `${analysis}: ${countText}.${fallback} ${persistence}`;
    }

    if (!this.stagedSources().length) {
      return 'Ready. Choose files or a folder to stage a source batch.';
    }

    const stats = this.sourceStats();
    if (this.limitedExtraction()) {
      return `${stats.files} source staged (${stats.bytes}). Browser extraction found ${stats.evidenceChars} evidence characters. Some files still need native exports or connector metadata for full object-level lineage, but the retrieved evidence is ready for analysis.`;
    }

    return `${stats.files} source staged (${stats.bytes}) with ${stats.readable} readable files and ${stats.evidenceChars} evidence characters. Ready to run analysis.`;
  });

  readonly reportSections = computed(() => {
    const generated = this.synthesisResult()?.canonicalDelta?.reportSections;
    if (generated?.length) {
      return generated
        .filter((section) => section?.title || section?.body)
        .map((section, index) => ({
          title: section.title || `Analysis Section ${index + 1}`,
          body: section.body || 'No narrative returned for this section.',
          confidence: typeof section.confidence === 'number' ? section.confidence : null,
          evidenceIds: Array.isArray(section.evidenceIds) ? section.evidenceIds : []
        }));
    }

    const delta = this.synthesisResult()?.canonicalDelta;
    if (delta) {
      return this.sectionsFromCanonicalDelta(delta);
    }

    return [
      {
        title: 'Executive Snapshot',
        body: `${this.model.processName} supports ${this.model.businessFunction}. Run a real source batch to replace this starter state with evidence-backed scope, lineage, blockers, confidence, and actions.`,
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Current-State Narrative',
        body: 'The generated report compresses the operating model into trigger, actor, input, processing step, validation, output, handoff, and exception path sections while preserving evidence references.',
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Lineage and Controls',
        body: `${this.model.relationships.length} starter node-edge relationships connect the source set, canonical graph, and generated outputs. Runs expand this into real lineage.`,
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Remediation Backlog',
        body: `${this.criticalBacklog().length} action records are prioritized for Fivetran ingestion, dbt rebuild, Snowpark controls, governance, and retirement decisions.`,
        confidence: null,
        evidenceIds: []
      }
    ];
  });

  constructor() {
    void this.refreshProductionState();
  }

  selectView(id: ViewId): void {
    this.activeView.set(id);
  }

  selectItem(id: string): void {
    this.selectedItemId.set(id);
  }

  updateSearch(value: string): void {
    this.searchTerm.set(value);
  }

  updateImportSourceKind(value: string): void {
    this.importSourceKind.set(value as SourceKind);
  }

  async refreshProductionState(): Promise<void> {
    await this.checkApiHealth();
    await this.loadRuns();
  }

  async checkApiHealth(): Promise<void> {
    try {
      this.apiError.set('');
      const response = await fetch('/api/health');
      const body = await this.readApiJson(response, 'Health check failed.');
      if (!response.ok) {
        throw new Error(body.error || 'Health check failed.');
      }
      this.apiHealth.set(body);
    } catch (error) {
      this.apiError.set(error instanceof Error ? error.message : 'Health check failed.');
    }
  }

  async loadRuns(): Promise<void> {
    try {
      const response = await fetch('/api/discovery/runs?limit=8');
      const body = await this.readApiJson(response, 'Could not load runs.');
      if (!response.ok) {
        throw new Error(body.error || 'Could not load runs.');
      }
      this.historicalRuns.set(body.runs || []);
    } catch {
      this.historicalRuns.set([]);
    }
  }

  async runSynthesis(): Promise<void> {
    if (!this.extractedText().trim()) {
      this.synthesisError.set('No evidence payload is staged. Choose files or paste evidence before running analysis.');
      return;
    }

    if (this.stagedSources().length > 0) {
      const confirmed = window.confirm('Run Uncle Kev\'s analysis? This sends extracted evidence from the selected sources to your configured backend and OpenAI model.');
      if (!confirmed) {
        return;
      }
    }

    this.isSynthesizing.set(true);
    this.synthesisError.set('');
    this.synthesisResult.set(null);

    try {
      const response = await fetch('/api/discovery/synthesize', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          sourceKind: this.importSourceKind(),
          sourceName: this.importSourceName(),
          knownArtifacts: this.knownArtifacts(),
          targetOutputs: this.targetOutputs(),
          extractedText: this.extractedText()
        })
      });
      const body = await this.readApiJson(response, 'Synthesis failed.');
      if (!response.ok) {
        throw new Error(body.error || 'Synthesis failed.');
      }
      this.synthesisResult.set(body);
      this.activeView.set('reports');
      await this.loadRuns();
    } catch (error) {
      this.synthesisError.set(error instanceof Error ? error.message : 'Synthesis failed.');
    } finally {
      this.isSynthesizing.set(false);
    }
  }

  private async readApiJson(response: Response, fallbackMessage: string): Promise<any> {
    const text = await response.text();
    try {
      return text ? JSON.parse(text) : {};
    } catch {
      const firstLine = text.split(/\r?\n/).find((line) => line.trim())?.trim();
      return {
        ok: false,
        error: firstLine || `${fallbackMessage} The server returned a non-JSON response.`
      };
    }
  }

  async stageFiles(fileList: FileList | null): Promise<void> {
    const files = Array.from(fileList || []);
    if (!files.length) {
      return;
    }

    this.isExtractingSources.set(true);
    this.stagedSources.set([]);

    const sources: StagedSource[] = [];
    for (const file of files) {
      const source = await this.extractFile(file);
      sources.push(source);
      this.stagedSources.set([...sources]);
    }

    this.importSourceKind.set(this.inferSourceKind(sources));
    this.importSourceName.set(files.length === 1 ? files[0].name : `${files.length} source files in the mash bill`);
    this.knownArtifactsText.set(sources.map((source) => source.path).join('\n'));
    this.extractedText.set(this.buildEvidencePayload(sources));
    this.synthesisResult.set(null);
    this.synthesisError.set('');
    this.isExtractingSources.set(false);
  }

  clearSources(fileInput?: HTMLInputElement, folderInput?: HTMLInputElement): void {
    this.stagedSources.set([]);
    this.extractedText.set('Choose files or a folder above. Extracted evidence will collect here before the still runs.');
    this.synthesisResult.set(null);
    this.synthesisError.set('');
    if (fileInput) {
      fileInput.value = '';
    }
    if (folderInput) {
      folderInput.value = '';
    }
  }

  async applyNeonSchema(): Promise<void> {
    try {
      this.apiError.set('');
      const response = await fetch('/api/admin/migrate', {
        method: 'POST'
      });
      const body = await this.readApiJson(response, 'Neon schema migration failed.');
      if (!response.ok) {
        throw new Error(body.error || 'Neon schema migration failed.');
      }
      await this.checkApiHealth();
    } catch (error) {
      this.apiError.set(error instanceof Error ? error.message : 'Neon schema migration failed.');
    }
  }

  databaseStatusText(health: ApiHealth): string {
    const database = health.database;
    if (!database.configured) {
      return 'not configured';
    }

    if (database.ready) {
      const tableText = database.tableCount && database.requiredTableCount
        ? ` (${database.tableCount}/${database.requiredTableCount} tables)`
        : '';
      return `${database.schema || 'distillery'} schema ready${tableText}`;
    }

    if (database.error) {
      return `error: ${database.error}`;
    }

    const missing = database.missingTables?.length
      ? `missing ${database.missingTables.join(', ')}`
      : 'not ready';
    return `${database.schema || 'distillery'} schema ${missing}`;
  }

  itemById(id: string): DiscoveryItem | undefined {
    return this.model.items.find((item) => item.id === id);
  }

  isFinished(item: DiscoveryItem): boolean {
    return item.evidence.length > 0 && item.confidence > 0 && item.recommendedAction.summary.length > 0;
  }

  criticalityRank(criticality: Criticality): number {
    const rank: Record<Criticality, number> = {
      critical: 4,
      high: 3,
      medium: 2,
      low: 1
    };
    return rank[criticality];
  }

  private traceUpstream(id: string, visited = new Set<string>()): DiscoveryItem[] {
    if (visited.has(id)) {
      return [];
    }

    visited.add(id);
    const item = this.itemById(id);
    if (!item) {
      return [];
    }

    const upstream = item.upstream.flatMap((upstreamId) => this.traceUpstream(upstreamId, visited));
    return [...upstream, item];
  }

  private toLines(value: string): string[] {
    return value
      .split(/\r?\n|,/)
      .map((line) => line.trim())
      .filter(Boolean);
  }

  private async extractFile(file: File): Promise<StagedSource> {
    const extension = this.extensionFor(file);
    const path = (file as File & { webkitRelativePath?: string }).webkitRelativePath || file.name;
    let status = 'metadata only';
    let text = '';

    try {
      if (extension === 'docx') {
        text = await this.extractDocx(file);
        status = text ? 'extracted docx' : 'empty docx';
      } else if (extension === 'xlsx' || extension === 'xlsm') {
        text = await this.extractXlsx(file);
        status = text ? 'extracted workbook' : 'empty workbook';
      } else if (this.isTextFile(file, extension)) {
        text = await file.text();
        status = 'extracted text';
      } else if (extension === 'accdb' || extension === 'mdb') {
        text = await this.extractAccessBinary(file);
        status = text.includes('Recovered strings:') ? 'extracted access strings' : 'access binary metadata';
      } else {
        text = `Binary source selected with MIME type ${file.type || 'unknown'}. Add a text/XML/CSV export or connector extraction for deep parsing.`;
      }
    } catch (error) {
      text = `Extraction failed: ${error instanceof Error ? error.message : 'unknown error'}`;
      status = 'extract failed';
    }

    return {
      name: file.name,
      path,
      size: file.size,
      extension,
      status,
      text
    };
  }

  private async extractDocx(file: File): Promise<string> {
    const { default: JSZip } = await import('jszip');
    const zip = await JSZip.loadAsync(file);
    const parts: string[] = [];
    const documentXml = await zip.file('word/document.xml')?.async('text');
    if (documentXml) {
      parts.push(`Document text: ${this.cleanXmlText(documentXml)}`);
    }

    const paths = Object.keys(zip.files).filter((path) => /word\/comments\.xml|docProps\//i.test(path)).slice(0, 30);
    for (const path of paths) {
      const text = await zip.file(path)?.async('text');
      if (text) {
        parts.push(`${path}: ${this.cleanXmlText(text)}`);
      }
    }

    return parts.join('\n\n');
  }

  private async extractXlsx(file: File): Promise<string> {
    const { default: JSZip } = await import('jszip');
    const zip = await JSZip.loadAsync(file);
    const parts: string[] = [];
    const workbookXml = await zip.file('xl/workbook.xml')?.async('text');
    if (workbookXml) {
      const sheetNames = Array.from(workbookXml.matchAll(/name="([^"]+)"/g)).map((match) => match[1]);
      parts.push(`Workbook sheets: ${sheetNames.join(', ') || 'not detected'}`);
    }

    const sharedStrings = await zip.file('xl/sharedStrings.xml')?.async('text');
    if (sharedStrings) {
      parts.push(`Shared strings sample: ${this.cleanXmlText(sharedStrings).slice(0, 12000)}`);
    }

    const worksheetPaths = Object.keys(zip.files).filter((path) => /xl\/worksheets\/sheet.*\.xml/i.test(path)).slice(0, 30);
    for (const path of worksheetPaths) {
      const xml = await zip.file(path)?.async('text');
      if (!xml) {
        continue;
      }
      const formulas = Array.from(xml.matchAll(/<f[^>]*>([\s\S]*?)<\/f>/g)).map((match) => match[1]).slice(0, 300);
      const dimensions = xml.match(/<dimension ref="([^"]+)"/)?.[1] || 'unknown';
      parts.push(`${path}: dimension ${dimensions}; formulas: ${formulas.join(' | ') || 'none detected'}`);
    }

    const metadataPaths = Object.keys(zip.files)
      .filter((path) => /connections|query|customXml|pivot|externalLink/i.test(path))
      .slice(0, 30);
    for (const path of metadataPaths) {
      const text = await zip.file(path)?.async('text');
      if (text) {
        parts.push(`${path}: ${this.cleanXmlText(text).slice(0, 10000)}`);
      }
    }

    return parts.join('\n\n');
  }

  private async extractAccessBinary(file: File): Promise<string> {
    const buffer = await file.arrayBuffer();
    const bytes = new Uint8Array(buffer);
    const strings = new Set<string>();
    let asciiRun = '';
    let utf16Run = '';

    const keep = (value: string): void => {
      const cleaned = value.replace(/\s+/g, ' ').trim();
      if (cleaned.length < 4 || cleaned.length > 180) {
        return;
      }
      if (!/[A-Za-z]/.test(cleaned)) {
        return;
      }
      const signal = /[A-Za-z0-9_.$# -]{4,}/.test(cleaned);
      if (!signal) {
        return;
      }
      strings.add(cleaned);
    };

    for (let index = 0; index < bytes.length; index += 1) {
      const byte = bytes[index];
      if (byte >= 32 && byte <= 126) {
        asciiRun += String.fromCharCode(byte);
      } else {
        keep(asciiRun);
        asciiRun = '';
      }

      if (index + 1 < bytes.length && bytes[index + 1] === 0 && byte >= 32 && byte <= 126) {
        utf16Run += String.fromCharCode(byte);
        index += 1;
      } else {
        keep(utf16Run);
        utf16Run = '';
      }

      if (strings.size >= 1800) {
        break;
      }
    }

    keep(asciiRun);
    keep(utf16Run);

    const ranked = [...strings]
      .filter((value) => !/^(Microsoft|Standard|General|Normal|Arial|Calibri)$/i.test(value))
      .sort((a, b) => this.accessStringScore(b) - this.accessStringScore(a))
      .slice(0, 900);

    if (!ranked.length) {
      return [
        'Access binary selected.',
        'Browser-side string recovery found no readable object names or business terms.',
        'Use exported Access object inventory, saved SQL, linked table metadata, VBA modules, forms/reports, or a native extractor for deep discovery.'
      ].join('\n');
    }

    return [
      'Access binary string recovery.',
      `File: ${file.name}`,
      `Size: ${file.size} bytes`,
      `Recovered string count: ${ranked.length}`,
      'Interpretation note: these recovered strings are source evidence hints, not a full Access catalog. OpenAI must analyze them, document confidence, and call out exact exports needed to finish object-level lineage.',
      '',
      'Recovered strings:',
      ranked.join('\n')
    ].join('\n').slice(0, 70000);
  }

  private buildEvidencePayload(sources: StagedSource[]): string {
    return sources.map((source, index) => [
      `--- SOURCE ${index + 1}: ${source.path} ---`,
      `name: ${source.name}`,
      `extension: ${source.extension || 'none'}`,
      `size: ${source.size} bytes`,
      `status: ${source.status}`,
      '',
      source.text
    ].join('\n')).join('\n\n').slice(0, 80000);
  }

  private extensionFor(file: File): string {
    const name = file.name.toLowerCase();
    const index = name.lastIndexOf('.');
    return index >= 0 ? name.slice(index + 1) : '';
  }

  private isTextFile(file: File, extension: string): boolean {
    const textExtensions = new Set([
      'txt', 'csv', 'tsv', 'sql', 'bas', 'vba', 'm', 'pq', 'json', 'xml',
      'md', 'log', 'ini', 'yml', 'yaml', 'html', 'css', 'js', 'ts'
    ]);
    return file.type.startsWith('text/') || textExtensions.has(extension);
  }

  private cleanXmlText(xml: string): string {
    return xml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  }

  private inferSourceKind(sources: StagedSource[]): SourceKind {
    const extensions = new Set(sources.map((source) => source.extension));
    if (extensions.has('accdb') || extensions.has('mdb')) {
      return 'access';
    }
    if (extensions.has('xlsx') || extensions.has('xlsm') || extensions.has('xls')) {
      return 'excel';
    }
    if (extensions.has('docx') || extensions.has('doc')) {
      return 'word';
    }
    if (['sql', 'csv', 'tsv', 'json', 'xml'].some((extension) => extensions.has(extension))) {
      return sources.length > 1 ? 'mixed' : 'database';
    }
    return sources.length > 1 ? 'mixed' : this.importSourceKind();
  }

  private sectionsFromCanonicalDelta(delta: NonNullable<SynthesisResponse['canonicalDelta']>) {
    return [
      {
        title: 'Executive Snapshot',
        body: `${delta.processName || this.importSourceName()} was analyzed for ${delta.businessFunction || 'data discovery and migration readiness'}. ${delta.recommendation || 'Review generated items, lineage, artifacts, and actions.'}`,
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Scope Coverage and Confidence',
        body: `${delta.items?.length || 0} canonical items, ${delta.relationships?.length || 0} relationships, ${delta.artifacts?.length || 0} artifacts, and ${delta.backlog?.length || 0} backlog actions were generated from the submitted evidence.`,
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Decisions Needed',
        body: delta.decisionRequired || 'Confirm ownership, unresolved evidence gaps, and migration action priority.',
        confidence: null,
        evidenceIds: []
      }
    ];
  }

  private accessStringScore(value: string): number {
    let score = 0;
    if (/[A-Za-z]+_[A-Za-z0-9_]+/.test(value)) {
      score += 8;
    }
    if (/\b(qry|tbl|frm|rpt|macro|module|select|insert|update|delete|join|where|from)\b/i.test(value)) {
      score += 12;
    }
    if (/[A-Z][a-z]+[A-Z][A-Za-z]+/.test(value)) {
      score += 6;
    }
    if (/\b(customer|invoice|order|claim|policy|account|product|payment|member|provider|vendor|employee|health|fettle)\b/i.test(value)) {
      score += 10;
    }
    return score + Math.min(value.length, 60) / 10;
  }

  private formatBytes(bytes: number): string {
    if (!bytes) {
      return '0 KB';
    }
    const units = ['B', 'KB', 'MB', 'GB'];
    let size = bytes;
    let unit = 0;
    while (size >= 1024 && unit < units.length - 1) {
      size /= 1024;
      unit += 1;
    }
    return `${size.toFixed(size >= 10 || unit === 0 ? 0 : 1)} ${units[unit]}`;
  }
}
