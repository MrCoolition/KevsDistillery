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
  queued?: boolean;
  needsPayload?: boolean;
  responseId?: string;
  responseStatus?: string;
  message?: string;
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
    systemsInScope?: string[];
    criticalOutputs?: string[];
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
  heading?: string;
  name?: string;
  section?: string;
  sectionTitle?: string;
  body?: string;
  content?: string;
  summary?: string;
  narrative?: string;
  confidence?: number;
  evidenceIds?: string[];
  evidence_ids?: string[];
  evidence?: unknown;
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

interface DownloadAsset {
  url: string;
  filename: string;
}

interface LiveArtifact {
  id: string;
  name: string;
  audience: string;
  purpose: string;
  progress: number;
  sourceModel: 'canonical graph';
}

interface PackageContext {
  result: SynthesisResponse;
  delta: NonNullable<SynthesisResponse['canonicalDelta']>;
  deltaRecord: Record<string, unknown>;
  items: Record<string, unknown>[];
  relationships: Record<string, unknown>[];
  artifacts: Record<string, unknown>[];
  backlog: Record<string, unknown>[];
  evidenceIndex: Record<string, unknown>[];
  lineageNodes: Record<string, unknown>[];
  lineageEdges: Record<string, unknown>[];
  failureRisks: Record<string, unknown>[];
  openQuestions: Record<string, unknown>[];
  sections: Array<{ title: string; body: string; confidence: number | null; evidenceIds: string[] }>;
  manifest: Record<string, unknown>;
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
  private readonly reportSectionTitles = [
    'Executive Snapshot',
    'Scope, Coverage, and Confidence',
    'Business Mission of the Process',
    'Current-State Operating Model',
    'System and Artifact Landscape',
    'Data Flow and Process Flow Summary',
    'Transformations and Business Logic',
    'Recursive Lineage and Source-of-Truth Assessment',
    'Controls, Exceptions, and Failure Modes',
    'Financial Impact and Business Exposure',
    'Recommendations and Action Plan',
    'Open Questions and Decisions Needed'
  ];

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
  readonly analysisStatus = signal('');
  readonly synthesisResult = signal<SynthesisResponse | null>(null);
  readonly synthesisError = signal('');
  readonly packMessage = signal('');
  readonly packError = signal('');
  readonly lastDownload = signal<DownloadAsset | null>(null);
  readonly isGeneratingPack = signal(false);
  readonly historicalRuns = signal<HistoricalRun[]>([]);
  readonly stagedSources = signal<StagedSource[]>([]);
  readonly isExtractingSources = signal(false);
  private activeDownloadUrl = '';

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
    if (this.synthesisResult()) {
      return 100;
    }
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

  readonly liveArtifacts = computed<LiveArtifact[]>(() => {
    const generatedArtifacts = this.synthesisResult()?.canonicalDelta?.artifacts;
    if (generatedArtifacts?.length) {
      return generatedArtifacts.map((artifact, index) => ({
        id: artifact.id || String(index + 1).padStart(2, '0'),
        name: artifact.name || 'Generated artifact',
        audience: artifact.audience || 'Distillery crew',
        purpose: artifact.purpose || artifact.type || 'Generated from canonical discovery model.',
        progress: this.synthesisResult() ? 100 : artifact.status === 'final' ? 100 : 76,
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
      return this.analysisStatus() || 'Starting gpt-5.5 background analysis against the staged evidence now.';
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

  async runSynthesis(): Promise<boolean> {
    if (!this.extractedText().trim()) {
      this.synthesisError.set('No evidence payload is staged. Choose files or paste evidence before running analysis.');
      return false;
    }

    if (this.stagedSources().length > 0) {
      const confirmed = window.confirm('Run Uncle Kev\'s analysis? This sends extracted evidence from the selected sources to your configured backend and OpenAI model.');
      if (!confirmed) {
        return false;
      }
    }

    this.isSynthesizing.set(true);
    this.synthesisError.set('');
    this.packError.set('');
    this.packMessage.set('');
    this.clearDownloadUrl();
    this.analysisStatus.set('Starting OpenAI background analysis.');
    this.synthesisResult.set(null);

    try {
      const payload = {
        sourceKind: this.importSourceKind(),
        sourceName: this.importSourceName(),
        knownArtifacts: this.knownArtifacts(),
        targetOutputs: this.targetOutputs(),
        extractedText: this.extractedText()
      };
      const response = await fetch('/api/discovery/synthesize', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
      });
      const body = await this.readApiJson(response, 'Synthesis failed.');
      if (!response.ok) {
        throw new Error(body.error || 'Synthesis failed.');
      }

      const completed = body.queued && body.responseId
        ? await this.pollSynthesis(body.responseId, payload)
        : body;

      this.synthesisResult.set(completed);
      this.activeView.set('reports');
      await this.loadRuns();
      return true;
    } catch (error) {
      this.synthesisError.set(error instanceof Error ? error.message : 'Synthesis failed.');
      return false;
    } finally {
      this.isSynthesizing.set(false);
      this.analysisStatus.set('');
    }
  }

  private async pollSynthesis(responseId: string, payload: Record<string, unknown>): Promise<SynthesisResponse> {
    const maxAttempts = 120;
    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      this.analysisStatus.set(`OpenAI background analysis ${attempt === 1 ? 'started' : 'still running'} (${attempt}/${maxAttempts}).`);
      await this.delay(3000);

      const response = await fetch('/api/discovery/status', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          responseId
        })
      });
      const body = await this.readApiJson(response, 'Could not check synthesis status.');
      if (response.status === 202 || body.queued) {
        this.analysisStatus.set(`OpenAI status: ${body.responseStatus || 'in progress'}. Uncle Kev is still distilling.`);
        continue;
      }
      if (!response.ok) {
        throw new Error(body.error || 'Could not check synthesis status.');
      }
      if (body.needsPayload) {
        this.analysisStatus.set('OpenAI finished. Persisting the canonical model to Neon.');
        return this.completeSynthesis(responseId, payload);
      }
      return body;
    }

    throw new Error('OpenAI analysis is still running. Refresh status and try again in a minute.');
  }

  private async completeSynthesis(responseId: string, payload: Record<string, unknown>): Promise<SynthesisResponse> {
    const response = await fetch('/api/discovery/status', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        responseId,
        ...payload
      })
    });
    const body = await this.readApiJson(response, 'Could not persist completed synthesis.');
    if (!response.ok || body.queued || body.needsPayload) {
      throw new Error(body.error || 'Could not persist completed synthesis.');
    }
    return body;
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => window.setTimeout(resolve, ms));
  }

  async generateActionPack(): Promise<void> {
    this.packError.set('');
    this.packMessage.set('');

    if (!this.synthesisResult()) {
      const hasEvidence = this.extractedText().trim() && !this.extractedText().startsWith('Choose files or a folder');
      if (!hasEvidence) {
        this.activeView.set('imports');
        this.synthesisError.set('Stage sources first. The Discovery Action Pack is generated after OpenAI analysis produces the canonical model.');
        return;
      }

      const analyzed = await this.runSynthesis();
      if (!analyzed) {
        return;
      }
    }

    this.isGeneratingPack.set(true);
    try {
      await this.downloadDiscoveryActionPack();
      this.packMessage.set('Discovery Action Pack generated. If the browser blocked the automatic save, use the download link below.');
      this.activeView.set('reports');
    } catch (error) {
      this.packError.set(error instanceof Error ? error.message : 'Could not generate the Discovery Action Pack.');
    } finally {
      this.isGeneratingPack.set(false);
    }
  }

  async downloadArtifact(artifact: LiveArtifact): Promise<void> {
    this.packError.set('');
    this.packMessage.set('');

    if (!this.synthesisResult()) {
      this.packError.set('Run analysis first. Artifact downloads are generated from the canonical discovery model.');
      this.activeView.set('imports');
      return;
    }

    this.isGeneratingPack.set(true);
    try {
      const generated = await this.buildArtifactDownload(artifact.name);
      this.downloadBlob(generated.blob, generated.filename);
      this.packMessage.set(`${generated.filename} generated. If the browser blocked the automatic save, use the download link below.`);
    } catch (error) {
      this.packError.set(error instanceof Error ? error.message : `Could not generate ${artifact.name}.`);
    } finally {
      this.isGeneratingPack.set(false);
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
    this.packError.set('');
    this.packMessage.set('');
    this.clearDownloadUrl();
    this.isExtractingSources.set(false);
  }

  clearSources(fileInput?: HTMLInputElement, folderInput?: HTMLInputElement): void {
    this.stagedSources.set([]);
    this.extractedText.set('Choose files or a folder above. Extracted evidence will collect here before the still runs.');
    this.synthesisResult.set(null);
    this.synthesisError.set('');
    this.packError.set('');
    this.packMessage.set('');
    this.clearDownloadUrl();
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
    const items = this.records(delta.items);
    const relationships = this.records(delta.relationships);
    const artifacts = this.records(delta.artifacts);
    const backlog = this.records(delta.backlog);
    const evidenceIndex = this.records((delta as Record<string, unknown>)['evidenceIndex']);
    const failureRisks = this.records((delta as Record<string, unknown>)['failureRisks']);
    const openQuestions = this.records((delta as Record<string, unknown>)['openQuestions']);
    const systems = delta.systemsInScope?.length ? delta.systemsInScope.join(', ') : this.knownArtifacts().join(', ');
    const criticalOutputs = delta.criticalOutputs?.length ? delta.criticalOutputs.join(', ') : this.targetOutputs().join(', ');

    return [
      {
        title: 'Executive Snapshot',
        body: `${delta.processName || this.importSourceName()} was analyzed for ${delta.businessFunction || 'data discovery and migration readiness'}. Risk is ${delta.overallRiskRating || 'unscored'}. ${this.asText(delta.recommendation) || 'Review generated items, lineage, artifacts, and actions.'}`,
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Scope Coverage and Confidence',
        body: `${items.length} canonical items, ${relationships.length} relationships, ${artifacts.length} artifacts, ${backlog.length} backlog actions, and ${evidenceIndex.length} evidence records were generated from submitted evidence. Systems in scope: ${systems || 'not specified'}.`,
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Business Mission of the Process',
        body: `${delta.businessFunction || 'The submitted sources support data discovery and modernization planning.'} Critical outputs: ${criticalOutputs || 'not specified'}.`,
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Current-State Operating Model',
        body: `Current state is represented by ${items.length} nodes and ${relationships.length} edges. The strongest discovered objects are ${items.slice(0, 4).map((item) => this.readString(item, ['name'])).filter(Boolean).join(', ') || 'not yet named'}.`,
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'System and Artifact Landscape',
        body: `Artifacts in scope: ${systems || 'source files and generated canonical evidence'}. Generated outputs cover executive reporting, current-state documentation, workbook inventory, diagrams, impact model, backlog, evidence archive, and manifest.`,
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Data Flow and Process Flow Summary',
        body: relationships.length
          ? relationships.slice(0, 4).map((relationship) => `${this.readString(relationship, ['fromId', 'from_id', 'from'])} to ${this.readString(relationship, ['toId', 'to_id', 'to'])}`).join('; ')
          : 'Data flow requires additional source metadata before edges can be confirmed.',
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Transformations and Business Logic',
        body: items.filter((item) => /query|macro|vba|sql|transform|rule|report/i.test(JSON.stringify(item))).slice(0, 4).map((item) => this.readString(item, ['name'])).filter(Boolean).join(', ') || 'Transformation logic requires query, macro, VBA, formula, or SQL exports to finish.',
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Recursive Lineage and Source-of-Truth Assessment',
        body: `Lineage has ${relationships.length} discovered relationships. Branches without exported object metadata remain inferred until terminal source systems, manual entry points, or approved blockers are documented.`,
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Controls, Exceptions, and Failure Modes',
        body: failureRisks.length
          ? failureRisks.slice(0, 3).map((risk) => this.readString(risk, ['scenario', 'risk', 'name'])).join('; ')
          : 'No complete control log was supplied. Control and exception evidence should be collected for every critical output.',
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Financial Impact and Business Exposure',
        body: this.asText(delta.estimatedDollarExposure) || 'Dollar exposure needs business volume, unit value, recovery labor, SLA, penalty, customer, and compliance inputs before it can be priced.',
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Recommendations and Action Plan',
        body: backlog.length
          ? backlog.slice(0, 4).map((action) => this.readString(action, ['title', 'action', 'summary'])).join('; ')
          : this.asText(delta.recommendation) || 'Confirm evidence gaps, export native metadata, and prioritize migration actions.',
        confidence: null,
        evidenceIds: []
      },
      {
        title: 'Open Questions and Decisions Needed',
        body: openQuestions.length
          ? openQuestions.slice(0, 4).map((question) => this.readString(question, ['question', 'text', 'decision'])).join('; ')
          : delta.decisionRequired || 'Confirm ownership, unresolved evidence gaps, and migration action priority.',
        confidence: null,
        evidenceIds: []
      }
    ];
  }

  private normalizeReportSection(section: ReportSection, index: number) {
    const record = section as Record<string, unknown>;
    const rawTitle = this.readString(record, ['title', 'heading', 'sectionTitle', 'section', 'name']);
    const fallbackTitle = this.reportSectionTitles[index] || `Report Section ${index + 1}`;
    const title = !rawTitle || /^analysis section\b/i.test(rawTitle) ? fallbackTitle : rawTitle;
    const body = this.readString(record, ['body', 'narrative', 'summary', 'content', 'text']) || 'No narrative returned for this section.';
    const confidence = this.readNumber(record, ['confidence', 'confidenceScore']);
    const evidenceIds = this.readStringArray(record, ['evidenceIds', 'evidence_ids', 'evidence']);
    return {
      title,
      body,
      confidence,
      evidenceIds
    };
  }

  private packageContext(): PackageContext {
    const result = this.synthesisResult();
    const delta = result?.canonicalDelta;
    if (!result || !delta) {
      throw new Error('No canonical model is available yet. Run analysis first.');
    }

    const deltaRecord = delta as Record<string, unknown>;
    const items = this.records(delta.items);
    const relationships = this.records(delta.relationships);
    const artifacts = this.records(delta.artifacts);
    const backlog = this.records(delta.backlog);
    const evidenceIndex = this.records(deltaRecord['evidenceIndex']);
    const lineageNodes = this.records(deltaRecord['lineageNodes']).length
      ? this.records(deltaRecord['lineageNodes'])
      : items.map((item) => ({
        node_id: this.readString(item, ['id', 'item_id']),
        node_type: this.readString(item, ['type']),
        name: this.readString(item, ['name']),
        criticality: this.readString(item, ['criticality']),
        owner: this.readString(item, ['owner']),
        confidence: this.readNumber(item, ['confidence'])
      }));
    const lineageEdges = this.records(deltaRecord['lineageEdges']).length
      ? this.records(deltaRecord['lineageEdges'])
      : relationships;
    const failureRisks = this.records(deltaRecord['failureRisks']);
    const openQuestions = this.records(deltaRecord['openQuestions']);
    const sections = this.reportSections();
    const manifest = {
      packageName: 'Discovery_Action_Pack',
      generatedAt: new Date().toISOString(),
      runId: result?.runId || '',
      model: result?.model || '',
      sourceKind: this.importSourceKind(),
      sourceName: this.importSourceName(),
      storedInNeon: Boolean(result?.stored),
      counts: this.generatedCounts(),
      artifacts: this.liveArtifacts().map((artifact) => artifact.name)
    };

    return {
      result,
      delta,
      deltaRecord,
      items,
      relationships,
      artifacts,
      backlog,
      evidenceIndex,
      lineageNodes,
      lineageEdges,
      failureRisks,
      openQuestions,
      sections,
      manifest
    };
  }

  private async buildArtifactDownload(artifactName: string): Promise<{ blob: Blob; filename: string }> {
    const context = this.packageContext();
    const normalizedName = artifactName.toLowerCase();

    if (normalizedName.includes('executive')) {
      return {
        filename: '01_Executive_Decision_Brief.pdf',
        blob: this.buildExecutiveBriefPdf(context)
      };
    }

    if (normalizedName.includes('current_state') || normalizedName.includes('architecture')) {
      return {
        filename: '02_Current_State_Architecture_Report.pdf',
        blob: this.buildArchitectureReportPdf(context)
      };
    }

    if (normalizedName.includes('technical_discovery') || normalizedName.includes('workbook')) {
      return {
        filename: '03_Technical_Discovery_Workbook.xlsx',
        blob: await this.buildTechnicalWorkbook(context)
      };
    }

    if (normalizedName.includes('auto_documentation')) {
      return {
        filename: '04_Auto_Documentation_Pack.zip',
        blob: await this.buildAutoDocumentationZip(context)
      };
    }

    if (normalizedName.includes('diagram')) {
      return {
        filename: '05_Diagram_Pack.zip',
        blob: await this.buildDiagramPackZip(context)
      };
    }

    if (normalizedName.includes('financial') || normalizedName.includes('impact_model')) {
      return {
        filename: '06_Financial_Impact_Model.xlsx',
        blob: await this.buildFinancialImpactWorkbook(context)
      };
    }

    if (normalizedName.includes('action_backlog') || normalizedName.includes('backlog')) {
      return {
        filename: '07_Action_Backlog.csv',
        blob: new Blob([this.toCsv(context.backlog)], { type: 'text/csv;charset=utf-8' })
      };
    }

    if (normalizedName.includes('evidence')) {
      return {
        filename: '08_Evidence_Archive.zip',
        blob: await this.buildEvidenceArchiveZip(context)
      };
    }

    if (normalizedName.includes('manifest') || normalizedName.includes('metadata')) {
      return {
        filename: '09_Metadata_Manifest.json',
        blob: new Blob([JSON.stringify(context.manifest, null, 2)], { type: 'application/json;charset=utf-8' })
      };
    }

    return {
      filename: `${this.safeFilename(artifactName) || 'artifact'}.json`,
      blob: new Blob([JSON.stringify(context.delta, null, 2)], { type: 'application/json;charset=utf-8' })
    };
  }

  private async downloadDiscoveryActionPack(): Promise<void> {
    const context = this.packageContext();
    const { default: JSZip } = await import('jszip');
    const zip = new JSZip();
    const root = zip.folder('Discovery_Action_Pack');
    if (!root) {
      throw new Error('Could not create package root.');
    }

    root.file('01_Executive_Decision_Brief.pdf', this.buildExecutiveBriefPdf(context));
    root.file('02_Current_State_Architecture_Report.pdf', this.buildArchitectureReportPdf(context));
    root.file('03_Technical_Discovery_Workbook.xlsx', await this.buildTechnicalWorkbook(context));
    await this.addAutoDocumentationPack(root, context);
    await this.addDiagramPack(root, context);
    root.file('06_Financial_Impact_Model.xlsx', await this.buildFinancialImpactWorkbook(context));
    root.file('07_Action_Backlog.csv', this.toCsv(context.backlog));
    await this.addEvidenceArchive(root, context);
    root.file('09_Metadata_Manifest.json', JSON.stringify(context.manifest, null, 2));

    const blob = await zip.generateAsync({ type: 'blob' });
    this.downloadBlob(blob, `Discovery_Action_Pack_${this.safeFilename(this.importSourceName()) || 'run'}.zip`);
  }

  private buildExecutiveBriefPdf(context: PackageContext): Blob {
    return this.buildPdf('Executive Decision Brief', [
      `Process: ${context.delta.processName || this.importSourceName()}`,
      `Business function: ${context.delta.businessFunction || 'Data discovery and migration readiness'}`,
      `Risk rating: ${context.delta.overallRiskRating || 'not specified'}`,
      `Recommendation: ${this.asText(context.delta.recommendation) || 'Review generated action plan.'}`,
      `Decision required: ${context.delta.decisionRequired || 'Confirm owners, blockers, and migration priority.'}`,
      'Top risks',
      ...context.failureRisks.slice(0, 5).map((risk) => this.readString(risk, ['scenario', 'risk', 'name']) || this.asText(risk)),
      'Executive report sections',
      ...context.sections.slice(0, 4).flatMap((section) => [section.title, section.body])
    ]);
  }

  private buildArchitectureReportPdf(context: PackageContext): Blob {
    return this.buildPdf('Current-State Architecture Report', [
      ...context.sections.flatMap((section) => [section.title, section.body]),
      'Canonical Items',
      ...context.items.map((item) => `${this.readString(item, ['id'])} ${this.readString(item, ['name'])}: ${this.readString(item, ['businessPurpose'])}`),
      'Relationships',
      ...context.relationships.map((relationship) => `${this.readString(relationship, ['fromId', 'from_id', 'from'])} -> ${this.readString(relationship, ['toId', 'to_id', 'to'])}: ${this.readString(relationship, ['type', 'relationship_type'])}`)
    ]);
  }

  private async buildTechnicalWorkbook(context: PackageContext): Promise<Blob> {
    return this.buildWorkbook([
      { name: '00_Manifest', rows: this.objectRows(context.manifest) },
      { name: '01_Artifacts', rows: this.tableRows(context.artifacts.length ? context.artifacts : this.liveArtifacts().map((artifact) => ({ ...artifact }))) },
      { name: '03_Process_Steps', rows: this.tableRows(this.records(context.deltaRecord['processSteps']).length ? this.records(context.deltaRecord['processSteps']) : context.items) },
      { name: '07_Data_Elements', rows: this.tableRows(this.records(context.deltaRecord['dataElements']).length ? this.records(context.deltaRecord['dataElements']) : context.items) },
      { name: '08_Lineage_Nodes', rows: this.tableRows(context.lineageNodes) },
      { name: '09_Lineage_Edges', rows: this.tableRows(context.lineageEdges) },
      { name: '10_Transforms_Rules', rows: this.tableRows(context.items.filter((item) => /query|macro|transform|vba|sql|power|formula|rule/i.test(JSON.stringify(item)))) },
      { name: '11_Controls_Exceptions', rows: this.tableRows(context.failureRisks.length ? context.failureRisks : context.backlog) },
      { name: '16_Impact_Model', rows: this.objectRows(context.delta.estimatedDollarExposure || {}) },
      { name: '17_Actions', rows: this.tableRows(context.backlog) },
      { name: '18_Open_Questions', rows: this.tableRows(context.openQuestions) },
      { name: '19_Evidence_Index', rows: this.tableRows(context.evidenceIndex.length ? context.evidenceIndex : this.evidenceFromItems(context.items)) }
    ]);
  }

  private async buildFinancialImpactWorkbook(context: PackageContext): Promise<Blob> {
    return this.buildWorkbook([
      { name: 'High_Level_Context', rows: this.objectRows(context.delta.estimatedDollarExposure || {}) },
      { name: 'Failure_Scenarios', rows: this.tableRows(context.failureRisks) },
      { name: 'Pricing_Inputs_Needed', rows: [
        ['Bucket', 'Input needed'],
        ['Revenue at risk', 'Units affected and dollar per unit'],
        ['Gross margin at risk', 'Margin percent by impacted output'],
        ['Cash timing impact', 'Billing, collection, close, or reporting delay value'],
        ['Rework labor cost', 'Recovery hours, loaded labor rate, frequency'],
        ['Compliance / SLA / customer impact', 'Penalty, credit, audit, churn, or commitment exposure']
      ] },
      { name: 'Actions', rows: this.tableRows(context.backlog) }
    ]);
  }

  private async buildAutoDocumentationZip(context: PackageContext): Promise<Blob> {
    const { default: JSZip } = await import('jszip');
    const zip = new JSZip();
    await this.addAutoDocumentationPack(zip, context);
    return zip.generateAsync({ type: 'blob' });
  }

  private async buildDiagramPackZip(context: PackageContext): Promise<Blob> {
    const { default: JSZip } = await import('jszip');
    const zip = new JSZip();
    await this.addDiagramPack(zip, context);
    return zip.generateAsync({ type: 'blob' });
  }

  private async buildEvidenceArchiveZip(context: PackageContext): Promise<Blob> {
    const { default: JSZip } = await import('jszip');
    const zip = new JSZip();
    await this.addEvidenceArchive(zip, context);
    return zip.generateAsync({ type: 'blob' });
  }

  private async addAutoDocumentationPack(zip: unknown, context: PackageContext): Promise<void> {
    const folder = (zip as { folder: (name: string) => { file: (path: string, data: string | Blob) => void } | null }).folder('04_Auto_Documentation_Pack');
    folder?.file('04a_System_Inventory.csv', this.toCsv(context.items.filter((item) => /system|database|file|workbook|access/i.test(JSON.stringify(item)))));
    folder?.file('04b_Object_Inventory.csv', this.toCsv(context.items));
    folder?.file('04c_Process_Steps.csv', this.toCsv(this.records(context.deltaRecord['processSteps']).length ? this.records(context.deltaRecord['processSteps']) : context.items));
    folder?.file('04d_Lineage_Nodes.csv', this.toCsv(context.lineageNodes));
    folder?.file('04e_Lineage_Edges.csv', this.toCsv(context.lineageEdges));
    folder?.file('04f_Transformations_Rules.csv', this.toCsv(context.items.filter((item) => /query|macro|transform|vba|sql|power|formula|rule/i.test(JSON.stringify(item)))));
    folder?.file('04g_Controls_Exceptions.csv', this.toCsv(context.failureRisks.length ? context.failureRisks : context.backlog));
    folder?.file('04h_Security_Access.csv', this.toCsv(context.items.filter((item) => /security|access|credential|pii|phi|privacy/i.test(JSON.stringify(item)))));
  }

  private async addDiagramPack(zip: unknown, context: PackageContext): Promise<void> {
    const folder = (zip as { folder: (name: string) => { file: (path: string, data: string | Blob) => void } | null }).folder('05_Diagram_Pack');
    const diagramNames = [
      'D01_Executive_Value_Stream.pdf',
      'D02_System_Context_Diagram.pdf',
      'D03_Business_Process_Swimlane.pdf',
      'D04_Detailed_Data_Flow.pdf',
      'D05_Recursive_Lineage_Graph.pdf',
      'D06_Object_Dependency_Map.pdf',
      'D07_Control_And_Exception_Map.pdf',
      'D08_Failure_Impact_Map.pdf',
      'D09_Schedule_And_Refresh_Timeline.pdf'
    ];
    for (const diagramName of diagramNames) {
      folder?.file(diagramName, this.buildPdf(diagramName.replace('.pdf', '').replace(/_/g, ' '), this.diagramLines(diagramName, context.items, context.relationships, context.failureRisks)));
    }
  }

  private async addEvidenceArchive(zip: unknown, context: PackageContext): Promise<void> {
    const folder = (zip as { folder: (name: string) => { file: (path: string, data: string | Blob) => void; folder: (path: string) => { file: (path: string, data: string | Blob) => void } | null } | null }).folder('08_Evidence_Archive');
    folder?.folder('Document_Extracts')?.file('Extracted_Evidence.txt', this.extractedText());
    folder?.folder('Document_Extracts')?.file('Canonical_Model.json', JSON.stringify(context.delta, null, 2));
    folder?.file('Evidence_Index.csv', this.toCsv(context.evidenceIndex.length ? context.evidenceIndex : this.evidenceFromItems(context.items)));
  }

  private buildPdf(title: string, lines: string[]): Blob {
    const safeLines = [title, '', ...lines]
      .flatMap((line) => this.wrapText(this.stripControlChars(this.asText(line)), 92))
      .slice(0, 360);
    const pages = this.chunk(safeLines, 46);
    const objects: string[] = [];
    objects[1] = '<< /Type /Catalog /Pages 2 0 R >>';
    objects[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>';
    const kids: string[] = [];

    pages.forEach((pageLines, index) => {
      const pageObject = 4 + index * 2;
      const contentObject = pageObject + 1;
      kids.push(`${pageObject} 0 R`);
      const content = [
        'BT',
        '/F1 10 Tf',
        '14 TL',
        '50 770 Td',
        ...pageLines.map((line) => `(${this.escapePdf(line)}) Tj T*`),
        'ET'
      ].join('\n');
      objects[pageObject] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents ${contentObject} 0 R >>`;
      objects[contentObject] = `<< /Length ${content.length} >>\nstream\n${content}\nendstream`;
    });

    objects[2] = `<< /Type /Pages /Kids [${kids.join(' ')}] /Count ${pages.length} >>`;

    let pdf = '%PDF-1.4\n';
    const offsets = [0];
    for (let index = 1; index < objects.length; index += 1) {
      offsets[index] = pdf.length;
      pdf += `${index} 0 obj\n${objects[index]}\nendobj\n`;
    }
    const xrefOffset = pdf.length;
    pdf += `xref\n0 ${objects.length}\n0000000000 65535 f \n`;
    for (let index = 1; index < objects.length; index += 1) {
      pdf += `${String(offsets[index]).padStart(10, '0')} 00000 n \n`;
    }
    pdf += `trailer\n<< /Size ${objects.length} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
    return new Blob([pdf], { type: 'application/pdf' });
  }

  private async buildWorkbook(sheets: Array<{ name: string; rows: string[][] }>): Promise<Blob> {
    const { default: JSZip } = await import('jszip');
    const zip = new JSZip();
    const sheetList = sheets.map((sheet, index) => ({ ...sheet, id: index + 1, safeName: sheet.name.slice(0, 31) || `Sheet${index + 1}` }));

    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
${sheetList.map((sheet) => `<Override PartName="/xl/worksheets/sheet${sheet.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}
</Types>`);
    zip.folder('_rels')?.file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
    zip.folder('xl')?.file('workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${sheetList.map((sheet) => `<sheet name="${this.escapeXml(sheet.safeName)}" sheetId="${sheet.id}" r:id="rId${sheet.id}"/>`).join('')}</sheets>
</workbook>`);
    zip.folder('xl')?.folder('_rels')?.file('workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheetList.map((sheet) => `<Relationship Id="rId${sheet.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${sheet.id}.xml"/>`).join('')}
</Relationships>`);
    const worksheets = zip.folder('xl')?.folder('worksheets');
    sheetList.forEach((sheet) => {
      worksheets?.file(`sheet${sheet.id}.xml`, this.sheetXml(sheet.rows));
    });

    return zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  private sheetXml(rows: string[][]): string {
    const body = rows.map((row, rowIndex) => {
      const cells = row.map((cell, columnIndex) => {
        const ref = `${this.columnName(columnIndex)}${rowIndex + 1}`;
        return `<c r="${ref}" t="inlineStr"><is><t>${this.escapeXml(cell)}</t></is></c>`;
      }).join('');
      return `<row r="${rowIndex + 1}">${cells}</row>`;
    }).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;
  }

  private objectRows(value: unknown): string[][] {
    const record = this.asRecord(value);
    const rows = [['Field', 'Value']];
    Object.keys(record).forEach((key) => rows.push([key, this.asText(record[key])]));
    return rows;
  }

  private tableRows(records: Record<string, unknown>[]): string[][] {
    const columns = this.columnsFor(records);
    return [columns, ...records.map((record) => columns.map((column) => this.asText(record[column])))];
  }

  private toCsv(records: Record<string, unknown>[]): string {
    const rows = this.tableRows(records);
    return rows.map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\r\n');
  }

  private columnsFor(records: Record<string, unknown>[]): string[] {
    const columns = new Set<string>();
    records.forEach((record) => Object.keys(record).forEach((key) => columns.add(key)));
    return [...columns].length ? [...columns] : ['status'];
  }

  private records(value: unknown): Record<string, unknown>[] {
    return Array.isArray(value)
      ? value.filter((item): item is Record<string, unknown> => Boolean(item) && typeof item === 'object').map((item) => item as Record<string, unknown>)
      : [];
  }

  private evidenceFromItems(items: Record<string, unknown>[]): Record<string, unknown>[] {
    return items.flatMap((item) => {
      const evidence = item['evidence'];
      if (!Array.isArray(evidence)) {
        return [];
      }
      return evidence.filter((entry): entry is Record<string, unknown> => Boolean(entry) && typeof entry === 'object');
    });
  }

  private diagramLines(name: string, items: Record<string, unknown>[], relationships: Record<string, unknown>[], risks: Record<string, unknown>[]): string[] {
    return [
      'Generated from canonical discovery model.',
      `Diagram: ${name.replace('.pdf', '').replace(/_/g, ' ')}`,
      '',
      'Key nodes:',
      ...items.slice(0, 12).map((item) => `${this.readString(item, ['id'])}: ${this.readString(item, ['name'])} (${this.readString(item, ['type'])})`),
      '',
      'Key relationships:',
      ...relationships.slice(0, 12).map((relationship) => `${this.readString(relationship, ['fromId', 'from_id', 'from'])} -> ${this.readString(relationship, ['toId', 'to_id', 'to'])} ${this.readString(relationship, ['type'])}`),
      '',
      'Failure / control notes:',
      ...risks.slice(0, 8).map((risk) => `${this.readString(risk, ['id'])}: ${this.readString(risk, ['scenario', 'risk', 'name'])}`)
    ];
  }

  private readString(record: Record<string, unknown>, keys: string[]): string {
    for (const key of keys) {
      const value = record[key];
      if (typeof value === 'string' && value.trim()) {
        return value.trim();
      }
      if (typeof value === 'number') {
        return String(value);
      }
    }
    return '';
  }

  private readNumber(record: Record<string, unknown>, keys: string[]): number | null {
    for (const key of keys) {
      const value = record[key];
      if (typeof value === 'number' && Number.isFinite(value)) {
        return value;
      }
      if (typeof value === 'string' && Number.isFinite(Number(value))) {
        return Number(value);
      }
    }
    return null;
  }

  private readStringArray(record: Record<string, unknown>, keys: string[]): string[] {
    for (const key of keys) {
      const value = record[key];
      if (Array.isArray(value)) {
        return value
          .map((item) => {
            if (typeof item === 'string') {
              return item;
            }
            if (item && typeof item === 'object') {
              return this.readString(item as Record<string, unknown>, ['id', 'evidenceId', 'evidence_id', 'name']);
            }
            return '';
          })
          .filter(Boolean);
      }
      if (typeof value === 'string' && value.trim()) {
        return value.split(/\s*,\s*/).filter(Boolean);
      }
    }
    return [];
  }

  private asRecord(value: unknown): Record<string, unknown> {
    return value && typeof value === 'object' ? value as Record<string, unknown> : {};
  }

  private asText(value: unknown): string {
    if (value === null || value === undefined) {
      return '';
    }
    if (typeof value === 'string') {
      return value;
    }
    if (typeof value === 'number' || typeof value === 'boolean') {
      return String(value);
    }
    return JSON.stringify(value);
  }

  private wrapText(value: string, width: number): string[] {
    const words = value.split(/\s+/).filter(Boolean);
    const lines: string[] = [];
    let current = '';
    for (const word of words) {
      if (`${current} ${word}`.trim().length > width) {
        lines.push(current);
        current = word;
      } else {
        current = `${current} ${word}`.trim();
      }
    }
    if (current) {
      lines.push(current);
    }
    return lines.length ? lines : [''];
  }

  private chunk<T>(values: T[], size: number): T[][] {
    const chunks: T[][] = [];
    for (let index = 0; index < values.length; index += size) {
      chunks.push(values.slice(index, index + size));
    }
    return chunks.length ? chunks : [[]];
  }

  private escapePdf(value: string): string {
    return value.replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)');
  }

  private escapeXml(value: string): string {
    return this.stripControlChars(String(value))
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  private stripControlChars(value: string): string {
    return value.replace(/[^\x20-\x7E]/g, ' ');
  }

  private columnName(index: number): string {
    let name = '';
    let current = index + 1;
    while (current > 0) {
      const remainder = (current - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      current = Math.floor((current - 1) / 26);
    }
    return name;
  }

  private safeFilename(value: string): string {
    return value.replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 80);
  }

  private downloadBlob(blob: Blob, filename: string): void {
    this.clearDownloadUrl();
    const url = URL.createObjectURL(blob);
    this.activeDownloadUrl = url;
    this.lastDownload.set({ url, filename });

    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    anchor.rel = 'noopener';
    anchor.style.display = 'none';
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
  }

  private clearDownloadUrl(): void {
    if (this.activeDownloadUrl) {
      URL.revokeObjectURL(this.activeDownloadUrl);
      this.activeDownloadUrl = '';
    }
    this.lastDownload.set(null);
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
