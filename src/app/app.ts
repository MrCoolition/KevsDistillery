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
  distillery?: {
    ready: boolean;
    label?: string;
  };
  workspace?: {
    configured: boolean;
    ready: boolean;
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
  engine?: string;
  model?: string;
  fallbackReason?: string | null;
  outputText?: string;
  canonicalDelta?: {
    processName?: string;
    businessFunction?: string;
    recommendation?: string;
    decisionRequired?: string;
    systemsInScope?: unknown[];
    criticalOutputs?: unknown[];
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

interface PdfBlock {
  type: 'section' | 'paragraph' | 'keyValues' | 'bullets' | 'note' | 'table';
  title?: string;
  text?: string;
  rows?: string[][];
  items?: string[];
}

type DiagramKind = 'valueStream' | 'context' | 'swimlane' | 'dataFlow' | 'lineage' | 'dependency' | 'control' | 'failure' | 'timeline';
type DiagramTone = 'source' | 'process' | 'graph' | 'output' | 'control' | 'risk' | 'action' | 'neutral';

interface DiagramSpec {
  id: string;
  fileBase: string;
  title: string;
  subtitle: string;
  kind: DiagramKind;
}

interface DiagramNode {
  id: string;
  label: string;
  sublabel: string;
  type: string;
  x: number;
  y: number;
  w: number;
  h: number;
  tone: DiagramTone;
  confidence: number | null;
}

interface DiagramEdge {
  id: string;
  from: string;
  to: string;
  label: string;
  tone: DiagramTone;
}

interface DiagramLane {
  label: string;
  x: number;
  y: number;
  w: number;
  h: number;
}

interface DiagramCallout {
  title: string;
  text: string;
  x: number;
  y: number;
  w: number;
  h: number;
  tone: DiagramTone;
}

interface DiagramScene {
  spec: DiagramSpec;
  width: number;
  height: number;
  nodes: DiagramNode[];
  edges: DiagramEdge[];
  lanes: DiagramLane[];
  callouts: DiagramCallout[];
}

interface InlineGraphNode {
  id: string;
  label: string;
  type: string;
  tone: DiagramTone;
  confidence: number | null;
  x: number;
  y: number;
}

interface InlineGraphEdge {
  id: string;
  label: string;
  tone: DiagramTone;
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  labelX: number;
  labelY: number;
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
  readonly analysisConfirmOpen = signal(false);
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
  private analysisConfirmResolver: ((confirmed: boolean) => void) | null = null;

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

  readonly reportPreviewNodes = computed<InlineGraphNode[]>(() => {
    const delta = this.synthesisResult()?.canonicalDelta;
    const items = delta?.items?.length
      ? this.records(delta.items)
      : this.model.items.map((item) => ({ ...item }) as unknown as Record<string, unknown>);
    const limited = items.slice(0, 10);
    const columns = limited.length <= 6 ? 3 : 4;
    const xStep = 100 / columns;
    const rows = Math.max(1, Math.ceil(limited.length / columns));
    const yStep = rows === 1 ? 0 : 56 / (rows - 1);

    return limited.map((item, index) => {
      const type = this.readString(item, ['type', 'node_type']) || 'object';
      const column = index % columns;
      const row = Math.floor(index / columns);
      return {
        id: this.readString(item, ['id', 'item_id', 'node_id']) || `N${index + 1}`,
        label: this.readString(item, ['name', 'object_name', 'title']) || `Node ${index + 1}`,
        type,
        tone: this.toneForType(type),
        confidence: this.readNumber(item, ['confidence']),
        x: 6 + column * xStep,
        y: rows === 1 ? 38 : 18 + row * yStep
      };
    });
  });

  readonly reportPreviewEdges = computed<InlineGraphEdge[]>(() => {
    const nodes = this.reportPreviewNodes();
    const nodeById = new Map(nodes.map((node) => [node.id, node]));
    const delta = this.synthesisResult()?.canonicalDelta;
    const relationships = delta?.relationships?.length
      ? this.records(delta.relationships)
      : this.model.relationships.map((relationship) => ({ ...relationship }) as unknown as Record<string, unknown>);
    const explicitEdges = relationships
      .map((relationship, index) => {
        const from = this.readString(relationship, ['fromId', 'from_id', 'from']);
        const to = this.readString(relationship, ['toId', 'to_id', 'to']);
        const fromNode = nodeById.get(from);
        const toNode = nodeById.get(to);
        if (!fromNode || !toNode) {
          return null;
        }
        return this.inlineGraphEdge(
          this.readString(relationship, ['id', 'relationship_id']) || `REL-${index + 1}`,
          fromNode,
          toNode,
          this.readString(relationship, ['type', 'edgeType', 'relationship_type']) || 'depends on',
          'graph'
        );
      })
      .filter((edge): edge is InlineGraphEdge => Boolean(edge));

    if (explicitEdges.length) {
      return explicitEdges.slice(0, 12);
    }

    return nodes.slice(0, -1).map((node, index) => {
      return this.inlineGraphEdge(`AUTO-${index + 1}`, node, nodes[index + 1], 'feeds', index < 2 ? 'source' : 'graph');
    });
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
      return this.analysisStatus() || 'The Distillery is analyzing the staged evidence now.';
    }

    const result = this.synthesisResult();
    if (result) {
      const counts = result.counts;
      const countText = counts
        ? `${counts.items} items, ${counts.relationships} relationships, ${counts.artifacts} artifacts, ${counts.backlog} actions`
        : 'canonical output returned';
      const persistence = result.stored
        ? 'Run saved to the workspace.'
        : `Workspace save needs attention: ${result.persistenceError || 'storage write failed.'}`;
      const fallback = result.fallbackReason ? ` ${result.fallbackReason}.` : '';
      const analysis = result.engine === 'The Distillery' || result.model
        ? 'The Distillery analyzed the staged evidence'
        : 'The Distillery generated a blocker-backed action pack';
      return `${analysis}: ${countText}.${fallback} ${persistence}`;
    }

    if (!this.stagedSources().length) {
      return 'Ready. Choose files or a folder to stage a source batch.';
    }

    const stats = this.sourceStats();
    if (this.limitedExtraction()) {
      return `${stats.files} source staged (${stats.bytes}). Browser extraction found ${stats.evidenceChars} evidence characters. Some files still need native exports or connector metadata for full object-level lineage, but the retrieved evidence is ready for The Distillery.`;
    }

    return `${stats.files} source staged (${stats.bytes}) with ${stats.readable} readable files and ${stats.evidenceChars} evidence characters. Ready to run The Distillery.`;
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

    if (this.hasEvidenceForTransmission()) {
      const confirmed = await this.confirmAnalysisTransmission();
      if (!confirmed) {
        return false;
      }
    }

    this.isSynthesizing.set(true);
    this.synthesisError.set('');
    this.packError.set('');
    this.packMessage.set('');
    this.clearDownloadUrl();
    this.analysisStatus.set('The Distillery run started.');
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

  approveAnalysisRun(): void {
    this.resolveAnalysisConfirmation(true);
  }

  cancelAnalysisRun(): void {
    this.resolveAnalysisConfirmation(false);
  }

  private hasEvidenceForTransmission(): boolean {
    const text = this.extractedText().trim();
    return Boolean(text) && !text.startsWith('Choose files or a folder');
  }

  private confirmAnalysisTransmission(): Promise<boolean> {
    if (this.analysisConfirmResolver) {
      return Promise.resolve(false);
    }

    this.analysisConfirmOpen.set(true);
    return new Promise((resolve) => {
      this.analysisConfirmResolver = resolve;
    });
  }

  private resolveAnalysisConfirmation(confirmed: boolean): void {
    const resolver = this.analysisConfirmResolver;
    this.analysisConfirmResolver = null;
    this.analysisConfirmOpen.set(false);
    resolver?.(confirmed);
  }

  private async pollSynthesis(responseId: string, payload: Record<string, unknown>): Promise<SynthesisResponse> {
    const maxAttempts = 120;
    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      this.analysisStatus.set(`The Distillery is ${attempt === 1 ? 'starting' : 'still running'} (${attempt}/${maxAttempts}).`);
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
        this.analysisStatus.set(`Distillery status: ${body.responseStatus || 'in progress'}. Uncle Kev is still distilling.`);
        continue;
      }
      if (!response.ok) {
        throw new Error(body.error || 'Could not check synthesis status.');
      }
      if (body.needsPayload) {
        this.analysisStatus.set('The Distillery finished. Saving the canonical model to the workspace.');
        return this.completeSynthesis(responseId, payload);
      }
      return body;
    }

    throw new Error('The Distillery is still running. Refresh status and try again in a minute.');
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
        this.synthesisError.set('Stage sources first. The Discovery Action Pack is generated after The Distillery produces the canonical model.');
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

  async verifyWorkspace(): Promise<void> {
    try {
      this.apiError.set('');
      const response = await fetch('/api/admin/migrate', {
        method: 'POST'
      });
      const body = await this.readApiJson(response, 'Workspace verification failed.');
      if (!response.ok) {
        throw new Error(body.error || 'Workspace verification failed.');
      }
      await this.checkApiHealth();
    } catch (error) {
      this.apiError.set(error instanceof Error ? error.message : 'Workspace verification failed.');
    }
  }

  distilleryReady(health: ApiHealth | null | undefined): boolean {
    return Boolean(health?.distillery?.ready);
  }

  workspaceReady(health: ApiHealth | null | undefined): boolean {
    return Boolean(health?.workspace?.ready);
  }

  workspaceStatusText(health: ApiHealth): string {
    const workspace = health.workspace;
    if (!workspace?.configured) {
      return 'not configured';
    }

    if (workspace.ready) {
      const tableText = workspace.tableCount && workspace.requiredTableCount
        ? ` (${workspace.tableCount}/${workspace.requiredTableCount} tables)`
        : '';
      return `ready${tableText}`;
    }

    if (workspace.error) {
      return `needs attention: ${workspace.error}`;
    }

    const missing = workspace.missingTables?.length
      ? `missing ${workspace.missingTables.length} workspace tables`
      : 'not ready';
    return missing;
  }

  sectionBodyLines(body: string): string[] {
    return body
      .split(/\n+/)
      .map((line) => line.trim())
      .filter(Boolean);
  }

  shouldVisualizeSection(title: string): boolean {
    return /data flow|process flow|lineage|source-of-truth|system and artifact/i.test(title);
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
      'Interpretation note: these recovered strings are source evidence hints, not a full Access catalog. The Distillery must analyze them, document confidence, and call out exact exports needed to finish object-level lineage.',
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
    const deltaRecord = delta as Record<string, unknown>;
    const systems = this.listText(deltaRecord['systemsInScope'], this.knownArtifacts().join(', '));
    const criticalOutputs = this.listText(deltaRecord['criticalOutputs'], this.targetOutputs().join(', '));
    const itemLabels = this.nodeLabelMap(items);

    return [
      {
        title: 'Executive Snapshot',
        body: this.cleanDisplayText(`${this.labelFor(delta.processName) || this.importSourceName()} was analyzed for ${this.labelFor(delta.businessFunction) || 'data discovery and migration readiness'}. Risk is ${this.labelFor(delta.overallRiskRating) || 'unscored'}. ${this.asText(delta.recommendation) || 'Review generated items, lineage, artifacts, and actions.'}`),
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
        body: `${this.labelFor(delta.businessFunction) || 'The submitted sources support data discovery and modernization planning.'} Critical outputs: ${criticalOutputs || 'not specified'}.`,
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
          ? relationships.slice(0, 5).map((relationship) => this.relationshipDisplay(relationship, itemLabels)).join('\n')
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
        body: this.exposureSummary(delta.estimatedDollarExposure) || 'Dollar exposure needs business volume, unit value, recovery labor, SLA, penalty, customer, and compliance inputs before it can be priced.',
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
    const body = this.cleanDisplayText(this.readString(record, ['body', 'narrative', 'summary', 'content', 'text']) || 'No narrative returned for this section.');
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
      engine: 'The Distillery',
      sourceKind: this.importSourceKind(),
      sourceName: this.importSourceName(),
      savedToWorkspace: Boolean(result?.stored),
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
    return this.buildScientificPdf(
      'Executive Decision Brief',
      `${this.labelFor(context.delta.processName) || this.importSourceName()} | Discovery Action Pack`,
      [
        {
          type: 'keyValues',
          title: 'Executive Snapshot',
          rows: [
            ['Process', this.labelFor(context.delta.processName) || this.importSourceName()],
            ['Business function', this.labelFor(context.delta.businessFunction) || 'Data discovery and migration readiness'],
            ['Overall risk rating', this.labelFor(context.delta.overallRiskRating) || 'Not specified'],
            ['Decision required', this.asText(context.delta.decisionRequired) || 'Confirm owners, blockers, and migration priority.'],
            ['Package confidence', `${this.averageConfidence()}% starter confidence; run-specific confidence appears by section.`]
          ]
        },
        {
          type: 'section',
          title: 'Recommendation',
          text: this.asText(context.delta.recommendation) || 'Review generated action plan.'
        },
        {
          type: 'bullets',
          title: 'Top Failure Risks',
          items: context.failureRisks.length
            ? context.failureRisks.slice(0, 6).map((risk) => this.readString(risk, ['scenario', 'risk', 'name']) || this.asText(risk))
            : ['No complete failure register was supplied. Collect control, exception, and recovery evidence for every critical output.']
        },
        {
          type: 'section',
          title: 'Financial Exposure Context',
          text: this.exposureSummary(context.delta.estimatedDollarExposure) || 'Dollar exposure needs business volume, unit value, recovery labor, SLA, penalty, customer, and compliance inputs before it can be priced.'
        },
        {
          type: 'table',
          title: 'Evidence-Backed Report Sections',
          rows: [
            ['Section', 'Evidence', 'Confidence'],
            ...context.sections.slice(0, 8).map((section) => [
              section.title,
              section.evidenceIds.length ? section.evidenceIds.join(', ') : 'Evidence pending',
              section.confidence === null ? 'Not scored' : `${section.confidence}%`
            ])
          ]
        }
      ]
    );
  }

  private buildArchitectureReportPdf(context: PackageContext): Blob {
    return this.buildScientificPdf(
      'Current-State Architecture Report',
      'Evidence-backed operating model, object landscape, lineage, controls, and migration actions',
      [
        ...context.sections.map((section) => ({
          type: 'section' as const,
          title: section.title,
          text: section.body
        })),
        {
          type: 'table',
          title: 'Canonical Items',
          rows: [
            ['ID', 'Type', 'Name', 'Purpose'],
            ...context.items.slice(0, 18).map((item) => [
              this.readString(item, ['id', 'item_id']),
              this.readString(item, ['type']),
              this.readString(item, ['name']),
              this.readString(item, ['businessPurpose', 'business_purpose', 'purpose'])
            ])
          ]
        },
        {
          type: 'table',
          title: 'Lineage Relationships',
          rows: [
            ['From', 'To', 'Edge', 'Confidence'],
            ...context.relationships.slice(0, 18).map((relationship) => [
              this.readString(relationship, ['fromId', 'from_id', 'from']),
              this.readString(relationship, ['toId', 'to_id', 'to']),
              this.readString(relationship, ['type', 'relationship_type', 'edgeType']),
              this.readString(relationship, ['confidence'])
            ])
          ]
        },
        {
          type: 'note',
          title: 'Quality Gate',
          text: 'A finding is not finished unless it carries evidence, confidence, owner, failure impact, dollar context, and a recommended next action.'
        }
      ]
    );
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
    const folder = (zip as { folder: (name: string) => any }).folder('05_Diagram_Pack');
    if (!folder) {
      return;
    }

    const svgFolder = folder.folder('SVG');
    const mermaidFolder = folder.folder('Mermaid_Source');
    const indexRows = [['Diagram', 'PDF', 'SVG', 'Mermaid source', 'Purpose']];
    for (const spec of this.diagramSpecs()) {
      const scene = this.buildDiagramScene(spec, context);
      folder.file(`${spec.fileBase}.pdf`, this.buildVisualDiagramPdf(scene));
      svgFolder?.file(`${spec.fileBase}.svg`, this.buildDiagramSvg(scene));
      mermaidFolder?.file(`${spec.fileBase}.mmd`, this.buildDiagramMermaid(scene));
      indexRows.push([
        spec.title,
        `${spec.fileBase}.pdf`,
        `SVG/${spec.fileBase}.svg`,
        `Mermaid_Source/${spec.fileBase}.mmd`,
        spec.subtitle
      ]);
    }
    folder.file('Diagram_Index.csv', indexRows.map((row) => row.map((cell) => `"${cell.replace(/"/g, '""')}"`).join(',')).join('\r\n'));
  }

  private diagramSpecs(): DiagramSpec[] {
    return [
      { id: 'D01', fileBase: 'D01_Executive_Value_Stream', title: 'D01 Executive Value Stream', subtitle: 'Source evidence to migration action, with business value and decision points.', kind: 'valueStream' },
      { id: 'D02', fileBase: 'D02_System_Context_Diagram', title: 'D02 System Context Diagram', subtitle: 'Systems, files, people, and outputs surrounding the canonical discovery graph.', kind: 'context' },
      { id: 'D03', fileBase: 'D03_Business_Process_Swimlane', title: 'D03 Business Process Swimlane', subtitle: 'End-to-end operating model by role, source, Distillery step, and engineering handoff.', kind: 'swimlane' },
      { id: 'D04', fileBase: 'D04_Detailed_Data_Flow', title: 'D04 Detailed Data Flow', subtitle: 'Data movements, transformations, storage points, and downstream outputs.', kind: 'dataFlow' },
      { id: 'D05', fileBase: 'D05_Recursive_Lineage_Graph', title: 'D05 Recursive Lineage Graph', subtitle: 'Upstream chain from critical outputs to source candidates and documented blockers.', kind: 'lineage' },
      { id: 'D06', fileBase: 'D06_Object_Dependency_Map', title: 'D06 Object Dependency Map', subtitle: 'Object-level dependency map for sources, queries, logic, reports, and outputs.', kind: 'dependency' },
      { id: 'D07', fileBase: 'D07_Control_And_Exception_Map', title: 'D07 Control And Exception Map', subtitle: 'Controls, exceptions, failure conditions, mitigations, and evidence confidence.', kind: 'control' },
      { id: 'D08', fileBase: 'D08_Failure_Impact_Map', title: 'D08 Failure Impact Map', subtitle: 'Failure scenarios mapped to impacted outputs, action owners, and exposure context.', kind: 'failure' },
      { id: 'D09', fileBase: 'D09_Schedule_And_Refresh_Timeline', title: 'D09 Schedule And Refresh Timeline', subtitle: 'Cadence, cutoffs, sequencing, dependencies, and timing risk.', kind: 'timeline' }
    ];
  }

  private buildDiagramScene(spec: DiagramSpec, context: PackageContext): DiagramScene {
    const width = 1200;
    const height = 760;
    const sourceName = this.labelFor(context.delta.processName) || this.importSourceName();
    const objects = this.diagramObjectNodes(context);
    const relationships = this.diagramRelationshipEdges(context, objects);
    const risks: Record<string, unknown>[] = context.failureRisks.length
      ? context.failureRisks
      : this.model.failureRisks.map((risk) => ({ ...risk }) as unknown as Record<string, unknown>);
    const actions: Record<string, unknown>[] = context.backlog.length
      ? context.backlog
      : this.model.backlog.map((action) => ({ ...action }) as unknown as Record<string, unknown>);

    if (spec.kind === 'valueStream') {
      const nodes = [
        this.diagramNode('SRC', 'Raw source set', sourceName, 'source', 70, 286, 150, 96, 'source', 85),
        this.diagramNode('EXT', 'Evidence extraction', 'files, docs, SQL, workbook XML', 'process', 270, 286, 166, 96, 'process', 82),
        this.diagramNode('GRAPH', 'Canonical proof graph', 'nodes, edges, evidence, confidence', 'graph', 488, 260, 188, 122, 'graph', 88),
        this.diagramNode('DOCS', 'Auto documentation', 'current-state inventory and narrative', 'output', 730, 205, 174, 96, 'output', 84),
        this.diagramNode('DIAGS', 'Diagram pack', 'value, context, flow, lineage, impact', 'output', 730, 370, 174, 96, 'output', 84),
        this.diagramNode('ACTIONS', 'Engineering backlog', 'Fivetran, dbt, SQL, Snowpark work', 'action', 965, 286, 170, 96, 'action', 82)
      ];
      return {
        spec,
        width,
        height,
        nodes,
        edges: [
          this.diagramEdge('E1', 'SRC', 'EXT', 'extracts', 'source'),
          this.diagramEdge('E2', 'EXT', 'GRAPH', 'normalizes', 'process'),
          this.diagramEdge('E3', 'GRAPH', 'DOCS', 'documents', 'output'),
          this.diagramEdge('E4', 'GRAPH', 'DIAGS', 'renders', 'output'),
          this.diagramEdge('E5', 'DOCS', 'ACTIONS', 'drives', 'action'),
          this.diagramEdge('E6', 'DIAGS', 'ACTIONS', 'prioritizes', 'action')
        ],
        lanes: [],
        callouts: [
          this.diagramCallout('Done standard', 'Every node maps back to evidence, confidence, owner, impact, and next action.', 450, 560, 300, 96, 'graph'),
          this.diagramCallout('Business exposure', this.exposureSummary(context.delta.estimatedDollarExposure) || 'Exposure model generated from process failure, lateness, wrong data, partial run, and auditability scenarios.', 805, 560, 320, 96, 'risk')
        ]
      };
    }

    if (spec.kind === 'context') {
      const center = this.diagramNode('GRAPH', 'Canonical discovery graph', 'single model driving all artifacts', 'graph', 494, 302, 212, 118, 'graph', 90);
      const around = objects.slice(0, 10).map((node, index) => {
        const angle = (-120 + index * (240 / Math.max(objects.slice(0, 10).length - 1, 1))) * Math.PI / 180;
        const radiusX = 405;
        const radiusY = 242;
        return {
          ...node,
          x: 600 + Math.cos(angle) * radiusX - 82,
          y: 360 + Math.sin(angle) * radiusY - 42,
          w: 164,
          h: 84,
          tone: index % 3 === 0 ? 'source' as DiagramTone : node.tone
        };
      });
      return {
        spec,
        width,
        height,
        nodes: [center, ...around],
        edges: around.map((node, index) => this.diagramEdge(`CTX-${index}`, node.id, 'GRAPH', node.type || 'feeds', node.tone)),
        lanes: [],
        callouts: [
          this.diagramCallout('Scope', `${objects.length} discovered objects and ${relationships.length} discovered relationships in the current run.`, 60, 595, 265, 78, 'neutral'),
          this.diagramCallout('Outputs', `${this.liveArtifacts().length} action-pack artifacts generated from the graph.`, 875, 595, 265, 78, 'output')
        ]
      };
    }

    if (spec.kind === 'swimlane') {
      const lanes = [
        this.diagramLane('Business / Operations', 58, 150, 1084, 116),
        this.diagramLane('Source Systems & Files', 58, 276, 1084, 116),
        this.diagramLane("Uncle Kev's Distillery", 58, 402, 1084, 116),
        this.diagramLane('Engineering Delivery', 58, 528, 1084, 116)
      ];
      const nodes = [
        this.diagramNode('TRIGGER', 'Trigger / SLA window', 'business asks for a trusted output', 'process', 110, 178, 150, 68, 'process', null),
        this.diagramNode('SELECT', 'Select sources', 'files, DBs, workbooks, docs', 'source', 310, 304, 150, 68, 'source', null),
        this.diagramNode('RUN', 'Run Distillery', 'extract, infer, score, synthesize', 'graph', 510, 430, 160, 68, 'graph', null),
        this.diagramNode('REVIEW', 'Review blockers', 'evidence gaps and confidence', 'control', 710, 430, 160, 68, 'control', null),
        this.diagramNode('PACK', 'Action pack', 'reports, diagrams, workbook, backlog', 'output', 910, 430, 160, 68, 'output', null),
        this.diagramNode('BUILD', 'Modernize', 'ingest, model, test, govern', 'action', 910, 556, 160, 68, 'action', null)
      ];
      return {
        spec,
        width,
        height,
        nodes,
        edges: [
          this.diagramEdge('S1', 'TRIGGER', 'SELECT', 'request', 'process'),
          this.diagramEdge('S2', 'SELECT', 'RUN', 'stage', 'source'),
          this.diagramEdge('S3', 'RUN', 'REVIEW', 'findings', 'graph'),
          this.diagramEdge('S4', 'REVIEW', 'PACK', 'approve', 'control'),
          this.diagramEdge('S5', 'PACK', 'BUILD', 'deliver', 'action')
        ],
        lanes,
        callouts: []
      };
    }

    if (spec.kind === 'control') {
      const riskNodes = risks.slice(0, 4).map((risk, index) => this.diagramNode(
        `RISK-${index + 1}`,
        this.readString(risk, ['scenario', 'risk', 'name']) || `Risk ${index + 1}`,
        this.readString(risk, ['detection', 'failureCondition', 'failure_condition']) || 'detection method pending',
        'failure mode',
        80,
        168 + index * 118,
        230,
        84,
        'risk',
        this.readNumber(risk, ['confidence'])
      ));
      const controlNodes = riskNodes.map((node, index) => this.diagramNode(
        `CTRL-${index + 1}`,
        index % 2 === 0 ? 'Preventive / detective control' : 'Exception path',
        'owner, mitigation, evidence required',
        'control',
        480,
        node.y,
        220,
        84,
        'control',
        null
      ));
      const actionNodes = controlNodes.map((node, index) => this.diagramNode(
        `ACT-${index + 1}`,
        this.readString(actions[index] || {}, ['title', 'action', 'summary']) || 'Remediation action',
        this.readString(actions[index] || {}, ['owner']) || 'owner pending',
        'action',
        880,
        node.y,
        220,
        84,
        'action',
        null
      ));
      return {
        spec,
        width,
        height,
        nodes: [...riskNodes, ...controlNodes, ...actionNodes],
        edges: riskNodes.flatMap((node, index) => [
          this.diagramEdge(`C${index}-A`, node.id, controlNodes[index].id, 'detects / mitigates', 'control'),
          this.diagramEdge(`C${index}-B`, controlNodes[index].id, actionNodes[index].id, 'remediates', 'action')
        ]),
        lanes: [],
        callouts: [
          this.diagramCallout('Control rule', 'Every material failure mode must have detection, recovery, owner, evidence, and acceptance criteria.', 330, 625, 540, 64, 'control')
        ]
      };
    }

    if (spec.kind === 'failure') {
      const scenarios = risks.slice(0, 5).map((risk, index) => this.diagramNode(
        `FAIL-${index + 1}`,
        this.readString(risk, ['scenario', 'risk', 'name']) || `Failure ${index + 1}`,
        this.readString(risk, ['effect', 'impactedOutput', 'impacted_output']) || 'impact pending',
        'failure',
        70,
        140 + index * 100,
        250,
        72,
        'risk',
        this.readNumber(risk, ['confidence'])
      ));
      const output = this.diagramNode('OUT', 'Critical outputs', this.listText(context.deltaRecord['criticalOutputs'], this.targetOutputs().join(', ')) || 'outputs pending', 'output', 480, 300, 240, 100, 'output', null);
      const action = this.diagramNode('ACTION', 'Action backlog', `${context.backlog.length || this.model.backlog.length} prioritized actions`, 'action', 890, 300, 230, 100, 'action', null);
      return {
        spec,
        width,
        height,
        nodes: [...scenarios, output, action],
        edges: [
          ...scenarios.map((node, index) => this.diagramEdge(`F${index}`, node.id, 'OUT', 'breaks / delays', 'risk')),
          this.diagramEdge('F-ACTION', 'OUT', 'ACTION', 'drives', 'action')
        ],
        lanes: [],
        callouts: [
          this.diagramCallout('Dollar lens', this.exposureSummary(context.delta.estimatedDollarExposure) || 'Revenue, margin, cash timing, rework, SLA, compliance, and customer exposure are priced when source evidence supports it.', 430, 560, 520, 86, 'risk')
        ]
      };
    }

    if (spec.kind === 'timeline') {
      const steps = [
        ['T1', 'Source selected', 'intake'],
        ['T2', 'Readable evidence extracted', 'local parse'],
        ['T3', 'Distillery run', 'synthesis'],
        ['T4', 'Graph saved', 'canonical model'],
        ['T5', 'Artifacts generated', 'action pack'],
        ['T6', 'Engineering actioned', 'delivery']
      ];
      const nodes = steps.map(([id, label, sublabel], index) => this.diagramNode(id, label, sublabel, 'timeline', 90 + index * 180, 336 + (index % 2 === 0 ? -88 : 88), 142, 66, index < 2 ? 'source' : index < 4 ? 'graph' : 'action', null));
      return {
        spec,
        width,
        height,
        nodes,
        edges: steps.slice(0, -1).map((step, index) => this.diagramEdge(`T${index}`, step[0], steps[index + 1][0], 'next', 'process')),
        lanes: [this.diagramLane('Run sequence and refresh dependency path', 80, 355, 1040, 52)],
        callouts: [
          this.diagramCallout('Timing risk', 'Late, wrong, partial, failed, or unauditable runs become priced failure scenarios in the impact model.', 330, 590, 540, 68, 'risk')
        ]
      };
    }

    const placed = this.placeObjectNodes(objects, spec.kind);
    const fallbackEdges = relationships.length
      ? relationships
      : placed.slice(0, -1).map((node, index) => this.diagramEdge(`AUTO-${index}`, node.id, placed[index + 1].id, 'depends on', 'neutral'));
    const callouts = spec.kind === 'lineage'
      ? [this.diagramCallout('Terminal condition', 'Each branch stops only at a source of record, third party, manual entry, blocker, duplicate, obsolete source, or approved stopping point.', 690, 590, 420, 82, 'graph')]
      : spec.kind === 'dependency'
        ? [this.diagramCallout('Object rule', 'Every object node maps to evidence and every edge maps to a dependency, transform, read/write, approval, or blocker.', 690, 590, 420, 82, 'graph')]
        : [this.diagramCallout('Flow rule', 'Every movement must declare source, target, transform, cadence, confidence, and downstream failure impact.', 690, 590, 420, 82, 'graph')];

    return {
      spec,
      width,
      height,
      nodes: placed,
      edges: fallbackEdges,
      lanes: [],
      callouts
    };
  }

  private diagramObjectNodes(context: PackageContext): DiagramNode[] {
    const records = context.items.length ? context.items : [
      { id: 'SRC-001', name: this.importSourceName(), type: this.importSourceKind(), confidence: 70 },
      { id: 'GRAPH-001', name: 'Canonical discovery graph', type: 'graph', confidence: 80 },
      { id: 'PACK-001', name: 'Discovery Action Pack', type: 'output', confidence: 80 }
    ];

    return records.slice(0, 12).map((item, index) => {
      const type = this.readString(item, ['type', 'node_type']) || 'object';
      return this.diagramNode(
        this.readString(item, ['id', 'item_id', 'node_id']) || `N${index + 1}`,
        this.readString(item, ['name', 'object_name']) || `Discovered object ${index + 1}`,
        type,
        type,
        0,
        0,
        170,
        78,
        this.toneForType(type),
        this.readNumber(item, ['confidence'])
      );
    });
  }

  private diagramRelationshipEdges(context: PackageContext, nodes: DiagramNode[]): DiagramEdge[] {
    const ids = new Set(nodes.map((node) => node.id));
    return context.relationships
      .map((relationship, index) => {
        const from = this.readString(relationship, ['fromId', 'from_id', 'from']);
        const to = this.readString(relationship, ['toId', 'to_id', 'to']);
        if (!ids.has(from) || !ids.has(to)) {
          return null;
        }
        return this.diagramEdge(
          this.readString(relationship, ['id', 'relationship_id']) || `REL-${index + 1}`,
          from,
          to,
          this.readString(relationship, ['type', 'relationship_type', 'edgeType']) || 'depends on',
          'neutral'
        );
      })
      .filter((edge): edge is DiagramEdge => Boolean(edge));
  }

  private placeObjectNodes(nodes: DiagramNode[], kind: DiagramKind): DiagramNode[] {
    if (kind === 'lineage') {
      const columns = [94, 344, 594, 844];
      return nodes.map((node, index) => ({
        ...node,
        x: columns[index % columns.length],
        y: 160 + Math.floor(index / columns.length) * 138,
        w: 184,
        h: 88
      }));
    }

    if (kind === 'dataFlow') {
      return nodes.map((node, index) => ({
        ...node,
        x: 80 + (index % 5) * 214,
        y: 190 + Math.floor(index / 5) * 166,
        w: 172,
        h: 88
      }));
    }

    return nodes.map((node, index) => ({
      ...node,
      x: 80 + (index % 4) * 270,
      y: 150 + Math.floor(index / 4) * 150,
      w: 210,
      h: 92
    }));
  }

  private diagramNode(id: string, label: string, sublabel: string, type: string, x: number, y: number, w: number, h: number, tone: DiagramTone, confidence: number | null): DiagramNode {
    return { id, label: this.cleanDisplayText(label), sublabel: this.cleanDisplayText(sublabel), type: this.cleanDisplayText(type), x, y, w, h, tone, confidence };
  }

  private diagramEdge(id: string, from: string, to: string, label: string, tone: DiagramTone): DiagramEdge {
    return { id, from, to, label: this.cleanDisplayText(label), tone };
  }

  private diagramLane(label: string, x: number, y: number, w: number, h: number): DiagramLane {
    return { label, x, y, w, h };
  }

  private diagramCallout(title: string, text: string, x: number, y: number, w: number, h: number, tone: DiagramTone): DiagramCallout {
    return { title: this.cleanDisplayText(title), text: this.cleanDisplayText(text), x, y, w, h, tone };
  }

  private toneForType(type: string): DiagramTone {
    if (/risk|failure|exception/i.test(type)) {
      return 'risk';
    }
    if (/control|approval|validation/i.test(type)) {
      return 'control';
    }
    if (/action|backlog|remediation/i.test(type)) {
      return 'action';
    }
    if (/report|output|artifact|pack/i.test(type)) {
      return 'output';
    }
    if (/process|step|macro|query|transform|logic|vba|sql/i.test(type)) {
      return 'process';
    }
    if (/graph|lineage|model/i.test(type)) {
      return 'graph';
    }
    if (/source|system|database|file|workbook|table|sheet|document|access|excel/i.test(type)) {
      return 'source';
    }
    return 'neutral';
  }

  private buildDiagramSvg(scene: DiagramScene): string {
    const nodeById = new Map(scene.nodes.map((node) => [node.id, node]));
    const edges = scene.edges.map((edge) => {
      const from = nodeById.get(edge.from);
      const to = nodeById.get(edge.to);
      if (!from || !to) {
        return '';
      }
      const points = this.edgePoints(from, to);
      const color = this.diagramTone(edge.tone).stroke;
      const labelX = (points.x1 + points.x2) / 2;
      const labelY = (points.y1 + points.y2) / 2 - 8;
      return `<g class="edge"><line x1="${points.x1}" y1="${points.y1}" x2="${points.x2}" y2="${points.y2}" stroke="${color}" stroke-width="3" marker-end="url(#arrow)"/><text x="${labelX}" y="${labelY}" text-anchor="middle">${this.escapeXml(edge.label)}</text></g>`;
    }).join('');
    const lanes = scene.lanes.map((lane) => `<g class="lane"><rect x="${lane.x}" y="${lane.y}" width="${lane.w}" height="${lane.h}" rx="18"/><text x="${lane.x + 22}" y="${lane.y + 30}">${this.escapeXml(lane.label)}</text></g>`).join('');
    const nodes = scene.nodes.map((node) => {
      const tone = this.diagramTone(node.tone);
      const titleLines = this.wrapText(node.label, Math.max(12, Math.floor(node.w / 9))).slice(0, 3);
      const subLines = this.wrapText(node.sublabel, Math.max(14, Math.floor(node.w / 8))).slice(0, 2);
      const title = titleLines.map((line, index) => `<tspan x="${node.x + 18}" dy="${index === 0 ? 0 : 20}">${this.escapeXml(line)}</tspan>`).join('');
      const sub = subLines.map((line, index) => `<tspan x="${node.x + 18}" dy="${index === 0 ? 26 : 15}">${this.escapeXml(line)}</tspan>`).join('');
      const confidence = node.confidence === null ? '' : `<text class="confidence" x="${node.x + node.w - 18}" y="${node.y + node.h - 18}" text-anchor="end">${node.confidence}%</text>`;
      return `<g class="node ${node.tone}"><rect x="${node.x}" y="${node.y}" width="${node.w}" height="${node.h}" rx="18" fill="${tone.fill}" stroke="${tone.stroke}"/><text class="node-id" x="${node.x + 18}" y="${node.y + 23}">${this.escapeXml(node.id)}</text><text class="node-title" x="${node.x + 18}" y="${node.y + 48}">${title}</text><text class="node-sub" x="${node.x + 18}" y="${node.y + node.h - 32}">${sub}</text>${confidence}</g>`;
    }).join('');
    const callouts = scene.callouts.map((callout) => {
      const tone = this.diagramTone(callout.tone);
      const lines = this.wrapText(callout.text, Math.floor(callout.w / 8.6)).slice(0, 4);
      return `<g class="callout"><rect x="${callout.x}" y="${callout.y}" width="${callout.w}" height="${callout.h}" rx="18" fill="${tone.soft}" stroke="${tone.stroke}"/><text class="callout-title" x="${callout.x + 18}" y="${callout.y + 28}">${this.escapeXml(callout.title)}</text><text class="callout-text" x="${callout.x + 18}" y="${callout.y + 53}">${lines.map((line, index) => `<tspan x="${callout.x + 18}" dy="${index === 0 ? 0 : 16}">${this.escapeXml(line)}</tspan>`).join('')}</text></g>`;
    }).join('');

    return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${scene.width}" height="${scene.height}" viewBox="0 0 ${scene.width} ${scene.height}" role="img" aria-label="${this.escapeXml(scene.spec.title)}">
  <defs>
    <linearGradient id="bg" x1="0" x2="1" y1="0" y2="1"><stop offset="0" stop-color="#0d0805"/><stop offset="0.55" stop-color="#130d09"/><stop offset="1" stop-color="#061611"/></linearGradient>
    <pattern id="grid" width="44" height="44" patternUnits="userSpaceOnUse"><path d="M 44 0 L 0 0 0 44" fill="none" stroke="#6f431d" stroke-opacity="0.18" stroke-width="1"/></pattern>
    <filter id="glow" x="-30%" y="-30%" width="160%" height="160%"><feDropShadow dx="0" dy="14" stdDeviation="13" flood-color="#000" flood-opacity="0.38"/></filter>
    <marker id="arrow" viewBox="0 0 12 12" refX="10" refY="6" markerWidth="10" markerHeight="10" orient="auto-start-reverse"><path d="M 1 1 L 11 6 L 1 11 z" fill="#d88b35"/></marker>
  </defs>
  <style>
    .title{font:900 36px Arial, sans-serif;fill:#fff5e6}.subtitle{font:700 16px Arial, sans-serif;fill:#d8b37d}.meta{font:800 12px Arial, sans-serif;fill:#b88a56;text-transform:uppercase;letter-spacing:.08em}.lane rect{fill:#ffffff08;stroke:#d88b3540;stroke-width:1.5}.lane text{font:900 14px Arial, sans-serif;fill:#d8b37d}.edge text{font:800 11px Arial, sans-serif;fill:#f3d5a8;paint-order:stroke;stroke:#130d09;stroke-width:5}.node{filter:url(#glow)}.node rect{stroke-width:2.5}.node-id{font:900 11px Arial, sans-serif;fill:#f7c987}.node-title{font:900 17px Arial, sans-serif;fill:#fff8ec}.node-sub{font:700 12px Arial, sans-serif;fill:#d9c3a6}.confidence{font:900 13px Arial, sans-serif;fill:#8ff0d2}.callout-title{font:900 14px Arial, sans-serif;fill:#fff2df}.callout-text{font:700 12px Arial, sans-serif;fill:#e8cfac}
  </style>
  <rect width="1200" height="760" fill="url(#bg)"/><rect width="1200" height="760" fill="url(#grid)"/>
  <rect x="44" y="34" width="1112" height="82" rx="24" fill="#1b1009" stroke="#d88b35" stroke-opacity="0.45"/>
  <text class="meta" x="72" y="66">Uncle Kev's Distillery / Discovery Action Pack</text>
  <text class="title" x="72" y="101">${this.escapeXml(scene.spec.title)}</text>
  <text class="subtitle" x="72" y="138">${this.escapeXml(scene.spec.subtitle)}</text>
  ${lanes}
  ${edges}
  ${nodes}
  ${callouts}
</svg>`;
  }

  private buildDiagramMermaid(scene: DiagramScene): string {
    const lines = ['flowchart LR'];
    for (const node of scene.nodes) {
      lines.push(`  ${this.mermaidId(node.id)}["${this.escapeMermaid(node.label)}<br/>${this.escapeMermaid(node.sublabel)}"]`);
    }
    for (const edge of scene.edges) {
      lines.push(`  ${this.mermaidId(edge.from)} -->|"${this.escapeMermaid(edge.label)}"| ${this.mermaidId(edge.to)}`);
    }
    lines.push('  classDef source fill:#2c1a10,stroke:#d88b35,color:#fff5e6');
    lines.push('  classDef process fill:#251a2f,stroke:#c59cff,color:#fff5e6');
    lines.push('  classDef graph fill:#073126,stroke:#34d7b4,color:#fff5e6');
    lines.push('  classDef output fill:#102641,stroke:#7aa7ff,color:#fff5e6');
    lines.push('  classDef control fill:#243018,stroke:#b4dc68,color:#fff5e6');
    lines.push('  classDef risk fill:#44181a,stroke:#ff6b6b,color:#fff5e6');
    lines.push('  classDef action fill:#322006,stroke:#f7b84b,color:#fff5e6');
    lines.push('  classDef neutral fill:#1d1916,stroke:#a68a68,color:#fff5e6');
    for (const node of scene.nodes) {
      lines.push(`  class ${this.mermaidId(node.id)} ${node.tone};`);
    }
    return lines.join('\n');
  }

  private buildVisualDiagramPdf(scene: DiagramScene): Blob {
    const pageWidth = 792;
    const pageHeight = 612;
    const scale = Math.min(pageWidth / scene.width, pageHeight / scene.height);
    const offsetX = (pageWidth - scene.width * scale) / 2;
    const offsetY = (pageHeight - scene.height * scale) / 2;
    const commands: string[] = [];
    const sx = (x: number): number => offsetX + x * scale;
    const sy = (y: number): number => pageHeight - (offsetY + y * scale);
    const rectY = (y: number, h: number): number => pageHeight - (offsetY + (y + h) * scale);
    const color = (rgb: [number, number, number]): string => rgb.map((part) => part.toFixed(3).replace(/0+$/, '').replace(/\.$/, '')).join(' ');
    const hexRgb = (hex: string): [number, number, number] => {
      const clean = hex.replace('#', '');
      return [
        parseInt(clean.slice(0, 2), 16) / 255,
        parseInt(clean.slice(2, 4), 16) / 255,
        parseInt(clean.slice(4, 6), 16) / 255
      ];
    };
    const rect = (x: number, y: number, w: number, h: number, fill: string, stroke = '', strokeWidth = 1): void => {
      const fillColor = color(hexRgb(fill));
      const strokePart = stroke ? `${color(hexRgb(stroke))} RG ${strokeWidth} w` : '';
      const op = stroke ? 'B' : 'f';
      commands.push(`q ${fillColor} rg ${strokePart} ${sx(x)} ${rectY(y, h)} ${w * scale} ${h * scale} re ${op} Q`);
    };
    const line = (x1: number, y1: number, x2: number, y2: number, stroke: string, width = 2): void => {
      commands.push(`q ${color(hexRgb(stroke))} RG ${width} w ${sx(x1)} ${sy(y1)} m ${sx(x2)} ${sy(y2)} l S Q`);
    };
    const polygon = (points: Array<[number, number]>, fill: string): void => {
      if (!points.length) {
        return;
      }
      const [first, ...rest] = points;
      commands.push(`q ${color(hexRgb(fill))} rg ${sx(first[0])} ${sy(first[1])} m ${rest.map((point) => `${sx(point[0])} ${sy(point[1])} l`).join(' ')} h f Q`);
    };
    const text = (x: number, y: number, value: string, size: number, font: 'F1' | 'F2', fill = '#fff5e6'): void => {
      commands.push(`BT /${font} ${Math.max(5, size * scale)} Tf ${color(hexRgb(fill))} rg ${sx(x)} ${sy(y)} Td (${this.escapePdf(this.stripControlChars(value))}) Tj ET`);
    };
    const wrappedText = (x: number, y: number, value: string, width: number, size: number, font: 'F1' | 'F2', fill = '#fff5e6', maxLines = 3): void => {
      this.wrapText(value, Math.max(10, Math.floor(width / (size * 0.56)))).slice(0, maxLines).forEach((lineText, index) => {
        text(x, y + index * (size + 4), lineText, size, font, fill);
      });
    };

    rect(0, 0, scene.width, scene.height, '#0d0805');
    for (let x = 0; x <= scene.width; x += 48) {
      line(x, 0, x, scene.height, '#2a1a10', 0.45);
    }
    for (let y = 0; y <= scene.height; y += 48) {
      line(0, y, scene.width, y, '#2a1a10', 0.45);
    }
    rect(44, 34, 1112, 82, '#1b1009', '#d88b35', 1.4);
    text(72, 66, "Uncle Kev's Distillery / Discovery Action Pack", 12, 'F2', '#d8b37d');
    text(72, 101, scene.spec.title, 34, 'F2', '#fff5e6');
    wrappedText(72, 136, scene.spec.subtitle, 900, 15, 'F1', '#d8b37d', 1);

    for (const lane of scene.lanes) {
      rect(lane.x, lane.y, lane.w, lane.h, '#15100c', '#5b3719', 0.7);
      text(lane.x + 22, lane.y + 30, lane.label, 14, 'F2', '#d8b37d');
    }

    const nodeById = new Map(scene.nodes.map((node) => [node.id, node]));
    for (const edge of scene.edges) {
      const from = nodeById.get(edge.from);
      const to = nodeById.get(edge.to);
      if (!from || !to) {
        continue;
      }
      const points = this.edgePoints(from, to);
      const tone = this.diagramTone(edge.tone);
      line(points.x1, points.y1, points.x2, points.y2, tone.stroke, 2.2);
      const angle = Math.atan2(points.y2 - points.y1, points.x2 - points.x1);
      const arrow = 14;
      polygon([
        [points.x2, points.y2],
        [points.x2 - Math.cos(angle - 0.45) * arrow, points.y2 - Math.sin(angle - 0.45) * arrow],
        [points.x2 - Math.cos(angle + 0.45) * arrow, points.y2 - Math.sin(angle + 0.45) * arrow]
      ], tone.stroke);
      wrappedText((points.x1 + points.x2) / 2 - 55, (points.y1 + points.y2) / 2 - 10, edge.label, 110, 10, 'F2', '#f3d5a8', 1);
    }

    for (const node of scene.nodes) {
      const tone = this.diagramTone(node.tone);
      rect(node.x, node.y, node.w, node.h, tone.fill, tone.stroke, 1.5);
      text(node.x + 16, node.y + 24, node.id, 11, 'F2', '#f7c987');
      wrappedText(node.x + 16, node.y + 50, node.label, node.w - 30, 17, 'F2', '#fff8ec', 2);
      wrappedText(node.x + 16, node.y + node.h - 30, node.sublabel, node.w - 30, 11, 'F1', '#d9c3a6', 1);
      if (node.confidence !== null) {
        text(node.x + node.w - 48, node.y + node.h - 17, `${node.confidence}%`, 12, 'F2', '#8ff0d2');
      }
    }

    for (const callout of scene.callouts) {
      const tone = this.diagramTone(callout.tone);
      rect(callout.x, callout.y, callout.w, callout.h, tone.soft, tone.stroke, 1.2);
      text(callout.x + 18, callout.y + 28, callout.title, 14, 'F2', '#fff2df');
      wrappedText(callout.x + 18, callout.y + 54, callout.text, callout.w - 36, 11, 'F1', '#e8cfac', 4);
    }

    const content = commands.join('\n');
    const objects: string[] = [];
    objects[1] = '<< /Type /Catalog /Pages 2 0 R >>';
    objects[2] = '<< /Type /Pages /Kids [5 0 R] /Count 1 >>';
    objects[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>';
    objects[4] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>';
    objects[5] = '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 792 612] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents 6 0 R >>';
    objects[6] = `<< /Length ${content.length} >>\nstream\n${content}\nendstream`;
    return this.pdfFromObjects(objects);
  }

  private edgePoints(from: DiagramNode, to: DiagramNode): { x1: number; y1: number; x2: number; y2: number } {
    const fromCenter = { x: from.x + from.w / 2, y: from.y + from.h / 2 };
    const toCenter = { x: to.x + to.w / 2, y: to.y + to.h / 2 };
    const dx = toCenter.x - fromCenter.x || 1;
    const dy = toCenter.y - fromCenter.y || 1;
    const trimFrom = Math.min(Math.abs((from.w / 2) / dx), Math.abs((from.h / 2) / dy)) * 0.88;
    const trimTo = Math.min(Math.abs((to.w / 2) / dx), Math.abs((to.h / 2) / dy)) * 0.88;
    return {
      x1: fromCenter.x + dx * trimFrom,
      y1: fromCenter.y + dy * trimFrom,
      x2: toCenter.x - dx * trimTo,
      y2: toCenter.y - dy * trimTo
    };
  }

  private diagramTone(tone: DiagramTone): { fill: string; stroke: string; soft: string } {
    const tones: Record<DiagramTone, { fill: string; stroke: string; soft: string }> = {
      source: { fill: '#2b1d14', stroke: '#d88b35', soft: '#21150d' },
      process: { fill: '#251a2f', stroke: '#c59cff', soft: '#1b1424' },
      graph: { fill: '#073126', stroke: '#34d7b4', soft: '#06241c' },
      output: { fill: '#102641', stroke: '#7aa7ff', soft: '#0c1c30' },
      control: { fill: '#243018', stroke: '#b4dc68', soft: '#192212' },
      risk: { fill: '#44181a', stroke: '#ff6b6b', soft: '#2a1011' },
      action: { fill: '#322006', stroke: '#f7b84b', soft: '#241704' },
      neutral: { fill: '#1d1916', stroke: '#a68a68', soft: '#171311' }
    };
    return tones[tone];
  }

  private inlineGraphEdge(id: string, from: InlineGraphNode, to: InlineGraphNode, label: string, tone: DiagramTone): InlineGraphEdge {
    const x1 = from.x + 8;
    const y1 = from.y + 8;
    const x2 = to.x + 8;
    const y2 = to.y + 8;
    return {
      id,
      label: this.cleanDisplayText(label),
      tone,
      x1,
      y1,
      x2,
      y2,
      labelX: (x1 + x2) / 2,
      labelY: (y1 + y2) / 2
    };
  }

  private nodeLabelMap(items: Record<string, unknown>[]): Map<string, string> {
    const labels = new Map<string, string>();
    items.forEach((item, index) => {
      const id = this.readString(item, ['id', 'item_id', 'node_id']) || `N${index + 1}`;
      const label = this.readString(item, ['name', 'object_name', 'title']) || id;
      labels.set(id, label);
    });
    return labels;
  }

  private relationshipDisplay(relationship: Record<string, unknown>, labels: Map<string, string>): string {
    const fromId = this.readString(relationship, ['fromId', 'from_id', 'from']);
    const toId = this.readString(relationship, ['toId', 'to_id', 'to']);
    const type = this.readString(relationship, ['type', 'edgeType', 'relationship_type']) || 'depends on';
    const confidence = this.readString(relationship, ['confidence']);
    const from = labels.get(fromId) || fromId || 'unresolved source';
    const to = labels.get(toId) || toId || 'unresolved target';
    return `${from} -> ${to} (${type}${confidence ? `, ${confidence}% confidence` : ''})`;
  }

  private mermaidId(value: string): string {
    return `N_${value.replace(/[^A-Za-z0-9_]+/g, '_')}`;
  }

  private escapeMermaid(value: string): string {
    return this.cleanDisplayText(value).replace(/"/g, "'");
  }

  private pdfFromObjects(objects: string[]): Blob {
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

  private async addEvidenceArchive(zip: unknown, context: PackageContext): Promise<void> {
    const folder = (zip as { folder: (name: string) => { file: (path: string, data: string | Blob) => void; folder: (path: string) => { file: (path: string, data: string | Blob) => void } | null } | null }).folder('08_Evidence_Archive');
    folder?.folder('Document_Extracts')?.file('Extracted_Evidence.txt', this.extractedText());
    folder?.folder('Document_Extracts')?.file('Canonical_Model.json', JSON.stringify(context.delta, null, 2));
    folder?.file('Evidence_Index.csv', this.toCsv(context.evidenceIndex.length ? context.evidenceIndex : this.evidenceFromItems(context.items)));
  }

  private buildPdf(title: string, lines: string[]): Blob {
    return this.buildScientificPdf(
      title,
      'Generated from the canonical Distillery model',
      this.linesToPdfBlocks(lines)
    );
  }

  private linesToPdfBlocks(lines: string[]): PdfBlock[] {
    const blocks: PdfBlock[] = [];
    let currentTitle = '';
    let currentText: string[] = [];
    const flush = (): void => {
      if (!currentTitle && !currentText.length) {
        return;
      }
      blocks.push({
        type: currentTitle ? 'section' : 'paragraph',
        title: currentTitle || undefined,
        text: currentText.join('\n')
      });
      currentTitle = '';
      currentText = [];
    };

    for (const line of lines) {
      const text = this.cleanDisplayText(this.asText(line));
      if (!text) {
        continue;
      }
      const looksLikeHeading = text.length < 72 && !/[.:]$/.test(text) && /^[A-Z0-9]/.test(text);
      if (looksLikeHeading) {
        flush();
        currentTitle = text;
      } else {
        currentText.push(text);
      }
    }
    flush();
    return blocks.length ? blocks : [{ type: 'paragraph', text: 'No report content generated.' }];
  }

  private buildScientificPdf(title: string, subtitle: string, blocks: PdfBlock[]): Blob {
    const pageWidth = 612;
    const pageHeight = 792;
    const margin = 48;
    const contentWidth = pageWidth - margin * 2;
    const bottom = 64;
    const bodyTop = 700;
    const bodyPages: string[][] = [];
    let page: string[] = [];
    let y = bodyTop;

    const newPage = (): void => {
      page = [];
      bodyPages.push(page);
      y = bodyTop;
    };

    const color = (rgb: [number, number, number]): string => rgb.map((part) => part.toFixed(3).replace(/0+$/, '').replace(/\.$/, '')).join(' ');
    const textCommand = (x: number, lineY: number, value: string, font: 'F1' | 'F2', size: number, rgb: [number, number, number]): string => {
      return `BT /${font} ${size} Tf ${color(rgb)} rg ${x} ${lineY} Td (${this.escapePdf(this.stripControlChars(value))}) Tj ET`;
    };
    const rectCommand = (x: number, rectY: number, width: number, height: number, rgb: [number, number, number]): string => {
      return `q ${color(rgb)} rg ${x} ${rectY} ${width} ${height} re f Q`;
    };
    const lineCommand = (x1: number, y1: number, x2: number, y2: number, rgb: [number, number, number], width = 0.75): string => {
      return `q ${color(rgb)} RG ${width} w ${x1} ${y1} m ${x2} ${y2} l S Q`;
    };
    const addText = (x: number, lineY: number, value: string, font: 'F1' | 'F2', size: number, rgb: [number, number, number] = [0.09, 0.07, 0.05]): void => {
      page.push(textCommand(x, lineY, value, font, size, rgb));
    };
    const addRect = (x: number, rectY: number, width: number, height: number, rgb: [number, number, number]): void => {
      page.push(rectCommand(x, rectY, width, height, rgb));
    };
    const addLine = (x1: number, y1: number, x2: number, y2: number, rgb: [number, number, number], width = 0.75): void => {
      page.push(lineCommand(x1, y1, x2, y2, rgb, width));
    };
    const ensure = (height: number): void => {
      if (y - height < bottom) {
        newPage();
      }
    };
    const wrapForWidth = (value: string, width: number, size: number): string[] => {
      return this.wrapText(this.cleanDisplayText(value), Math.max(18, Math.floor(width / (size * 0.52))));
    };
    const addWrapped = (value: string, size = 10.5, font: 'F1' | 'F2' = 'F1', x = margin, width = contentWidth, rgb: [number, number, number] = [0.12, 0.09, 0.07]): void => {
      const paragraphs = this.cleanDisplayText(value).split(/\n+/).map((part) => part.trim()).filter(Boolean);
      const lineHeight = size + 4.2;
      for (const paragraph of paragraphs.length ? paragraphs : ['']) {
        const lines = wrapForWidth(paragraph, width, size);
        ensure(lines.length * lineHeight + 4);
        for (const line of lines) {
          addText(x, y, line, font, size, rgb);
          y -= lineHeight;
        }
        y -= 3;
      }
    };
    const addSectionTitle = (heading: string): void => {
      ensure(44);
      y -= 6;
      addText(margin, y, this.cleanDisplayText(heading), 'F2', 13.5, [0.32, 0.18, 0.08]);
      y -= 7;
      addLine(margin, y, margin + contentWidth, y, [0.78, 0.48, 0.2], 0.9);
      y -= 16;
    };
    const addKeyValues = (block: PdfBlock): void => {
      if (block.title) {
        addSectionTitle(block.title);
      }
      for (const row of block.rows || []) {
        const label = this.cleanDisplayText(row[0] || '');
        const value = this.cleanDisplayText(row[1] || '');
        const valueLines = wrapForWidth(value, contentWidth - 170, 9.4);
        const rowHeight = Math.max(30, valueLines.length * 12 + 14);
        ensure(rowHeight + 6);
        addRect(margin, y - rowHeight + 8, contentWidth, rowHeight, [0.965, 0.94, 0.9]);
        addText(margin + 12, y - 10, label, 'F2', 8.5, [0.42, 0.26, 0.12]);
        let valueY = y - 10;
        for (const line of valueLines) {
          addText(margin + 158, valueY, line, 'F1', 9.4, [0.12, 0.09, 0.07]);
          valueY -= 12;
        }
        y -= rowHeight + 5;
      }
      y -= 5;
    };
    const addBullets = (block: PdfBlock): void => {
      if (block.title) {
        addSectionTitle(block.title);
      }
      for (const item of block.items || []) {
        const lines = wrapForWidth(item, contentWidth - 20, 10);
        ensure(lines.length * 13 + 4);
        addText(margin, y, '-', 'F2', 10, [0.78, 0.48, 0.2]);
        let lineY = y;
        for (const line of lines) {
          addText(margin + 18, lineY, line, 'F1', 10, [0.12, 0.09, 0.07]);
          lineY -= 13;
        }
        y = lineY - 3;
      }
      y -= 4;
    };
    const addTable = (block: PdfBlock): void => {
      if (block.title) {
        addSectionTitle(block.title);
      }
      const rows = (block.rows || []).filter((row) => row.length);
      if (!rows.length) {
        addWrapped('No table rows generated.');
        return;
      }
      const columns = rows[0].slice(0, 4);
      const widths = columns.length === 2
        ? [150, contentWidth - 150]
        : columns.length === 3
          ? [150, 230, contentWidth - 380]
          : [72, 92, 150, contentWidth - 314];
      const renderRow = (row: string[], header = false): void => {
        const cells = columns.map((_, index) => this.cleanDisplayText(row[index] || ''));
        const wrapped = cells.map((cell, index) => wrapForWidth(cell, widths[index] - 12, header ? 8.2 : 8));
        const rowHeight = Math.max(header ? 26 : 34, Math.max(...wrapped.map((lines) => lines.length)) * 10 + 14);
        ensure(rowHeight + 4);
        addRect(margin, y - rowHeight + 8, contentWidth, rowHeight, header ? [0.16, 0.11, 0.07] : [0.985, 0.97, 0.94]);
        let x = margin;
        cells.forEach((_, index) => {
          let lineY = y - 10;
          wrapped[index].forEach((line) => {
            addText(x + 6, lineY, line, header ? 'F2' : 'F1', header ? 8.2 : 8, header ? [1, 0.95, 0.86] : [0.12, 0.09, 0.07]);
            lineY -= 10;
          });
          if (index < cells.length - 1) {
            addLine(x + widths[index], y + 5, x + widths[index], y - rowHeight + 10, header ? [0.34, 0.24, 0.16] : [0.85, 0.78, 0.67], 0.4);
          }
          x += widths[index];
        });
        y -= rowHeight + 4;
      };
      renderRow(rows[0], true);
      rows.slice(1, 16).forEach((row) => renderRow(row));
      if (rows.length > 16) {
        addWrapped(`${rows.length - 16} additional rows are available in the workbook and machine-readable package.`, 8.8, 'F1', margin, contentWidth, [0.42, 0.26, 0.12]);
      }
      y -= 6;
    };
    const addNote = (block: PdfBlock): void => {
      const text = this.cleanDisplayText(block.text || '');
      const lines = wrapForWidth(text, contentWidth - 24, 9.8);
      const height = Math.max(48, lines.length * 12 + 34);
      ensure(height + 8);
      addRect(margin, y - height + 8, contentWidth, height, [1, 0.956, 0.88]);
      if (block.title) {
        addText(margin + 12, y - 12, block.title, 'F2', 10.4, [0.42, 0.24, 0.1]);
      }
      let lineY = y - 28;
      for (const line of lines) {
        addText(margin + 12, lineY, line, 'F1', 9.8, [0.16, 0.11, 0.07]);
        lineY -= 12;
      }
      y -= height + 10;
    };

    newPage();
    addRect(margin, y - 66, contentWidth, 64, [0.12, 0.08, 0.05]);
    addText(margin + 14, y - 24, title, 'F2', 20, [1, 0.96, 0.88]);
    addText(margin + 14, y - 45, this.cleanDisplayText(subtitle), 'F1', 9.4, [0.87, 0.72, 0.56]);
    y -= 88;

    for (const block of blocks) {
      if (block.type === 'keyValues') {
        addKeyValues(block);
      } else if (block.type === 'bullets') {
        addBullets(block);
      } else if (block.type === 'table') {
        addTable(block);
      } else if (block.type === 'note') {
        addNote(block);
      } else if (block.type === 'section') {
        if (block.title) {
          addSectionTitle(block.title);
        }
        addWrapped(block.text || '');
        y -= 4;
      } else {
        addWrapped(block.text || '');
      }
    }

    const pages = bodyPages.length ? bodyPages : [[]];
    const objects: string[] = [];
    objects[1] = '<< /Type /Catalog /Pages 2 0 R >>';
    objects[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>';
    objects[4] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>';
    const kids: string[] = [];

    pages.forEach((pageCommands, index) => {
      const pageObject = 5 + index * 2;
      const contentObject = pageObject + 1;
      kids.push(`${pageObject} 0 R`);
      const content = [
        rectCommand(0, pageHeight - 44, pageWidth, 44, [0.08, 0.06, 0.04]),
        textCommand(margin, pageHeight - 27, this.cleanDisplayText(title).slice(0, 64), 'F2', 9.5, [1, 0.96, 0.88]),
        textCommand(pageWidth - 178, pageHeight - 27, "Uncle Kev's Distillery", 'F1', 8.4, [0.87, 0.72, 0.56]),
        lineCommand(margin, pageHeight - 49, pageWidth - margin, pageHeight - 49, [0.78, 0.48, 0.2], 0.9),
        ...pageCommands,
        lineCommand(margin, 47, pageWidth - margin, 47, [0.78, 0.48, 0.2], 0.45),
        textCommand(margin, 30, `Generated ${new Date().toLocaleDateString()} from canonical discovery evidence`, 'F1', 7.8, [0.42, 0.34, 0.26]),
        textCommand(pageWidth - 118, 30, `Page ${index + 1} of ${pages.length}`, 'F1', 7.8, [0.42, 0.34, 0.26])
      ].join('\n');
      objects[pageObject] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents ${contentObject} 0 R >>`;
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

  private readString(record: Record<string, unknown>, keys: string[]): string {
    for (const key of keys) {
      const value = record[key];
      const label = this.labelFor(value);
      if (label) {
        return label;
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
          .map((item) => this.labelFor(item))
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

  private listText(value: unknown, fallback = ''): string {
    const values = Array.isArray(value) ? value : [value];
    const text = values
      .map((item) => this.labelFor(item) || this.asText(item))
      .map((item) => this.cleanDisplayText(item))
      .filter(Boolean)
      .join(', ');
    return text || fallback;
  }

  private labelFor(value: unknown): string {
    if (value === null || value === undefined) {
      return '';
    }
    if (typeof value === 'string') {
      return this.cleanDisplayText(value);
    }
    if (typeof value === 'number' || typeof value === 'boolean') {
      return String(value);
    }
    if (Array.isArray(value)) {
      return this.listText(value);
    }
    if (typeof value === 'object') {
      const record = value as Record<string, unknown>;
      const preferredKeys = [
        'name',
        'title',
        'output',
        'outputName',
        'criticalOutput',
        'label',
        'id',
        'evidenceId',
        'evidence_id',
        'object_name',
        'artifact_name',
        'summary',
        'scenario',
        'question'
      ];
      for (const key of preferredKeys) {
        const candidate = record[key];
        if (typeof candidate === 'string' && candidate.trim()) {
          return this.cleanDisplayText(candidate);
        }
        if (typeof candidate === 'number' || typeof candidate === 'boolean') {
          return String(candidate);
        }
      }
      return this.compactObjectText(record);
    }
    return '';
  }

  private compactObjectText(record: Record<string, unknown>): string {
    const exposure = this.exposureBucketText(record);
    if (exposure) {
      return exposure;
    }
    return Object.entries(record)
      .filter(([, value]) => value !== null && value !== undefined && typeof value !== 'object')
      .slice(0, 4)
      .map(([key, value]) => `${this.displayKey(key)}: ${this.cleanDisplayText(String(value))}`)
      .join('; ');
  }

  private asText(value: unknown): string {
    if (value === null || value === undefined) {
      return '';
    }
    if (typeof value === 'string') {
      return this.cleanDisplayText(value);
    }
    if (typeof value === 'number' || typeof value === 'boolean') {
      return String(value);
    }
    if (Array.isArray(value)) {
      return value
        .map((item) => this.labelFor(item) || this.asText(item))
        .map((item) => this.cleanDisplayText(item))
        .filter(Boolean)
        .join('; ');
    }
    if (typeof value === 'object') {
      const record = value as Record<string, unknown>;
      const exposure = this.exposureSummary(record);
      if (exposure) {
        return exposure;
      }
      return Object.entries(record)
        .map(([key, entry]) => {
          const text = this.asText(entry);
          return text ? `${this.displayKey(key)}: ${text}` : '';
        })
        .filter(Boolean)
        .join(' ');
    }
    return '';
  }

  private exposureSummary(value: unknown): string {
    const record = this.asRecord(value);
    if (!Object.keys(record).length) {
      return '';
    }

    const singleBucket = this.exposureBucketText(record);
    if (singleBucket) {
      return singleBucket;
    }

    const bucketLines = Object.entries(record)
      .map(([key, entry]) => {
        if (entry && typeof entry === 'object' && !Array.isArray(entry)) {
          const bucket = this.exposureBucketText(entry as Record<string, unknown>);
          if (bucket) {
            return `${this.displayKey(key)}: ${bucket}`;
          }
        }
        const text = this.asText(entry);
        return text ? `${this.displayKey(key)}: ${text}` : '';
      })
      .filter(Boolean);
    return bucketLines.join('\n');
  }

  private exposureBucketText(record: Record<string, unknown>): string {
    const hasExposureValues = ['low', 'base', 'high'].some((key) => record[key] !== undefined);
    if (!hasExposureValues) {
      return '';
    }

    const values = ['low', 'base', 'high']
      .filter((key) => record[key] !== undefined)
      .map((key) => `${this.displayKey(key)} ${this.formatExposureValue(record[key])}`)
      .join(', ');
    const assumptions = this.labelFor(record['assumptions']) || this.asText(record['assumptions']);
    return this.cleanDisplayText(`${values || 'Estimate pending'}${assumptions ? `. Assumptions: ${assumptions}` : ''}.`);
  }

  private formatExposureValue(value: unknown): string {
    if (typeof value === 'number' && Number.isFinite(value)) {
      return value === 0 ? '$0' : value.toLocaleString(undefined, { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });
    }
    if (typeof value === 'string' && value.trim()) {
      const numeric = Number(value);
      if (Number.isFinite(numeric)) {
        return numeric === 0 ? '$0' : numeric.toLocaleString(undefined, { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });
      }
      return this.cleanDisplayText(value);
    }
    return this.labelFor(value) || 'pending';
  }

  private displayKey(key: string): string {
    return key
      .replace(/[_-]+/g, ' ')
      .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/^./, (letter) => letter.toUpperCase());
  }

  private cleanDisplayText(value: string): string {
    return String(value)
      .replace(/\r\n?/g, '\n')
      .replace(/[^\x20-\x7E\n]/g, ' ')
      .replace(/\[object Object\]/g, 'unlabeled object')
      .replace(/[ \t]+/g, ' ')
      .replace(/ *\n */g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
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
