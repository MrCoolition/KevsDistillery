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
    error?: string;
  };
}

interface SynthesisResponse {
  ok: boolean;
  runId?: string;
  stored?: boolean;
  counts?: {
    items: number;
    relationships: number;
    artifacts: number;
    backlog: number;
  } | null;
  model?: string;
  outputText?: string;
  canonicalDelta?: {
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
  readonly selectedItemId = signal<string>('OUT-001');
  readonly searchTerm = signal('');
  readonly importSourceKind = signal<SourceKind>('excel');
  readonly importSourceName = signal('Selected source set');
  readonly knownArtifactsText = signal('');
  readonly targetOutputsText = signal('01_Executive_Decision_Brief.pdf\n02_Current_State_Architecture_Report.pdf\n03_Technical_Discovery_Workbook.xlsx\n05_Diagram_Pack\n07_Action_Backlog.csv');
  readonly extractedText = signal('Choose files or a folder above. Extracted evidence will appear here before execution.');
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
      description: 'Multi-source evidence batches for full current-state synthesis.'
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
        audience: artifact.audience || 'Discovery team',
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

  readonly reportSections = computed(() => [
    {
      title: 'Executive Snapshot',
      body: `${this.model.processName} supports ${this.model.businessFunction}. Execute a real source set to replace this starter state with evidence-backed scope, lineage, blockers, confidence, and actions.`
    },
    {
      title: 'Current-State Narrative',
      body: 'The generated report compresses the operating model into trigger, actor, input, processing step, validation, output, handoff, and exception path sections while preserving evidence references.'
    },
    {
      title: 'Lineage and Controls',
      body: `${this.model.relationships.length} starter node-edge relationships connect source selection, canonical graph generation, and artifact output. Executed runs expand this into real lineage.`
    },
    {
      title: 'Remediation Backlog',
      body: `${this.criticalBacklog().length} action records are prioritized for Fivetran ingestion, dbt rebuild, Snowpark controls, governance, and retirement decisions.`
    }
  ]);

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
      const body = await response.json();
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
      const body = await response.json();
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
      this.synthesisError.set('No evidence payload is staged. Choose files or paste evidence first.');
      return;
    }

    if (this.stagedSources().length > 0) {
      const confirmed = window.confirm('Execute Distillery synthesis? This sends extracted evidence from the selected sources to your configured backend and OpenAI model.');
      if (!confirmed) {
        return;
      }
    }

    this.isSynthesizing.set(true);
    this.synthesisError.set('');

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
      const body = await response.json();
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

    this.importSourceName.set(files.length === 1 ? files[0].name : `${files.length} source files`);
    this.knownArtifactsText.set(sources.map((source) => source.path).join('\n'));
    this.extractedText.set(this.buildEvidencePayload(sources));
    this.isExtractingSources.set(false);
  }

  clearSources(fileInput?: HTMLInputElement, folderInput?: HTMLInputElement): void {
    this.stagedSources.set([]);
    this.extractedText.set('');
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
      const body = await response.json();
      if (!response.ok) {
        throw new Error(body.error || 'Neon schema migration failed.');
      }
      await this.checkApiHealth();
    } catch (error) {
      this.apiError.set(error instanceof Error ? error.message : 'Neon schema migration failed.');
    }
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
        text = 'Access binary selected. Browser security blocks direct ACCDB introspection; use an exported object inventory, saved SQL export, VBA export, or database connector extract for deep object discovery.';
        status = 'access binary metadata';
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

  private buildEvidencePayload(sources: StagedSource[]): string {
    return sources.map((source, index) => [
      `--- SOURCE ${index + 1}: ${source.path} ---`,
      `name: ${source.name}`,
      `extension: ${source.extension || 'none'}`,
      `size: ${source.size} bytes`,
      `status: ${source.status}`,
      '',
      source.text
    ].join('\n')).join('\n\n').slice(0, 240000);
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
