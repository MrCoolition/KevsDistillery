import { CommonModule } from '@angular/common';
import { Component, computed, signal } from '@angular/core';

@Component({
  selector: 'app-root',
  imports: [CommonModule],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class App {
  protected readonly uploadInstruction =
    'Upload the source file(s) for discovery. Supported sources include Access databases, Excel workbooks, Word documents, CSV/text files, SQL/script files, and related supporting files. For best results, upload all known upstream/downstream files together. If upstream files are missing, the dossier will document them as lineage blockers and create action items to resolve them.';

  protected readonly runPrompt =
    'Start a fresh elite Data Source Discovery Dossier for the uploaded file(s). Use the full dossier standard. Generate the complete package with the required folder structure, executive brief, architecture report, technical workbook, diagram pack with legends, evidence archive, auto-documentation pack, metadata manifest, action backlog, and financial impact model. Run QA before delivery and clearly document any blockers or assumptions.';

  protected readonly selectedFiles = signal<File[]>([]);
  protected readonly isRunning = signal(false);
  protected readonly status = signal('Ready for source files');
  protected readonly error = signal('');
  protected readonly downloadUrl = signal('');
  protected readonly downloadName = signal('');
  protected readonly packageSummary = signal<PackageSummary | null>(null);
  protected readonly elapsedSeconds = signal(0);
  private runTimer: number | undefined;

  protected readonly fileCount = computed(() => this.selectedFiles().length);
  protected readonly totalSize = computed(() =>
    this.selectedFiles().reduce((sum, file) => sum + file.size, 0),
  );

  protected onFilesSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.selectedFiles.set(Array.from(input.files ?? []));
    this.error.set('');
    this.downloadUrl.set('');
    this.downloadName.set('');
    this.packageSummary.set(null);
    this.elapsedSeconds.set(0);
    this.status.set(this.selectedFiles().length ? 'Files staged for a fresh discovery run' : 'Ready for source files');
  }

  protected removeFile(index: number): void {
    this.selectedFiles.update((files) => files.filter((_, fileIndex) => fileIndex !== index));
  }

  protected async generateDossier(): Promise<void> {
    if (!this.selectedFiles().length || this.isRunning()) {
      return;
    }

    this.isRunning.set(true);
    this.error.set('');
    this.downloadUrl.set('');
    this.downloadName.set('');
    this.packageSummary.set(null);
    this.elapsedSeconds.set(0);
    this.status.set('Running trusted desktop deep discovery, recursive lineage, AI synthesis, and package QA');
    this.runTimer = window.setInterval(() => this.elapsedSeconds.update((seconds) => seconds + 1), 1000);

    try {
      const body = new FormData();
      for (const file of this.selectedFiles()) {
        body.append('sources', file, file.name);
      }
      body.append('runPrompt', this.runPrompt);

      const response = await fetch('/api/dossiers', {
        method: 'POST',
        body,
      });

      if (!response.ok) {
        const problem = await response.text();
        throw new Error(problem || `Dossier generation failed with ${response.status}`);
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const disposition = response.headers.get('content-disposition') ?? '';
      const nameMatch = disposition.match(/filename="?([^"]+)"?/i);
      const summaryHeader = response.headers.get('x-dossier-summary');

      this.downloadUrl.set(url);
      this.downloadName.set(nameMatch?.[1] ?? 'Data_Source_Discovery_Dossier.zip');
      this.packageSummary.set(summaryHeader ? JSON.parse(decodeURIComponent(summaryHeader)) as PackageSummary : null);
      this.status.set('QA complete. Package ready.');
    } catch (cause) {
      const message = cause instanceof Error ? cause.message : 'Unexpected generation failure';
      this.error.set(
        message === 'Failed to fetch'
          ? 'The browser could not reach the dossier API. Confirm the combined dev server is running with npm start, then retry. Large Access files now run with bounded extraction and should return a package with explicit blockers instead of waiting indefinitely.'
          : message,
      );
      this.status.set('Generation stopped');
    } finally {
      if (this.runTimer !== undefined) {
        window.clearInterval(this.runTimer);
        this.runTimer = undefined;
      }
      this.isRunning.set(false);
    }
  }

  protected formatBytes(bytes: number): string {
    if (!bytes) {
      return '0 B';
    }

    const units = ['B', 'KB', 'MB', 'GB'];
    const exponent = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
    return `${(bytes / 1024 ** exponent).toFixed(exponent === 0 ? 0 : 1)} ${units[exponent]}`;
  }

  protected formatNumber(value: number): string {
    return new Intl.NumberFormat('en-US').format(Math.round(value || 0));
  }

  protected formatCurrency(value: number): string {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: value > 0 && value < 0.01 ? 4 : 2,
      maximumFractionDigits: value > 0 && value < 0.01 ? 4 : 2,
    }).format(value || 0);
  }

  protected formatPercent(value: number): string {
    return new Intl.NumberFormat('en-US', {
      style: 'percent',
      minimumFractionDigits: 0,
      maximumFractionDigits: 1,
    }).format(value || 0);
  }
}

type PackageSummary = {
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

type AiCostSummary = {
  pricingSource: string;
  totalRequests: number;
  totalInputTokens: number;
  totalCachedInputTokens: number;
  totalBillableInputTokens: number;
  totalOutputTokens: number;
  totalReasoningTokens: number;
  totalTokens: number;
  estimatedTotalCostUsd: number;
  estimatedCacheSavingsUsd: number;
  cacheHitRate: number;
  optimizationNote: string;
  models: Array<{
    model: string;
    requests: number;
    inputTokens: number;
    cachedInputTokens: number;
    billableInputTokens: number;
    outputTokens: number;
    reasoningTokens: number;
    totalTokens: number;
    inputCostUsd: number;
    cachedInputCostUsd: number;
    outputCostUsd: number;
    totalCostUsd: number;
    cacheSavingsUsd: number;
    cacheHitRate: number;
    pricing: {
      inputPerMillion: number;
      cachedInputPerMillion: number;
      outputPerMillion: number;
      pricingAvailable: boolean;
    };
  }>;
};
