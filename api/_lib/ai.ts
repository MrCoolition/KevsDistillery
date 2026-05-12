import OpenAI from 'openai';
import { MASTER_DOSSIER_STANDARD, MASTER_DOSSIER_STANDARD_VERSION } from './contract.js';
import type { AiCostSummary, AiModelCostSummary, AiNarrative, DiscoveryModel } from './types.js';

type TokenPricing = {
  inputPerMillion: number;
  cachedInputPerMillion: number;
  outputPerMillion: number;
};

const DEFAULT_OPENAI_PRICING: Record<string, TokenPricing> = {
  'gpt-5.5': { inputPerMillion: 5, cachedInputPerMillion: 0.5, outputPerMillion: 30 },
  'gpt-5.4-mini': { inputPerMillion: 0.75, cachedInputPerMillion: 0.075, outputPerMillion: 4.5 },
  'gpt-5.4': { inputPerMillion: 2.5, cachedInputPerMillion: 0.25, outputPerMillion: 15 },
  'gpt-5.2': { inputPerMillion: 1.75, cachedInputPerMillion: 0.175, outputPerMillion: 14 },
  'gpt-5.1': { inputPerMillion: 1.25, cachedInputPerMillion: 0.125, outputPerMillion: 10 },
  'gpt-5-mini': { inputPerMillion: 0.25, cachedInputPerMillion: 0.025, outputPerMillion: 2 },
  'gpt-5-nano': { inputPerMillion: 0.05, cachedInputPerMillion: 0.005, outputPerMillion: 0.4 },
  'gpt-5': { inputPerMillion: 1.25, cachedInputPerMillion: 0.125, outputPerMillion: 10 },
  'gpt-4.1-mini': { inputPerMillion: 0.4, cachedInputPerMillion: 0.1, outputPerMillion: 1.6 },
  'gpt-4.1-nano': { inputPerMillion: 0.1, cachedInputPerMillion: 0.025, outputPerMillion: 0.4 },
  'gpt-4.1': { inputPerMillion: 2, cachedInputPerMillion: 0.5, outputPerMillion: 8 },
  'gpt-4o-mini': { inputPerMillion: 0.15, cachedInputPerMillion: 0.075, outputPerMillion: 0.6 },
  'gpt-4o': { inputPerMillion: 2.5, cachedInputPerMillion: 1.25, outputPerMillion: 10 },
};

export async function enrichWithOpenAI(model: DiscoveryModel): Promise<void> {
  const apiKey = process.env.OPENAI_API_KEY;
  const modelName = process.env.OPENAI_MODEL ?? 'gpt-5.5';

  if (!apiKey) {
    model.aiNarrative = {
      enabled: false,
      model: modelName,
      limitations: ['OPENAI_API_KEY is not configured. Deterministic dossier generation completed without AI synthesis.'],
      cost: emptyCostSummary(modelName, 'OPENAI_API_KEY not configured; no OpenAI token usage recorded.'),
    };
    model.limitations.push('OpenAI synthesis was skipped because OPENAI_API_KEY is not configured.');
    return;
  }

  try {
    const client = new OpenAI({ apiKey });
    const summaryPayload = {
      masterDossierStandardVersion: MASTER_DOSSIER_STANDARD_VERSION,
      packageName: model.packageName,
      sourceFiles: model.sourceFiles,
      sourceTypeCounts: model.sourceTypeCounts,
      nodeCounts: summarizeBy(model.nodes, 'node_type'),
      edgeCounts: summarizeBy(model.edges, 'edge_type'),
      dataQualityFindings: model.dataQualityFindings.slice(0, 20),
      openQuestions: model.openQuestions.slice(0, 20),
      actions: model.actions.slice(0, 20),
      financialExposure: model.financialExposure,
      limitations: model.limitations,
      blockedSources: model.blockedSources,
      highLevelFilePurposeHint: createFilePurposeHint(model),
      dataElements: model.dataElements.slice(0, 120),
      accessMacroActions: model.access.Access_Macro_Action_Sequence.slice(0, 50),
      accessVbaModules: model.access.Access_Module_VBA_Register.slice(0, 50),
      excelVbaProcedures: model.excel.Excel_VBA_Register.slice(0, 80),
      excelPowerQueryArtifacts: model.excel.Excel_PowerQuery_Register.slice(0, 50),
      textSignals: {
        documents: model.nodes
          .filter((node) => node.node_type === 'document')
          .slice(0, 20)
          .map((node) => ({ nodeId: node.node_id, name: node.name, description: node.description, evidenceId: node.evidence_id })),
        sections: model.nodes
          .filter((node) => node.node_type === 'document section')
          .slice(0, 30)
          .map((node) => ({ nodeId: node.node_id, name: node.name, description: node.description, confidence: node.confidence })),
        keyValueSignals: model.dataElements
          .filter((element) => element.asset.endsWith('.txt'))
          .slice(0, 40)
          .map((element) => ({ field: element.field_name, type: element.inferred_type, sample: element.sample_values, sensitivity: element.sensitive_indicator })),
        processSteps: model.processSteps.slice(0, 20),
      },
    };

    const response = await client.responses.create({
      model: modelName,
      instructions: MASTER_DOSSIER_STANDARD,
      input: [
        {
          role: 'user',
          content: [
            {
              type: 'input_text',
              text:
                'Synthesize leadership-safe narrative from this deterministic canonical discovery model. The first and most important field is filePurpose: a 30,000-foot plain-English statement of what the uploaded file is and what business/process function it appears to serve, using evidence first and marking inference plainly. For flat files, use the column names, sample values, data elements, and semantic profile; if the evidence shows roster/person/team/contact/role fields, plainly describe it as a roster or people/team reference dataset and explain the likely operational use. Do not invent evidence, owners, finance-certified values, inaccessible lineage, or unsupported claims. Return concise JSON only.',
            },
            {
              type: 'input_text',
              text: JSON.stringify(summaryPayload),
            },
          ],
        },
      ],
      text: {
        format: {
          type: 'json_schema',
          name: 'dossier_narrative',
          schema: {
            type: 'object',
            additionalProperties: false,
            properties: {
              filePurpose: { type: 'string' },
              executiveSummary: { type: 'string' },
              architectureSummary: { type: 'string' },
              recommendedPath: { type: 'string' },
              limitations: {
                type: 'array',
                items: { type: 'string' },
              },
            },
            required: ['filePurpose', 'executiveSummary', 'architectureSummary', 'recommendedPath', 'limitations'],
          },
          strict: true,
        },
      },
    });

    const parsed = JSON.parse(response.output_text) as Pick<
      AiNarrative,
      'filePurpose' | 'executiveSummary' | 'architectureSummary' | 'recommendedPath' | 'limitations'
    >;

    model.aiNarrative = {
      enabled: true,
      model: modelName,
      ...parsed,
      cost: createCostSummary(modelName, response.usage),
    };
  } catch (error) {
    const message = sanitizeOpenAiError(error instanceof Error ? error.message : 'unknown OpenAI error');
    model.aiNarrative = {
      enabled: false,
      model: modelName,
      error: message,
      limitations: [`OpenAI synthesis failed: ${message}`],
      cost: emptyCostSummary(modelName, `OpenAI synthesis failed before billable usage could be recorded: ${message}`),
    };
    model.limitations.push(`OpenAI synthesis failed: ${message}`);
  }
}

function createFilePurposeHint(model: DiscoveryModel): Record<string, unknown> {
  return {
    sourceProcessName: model.sourceProcessName,
    sourceFiles: model.sourceFiles.map((source) => ({
      fileName: source.file_name,
      fileType: source.file_type,
      sizeBytes: source.file_size_bytes,
    })),
    sourceTypeCounts: model.sourceTypeCounts,
    nodeCounts: summarizeBy(model.nodes, 'node_type'),
    flatFileSignals: {
      rowCountSummary: model.rowCountSummary,
      dataElements: model.dataElements.slice(0, 80).map((element) => ({
        asset: element.asset,
        field: element.field_name,
        type: element.inferred_type,
        samples: element.sample_values,
        sensitivity: element.sensitive_indicator,
      })),
      tableNodes: model.nodes
        .filter((node) => node.node_type === 'table' && model.sourceFiles.some((source) => source.file_type === 'flat-file' && node.source_file === source.file_name))
        .map((node) => ({
          name: node.name,
          description: node.description,
          businessPurpose: node.business_purpose,
          confidence: node.confidence,
          criticality: node.criticality,
          recommendedAction: node.recommended_action,
        })),
    },
    workbookSignals: {
      sheets: model.excel.Excel_Sheet_Inventory.length,
      tablesAndNamedRanges: model.excel.Excel_Table_NamedRange_Register.length,
      formulaAreas: model.excel.Excel_Formula_Register.length,
      powerQueries: model.excel.Excel_PowerQuery_Register.length,
      vbaProcedures: model.excel.Excel_VBA_Register.length,
      vbaProcedureNames: model.excel.Excel_VBA_Register.slice(0, 20).map((row) => row.procedure_name ?? row.name),
    },
    accessSignals: {
      tables: model.access.Access_Table_Register.length,
      linkedTables: model.access.Access_Linked_Table_Register.length,
      queries: model.access.Access_Query_Register.length,
      macros: model.access.Access_Macro_Register.length,
      modules: model.access.Access_Module_VBA_Register.length,
      formsReports: model.access.Access_Form_Report_Register.length,
    },
    outputs: model.nodes
      .filter((node) => node.node_type === 'output')
      .slice(0, 10)
      .map((node) => ({ nodeId: node.node_id, name: node.name, description: node.description })),
  };
}

function sanitizeOpenAiError(message: string): string {
  return message.replace(/sk-[^\s"'`]+/g, 'sk-***');
}

function summarizeBy<T extends Record<string, unknown>>(rows: T[], key: keyof T): Record<string, number> {
  return rows.reduce<Record<string, number>>((counts, row) => {
    const value = String(row[key] ?? 'unknown');
    counts[value] = (counts[value] ?? 0) + 1;
    return counts;
  }, {});
}

function createCostSummary(modelName: string, usage: unknown): AiCostSummary {
  const usageRecord = (usage ?? {}) as Record<string, any>;
  const inputTokens = numberValue(usageRecord.input_tokens);
  const cachedInputTokens = Math.min(
    inputTokens,
    numberValue(
      usageRecord.input_tokens_details?.cached_tokens ??
        usageRecord.input_token_details?.cached_tokens ??
        usageRecord.prompt_tokens_details?.cached_tokens,
    ),
  );
  const outputTokens = numberValue(usageRecord.output_tokens);
  const reasoningTokens = numberValue(
    usageRecord.output_tokens_details?.reasoning_tokens ?? usageRecord.completion_tokens_details?.reasoning_tokens,
  );
  const totalTokens = numberValue(usageRecord.total_tokens) || inputTokens + outputTokens;
  const billableInputTokens = Math.max(0, inputTokens - cachedInputTokens);
  const pricing = resolvePricing(modelName);
  const inputCostUsd = costForTokens(billableInputTokens, pricing.pricing.inputPerMillion);
  const cachedInputCostUsd = costForTokens(cachedInputTokens, pricing.pricing.cachedInputPerMillion);
  const outputCostUsd = costForTokens(outputTokens, pricing.pricing.outputPerMillion);
  const cacheSavingsUsd = costForTokens(
    cachedInputTokens,
    Math.max(0, pricing.pricing.inputPerMillion - pricing.pricing.cachedInputPerMillion),
  );
  const modelSummary: AiModelCostSummary = {
    model: modelName,
    requests: usage ? 1 : 0,
    inputTokens,
    cachedInputTokens,
    billableInputTokens,
    outputTokens,
    reasoningTokens,
    totalTokens,
    inputCostUsd,
    cachedInputCostUsd,
    outputCostUsd,
    totalCostUsd: inputCostUsd + cachedInputCostUsd + outputCostUsd,
    cacheSavingsUsd,
    cacheHitRate: inputTokens ? cachedInputTokens / inputTokens : 0,
    pricing: {
      ...pricing.pricing,
      pricingAvailable: pricing.available,
    },
  };

  return {
    pricingSource: pricing.source,
    totalRequests: modelSummary.requests,
    totalInputTokens: inputTokens,
    totalCachedInputTokens: cachedInputTokens,
    totalBillableInputTokens: billableInputTokens,
    totalOutputTokens: outputTokens,
    totalReasoningTokens: reasoningTokens,
    totalTokens,
    estimatedTotalCostUsd: modelSummary.totalCostUsd,
    estimatedCacheSavingsUsd: cacheSavingsUsd,
    cacheHitRate: modelSummary.cacheHitRate,
    optimizationNote:
      cachedInputTokens > 0
        ? `OpenAI prompt caching reused ${cachedInputTokens.toLocaleString()} input token(s), reducing estimated spend by ${formatUsd(cacheSavingsUsd)} for this run.`
        : 'No cached input tokens were reported for this run. Deterministic extraction still limits spend by sending only the canonical discovery summary to AI.',
    models: [modelSummary],
  };
}

function emptyCostSummary(modelName: string, optimizationNote: string): AiCostSummary {
  const pricing = resolvePricing(modelName);
  return {
    pricingSource: pricing.source,
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
    optimizationNote,
    models: [
      {
        model: modelName,
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
          ...pricing.pricing,
          pricingAvailable: pricing.available,
        },
      },
    ],
  };
}

function resolvePricing(modelName: string): { pricing: TokenPricing; available: boolean; source: string } {
  const pricingTable = loadPricingTable();
  const normalized = modelName.toLowerCase();
  const exact = pricingTable[normalized];
  if (exact) {
    return {
      pricing: exact,
      available: true,
      source: 'OpenAI API pricing defaults reviewed 2026-05-08; override with OPENAI_PRICING_JSON for account-specific rates.',
    };
  }

  const prefix = Object.keys(pricingTable)
    .sort((left, right) => right.length - left.length)
    .find((candidate) => normalized.startsWith(candidate));
  if (prefix) {
    return {
      pricing: pricingTable[prefix],
      available: true,
      source: `OpenAI API pricing defaults reviewed 2026-05-08; matched ${prefix}. Override with OPENAI_PRICING_JSON for account-specific rates.`,
    };
  }

  return {
    pricing: { inputPerMillion: 0, cachedInputPerMillion: 0, outputPerMillion: 0 },
    available: false,
    source: `No pricing table entry for ${modelName}. Set OPENAI_PRICING_JSON to estimate spend for this model.`,
  };
}

function loadPricingTable(): Record<string, TokenPricing> {
  const configured = process.env.OPENAI_PRICING_JSON;
  if (!configured) {
    return DEFAULT_OPENAI_PRICING;
  }
  try {
    const parsed = JSON.parse(configured) as Record<string, TokenPricing>;
    return {
      ...DEFAULT_OPENAI_PRICING,
      ...Object.fromEntries(Object.entries(parsed).map(([model, pricing]) => [model.toLowerCase(), pricing])),
    };
  } catch {
    return DEFAULT_OPENAI_PRICING;
  }
}

function numberValue(value: unknown): number {
  const number = Number(value ?? 0);
  return Number.isFinite(number) && number > 0 ? number : 0;
}

function costForTokens(tokens: number, pricePerMillion: number): number {
  return (tokens / 1_000_000) * pricePerMillion;
}

function formatUsd(value: number): string {
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 4 }).format(value);
}
