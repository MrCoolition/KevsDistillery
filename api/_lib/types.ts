import type { EDGE_TYPES, NODE_TYPES } from './contract.js';

export type SourceType = 'access' | 'excel' | 'word' | 'flat-file' | 'sql' | 'script' | 'unknown';
export type NodeType = (typeof NODE_TYPES)[number];
export type EdgeType = (typeof EDGE_TYPES)[number];
export type Confidence = 'confirmed' | 'inferred' | 'blocked' | 'unknown' | 'partial' | 'owner-confirmation-required';
export type Criticality = 'P0' | 'P1' | 'P2' | 'P3' | 'unknown';

export type UploadedSource = {
  originalName: string;
  mimeType?: string;
  size: number;
  buffer: Buffer;
  tempPath?: string;
  sourcePath?: string;
  discoveredFrom?: string;
};

export type SourceFileMeta = {
  source_id: string;
  file_name: string;
  file_type: SourceType;
  extension: string;
  file_size_bytes: number;
  sha256: string;
  evidence_id: string;
};

export type DiscoveryNode = {
  node_id: string;
  node_type: NodeType;
  name: string;
  description: string;
  source_file: string;
  business_purpose: string;
  owner_status: string;
  criticality: Criticality;
  confidence: Confidence;
  evidence_id: string;
  recommended_action: string;
  failure_impact: string;
  dollar_exposure: string;
};

export type DiscoveryEdge = {
  edge_id: string;
  from_node_id: string;
  to_node_id: string;
  edge_type: EdgeType;
  description: string;
  automated_flag: 'automated' | 'manual' | 'unknown';
  transformation_id: string;
  cadence: string;
  confidence: Confidence;
  evidence_id: string;
};

export type EvidenceItem = {
  evidence_id: string;
  title: string;
  category: string;
  relative_path: string;
  summary: string;
  source_file: string;
  confidence: Confidence;
  content: string | Buffer;
};

export type ProcessStep = {
  process_step_id: string;
  step_name: string;
  actor_or_role: string;
  trigger: string;
  description: string;
  manual_or_automated: 'manual' | 'automated' | 'mixed' | 'unknown';
  input_node_id: string;
  output_node_id: string;
  evidence_id: string;
  confidence: Confidence;
  recommended_action: string;
};

export type TransformationRule = {
  transformation_id: string;
  source_asset: string;
  rule_type: string;
  rule_description: string;
  input_fields: string;
  output_fields: string;
  evidence_id: string;
  confidence: Confidence;
  recommended_action: string;
};

export type DataElement = {
  data_element_id: string;
  asset: string;
  field_name: string;
  inferred_type: string;
  null_count: number;
  sample_values: string;
  sensitive_indicator: string;
  evidence_id: string;
  confidence: Confidence;
  recommended_action: string;
};

export type DataQualityFinding = {
  finding_id: string;
  asset: string;
  field: string;
  issue: string;
  example: string;
  severity: Criticality;
  business_impact: string;
  recommended_fix: string;
  evidence_id: string;
  confidence: Confidence;
};

export type ControlException = {
  control_id: string;
  asset: string;
  control_type: string;
  description: string;
  status: string;
  evidence_id: string;
  confidence: Confidence;
  recommended_action: string;
};

export type SecurityAccessFinding = {
  security_id: string;
  asset: string;
  concern: string;
  status: string;
  impact: string;
  evidence_id: string;
  confidence: Confidence;
  recommended_action: string;
};

export type OpenQuestion = {
  question_id: string;
  asset: string;
  question: string;
  owner_role: string;
  blocker_type: string;
  priority: Criticality;
  evidence_id: string;
  status: string;
};

export type ActionItem = {
  action_id: string;
  title: string;
  description: string;
  source_asset: string;
  owner_role: string;
  recommended_owner: string;
  action_type: string;
  priority: Criticality;
  severity: Criticality;
  dependency: string;
  due_date_or_phase: string;
  acceptance_criteria: string;
  evidence_id: string;
  related_risk: string;
  expected_business_value: string;
  status: string;
};

export type FinancialExposure = {
  process_or_output: string;
  failure_scenario: string;
  frequency: string;
  units_affected: number;
  dollar_per_unit: number;
  revenue_at_risk: number;
  margin_percent: number;
  margin_at_risk: number;
  rework_hours: number;
  labor_rate: number;
  labor_recovery_cost: number;
  customer_sla_exposure: number;
  compliance_exposure: number;
  cash_timing_cost: number;
  low_impact: number;
  base_impact: number;
  high_impact: number;
  annualized_low: number;
  annualized_base: number;
  annualized_high: number;
  confidence: Confidence;
  assumptions: string;
  evidence_id: string;
  finance_validation_needed: string;
};

export type QaRecord = {
  qa_id: string;
  check: string;
  status: 'PASS' | 'PASS_WITH_LIMITATION' | 'FAIL';
  evidence_id: string;
  notes: string;
};

export type AiNarrative = {
  enabled: boolean;
  model: string;
  filePurpose?: string;
  executiveSummary?: string;
  architectureSummary?: string;
  recommendedPath?: string;
  limitations?: string[];
  error?: string;
  cost?: AiCostSummary;
};

export type AiCostSummary = {
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
  models: AiModelCostSummary[];
};

export type AiModelCostSummary = {
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
};

export type DiscoveryModel = {
  runId: string;
  packageName: string;
  sourceProcessName: string;
  generatedDate: string;
  analysisVersion: string;
  packageVersion: string;
  sourceFiles: SourceFileMeta[];
  sourceTypeCounts: Record<SourceType, number>;
  nodes: DiscoveryNode[];
  edges: DiscoveryEdge[];
  evidence: EvidenceItem[];
  processSteps: ProcessStep[];
  transformations: TransformationRule[];
  dataElements: DataElement[];
  dataQualityFindings: DataQualityFinding[];
  controls: ControlException[];
  securityAccess: SecurityAccessFinding[];
  openQuestions: OpenQuestion[];
  actions: ActionItem[];
  financialExposure: FinancialExposure[];
  qaRecords: QaRecord[];
  access: Record<string, Record<string, unknown>[]>;
  excel: Record<string, Record<string, unknown>[]>;
  word: Record<string, Record<string, unknown>[]>;
  dependencyUsage: Record<string, unknown>[];
  modernization: Record<string, unknown>[];
  scheduleSla: Record<string, unknown>[];
  failureModes: Record<string, unknown>[];
  rowCountSummary: Record<string, number>;
  aiNarrative: AiNarrative;
  limitations: string[];
  blockedSources: string[];
  assumptions: string[];
};
