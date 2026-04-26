export type DiscoveryItemType =
  | 'system'
  | 'database'
  | 'file'
  | 'workbook'
  | 'sheet'
  | 'table'
  | 'query'
  | 'macro'
  | 'form'
  | 'report'
  | 'document'
  | 'process step'
  | 'data element'
  | 'output'
  | 'control'
  | 'person / role';

export type RelationshipType =
  | 'reads_from'
  | 'writes_to'
  | 'transforms'
  | 'refreshes'
  | 'triggers'
  | 'approves'
  | 'sends'
  | 'depends_on'
  | 'documented_by'
  | 'manually_keys'
  | 'exports_to'
  | 'imports_from';

export type Criticality = 'critical' | 'high' | 'medium' | 'low';
export type LineageStatus = 'confirmed' | 'inferred' | 'blocked' | 'obsolete' | 'duplicate' | 'partial';
export type ActionMode = 'retire' | 'stabilize' | 'govern' | 'rebuild' | 'automate' | 'migrate' | 'leave as-is temporarily';

export interface EvidenceRef {
  id: string;
  type: 'SQL' | 'PowerQuery_M' | 'VBA' | 'Screenshot' | 'Interview_Note' | 'Document_Extract' | 'Profile' | 'Control_Log';
  location: string;
  description: string;
}

export interface DollarExposure {
  low: number;
  base: number;
  high: number;
  assumptions: string;
}

export interface RecommendedAction {
  mode: ActionMode;
  summary: string;
  owner: string;
  priority: 'P0' | 'P1' | 'P2' | 'P3';
  acceptanceCriteria: string;
}

export interface DiscoveryRelationship {
  id: string;
  fromId: string;
  toId: string;
  type: RelationshipType;
  automated: boolean;
  cadence: string;
  confidence: number;
  transformId?: string;
  evidenceId: string;
}

export interface DiscoveryItem {
  id: string;
  type: DiscoveryItemType;
  name: string;
  businessPurpose: string;
  owner: string;
  evidence: EvidenceRef[];
  confidence: number;
  criticality: Criticality;
  upstream: string[];
  downstream: string[];
  failureImpact: string;
  dollarExposure: DollarExposure;
  recommendedAction: RecommendedAction;
  status: LineageStatus;
  tags: string[];
}

export interface ArtifactPackageItem {
  id: string;
  name: string;
  audience: string;
  purpose: string;
  progress: number;
  sourceModel: 'canonical graph';
}

export interface ConfidenceArea {
  area: string;
  reviewed: string;
  coverage: number;
  confidence: number;
  blocked: string;
}

export interface FailureRisk {
  id: string;
  scenario: string;
  impactedOutput: string;
  detection: string;
  recovery: string;
  exposure: DollarExposure;
  confidence: number;
}

export interface BacklogAction {
  actionId: string;
  title: string;
  owner: string;
  priority: 'P0' | 'P1' | 'P2' | 'P3';
  dependency: string;
  dueDate: string;
  acceptanceCriteria: string;
  linkedItemId: string;
  mode: ActionMode;
}

export interface ExtractionCapability {
  source: 'Access' | 'Excel' | 'Word' | 'Database' | 'Interview';
  autoExtracts: string[];
  currentStateOutputs: string[];
  readiness: number;
}

export interface DiscoveryModel {
  packageName: string;
  processName: string;
  businessFunction: string;
  recommendation: string;
  decisionRequired: string;
  systemsInScope: string[];
  criticalOutputs: string[];
  overallRiskRating: 'severe' | 'high' | 'medium' | 'low';
  estimatedDollarExposure: DollarExposure;
  items: DiscoveryItem[];
  relationships: DiscoveryRelationship[];
  artifacts: ArtifactPackageItem[];
  confidenceAreas: ConfidenceArea[];
  failureRisks: FailureRisk[];
  backlog: BacklogAction[];
  extractionCapabilities: ExtractionCapability[];
}
