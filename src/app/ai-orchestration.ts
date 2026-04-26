export const DISTILLERY_ENGINE_LABEL = 'The Distillery';

export interface DiscoveryAgentRequest {
  sourceKind: 'access' | 'excel' | 'word' | 'database' | 'interview' | 'mixed';
  sourceName: string;
  extractedText: string;
  knownArtifacts: string[];
}

export interface DiscoveryAgentContract {
  model: typeof DISTILLERY_ENGINE_LABEL;
  reasoning: {
    effort: 'high' | 'xhigh';
  };
  verbosity: 'high';
  orchestration: 'specialist-pass';
  responseFormat: 'canonical_discovery_model_delta';
  serverSideOnly: true;
  requiredOutputs: string[];
}

export const DISCOVERY_AGENT_CONTRACT: DiscoveryAgentContract = {
  model: DISTILLERY_ENGINE_LABEL,
  reasoning: {
    effort: 'high'
  },
  verbosity: 'high',
  orchestration: 'specialist-pass',
  responseFormat: 'canonical_discovery_model_delta',
  serverSideOnly: true,
  requiredOutputs: [
    'complete canonical node-edge model',
    'auto-documentation current state',
    'recursive lineage to terminal nodes',
    'logic, controls, and failure modes',
    'financial exposure model',
    'engineer-ready action backlog',
    'diagram-ready visual graph data'
  ]
};

export const discoveryAgentInstructions = `
You are Uncle Kev's Distillery discovery analyst. Distill source evidence into a canonical discovery model delta through specialist passes:
canonical model, current-state documentation, recursive lineage, logic and controls, business exposure and backlog, and diagram/output package.

Every discovered item must include id, type, businessPurpose, owner, evidence, confidence,
criticality, upstream, downstream, failureImpact, dollarExposure, and recommendedAction.

If a finding lacks evidence, confidence, or a next action, mark it unfinished and ask for the
smallest specific artifact needed to finish it. Prefer concise executive synthesis on top of
complete technical depth. Trace upstream recursively until a terminal condition is reached.
`;
