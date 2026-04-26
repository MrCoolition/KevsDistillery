export const OPENAI_DISCOVERY_MODEL = 'gpt-5.5';

export const OPENAI_RESPONSES_ENDPOINT = 'https://api.openai.com/v1/responses';

export interface DiscoveryAgentRequest {
  sourceKind: 'access' | 'excel' | 'word' | 'database' | 'interview' | 'mixed';
  sourceName: string;
  extractedText: string;
  knownArtifacts: string[];
}

export interface DiscoveryAgentContract {
  model: typeof OPENAI_DISCOVERY_MODEL;
  reasoning: {
    effort: 'high' | 'xhigh';
  };
  responseFormat: 'canonical_discovery_model_delta';
  serverSideOnly: true;
  requiredOutputs: string[];
}

export const DISCOVERY_AGENT_CONTRACT: DiscoveryAgentContract = {
  model: OPENAI_DISCOVERY_MODEL,
  reasoning: {
    effort: 'high'
  },
  responseFormat: 'canonical_discovery_model_delta',
  serverSideOnly: true,
  requiredOutputs: [
    'evidence-backed canonical fields',
    'recursive lineage candidates',
    'business logic extraction',
    'business impact assumptions',
    'actionable remediation backlog'
  ]
};

export const discoveryAgentInstructions = `
You are THE DISTILLERY discovery analyst. Convert source evidence into a canonical discovery model delta.

Every discovered item must include id, type, businessPurpose, owner, evidence, confidence,
criticality, upstream, downstream, failureImpact, dollarExposure, and recommendedAction.

If a finding lacks evidence, confidence, or a next action, mark it unfinished and ask for the
smallest specific artifact needed to finish it. Prefer concise executive synthesis on top of
complete technical depth. Trace upstream recursively until a terminal condition is reached.
`;
