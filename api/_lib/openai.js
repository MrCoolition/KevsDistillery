const { createHash } = require('node:crypto');

const model = process.env.OPENAI_MODEL || 'gpt-5.5';
const reasoningEffort = process.env.OPENAI_REASONING_EFFORT || 'high';
const responseVerbosity = process.env.OPENAI_VERBOSITY || 'high';
const OPENAI_TIMEOUT_MS = Number(process.env.OPENAI_TIMEOUT_MS || 59000);
const OPENAI_START_TIMEOUT_MS = Number(process.env.OPENAI_START_TIMEOUT_MS || 25000);
const OPENAI_MAX_OUTPUT_TOKENS = Number(process.env.OPENAI_MAX_OUTPUT_TOKENS || 12000);
const OPENAI_SPECIALIST_OUTPUT_TOKENS = Number(process.env.OPENAI_SPECIALIST_OUTPUT_TOKENS || 6500);
const DISTILLERY_SINGLE_PASS = process.env.DISTILLERY_SINGLE_PASS === 'true';
const PENDING_RESPONSE_STATUSES = new Set(['queued', 'in_progress']);

const REPORT_SECTION_TITLES = [
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

const REQUIRED_ARTIFACTS = [
  {
    id: '01',
    name: '01_Executive_Decision_Brief.pdf',
    type: 'report',
    audience: 'Leadership',
    purpose: 'Decision summary covering mission, risk, dollar exposure, top blockers, and approvals needed.'
  },
  {
    id: '02',
    name: '02_Current_State_Architecture_Report.pdf',
    type: 'report',
    audience: 'Business and IT',
    purpose: 'Evidence-backed operating model, source landscape, lineage, controls, failure modes, and migration recommendation.'
  },
  {
    id: '03',
    name: '03_Technical_Discovery_Workbook.xlsx',
    type: 'workbook',
    audience: 'Engineers and analysts',
    purpose: 'Structured inventory of artifacts, objects, process steps, data elements, transformations, lineage, risks, and actions.'
  },
  {
    id: '04',
    name: '04_Auto_Documentation_Pack',
    type: 'machine_documentation',
    audience: 'Automation and governance',
    purpose: 'Machine-readable generated current-state documentation from the canonical discovery model.'
  },
  {
    id: '05',
    name: '05_Diagram_Pack',
    type: 'diagram_pack',
    audience: 'All audiences',
    purpose: 'Visual process flow, data flow, lineage, object dependencies, controls, failure impact, and schedule timing.'
  },
  {
    id: '06',
    name: '06_Financial_Impact_Model.xlsx',
    type: 'workbook',
    audience: 'Leadership and finance',
    purpose: 'Low/base/high exposure model for no-run, late-run, wrong-data, partial-run, and unauditable-run scenarios.'
  },
  {
    id: '07',
    name: '07_Action_Backlog.csv',
    type: 'backlog',
    audience: 'Delivery teams',
    purpose: 'Prioritized remediation, stabilization, governance, extraction, migration, and retirement actions.'
  },
  {
    id: '08',
    name: '08_Evidence_Archive',
    type: 'evidence',
    audience: 'Audit and project team',
    purpose: 'Traceable proof for every finding, blocker, assumption, extracted object, and recommendation.'
  },
  {
    id: '09',
    name: '09_Metadata_Manifest.json',
    type: 'manifest',
    audience: 'Automation',
    purpose: 'Machine-readable package manifest and canonical model summary.'
  }
];

const SPECIALIST_PASSES = [
  {
    key: 'canonical',
    title: 'Canonical Discovery Architect',
    maxOutputTokens: OPENAI_MAX_OUTPUT_TOKENS,
    outputKeys: ['processName', 'businessFunction', 'recommendation', 'decisionRequired', 'systemsInScope', 'criticalOutputs', 'overallRiskRating', 'estimatedDollarExposure', 'executiveBrief', 'reportSections', 'items', 'relationships', 'artifacts', 'backlog', 'evidenceIndex', 'lineageNodes', 'lineageEdges', 'failureRisks', 'openQuestions'],
    reportSectionFocus: REPORT_SECTION_TITLES,
    focus: [
      'Build the strongest possible canonical discovery model from all evidence.',
      'Identify every material source, object, process step, output, control, blocker, and engineering action that evidence supports.',
      'Normalize duplicate names, assign stable IDs, and keep all downstream outputs tied to one proof graph.'
    ]
  },
  {
    key: 'documentation',
    title: 'Current-State Auto Documentation Lead',
    maxOutputTokens: OPENAI_SPECIALIST_OUTPUT_TOKENS,
    outputKeys: ['processName', 'businessFunction', 'systemsInScope', 'criticalOutputs', 'reportSections', 'peopleRoles', 'processSteps', 'accessObjects', 'excelObjects', 'wordExtracts', 'evidenceIndex', 'openQuestions'],
    reportSectionFocus: [
      'Scope, Coverage, and Confidence',
      'Business Mission of the Process',
      'Current-State Operating Model',
      'System and Artifact Landscape',
      'Data Flow and Process Flow Summary'
    ],
    focus: [
      'Generate auto-documentation of the current state from evidence, not generic commentary.',
      'Populate processSteps, peopleRoles, systemsInScope, artifacts, reportSections, and openQuestions.',
      'Write operating-model narrative as trigger, actor, input, step, validation, output, handoff, exception path, SLA, and owner.'
    ]
  },
  {
    key: 'lineage',
    title: 'Recursive Lineage Principal',
    maxOutputTokens: OPENAI_SPECIALIST_OUTPUT_TOKENS,
    outputKeys: ['criticalOutputs', 'items', 'relationships', 'lineageNodes', 'lineageEdges', 'reportSections', 'evidenceIndex', 'openQuestions'],
    reportSectionFocus: [
      'Data Flow and Process Flow Summary',
      'Recursive Lineage and Source-of-Truth Assessment'
    ],
    focus: [
      'Trace every critical output upstream recursively until a terminal condition or explicit blocker is reached.',
      'Populate lineageNodes, lineageEdges, relationships, source-of-truth candidates, unresolved nodes, branch statuses, and confidence.',
      'Classify terminal nodes as authoritative system of record, third-party source, manual entry, blocked, obsolete, duplicate, or approved stopping point.'
    ]
  },
  {
    key: 'logic_controls',
    title: 'Logic, Controls, and Failure-Mode Examiner',
    maxOutputTokens: OPENAI_SPECIALIST_OUTPUT_TOKENS,
    outputKeys: ['items', 'relationships', 'transformationsRules', 'controlsExceptions', 'dataQuality', 'securityAccess', 'scheduleSla', 'failureRisks', 'failureModes', 'reportSections', 'evidenceIndex', 'openQuestions'],
    reportSectionFocus: [
      'Transformations and Business Logic',
      'Controls, Exceptions, and Failure Modes'
    ],
    focus: [
      'Extract Access, Excel, Word, SQL, VBA, Power Query, formula, macro, manual override, and tribal-rule evidence.',
      'Populate transformationsRules, controlsExceptions, failureRisks, dataQuality, securityAccess, and scheduleSla.',
      'Describe how failures are detected, who recovers, what outputs break, and what native source export would prove the logic.'
    ]
  },
  {
    key: 'impact_backlog',
    title: 'Business Exposure and Action Backlog Strategist',
    maxOutputTokens: OPENAI_SPECIALIST_OUTPUT_TOKENS,
    outputKeys: ['overallRiskRating', 'estimatedDollarExposure', 'financialModel', 'failureRisks', 'backlog', 'recommendation', 'decisionRequired', 'reportSections', 'evidenceIndex', 'openQuestions'],
    reportSectionFocus: [
      'Executive Snapshot',
      'Financial Impact and Business Exposure',
      'Recommendations and Action Plan',
      'Open Questions and Decisions Needed'
    ],
    focus: [
      'Build a high-level financial exposure model for no-run, late-run, wrong-data, partial-run, and unauditable-run scenarios.',
      'Populate financialModel, estimatedDollarExposure, failureRisks, backlog, recommendations, decisions, and acceptance criteria.',
      'When dollar inputs are missing, state exact missing business inputs and create actions to collect them.'
    ]
  },
  {
    key: 'diagram_pack',
    title: 'Diagram and Output Package Architect',
    maxOutputTokens: OPENAI_SPECIALIST_OUTPUT_TOKENS,
    outputKeys: ['artifacts', 'items', 'relationships', 'lineageNodes', 'lineageEdges', 'controlsExceptions', 'failureRisks', 'scheduleSla', 'reportSections', 'evidenceIndex'],
    reportSectionFocus: [
      'System and Artifact Landscape',
      'Data Flow and Process Flow Summary',
      'Recursive Lineage and Source-of-Truth Assessment',
      'Controls, Exceptions, and Failure Modes'
    ],
    focus: [
      'Prepare diagram-ready nodes and edges for executive value stream, system context, swimlane, data flow, recursive lineage, dependency, controls, failure, and timeline views.',
      'Ensure every visual node maps back to a canonical ID and includes owner, cadence, manual/automated, criticality, confidence, and dollar sensitivity when known.',
      'Populate artifacts, relationships, lineageNodes, lineageEdges, controls, failureRisks, and schedule/timing details that the UI can render into diagrams.'
    ]
  }
];

function buildInstructions(sourceKind, sourceName, pass = null) {
  const isCanonical = !pass || pass.key === 'canonical';
  return [
    'You are Uncle Kev\'s Distillery principal data discovery analyst.',
    `Source kind: ${sourceKind}. Batch name: ${sourceName}.`,
    'Return one valid JSON object only. No markdown, no prose outside JSON, no placeholder strings.',
    'This is production data discovery for engineers migrating current-state processes into Snowflake using Fivetran, dbt, SQL, and Snowpark.',
    'Analyze retrieved evidence, not just filenames. Treat filenames as weak evidence unless supported by extracted content. Never invent facts.',
    'If evidence is incomplete, create a blocker-backed finding with the exact smallest source artifact needed to finish the discovery.',
    'A finding is unfinished unless it has evidence, confidence, and next action.',
    'Use one canonical node-edge model to drive the report, auto-documentation, diagrams, recursive lineage, financial exposure, and remediation backlog.',
    isCanonical
      ? 'Required top-level keys for this pass: processName, businessFunction, recommendation, decisionRequired, systemsInScope, criticalOutputs, overallRiskRating, estimatedDollarExposure, executiveBrief, reportSections, items, relationships, artifacts, backlog, evidenceIndex, lineageNodes, lineageEdges, failureRisks, openQuestions.'
      : 'This is a focused specialist pass. Return only the requested outputKeys from the user payload plus any minimal IDs needed for traceability. Omit empty arrays and do not restate the full canonical object.',
    'Also include any supported technical arrays: peopleRoles, processSteps, accessObjects, excelObjects, wordExtracts, dataElements, transformationsRules, controlsExceptions, dataQuality, securityAccess, scheduleSla, failureModes, financialModel.',
    isCanonical
      ? 'reportSections: exactly 12 sections with the title property using these exact titles in order: Executive Snapshot; Scope, Coverage, and Confidence; Business Mission of the Process; Current-State Operating Model; System and Artifact Landscape; Data Flow and Process Flow Summary; Transformations and Business Logic; Recursive Lineage and Source-of-Truth Assessment; Controls, Exceptions, and Failure Modes; Financial Impact and Business Exposure; Recommendations and Action Plan; Open Questions and Decisions Needed. Body should be concise but specific, usually 60 to 130 words. Include confidence and evidenceIds.'
      : 'reportSections: include only the reportSectionFocus titles from the user payload. Body should be concise but specific, usually 60 to 120 words. Include confidence and evidenceIds.',
    'items: include every material canonical item supported by evidence or required as a documented blocker. Each item requires id, type, name, businessPurpose, owner, evidence, confidence, criticality, upstream, downstream, failureImpact, dollarExposure, recommendedAction, status.',
    'recommendedAction requires mode, summary, owner, priority, acceptanceCriteria. dollarExposure requires low, base, high, assumptions. evidence requires id, type, location, description.',
    'relationships: include every material discovered or inferred edge with id, fromId, toId, type, automated, cadence, transformId, evidenceId, confidence, and status.',
    'lineageNodes and lineageEdges must use stable IDs that correspond to items and relationships whenever possible.',
    'Recursive lineage rule: when an upstream source is found, treat it as a new discovery target and trace until authoritative system of record, third-party source, manual entry, blocked, obsolete, duplicate, or approved practical stopping point.',
    'Access evidence to mine: database metadata, split front/back end, linked tables, row counts, columns, keys, indexes, relationships, saved SQL, macros, VBA, forms, reports, startup objects, imports, exports, hidden objects, and downstream outputs.',
    'Excel evidence to mine: workbook metadata, hidden sheets, tables, named ranges, formula regions, Power Query M, external links, data model relationships, measures, pivots, slicers, dashboards, VBA, workbook events, refresh order, manual zones, hardcoded overrides, and outputs.',
    'Word/process evidence to mine: headings, actors, process steps, inputs, outputs, approvals, rules, exceptions, systems, data elements, controls, signoffs, deadlines, and SLAs.',
    'estimatedDollarExposure must cover revenue at risk, gross margin at risk, cash timing impact, rework labor cost, and compliance/SLA/customer exposure. Include low/base/high estimates with assumptions. If pricing evidence is missing, use zero low/base/high and state exact business inputs needed.',
    'For non-transaction processes, use proxy value fields such as spend influenced, revenue influenced, inventory value managed, payroll affected, compliance exposure, decisions delayed, and executive reporting dependency.',
    'backlog actions must be ready to work: owner, priority, dependency, due date when inferable, linked item, mode, summary, and acceptanceCriteria.',
    isCanonical || pass?.key === 'diagram_pack'
      ? 'artifacts must list all 9 Discovery_Action_Pack outputs with purpose, audience, and status: 01_Executive_Decision_Brief.pdf, 02_Current_State_Architecture_Report.pdf, 03_Technical_Discovery_Workbook.xlsx, 04_Auto_Documentation_Pack, 05_Diagram_Pack, 06_Financial_Impact_Model.xlsx, 07_Action_Backlog.csv, 08_Evidence_Archive, 09_Metadata_Manifest.json.'
      : 'Do not list all 9 artifacts in this specialist pass unless directly relevant; the merge layer owns the final package artifact list.',
    'Every critical output must have current-state narrative, diagram coverage, recursive lineage, business logic extraction, financial exposure, and a clear action recommendation or blocker.',
    'Do not use arbitrary low item caps. Be comprehensive in your focused area, deduplicate aggressively, and prefer evidence-backed depth over generic summaries.',
    'The user payload includes analysisPassTitle, specialistFocus, outputKeys, and reportSectionFocus. Obey those fields exactly to keep this pass deep and bounded.'
  ].join('\n');
}

function extractOutputText(result) {
  if (typeof result.output_text === 'string') {
    return result.output_text;
  }

  const chunks = [];
  for (const item of result.output || []) {
    for (const content of item.content || []) {
      if (content.type === 'output_text' && typeof content.text === 'string') {
        chunks.push(content.text);
      }
    }
  }

  return chunks.join('\n').trim();
}

function parseJsonOutput(outputText) {
  if (!outputText) {
    return null;
  }

  try {
    return JSON.parse(outputText);
  } catch {
    const fenced = outputText.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (fenced) {
      try {
        return JSON.parse(fenced[1]);
      } catch {
        return null;
      }
    }

    const firstBrace = outputText.indexOf('{');
    const lastBrace = outputText.lastIndexOf('}');
    if (firstBrace >= 0 && lastBrace > firstBrace) {
      try {
        return JSON.parse(outputText.slice(firstBrace, lastBrace + 1));
      } catch {
        return null;
      }
    }
  }

  return null;
}

function textOptions() {
  return {
    verbosity: responseVerbosity,
    format: {
      type: 'json_object'
    }
  };
}

function buildInputPayload(payload, pass) {
  const {
    sourceKind,
    sourceName,
    extractedText,
    knownArtifacts = [],
    targetOutputs = []
  } = payload;

  return {
    sourceKind,
    sourceName,
    knownArtifacts,
    targetOutputs,
    outputPackageStandard: REQUIRED_ARTIFACTS.map((artifact) => artifact.name),
    requiredReportSections: REPORT_SECTION_TITLES,
    extractedText,
    analysisPass: pass?.key || 'canonical',
    analysisPassTitle: pass?.title || 'Canonical Discovery Architect',
    specialistFocus: pass?.focus || [],
    outputKeys: pass?.outputKeys || [],
    reportSectionFocus: pass?.reportSectionFocus || REPORT_SECTION_TITLES
  };
}

function promptCacheKey(payload) {
  const hash = createHash('sha256')
    .update(String(payload.sourceKind || 'mixed'))
    .update('\n')
    .update(String(payload.sourceName || 'source'))
    .update('\n')
    .update(String(payload.extractedText || ''))
    .digest('hex')
    .slice(0, 32);
  return `distillery:${hash}`;
}

function buildRequestBody(payload, background = false, pass = SPECIALIST_PASSES[0]) {
  const { sourceKind, sourceName } = payload;

  return {
    model,
    reasoning: {
      effort: reasoningEffort
    },
    instructions: buildInstructions(sourceKind, sourceName, pass),
    max_output_tokens: pass?.maxOutputTokens || OPENAI_MAX_OUTPUT_TOKENS,
    store: true,
    background,
    prompt_cache_key: promptCacheKey(payload),
    prompt_cache_retention: '24h',
    text: textOptions(),
    input: [
      {
        role: 'user',
        content: [
          {
            type: 'input_text',
            text: JSON.stringify(buildInputPayload(payload, pass))
          }
        ]
      }
    ]
  };
}

async function synthesizeWithOpenAI(payload) {
  const result = await openAIJson('/responses', {
    method: 'POST',
    timeoutMs: OPENAI_TIMEOUT_MS,
    body: buildRequestBody(payload, false, SPECIALIST_PASSES[0])
  });

  const synthesis = responseToSynthesis(result, SPECIALIST_PASSES[0]);
  if (synthesis.pending) {
    const error = new Error('The Distillery started a background run. Poll for completion.');
    error.statusCode = 202;
    error.detail = synthesis;
    throw error;
  }
  return synthesis;
}

async function openAIJson(path, options = {}) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    const error = new Error('The Distillery engine is not configured.');
    error.statusCode = 500;
    throw error;
  }

  const abortController = new AbortController();
  const timeout = setTimeout(() => abortController.abort(), options.timeoutMs || OPENAI_START_TIMEOUT_MS);

  let response;
  try {
    response = await fetch(`https://api.openai.com/v1${path}`, {
      method: options.method || 'GET',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`
      },
      signal: abortController.signal,
      body: options.body ? JSON.stringify(options.body) : undefined
    });
  } catch (error) {
    if (error?.name === 'AbortError') {
      const timeoutError = new Error('The Distillery did not acknowledge the background run before the request window closed.');
      timeoutError.statusCode = 504;
      throw timeoutError;
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }

  const responseText = await response.text();
  let result;
  try {
    result = responseText ? JSON.parse(responseText) : {};
  } catch {
    const error = new Error('The Distillery returned a non-JSON response.');
    error.statusCode = 502;
    error.detail = responseText.slice(0, 1000);
    throw error;
  }

  if (!response.ok) {
    const error = new Error('The Distillery request failed.');
    error.statusCode = response.status;
    error.detail = result;
    throw error;
  }

  return result;
}

function responseToSynthesis(result, pass = null) {
  if (PENDING_RESPONSE_STATUSES.has(result.status)) {
    return {
      pending: true,
      responseId: result.id,
      responseStatus: result.status,
      model,
      passKey: pass?.key || 'canonical',
      passTitle: pass?.title || 'Canonical Discovery Architect',
      orchestration: pass ? 'specialist-pass' : 'single-pass'
    };
  }

  if (result.status === 'incomplete') {
    const outputText = extractOutputText(result);
    const canonicalDelta = parseJsonOutput(outputText);
    if (canonicalDelta && typeof canonicalDelta === 'object') {
      return {
        pending: false,
        responseId: result.id,
        responseStatus: result.status,
        model,
        passKey: pass?.key || 'canonical',
        passTitle: pass?.title || 'Canonical Discovery Architect',
        orchestration: pass ? 'specialist-pass' : 'single-pass',
        fallbackReason: incompleteReason(result, pass),
        outputText,
        canonicalDelta,
        raw: result
      };
    }

    return incompletePassSynthesis(result, pass, outputText);
  }

  if (result.status && result.status !== 'completed') {
    const error = new Error(result.error?.message || `The Distillery run ended with status ${result.status}.`);
    error.statusCode = 502;
    error.detail = result.error || result.incomplete_details || result;
    throw error;
  }

  const outputText = extractOutputText(result);
  const canonicalDelta = parseJsonOutput(outputText);
  if (!canonicalDelta || typeof canonicalDelta !== 'object') {
    const error = new Error('The Distillery did not return a canonical discovery JSON object.');
    error.statusCode = 502;
    error.detail = outputText.slice(0, 1000);
    throw error;
  }

  return {
    pending: false,
    responseId: result.id,
    responseStatus: result.status || 'completed',
    model,
    passKey: pass?.key || 'canonical',
    passTitle: pass?.title || 'Canonical Discovery Architect',
    orchestration: pass ? 'specialist-pass' : 'single-pass',
    outputText,
    canonicalDelta,
    raw: result
  };
}

function incompleteReason(result, pass = null) {
  const reason = result.incomplete_details?.reason || result.incomplete_details?.type || 'response_limit';
  const title = pass?.title || 'Distillery pass';
  return `${title} ended incomplete (${reason}) but returned parseable discovery JSON. The output was preserved and merged with the remaining specialist passes.`;
}

function incompletePassSynthesis(result, pass = null, outputText = '') {
  const passKey = pass?.key || 'canonical';
  const passTitle = pass?.title || 'Distillery pass';
  const reason = result.incomplete_details?.reason || result.incomplete_details?.type || 'response_limit';
  const evidenceId = `EV-${passKey.toUpperCase()}-INCOMPLETE`;
  const actionId = `ACT-${passKey.toUpperCase()}-RERUN`;
  const fallbackReason = `${passTitle} ended incomplete (${reason}) before it returned valid JSON. The Distillery preserved the run and created a blocker-backed action instead of failing the whole package.`;

  return {
    pending: false,
    responseId: result.id,
    responseStatus: 'incomplete',
    model,
    passKey,
    passTitle,
    orchestration: pass ? 'specialist-pass' : 'single-pass',
    fallbackReason,
    outputText: outputText || fallbackReason,
    canonicalDelta: {
      processName: 'Distillery discovery run',
      businessFunction: 'Data discovery and migration readiness',
      recommendation: 'Use the completed specialist passes and rerun the incomplete pass if more depth is needed.',
      decisionRequired: 'Confirm whether to rerun the incomplete specialist pass or proceed with available evidence.',
      systemsInScope: [],
      criticalOutputs: [],
      overallRiskRating: 'High',
      estimatedDollarExposure: {},
      executiveBrief: {},
      reportSections: [],
      items: [],
      relationships: [],
      artifacts: [],
      backlog: [
        {
          actionId,
          title: `Rerun ${passTitle}`,
          mode: 'stabilize',
          owner: 'Discovery owner',
          priority: 'P0',
          dependency: 'Focused evidence payload and bounded output keys',
          acceptanceCriteria: `${passTitle} returns valid canonical JSON with evidence, confidence, and next actions.`,
          linkedItemId: evidenceId,
          summary: fallbackReason
        }
      ],
      evidenceIndex: [
        {
          id: evidenceId,
          type: 'distillery_run_status',
          location: result.id || passKey,
          description: fallbackReason,
          relatedObject: passKey
        }
      ],
      lineageNodes: [],
      lineageEdges: [],
      failureRisks: [
        {
          id: `RISK-${passKey.toUpperCase()}-INCOMPLETE`,
          scenario: `${passTitle} did not complete`,
          trigger: reason,
          effect: 'The merged action pack may be shallower in this specialist area.',
          detection: 'Distillery response status incomplete',
          recovery: 'Rerun focused pass with narrower source evidence or exported metadata.',
          impactedOutput: passTitle,
          confidence: 100
        }
      ],
      openQuestions: [
        {
          id: `Q-${passKey.toUpperCase()}-INCOMPLETE`,
          question: `Should ${passTitle} be rerun with narrower evidence or native metadata exports?`,
          owner: 'Discovery owner',
          impactIfUnanswered: 'The Discovery Action Pack remains complete enough to use but contains a documented blocker for this specialist pass.'
        }
      ]
    },
    raw: result
  };
}

async function startBackgroundSynthesis(payload) {
  if (DISTILLERY_SINGLE_PASS) {
    const result = await openAIJson('/responses', {
      method: 'POST',
      timeoutMs: OPENAI_START_TIMEOUT_MS,
      body: buildRequestBody(payload, true, SPECIALIST_PASSES[0])
    });

    return responseToSynthesis(result, SPECIALIST_PASSES[0]);
  }

  const started = await Promise.all(SPECIALIST_PASSES.map(async (pass) => {
    const result = await openAIJson('/responses', {
      method: 'POST',
      timeoutMs: OPENAI_START_TIMEOUT_MS,
      body: buildRequestBody(payload, true, pass)
    });
    const synthesis = responseToSynthesis(result, pass);
    return {
      pass,
      synthesis
    };
  }));

  if (started.every((entry) => !entry.synthesis.pending)) {
    return mergeSpecialistSyntheses(started.map((entry) => entry.synthesis));
  }

  return {
    pending: true,
    responseIds: started.map((entry) => ({
      key: entry.pass.key,
      title: entry.pass.title,
      responseId: entry.synthesis.responseId,
      responseStatus: entry.synthesis.responseStatus
    })),
    responseStatus: summarizeStatuses(started.map((entry) => entry.synthesis.responseStatus)),
    model,
    orchestration: 'specialist-pass',
    passCount: started.length
  };
}

async function retrieveBackgroundSynthesis(responseRef) {
  const refs = normalizeResponseRefs(responseRef);
  if (!refs.length) {
    const error = new Error('responseId or responseIds is required.');
    error.statusCode = 400;
    throw error;
  }

  if (refs.length === 1 && !refs[0].key) {
    const result = await openAIJson(`/responses/${encodeURIComponent(refs[0].responseId)}`, {
      method: 'GET',
      timeoutMs: OPENAI_START_TIMEOUT_MS
    });

    return responseToSynthesis(result, SPECIALIST_PASSES[0]);
  }

  const retrieved = await Promise.all(refs.map(async (ref) => {
    const pass = passForKey(ref.key);
    const result = await openAIJson(`/responses/${encodeURIComponent(ref.responseId)}`, {
      method: 'GET',
      timeoutMs: OPENAI_START_TIMEOUT_MS
    });
    return {
      ref,
      pass,
      synthesis: responseToSynthesis(result, pass)
    };
  }));

  const responseIds = retrieved.map((entry) => ({
    key: entry.pass.key,
    title: entry.pass.title,
    responseId: entry.synthesis.responseId,
    responseStatus: entry.synthesis.responseStatus
  }));

  if (retrieved.some((entry) => entry.synthesis.pending)) {
    return {
      pending: true,
      responseIds,
      responseStatus: summarizeStatuses(retrieved.map((entry) => entry.synthesis.responseStatus)),
      model,
      orchestration: 'specialist-pass',
      passCount: retrieved.length
    };
  }

  return mergeSpecialistSyntheses(retrieved.map((entry) => entry.synthesis));
}

function normalizeResponseRefs(value) {
  if (Array.isArray(value)) {
    return value
      .map((entry) => {
        if (typeof entry === 'string') {
          return { responseId: entry };
        }
        if (entry && typeof entry === 'object' && entry.responseId) {
          return {
            key: entry.key || null,
            title: entry.title || null,
            responseId: entry.responseId
          };
        }
        return null;
      })
      .filter(Boolean);
  }

  if (typeof value === 'string') {
    return [{ responseId: value }];
  }

  if (value && typeof value === 'object' && value.responseId) {
    return [{
      key: value.key || null,
      title: value.title || null,
      responseId: value.responseId
    }];
  }

  return [];
}

function passForKey(key) {
  return SPECIALIST_PASSES.find((pass) => pass.key === key) || SPECIALIST_PASSES[0];
}

function summarizeStatuses(statuses) {
  const cleaned = statuses.filter(Boolean);
  if (!cleaned.length) {
    return 'in_progress';
  }
  if (cleaned.every((status) => status === 'completed')) {
    return 'completed';
  }
  if (cleaned.some((status) => status === 'in_progress')) {
    return 'in_progress';
  }
  return cleaned[0];
}

function mergeSpecialistSyntheses(syntheses) {
  const completed = syntheses.filter((synthesis) => synthesis?.canonicalDelta);
  const deltas = completed.map((synthesis) => synthesis.canonicalDelta);
  const canonical = deltas.find((delta, index) => completed[index]?.passKey === 'canonical') || deltas[0] || {};

  const merged = {
    processName: chooseBestText(deltas, 'processName') || canonical.processName || 'Distillery discovery run',
    businessFunction: chooseBestText(deltas, 'businessFunction') || canonical.businessFunction || 'Data discovery and migration readiness',
    recommendation: chooseBestText(deltas, 'recommendation') || canonical.recommendation || 'Complete evidence-backed discovery, close blockers, and generate the action pack.',
    decisionRequired: chooseBestText(deltas, 'decisionRequired') || canonical.decisionRequired || 'Confirm owners, source access, and remediation priorities.',
    systemsInScope: mergePrimitiveArrays(deltas.map((delta) => safeArray(delta.systemsInScope))),
    criticalOutputs: mergePrimitiveArrays(deltas.map((delta) => safeArray(delta.criticalOutputs))),
    overallRiskRating: chooseHighestRisk(deltas) || canonical.overallRiskRating || 'High',
    estimatedDollarExposure: deltas.reduce((accumulator, delta) => deepMerge(accumulator, delta.estimatedDollarExposure || {}), {}),
    executiveBrief: deltas.reduce((accumulator, delta) => deepMerge(accumulator, delta.executiveBrief || {}), {}),
    reportSections: mergeReportSections(deltas),
    items: mergeArraysByKey(deltas.map((delta) => safeArray(delta.items)), itemKey),
    relationships: mergeArraysByKey(deltas.map((delta) => safeArray(delta.relationships)), relationshipKey),
    artifacts: ensureDiscoveryArtifacts(mergeArraysByKey(deltas.map((delta) => safeArray(delta.artifacts)), artifactKey)),
    backlog: mergeArraysByKey(deltas.map((delta) => safeArray(delta.backlog)), actionKey),
    evidenceIndex: mergeArraysByKey(deltas.map((delta) => safeArray(delta.evidenceIndex)), evidenceKey),
    lineageNodes: mergeArraysByKey(deltas.map((delta) => safeArray(delta.lineageNodes)), lineageNodeKey),
    lineageEdges: mergeArraysByKey(deltas.map((delta) => safeArray(delta.lineageEdges)), lineageEdgeKey),
    failureRisks: mergeArraysByKey(deltas.map((delta) => safeArray(delta.failureRisks || delta.failure_risks || delta.failureModes)), failureKey),
    openQuestions: mergeArraysByKey(deltas.map((delta) => safeArray(delta.openQuestions || delta.open_questions)), questionKey)
  };

  for (const key of [
    'peopleRoles',
    'processSteps',
    'accessObjects',
    'excelObjects',
    'wordExtracts',
    'dataElements',
    'transformationsRules',
    'controlsExceptions',
    'dataQuality',
    'securityAccess',
    'scheduleSla',
    'failureModes',
    'financialModel'
  ]) {
    const mergedArray = mergeArraysByKey(deltas.map((delta) => safeArray(delta[key])), genericRecordKey(key));
    if (mergedArray.length) {
      merged[key] = mergedArray;
    }
  }

  if (!merged.criticalOutputs.length && merged.items.length) {
    merged.criticalOutputs = merged.items
      .filter((item) => /output|report|dashboard|extract|file/i.test(String(item.type || item.name || '')))
      .map((item) => item.name || item.id)
      .filter(Boolean)
      .slice(0, 25);
  }

  if (!merged.systemsInScope.length && merged.items.length) {
    merged.systemsInScope = merged.items
      .filter((item) => /system|database|file|workbook|document|access|excel|word/i.test(String(item.type || item.name || '')))
      .map((item) => item.name || item.id)
      .filter(Boolean)
      .slice(0, 50);
  }

  const passSummaries = completed.map((synthesis) => ({
    passKey: synthesis.passKey,
    passTitle: synthesis.passTitle,
    responseId: synthesis.responseId,
    responseStatus: synthesis.responseStatus,
    fallbackReason: synthesis.fallbackReason || null,
    items: safeArray(synthesis.canonicalDelta.items).length,
    relationships: safeArray(synthesis.canonicalDelta.relationships).length,
    backlog: safeArray(synthesis.canonicalDelta.backlog).length
  }));
  const fallbackReasons = completed
    .map((synthesis) => synthesis.fallbackReason)
    .filter(Boolean);
  const mergedStatus = completed.some((synthesis) => synthesis.responseStatus === 'incomplete')
    ? 'completed_with_blockers'
    : 'completed';

  return {
    pending: false,
    responseIds: completed.map((synthesis) => ({
      key: synthesis.passKey,
      title: synthesis.passTitle,
      responseId: synthesis.responseId,
      responseStatus: synthesis.responseStatus
    })),
    responseStatus: mergedStatus,
    model,
    orchestration: 'specialist-pass',
    passCount: completed.length,
    fallbackReason: fallbackReasons.join(' '),
    outputText: JSON.stringify({
      orchestration: 'specialist-pass',
      reasoningEffort,
      verbosity: responseVerbosity,
      passes: passSummaries,
      canonicalDelta: merged
    }),
    canonicalDelta: merged,
    raw: passSummaries
  };
}

function mergeReportSections(deltas) {
  const byTitle = new Map(REPORT_SECTION_TITLES.map((title) => [title, {
    title,
    body: '',
    confidence: null,
    evidenceIds: []
  }]));

  for (const delta of deltas) {
    const sections = safeArray(delta.reportSections);
    sections.forEach((section, index) => {
      const normalized = normalizeReportSection(section, index);
      const title = REPORT_SECTION_TITLES.includes(normalized.title)
        ? normalized.title
        : REPORT_SECTION_TITLES[index] || normalized.title;
      const existing = byTitle.get(title) || {
        title,
        body: '',
        confidence: null,
        evidenceIds: []
      };
      byTitle.set(title, {
        title,
        body: chooseBetterBody(existing.body, normalized.body),
        confidence: chooseBetterNumber(existing.confidence, normalized.confidence),
        evidenceIds: mergePrimitiveArrays([[...existing.evidenceIds], [...normalized.evidenceIds]])
      });
    });
  }

  return REPORT_SECTION_TITLES.map((title) => {
    const section = byTitle.get(title);
    if (section?.body) {
      return section;
    }
    return {
      title,
      body: 'Discovery evidence for this section is pending. The Distillery created a blocker so the next source export can complete it.',
      confidence: null,
      evidenceIds: []
    };
  });
}

function normalizeReportSection(section, index) {
  const record = asRecord(section);
  const rawTitle = firstText(record, ['title', 'heading', 'sectionTitle', 'section', 'name']);
  const title = !rawTitle || /^analysis section\b/i.test(rawTitle)
    ? REPORT_SECTION_TITLES[index] || `Report Section ${index + 1}`
    : rawTitle;
  return {
    title,
    body: firstText(record, ['body', 'narrative', 'summary', 'content', 'text']) || '',
    confidence: numberOrNull(record.confidence ?? record.confidenceScore),
    evidenceIds: stringArray(record.evidenceIds || record.evidence_ids || record.evidence)
  };
}

function ensureDiscoveryArtifacts(artifacts) {
  const byName = new Map();
  for (const artifact of artifacts) {
    const name = String(artifact.name || artifact.artifact_name || artifact.id || '').trim();
    if (name) {
      byName.set(name.toLowerCase(), artifact);
    }
  }

  return REQUIRED_ARTIFACTS.map((required) => {
    const existing = byName.get(required.name.toLowerCase()) || {};
    return {
      ...required,
      ...existing,
      id: existing.id || required.id,
      name: existing.name || required.name,
      type: existing.type || required.type,
      audience: existing.audience || required.audience,
      purpose: existing.purpose || required.purpose,
      status: existing.status || 'ready',
      progress: existing.progress ?? 100
    };
  });
}

function mergeArraysByKey(arrayGroups, keyFn) {
  const byKey = new Map();
  for (const group of arrayGroups) {
    for (const value of group) {
      const record = asRecord(value);
      if (!Object.keys(record).length) {
        continue;
      }
      const key = keyFn(record, byKey.size);
      const existing = byKey.get(key);
      byKey.set(key, existing ? mergeRecord(existing, record) : record);
    }
  }
  return [...byKey.values()];
}

function mergeRecord(existing, next) {
  const merged = { ...existing, ...next };
  for (const key of ['evidence', 'evidenceIds', 'evidence_ids', 'upstream', 'downstream', 'tags', 'sourceObjects', 'targetObjects']) {
    if (Array.isArray(existing[key]) || Array.isArray(next[key])) {
      merged[key] = mergePrimitiveArrays([safeArray(existing[key]), safeArray(next[key])]);
    }
  }
  return merged;
}

function deepMerge(target, source) {
  const left = asRecord(target);
  const right = asRecord(source);
  const merged = { ...left };
  for (const [key, value] of Object.entries(right)) {
    if (Array.isArray(value)) {
      merged[key] = mergePrimitiveArrays([safeArray(merged[key]), value]);
    } else if (value && typeof value === 'object') {
      merged[key] = deepMerge(merged[key], value);
    } else if (value !== null && value !== undefined && value !== '') {
      merged[key] = value;
    }
  }
  return merged;
}

function mergePrimitiveArrays(arrayGroups) {
  const seen = new Set();
  const merged = [];
  for (const group of arrayGroups) {
    for (const value of group) {
      if (value === null || value === undefined || value === '') {
        continue;
      }
      const key = typeof value === 'object' ? JSON.stringify(value) : String(value).toLowerCase();
      if (!seen.has(key)) {
        seen.add(key);
        merged.push(value);
      }
    }
  }
  return merged;
}

function chooseBestText(deltas, key) {
  const candidates = deltas
    .map((delta) => delta?.[key])
    .filter((value) => typeof value === 'string' && value.trim())
    .map((value) => value.trim());
  if (!candidates.length) {
    return '';
  }
  return [...candidates].sort((a, b) => b.length - a.length)[0];
}

function chooseBetterBody(left, right) {
  if (!right) {
    return left || '';
  }
  if (!left) {
    return right;
  }
  const leftGeneric = /pending|generated from canonical|no narrative|analysis section/i.test(left);
  const rightGeneric = /pending|generated from canonical|no narrative|analysis section/i.test(right);
  if (leftGeneric && !rightGeneric) {
    return right;
  }
  if (!leftGeneric && rightGeneric) {
    return left;
  }
  return right.length > left.length ? right : left;
}

function chooseBetterNumber(left, right) {
  if (Number.isFinite(right) && !Number.isFinite(left)) {
    return right;
  }
  if (Number.isFinite(right) && Number.isFinite(left)) {
    return Math.max(left, right);
  }
  return Number.isFinite(left) ? left : null;
}

function chooseHighestRisk(deltas) {
  const rank = {
    critical: 4,
    high: 3,
    medium: 2,
    moderate: 2,
    low: 1
  };
  const risks = deltas
    .map((delta) => String(delta?.overallRiskRating || '').trim())
    .filter(Boolean);
  if (!risks.length) {
    return '';
  }
  return risks.sort((a, b) => (rank[b.toLowerCase()] || 0) - (rank[a.toLowerCase()] || 0))[0];
}

function itemKey(record, index) {
  return stableKey(record, ['id', 'item_id', 'node_id'], ['type', 'name'], `item-${index}`);
}

function relationshipKey(record, index) {
  return stableKey(record, ['id', 'relationshipId', 'relationship_id', 'edge_id'], ['fromId', 'from_id', 'from', 'toId', 'to_id', 'to', 'type', 'edgeType'], `relationship-${index}`);
}

function artifactKey(record, index) {
  return stableKey(record, ['id', 'artifact_id'], ['name', 'artifact_name'], `artifact-${index}`);
}

function actionKey(record, index) {
  return stableKey(record, ['actionId', 'action_id', 'id'], ['title', 'summary', 'linkedItemId'], `action-${index}`);
}

function evidenceKey(record, index) {
  return stableKey(record, ['evidenceId', 'evidence_id', 'id'], ['type', 'location', 'description'], `evidence-${index}`);
}

function lineageNodeKey(record, index) {
  return stableKey(record, ['nodeId', 'node_id', 'id'], ['node_type', 'type', 'name'], `lineage-node-${index}`);
}

function lineageEdgeKey(record, index) {
  return stableKey(record, ['edgeId', 'edge_id', 'id'], ['fromId', 'from_id', 'toId', 'to_id', 'edge_type', 'type'], `lineage-edge-${index}`);
}

function failureKey(record, index) {
  return stableKey(record, ['scenarioId', 'scenario_id', 'id'], ['scenario', 'trigger', 'effect', 'impactedOutput'], `failure-${index}`);
}

function questionKey(record, index) {
  return stableKey(record, ['questionId', 'question_id', 'id'], ['question'], `question-${index}`);
}

function genericRecordKey(prefix) {
  return (record, index) => stableKey(record, ['id', `${prefix}Id`, `${prefix}_id`, 'object_id', 'step_id', 'control_id', 'element_id', 'transform_id', 'scenario_id'], ['name', 'object_name', 'title', 'location', 'description'], `${prefix}-${index}`);
}

function stableKey(record, idKeys, compoundKeys, fallback) {
  for (const key of idKeys) {
    const value = record[key];
    if (value !== null && value !== undefined && String(value).trim()) {
      return String(value).trim().toLowerCase();
    }
  }
  const compound = compoundKeys
    .map((key) => record[key])
    .filter((value) => value !== null && value !== undefined && String(value).trim())
    .map((value) => String(value).trim().toLowerCase())
    .join('|');
  return compound || fallback;
}

function safeArray(value) {
  return Array.isArray(value) ? value : [];
}

function asRecord(value) {
  return value && typeof value === 'object' && !Array.isArray(value) ? value : {};
}

function numberOrNull(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function firstText(record, keys) {
  for (const key of keys) {
    const value = record[key];
    if (typeof value === 'string' && value.trim()) {
      return value.trim();
    }
    if (Array.isArray(value)) {
      const text = value.map((item) => String(item || '').trim()).filter(Boolean).join(', ');
      if (text) {
        return text;
      }
    }
  }
  return '';
}

function stringArray(value) {
  if (Array.isArray(value)) {
    return value.map((item) => {
      if (typeof item === 'string') {
        return item;
      }
      if (item && typeof item === 'object') {
        return String(item.id || item.evidenceId || item.evidence_id || '').trim();
      }
      return String(item || '').trim();
    }).filter(Boolean);
  }
  if (typeof value === 'string' && value.trim()) {
    return value.split(/\s*,\s*/).filter(Boolean);
  }
  if (value && typeof value === 'object') {
    const id = value.id || value.evidenceId || value.evidence_id;
    return id ? [String(id)] : [];
  }
  return [];
}

module.exports = {
  model,
  parseJsonOutput,
  retrieveBackgroundSynthesis,
  startBackgroundSynthesis,
  synthesizeWithOpenAI
};
