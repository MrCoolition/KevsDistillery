const { createHash } = require('node:crypto');

const model = process.env.OPENAI_MODEL || 'gpt-5.5';
const reasoningEffort = process.env.OPENAI_REASONING_EFFORT || 'high';
const responseVerbosity = process.env.OPENAI_VERBOSITY || 'high';
const OPENAI_TIMEOUT_MS = Number(process.env.OPENAI_TIMEOUT_MS || 59000);
const OPENAI_START_TIMEOUT_MS = Number(process.env.OPENAI_START_TIMEOUT_MS || 25000);
const OPENAI_MAX_OUTPUT_TOKENS = Number(process.env.OPENAI_MAX_OUTPUT_TOKENS || 16000);
const OPENAI_SPECIALIST_OUTPUT_TOKENS = Number(process.env.OPENAI_SPECIALIST_OUTPUT_TOKENS || 10000);
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
    'Systematic extraction comes first: the payload may include a SYSTEMATIC DISCOVERY INVENTORY, workbook XML objects, connection definitions, query-table metadata, Power Query candidates, VBA module source excerpts, and code-scan markers. Treat those deterministic objects as primary evidence.',
    'Never treat generated package artifact names as source objects, business outputs, or critical outputs. 01_Executive_Decision_Brief.pdf, Diagram Pack, Action Backlog, Evidence Archive, and Metadata Manifest are deliverables, not discovered source nodes.',
    'For Excel workbooks, sheet nodes may only come from explicit Workbook sheet map / workbook <sheet> evidence. Namespace/function tokens such as microsoft.com:RD, microsoft.com:Single, microsoft.com:FV, LET_WF, LAMBDA_WF, ARRAYTEXT_WF, and _xlfn.* are not worksheet names.',
    'External connections, external links, query tables, Power Query candidates, VBA projects/procedures, defined names, pivots, formulas, hyperlinks, and worksheet dimensions must be separate object types with their own evidence and next action.',
    'When VBA source excerpts are present, inspect the code. Extract procedures, call graph hints, SQL strings, file/path/URL references, RefreshAll/query-table/connection calls, Range/Cells writes, formula writes, output exports, error handlers, event handlers, controls, and transformation rules as canonical items and transformationsRules.',
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
    'Keep JSON compact enough to complete: no markdown prose, no repeated boilerplate, no duplicate objects, no generated artifact nodes, and no run-status/rerun actions in the delivery backlog.',
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
    return mergeSpecialistSyntheses(started.map((entry) => entry.synthesis), payload);
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

async function retrieveBackgroundSynthesis(responseRef, requestPayload = {}) {
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
    const finished = retrieved.filter((entry) => !entry.synthesis.pending && entry.synthesis.canonicalDelta);
    const partial = finished.length
      ? mergeSpecialistSyntheses(finished.map((entry) => entry.synthesis), requestPayload)
      : null;
    return {
      pending: true,
      responseIds,
      responseStatus: summarizeStatuses(retrieved.map((entry) => entry.synthesis.responseStatus)),
      model,
      orchestration: 'specialist-pass',
      passCount: retrieved.length,
      passProgress: summarizePassProgress(retrieved),
      partialCanonicalDelta: partial?.canonicalDelta || null,
      partialCounts: partial?.canonicalDelta ? canonicalCounts(partial.canonicalDelta) : null
    };
  }

  return mergeSpecialistSyntheses(retrieved.map((entry) => entry.synthesis), requestPayload);
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

function summarizePassProgress(entries) {
  const finished = entries.filter((entry) => !entry.synthesis.pending);
  const running = entries.filter((entry) => entry.synthesis.pending);
  return {
    total: entries.length,
    finished: finished.length,
    running: running.length,
    completed: entries.filter((entry) => entry.synthesis.responseStatus === 'completed').length,
    incomplete: entries.filter((entry) => entry.synthesis.responseStatus === 'incomplete').length,
    runningTitles: running.map((entry) => entry.pass.title),
    finishedTitles: finished.map((entry) => entry.pass.title)
  };
}

function canonicalCounts(delta = {}) {
  return {
    items: safeArray(delta.items).length,
    relationships: safeArray(delta.relationships).length,
    artifacts: safeArray(delta.artifacts).length,
    backlog: safeArray(delta.backlog).length
  };
}

function mergeSpecialistSyntheses(syntheses, requestPayload = {}) {
  const completed = syntheses.filter((synthesis) => synthesis?.canonicalDelta);
  const deltas = completed.map((synthesis) => synthesis.canonicalDelta);
  const canonical = deltas.find((delta, index) => completed[index]?.passKey === 'canonical') || deltas[0] || {};
  const sourceFallback = buildSourceEvidenceDelta(requestPayload);
  const processName = chooseSourceAwareText(deltas, 'processName', canonical.processName, sourceFallback.processName, 'Source discovery run');
  const businessFunction = chooseSourceAwareText(deltas, 'businessFunction', canonical.businessFunction, sourceFallback.businessFunction, 'Data discovery and migration readiness');
  const recommendation = chooseSourceAwareText(deltas, 'recommendation', canonical.recommendation, sourceFallback.recommendation, 'Complete evidence-backed discovery, close blockers, and generate the action pack.');
  const decisionRequired = chooseSourceAwareText(deltas, 'decisionRequired', canonical.decisionRequired, sourceFallback.decisionRequired, 'Confirm owners, source access, and remediation priorities.');

  const merged = {
    processName,
    businessFunction,
    recommendation,
    decisionRequired,
    systemsInScope: mergePrimitiveArrays([...deltas.map((delta) => safeArray(delta.systemsInScope)), safeArray(sourceFallback.systemsInScope)]),
    criticalOutputs: mergePrimitiveArrays([...deltas.map((delta) => safeArray(delta.criticalOutputs)), safeArray(sourceFallback.criticalOutputs)]),
    overallRiskRating: chooseHighestRisk(deltas) || canonical.overallRiskRating || 'High',
    estimatedDollarExposure: deltas.reduce((accumulator, delta) => deepMerge(accumulator, delta.estimatedDollarExposure || {}), {}),
    executiveBrief: deltas.reduce((accumulator, delta) => deepMerge(accumulator, delta.executiveBrief || {}), {}),
    reportSections: mergeReportSections([...deltas, sourceFallback]),
    items: mergeArraysByKey([...deltas.map((delta) => safeArray(delta.items)), safeArray(sourceFallback.items)], itemKey).filter((item) => !isInternalDiscoveryRecord(item)),
    relationships: mergeArraysByKey([...deltas.map((delta) => safeArray(delta.relationships)), safeArray(sourceFallback.relationships)], relationshipKey),
    artifacts: ensureDiscoveryArtifacts(mergeArraysByKey(deltas.map((delta) => safeArray(delta.artifacts)), artifactKey)),
    backlog: mergeArraysByKey([...deltas.map((delta) => safeArray(delta.backlog)), safeArray(sourceFallback.backlog)], actionKey).filter((action) => !isRunDiagnosticAction(action)),
    evidenceIndex: mergeArraysByKey([...deltas.map((delta) => safeArray(delta.evidenceIndex)), safeArray(sourceFallback.evidenceIndex)], evidenceKey).filter((evidence) => !isRunDiagnosticEvidence(evidence)),
    lineageNodes: mergeArraysByKey([...deltas.map((delta) => safeArray(delta.lineageNodes)), safeArray(sourceFallback.lineageNodes)], lineageNodeKey),
    lineageEdges: mergeArraysByKey([...deltas.map((delta) => safeArray(delta.lineageEdges)), safeArray(sourceFallback.lineageEdges)], lineageEdgeKey),
    failureRisks: mergeArraysByKey(deltas.map((delta) => safeArray(delta.failureRisks || delta.failure_risks || delta.failureModes)), failureKey).filter((risk) => !isRunDiagnosticRisk(risk)),
    openQuestions: mergeArraysByKey(deltas.map((delta) => safeArray(delta.openQuestions || delta.open_questions)), questionKey).filter((question) => !isRunDiagnosticQuestion(question))
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
    const mergedArray = mergeArraysByKey([...deltas.map((delta) => safeArray(delta[key])), safeArray(sourceFallback[key])], genericRecordKey(key));
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

function buildSourceEvidenceDelta(payload = {}) {
  const extractedText = String(payload.extractedText || '');
  const sourceName = String(payload.sourceName || firstSourceName(extractedText) || 'Uploaded source');
  const sourceKind = String(payload.sourceKind || inferSourceKindFromEvidence(extractedText, sourceName));
  if (!extractedText && !sourceName) {
    return {};
  }

  const evidenceId = 'EV-SOURCE-INTAKE';
  const workbookSheets = extractListAfter(extractedText, /Workbook sheets:\s*([^\n\r]+)/i, 40)
    .filter(isLikelyWorkbookSheetName);
  const legacyWorksheetMatches = [...extractedText.matchAll(/^(xl\/worksheets\/[^:\s)]+):\s*dimension\s*([^;]+);\s*formulas:\s*([^\n\r]+)/gmi)]
    .map((match) => ({
      name: match[1].replace(/^.*\//, '').replace(/\.xml$/i, ''),
      path: match[1],
      dimension: match[2],
      formulas: match[3]
    }));
  const namedWorksheetMatches = [...extractedText.matchAll(/Worksheet\s+([^\n\r(]+)\s+\((xl\/worksheets\/[^)]+)\):\s*dimension\s*([^;]+);\s*formulas:\s*([^\n\r;]+)/gi)]
    .map((match) => ({
      name: cleanName(match[1]),
      path: match[2],
      dimension: match[3],
      formulas: match[4]
    }));
  const worksheetMatches = [...legacyWorksheetMatches, ...namedWorksheetMatches]
    .filter((sheet) => isLikelyWorkbookSheetName(sheet.name))
    .slice(0, 30)
    .map((sheet, index) => ({
      id: `WS-${String(index + 1).padStart(3, '0')}`,
      ...sheet
    }));
  const metadataPaths = uniqueMatches(extractedText, /\b(xl\/(?:connections|queryTables|queries|pivotTables|pivotCache|externalLinks|customXml|customData|tables|slicers)\/[^\s:]+)\b/gi, 35);
  const connectionEvidence = uniqueMatches(extractedText, /^Connection\s+\d+:\s*([^\n\r]+)/gmi, 20);
  const powerQueryEvidence = uniqueMatches(extractedText, /^Power Query candidate\s+\d+:\s*([^\n\r]+)/gmi, 20);
  const externalLinkEvidence = [
    ...uniqueMatches(extractedText, /^External link\s+\d+:\s*([^\n\r]+)/gmi, 20),
    ...uniqueMatches(extractedText, /^External URL reference\s+\d+:\s*([^\n\r]+)/gmi, 30)
  ].slice(0, 30);
  const queryTableEvidence = uniqueMatches(extractedText, /^Query table for\s+[^:]+:\s*([^\n\r]+)/gmi, 20);
  const vbaProcedures = extractListAfter(extractedText, /VBA procedures:\s*([^\n\r]+)/i, 40).filter((value) => !/^none recovered|not detected|unknown$/i.test(value));
  const vbaModuleEvidence = parseVbaModuleEvidence(extractedText, 30);
  const vbaProcedureNames = mergePrimitiveArrays([vbaProcedures, vbaModuleEvidence.flatMap((module) => module.procedures)]).slice(0, 80);
  const definedNames = extractListAfter(extractedText, /Defined names:\s*([^\n\r]+)/i, 40);
  const hasVbaProject = /VBA project:\s*present/i.test(extractedText);
  const recoveredTerms = recoveredEvidenceTerms(extractedText, 40);
  const sheetNames = workbookSheets.length
    ? workbookSheets
    : worksheetMatches.map((sheet) => sheet.name);
  const formulaSheets = worksheetMatches.filter((sheet) => !/^none detected$/i.test(sheet.formulas.trim())).slice(0, 12);
  const criticalOutputs = deriveCriticalOutputs(payload, sourceName, sheetNames, recoveredTerms);

  const sourceNode = canonicalItem({
    id: 'SRC-001',
    type: sourceKind === 'excel' ? 'workbook' : sourceKind || 'file',
    name: sourceName,
    businessPurpose: `${sourceName} is the submitted source artifact for this discovery run.`,
    evidenceId,
    confidence: extractedText ? 86 : 65,
    criticality: 'high',
    downstream: ['INV-001'],
    failureImpact: 'If this source is misunderstood, engineers will migrate the wrong process, miss dependencies, or rebuild incomplete logic.',
    action: 'Inventory workbook structure, formulas, connections, controls, outputs, and unresolved blockers from extracted source evidence.'
  });
  const inventoryNode = canonicalItem({
    id: 'INV-001',
    type: 'source_inventory',
    name: `${sourceName} extracted inventory`,
    businessPurpose: 'Browser-side extraction of workbook sheets, dimensions, formulas, strings, and package metadata used as the first-pass discovery inventory.',
    evidenceId,
    confidence: extractedText ? 82 : 60,
    criticality: 'medium',
    upstream: ['SRC-001'],
    downstream: formulaSheets.length ? formulaSheets.map((_, index) => `FORM-${String(index + 1).padStart(3, '0')}`) : ['OUT-001'],
    failureImpact: 'If the inventory is incomplete, lineage and transformation reconstruction remains partial until native metadata exports are supplied.',
    action: 'Validate extracted inventory against native workbook metadata, named ranges, tables, Power Query, VBA modules, and sample outputs.'
  });

  const sheetNodes = sheetNames.slice(0, 20).map((sheet, index) => canonicalItem({
    id: `SHEET-${String(index + 1).padStart(3, '0')}`,
    type: 'sheet',
    name: cleanName(sheet),
    businessPurpose: `Worksheet discovered in ${sourceName}.`,
    evidenceId,
    confidence: 84,
    criticality: index < 5 ? 'medium' : 'low',
    upstream: ['SRC-001'],
    downstream: formulaSheets.length ? [`FORM-${String(Math.min(index + 1, formulaSheets.length)).padStart(3, '0')}`] : ['OUT-001'],
    failureImpact: 'Incorrect sheet interpretation can break workbook lineage, refresh order, output mapping, or migration requirements.',
    action: 'Classify sheet as input, staging, calculation, control, or output and capture owner and refresh cadence.'
  }));

  const formulaNodes = formulaSheets.map((sheet, index) => canonicalItem({
    id: `FORM-${String(index + 1).padStart(3, '0')}`,
    type: 'formula_block',
    name: `${sheet.path.replace(/^.*\//, '').replace(/\.xml$/i, '')} formulas`,
    businessPurpose: 'Formula logic discovered in workbook XML and requiring business interpretation before migration.',
    evidenceId,
    confidence: 78,
    criticality: 'high',
    upstream: ['INV-001', `SHEET-${String(index + 1).padStart(3, '0')}`],
    downstream: ['OUT-001'],
    failureImpact: 'If formulas are rebuilt incorrectly, migrated outputs may silently diverge from the current workbook.',
    action: 'Extract full formulas with cell addresses, named ranges, precedents, dependents, and expected output examples.'
  }));

  const metadataNodes = metadataPaths
    .slice(0, 12)
    .map((path, index) => canonicalItem({
      id: `META-${String(index + 1).padStart(3, '0')}`,
      type: metadataType(path),
      name: path,
      businessPurpose: `Workbook package metadata discovered in ${sourceName}.`,
      evidenceId,
      confidence: 76,
      criticality: /connection|query|external/i.test(path) ? 'high' : 'medium',
      upstream: ['SRC-001'],
      downstream: ['INV-001'],
      failureImpact: 'Unresolved workbook metadata can hide external sources, refresh dependencies, pivots, or object-level lineage.',
      action: 'Export and inspect full workbook connection, query, pivot, external link, and VBA metadata.'
    }));

  const connectionNodes = connectionEvidence.slice(0, 12).map((text, index) => canonicalItem({
    id: `CONN-${String(index + 1).padStart(3, '0')}`,
    type: 'external_connection',
    name: text,
    businessPurpose: `External connection discovered in ${sourceName}.`,
    evidenceId,
    confidence: 86,
    criticality: 'high',
    upstream: ['SRC-001'],
    downstream: ['INV-001'],
    failureImpact: 'If this connection is missed, upstream source-of-truth and refresh lineage are incomplete.',
    action: 'Confirm connection target, credential owner, refresh cadence, source system, and replacement ingestion pattern.'
  }));

  const powerQueryNodes = powerQueryEvidence.slice(0, 12).map((text, index) => canonicalItem({
    id: `PQ-${String(index + 1).padStart(3, '0')}`,
    type: 'power_query',
    name: text,
    businessPurpose: `Power Query or query-table logic candidate discovered in ${sourceName}.`,
    evidenceId,
    confidence: 82,
    criticality: 'high',
    upstream: connectionNodes.length ? connectionNodes.map((node) => node.id).slice(0, 3) : ['SRC-001'],
    downstream: ['OUT-001'],
    failureImpact: 'If Power Query logic is omitted, Fivetran/dbt/Snowpark rebuilds may miss source filters, joins, type changes, or refresh dependencies.',
    action: 'Extract full M/query-table definition, source credentials, parameters, refresh order, and output worksheet target.'
  }));

  const externalLinkNodes = externalLinkEvidence.slice(0, 12).map((text, index) => canonicalItem({
    id: `XLINK-${String(index + 1).padStart(3, '0')}`,
    type: 'external_link',
    name: text,
    businessPurpose: `External workbook link discovered in ${sourceName}.`,
    evidenceId,
    confidence: 84,
    criticality: 'high',
    upstream: ['SRC-001'],
    downstream: ['INV-001'],
    failureImpact: 'External workbook links can make lineage dependent on unmanaged files, network paths, or manually maintained sources.',
    action: 'Confirm external workbook owner, file path, refresh cadence, and whether it is authoritative, duplicate, or obsolete.'
  }));

  const queryTableNodes = queryTableEvidence.slice(0, 12).map((text, index) => canonicalItem({
    id: `QT-${String(index + 1).padStart(3, '0')}`,
    type: 'query_table',
    name: text,
    businessPurpose: `Worksheet query table discovered in ${sourceName}.`,
    evidenceId,
    confidence: 82,
    criticality: 'high',
    upstream: connectionNodes.length ? connectionNodes.map((node) => node.id).slice(0, 3) : ['SRC-001'],
    downstream: ['OUT-001'],
    failureImpact: 'Query table refresh behavior may control what data arrives on sheets before formulas/macros run.',
    action: 'Extract query text, connection target, refresh settings, destination range, and downstream formula/output dependencies.'
  }));

  const vbaProjectNode = hasVbaProject || vbaModuleEvidence.length ? [
    canonicalItem({
      id: 'VBA-001',
      type: 'vba_project',
      name: `${sourceName} VBA project`,
      businessPurpose: `Macro project detected in ${sourceName}.`,
      evidenceId,
      confidence: vbaModuleEvidence.length ? 88 : 80,
      criticality: 'high',
      upstream: ['SRC-001'],
      downstream: vbaModuleEvidence.length ? ['VBA-MOD-001'] : vbaProcedureNames.length ? ['VBA-PROC-001'] : ['OUT-001'],
      failureImpact: 'VBA can hide refresh orchestration, external calls, file exports, user actions, and business rules.',
      action: vbaModuleEvidence.length
        ? 'Analyze extracted module source for transformations, file IO, SQL, refresh order, event triggers, call graph, and generated outputs.'
        : 'Export all modules, worksheet/workbook events, forms, button bindings, references, and macro execution order.'
    })
  ] : [];

  const vbaModuleNodes = vbaModuleEvidence.slice(0, 20).map((module, index) => canonicalItem({
    id: `VBA-MOD-${String(index + 1).padStart(3, '0')}`,
    type: 'vba_module',
    name: module.name,
    businessPurpose: `Extracted VBA module source from ${sourceName}; deterministic code scan found ${module.procedures.length} procedures, ${module.markers.length} transformation/automation markers, ${module.sqlStrings.length} SQL string candidates, and ${module.fileRefs.length} file or URL references.`,
    evidenceId,
    confidence: 88,
    criticality: 'high',
    upstream: ['VBA-001'],
    downstream: module.procedures.length ? [`VBA-PROC-${String(index + 1).padStart(3, '0')}-001`] : ['OUT-001'],
    failureImpact: 'Module code can mutate worksheets, refresh data, write files, run SQL, call external systems, and encode undocumented business rules.',
    action: 'Review module code, build a call graph, map touched sheets/ranges/files/queries, classify transformations, and translate rebuildable logic into dbt/Snowpark tests.'
  }));

  const vbaProcedureNodes = vbaModuleEvidence.length
    ? vbaModuleEvidence.slice(0, 20).flatMap((module, moduleIndex) => module.procedures.slice(0, 16).map((procedure, procIndex) => canonicalItem({
      id: `VBA-PROC-${String(moduleIndex + 1).padStart(3, '0')}-${String(procIndex + 1).padStart(3, '0')}`,
      type: 'vba_procedure',
      name: procedure,
      businessPurpose: `Procedure ${procedure} recovered from VBA module ${module.name}. Markers: ${module.markers.slice(0, 4).join(' | ') || 'none in excerpt'}.`,
      evidenceId,
      confidence: 86,
      criticality: 'high',
      upstream: [`VBA-MOD-${String(moduleIndex + 1).padStart(3, '0')}`],
      downstream: ['OUT-001'],
      failureImpact: 'Unmapped macro procedures can silently alter workbook state, outputs, files, SQL sources, and refresh order.',
      action: 'Map procedure callers, touched ranges/sheets, external files, SQL strings, error handlers, and output effects.'
    })))
    : vbaProcedureNames.slice(0, 20).map((procedure, index) => canonicalItem({
      id: `VBA-PROC-${String(index + 1).padStart(3, '0')}`,
      type: 'vba_procedure',
      name: procedure,
      businessPurpose: `Recovered VBA procedure candidate from ${sourceName}.`,
      evidenceId,
      confidence: 72,
      criticality: 'high',
      upstream: ['VBA-001'],
      downstream: ['OUT-001'],
      failureImpact: 'Unmapped macro procedures can silently alter workbook state, outputs, files, and refresh order.',
      action: 'Export procedure body, call graph, triggered event, affected sheets/ranges/files, and replacement design.'
    }));

  const vbaRuleNodes = vbaModuleEvidence.slice(0, 20).flatMap((module, moduleIndex) => [
    ...module.markers.slice(0, 12).map((marker, markerIndex) => canonicalItem({
      id: `VBA-RULE-${String(moduleIndex + 1).padStart(3, '0')}-${String(markerIndex + 1).padStart(3, '0')}`,
      type: 'vba_transformation_marker',
      name: `${module.name}: ${marker.slice(0, 80)}`,
      businessPurpose: `Deterministic VBA code scan found transformation/automation marker in ${module.name}.`,
      evidenceId,
      confidence: 84,
      criticality: 'high',
      upstream: [`VBA-MOD-${String(moduleIndex + 1).padStart(3, '0')}`],
      downstream: ['OUT-001'],
      failureImpact: 'Transformation marker may represent hidden writeback, refresh, file export, calculation, or data mutation logic.',
      action: 'Classify the marker as transform, refresh, file IO, control, or output generation and map source/target objects.'
    })),
    ...module.sqlStrings.slice(0, 8).map((sql, sqlIndex) => canonicalItem({
      id: `VBA-SQL-${String(moduleIndex + 1).padStart(3, '0')}-${String(sqlIndex + 1).padStart(3, '0')}`,
      type: 'vba_sql_string',
      name: `${module.name}: SQL candidate ${sqlIndex + 1}`,
      businessPurpose: `SQL string recovered from VBA module ${module.name}.`,
      evidenceId,
      confidence: 84,
      criticality: 'high',
      upstream: [`VBA-MOD-${String(moduleIndex + 1).padStart(3, '0')}`],
      downstream: ['OUT-001'],
      failureImpact: 'SQL embedded in VBA can define source filters, joins, deletes, updates, or output creation that must be rebuilt and tested.',
      action: 'Extract complete SQL text, identify source/target tables, classify DML/DDL/read query, and translate into governed migration logic.'
    })),
    ...module.fileRefs.slice(0, 8).map((ref, refIndex) => canonicalItem({
      id: `VBA-REF-${String(moduleIndex + 1).padStart(3, '0')}-${String(refIndex + 1).padStart(3, '0')}`,
      type: 'vba_file_reference',
      name: `${module.name}: ${ref.slice(0, 90)}`,
      businessPurpose: `File, URL, or path reference recovered from VBA module ${module.name}.`,
      evidenceId,
      confidence: 82,
      criticality: 'high',
      upstream: [`VBA-MOD-${String(moduleIndex + 1).padStart(3, '0')}`],
      downstream: ['OUT-001'],
      failureImpact: 'File/path references can hide upstream extracts, generated outputs, manually maintained dependencies, or audit evidence.',
      action: 'Confirm file owner, path, cadence, read/write behavior, retention, and migration treatment.'
    }))
  ]).slice(0, 80);

  const definedNameNodes = definedNames.slice(0, 20).map((definedName, index) => canonicalItem({
    id: `NAME-${String(index + 1).padStart(3, '0')}`,
    type: 'named_range',
    name: definedName,
    businessPurpose: `Workbook defined name discovered in ${sourceName}.`,
    evidenceId,
    confidence: 80,
    criticality: 'medium',
    upstream: ['SRC-001'],
    downstream: formulaSheets.length ? [`FORM-${String(Math.min(index + 1, formulaSheets.length)).padStart(3, '0')}`] : ['OUT-001'],
    failureImpact: 'Named ranges can hide calculation inputs, parameters, output areas, or workbook control flags.',
    action: 'Map defined name scope, formula/reference, consumers, and migration equivalent.'
  }));

  const outputNodes = criticalOutputs.slice(0, 8).map((output, index) => canonicalItem({
    id: `OUT-${String(index + 1).padStart(3, '0')}`,
    type: 'output',
    name: cleanName(output),
    businessPurpose: `Business output candidate inferred from ${sourceName}.`,
    evidenceId,
    confidence: index === 0 ? 72 : 64,
    criticality: 'high',
    upstream: formulaNodes.length
      ? formulaNodes.map((node) => node.id)
      : vbaRuleNodes.length
        ? vbaRuleNodes.slice(0, 10).map((node) => node.id)
        : vbaProcedureNodes.length
          ? vbaProcedureNodes.slice(0, 10).map((node) => node.id)
          : sheetNodes.slice(0, 5).map((node) => node.id),
    failureImpact: 'If this output is late, wrong, partial, or unauditable, downstream decisions and migration validation are at risk.',
    action: 'Confirm output owner, SLA, consumers, refresh cadence, dollar exposure inputs, and acceptance criteria.'
  }));

  const vbaNodes = [...vbaProjectNode, ...vbaModuleNodes, ...vbaProcedureNodes, ...vbaRuleNodes];
  const items = [sourceNode, inventoryNode, ...sheetNodes, ...metadataNodes, ...connectionNodes, ...externalLinkNodes, ...queryTableNodes, ...powerQueryNodes, ...definedNameNodes, ...vbaNodes, ...formulaNodes, ...outputNodes];
  const relationships = [
    relationship('REL-SRC-INV', 'SRC-001', 'INV-001', 'contains', evidenceId, 86),
    ...sheetNodes.map((node, index) => relationship(`REL-SRC-SHEET-${index + 1}`, 'SRC-001', node.id, 'contains_sheet', evidenceId, 84)),
    ...[...metadataNodes, ...connectionNodes, ...externalLinkNodes, ...queryTableNodes, ...powerQueryNodes, ...definedNameNodes, ...vbaProjectNode].map((node, index) => relationship(`REL-META-INV-${index + 1}`, node.id, 'INV-001', 'documents_metadata', evidenceId, 76)),
    ...vbaModuleNodes.map((node, index) => relationship(`REL-VBA-MOD-${index + 1}`, 'VBA-001', node.id, 'contains_vba_module', evidenceId, 88)),
    ...vbaProcedureNodes.map((node, index) => relationship(`REL-VBA-PROC-${index + 1}`, safeArray(node.upstream)[0] || 'VBA-001', node.id, 'contains_vba_procedure', evidenceId, 86)),
    ...vbaRuleNodes.map((node, index) => relationship(`REL-VBA-RULE-${index + 1}`, safeArray(node.upstream)[0] || 'VBA-001', node.id, 'documents_vba_logic', evidenceId, 84)),
    ...formulaNodes.map((node, index) => relationship(`REL-INV-FORM-${index + 1}`, 'INV-001', node.id, 'extracts_formula_logic', evidenceId, 78)),
    ...outputNodes.flatMap((output, outputIndex) => {
      const upstream = safeArray(output.upstream);
      return upstream.slice(0, 8).map((from, index) => relationship(`REL-OUT-${outputIndex + 1}-${index + 1}`, from, output.id, 'feeds_output', evidenceId, 72));
    })
  ];

  const reportSections = [
    reportSection('Executive Snapshot', `${sourceName} was analyzed from uploaded source evidence. The first-pass model identifies ${sheetNodes.length} real workbook sheets, ${formulaNodes.length} formula blocks, ${connectionNodes.length} external connections, ${powerQueryNodes.length + queryTableNodes.length} query/Power Query objects, ${externalLinkNodes.length} external links, ${vbaNodes.length ? 'a VBA project' : 'no VBA project evidence'}, ${definedNameNodes.length} defined names, ${metadataNodes.length} metadata objects, and ${outputNodes.length} output candidates. Confidence is evidence-backed but remains blocker-aware until full native exports and sample outputs are confirmed.`, 84, [evidenceId]),
    reportSection('System and Artifact Landscape', `${sourceName} is the source artifact in scope. Discovered sheet evidence includes ${sheetNames.slice(0, 10).join(', ') || 'workbook XML'}${connectionNodes.length ? `; external connections include ${connectionNodes.slice(0, 3).map((node) => node.name).join('; ')}` : ''}${powerQueryNodes.length ? `; Power Query candidates include ${powerQueryNodes.slice(0, 3).map((node) => node.name).join('; ')}` : ''}${vbaNodes.length ? '; VBA project evidence is present' : ''}.`, 84, [evidenceId]),
    reportSection('Data Flow and Process Flow Summary', relationships.slice(0, 8).map((edge) => `${edge.fromId} ${edge.type} ${edge.toId}`).join('; ') || `${sourceName} contains discovered workbook structures that require confirmed output mapping.`, 78, [evidenceId]),
    reportSection('Transformations and Business Logic', `${formulaNodes.length} formula blocks, ${definedNameNodes.length} defined names, ${powerQueryNodes.length} Power Query candidates, ${queryTableNodes.length} query tables, and ${vbaNodes.length} VBA nodes were detected from workbook package evidence. Full source definitions, macro bodies, query text, refresh order, and expected outputs are required to finish transformation lineage.`, 78, [evidenceId]),
    reportSection('Recursive Lineage and Source-of-Truth Assessment', `Lineage starts at ${sourceName}, proceeds through extracted workbook inventory, sheets, metadata, and formula blocks, and currently terminates at inferred output candidates until external connections, source extracts, named ranges, and owner-confirmed outputs are supplied.`, 76, [evidenceId])
  ];

  const backlog = [
    {
      actionId: 'ACT-SOURCE-METADATA-EXPORT',
      title: `Export native metadata for ${sourceName}`,
      mode: 'stabilize',
      owner: 'Source owner',
      priority: 'P0',
      dependency: sourceName,
      acceptanceCriteria: 'Workbook sheets, tables, named ranges, formulas with addresses, Power Query, VBA, pivots, external links, refresh order, sample outputs, owner, SLA, and control evidence are attached to the discovery run.',
      linkedItemId: 'SRC-001',
      summary: 'Complete source metadata is required to move from inferred workbook discovery to migration-grade lineage.'
    },
    ...(externalLinkNodes.length ? [{
      actionId: 'ACT-EXTERNAL-REFERENCE-CATALOG',
      title: 'Catalog every external URL, workbook link, and source reference',
      mode: 'govern',
      owner: 'Source owner',
      priority: 'P0',
      dependency: externalLinkNodes.map((node) => node.id).join(', '),
      acceptanceCriteria: 'Each external reference has owner, source system, credential/access path, refresh cadence, business purpose, authoritative-source decision, and migration treatment.',
      linkedItemId: externalLinkNodes[0].id,
      summary: 'External references are upstream lineage nodes, not incidental text; they must be classified before engineering rebuild.'
    }] : []),
    ...(connectionNodes.length || queryTableNodes.length ? [{
      actionId: 'ACT-CONNECTION-QUERY-TRACE',
      title: 'Trace workbook connections and query tables to terminal sources',
      mode: 'migrate',
      owner: 'Data engineering',
      priority: 'P0',
      dependency: [...connectionNodes, ...queryTableNodes].map((node) => node.id).slice(0, 8).join(', '),
      acceptanceCriteria: 'Every connection/query table has source, command/query text, destination range/table, refresh behavior, credentials owner, and downstream outputs mapped.',
      linkedItemId: connectionNodes[0]?.id || queryTableNodes[0]?.id || 'SRC-001',
      summary: 'Connection lineage drives Fivetran ingestion scope, dbt source definitions, Snowpark exception handling, and acceptance tests.'
    }] : []),
    ...(powerQueryNodes.length ? [{
      actionId: 'ACT-POWER-QUERY-DECOMPILE',
      title: 'Extract and rebuild Power Query logic',
      mode: 'rebuild',
      owner: 'Analytics engineering',
      priority: 'P0',
      dependency: powerQueryNodes.map((node) => node.id).slice(0, 8).join(', '),
      acceptanceCriteria: 'Full M code, parameters, source queries, joins, type conversions, refresh order, and output targets are documented and translated into dbt/Snowpark tests.',
      linkedItemId: powerQueryNodes[0].id,
      summary: 'Power Query steps are transformation logic and must be treated as production code.'
    }] : []),
    ...(vbaNodes.length ? [{
      actionId: 'ACT-VBA-MACRO-CALLGRAPH',
      title: 'Export VBA modules and build macro call graph',
      mode: 'rebuild',
      owner: 'Application owner',
      priority: 'P0',
      dependency: 'VBA-001',
      acceptanceCriteria: 'All modules, procedures, workbook/worksheet events, button bindings, references, file IO, refresh calls, error handlers, and generated outputs are captured with evidence IDs.',
      linkedItemId: 'VBA-001',
      summary: 'Macro code can orchestrate refresh, mutate data, produce files, and hide business rules; migration cannot be signed off without the call graph.'
    }] : []),
    ...(formulaNodes.length || definedNameNodes.length ? [{
      actionId: 'ACT-FORMULA-NAMED-RANGE-MAP',
      title: 'Map formulas, named ranges, precedents, and output cells',
      mode: 'rebuild',
      owner: 'Analytics engineering',
      priority: 'P1',
      dependency: [...formulaNodes, ...definedNameNodes].map((node) => node.id).slice(0, 10).join(', '),
      acceptanceCriteria: 'Formula blocks and named ranges have cell addresses, source precedents, downstream dependents, business meaning, expected values, and dbt/Snowpark rebuild recommendations.',
      linkedItemId: formulaNodes[0]?.id || definedNameNodes[0]?.id || 'INV-001',
      summary: 'Workbook calculations need source-controlled equivalent logic and reconciliation tests.'
    }] : []),
    {
      actionId: 'ACT-OUTPUT-OWNER-SLA-SIGNOFF',
      title: 'Confirm business outputs, owners, SLA, and failure impact',
      mode: 'govern',
      owner: 'Process owner',
      priority: 'P1',
      dependency: outputNodes.map((node) => node.id).join(', ') || 'SRC-001',
      acceptanceCriteria: 'Every critical output has owner, audience, cadence, SLA window, downstream decision/process, failure mode, recovery method, and low/base/high exposure assumptions.',
      linkedItemId: outputNodes[0]?.id || 'SRC-001',
      summary: 'Engineering actionability requires confirmed outputs and business consequences, not just workbook object inventory.'
    }
  ];

  return {
    processName: sourceName,
    businessFunction: `${sourceKindLabel(sourceKind)} discovery and migration readiness`,
    recommendation: 'Use the source-specific discovery graph to validate workbook objects, export hidden logic, and confirm output ownership before migration.',
    decisionRequired: 'Approve native metadata export for tables, named ranges, formulas with addresses, Power Query, VBA, pivots, connections, sample outputs, owners, and SLAs.',
    systemsInScope: [sourceName, ...sheetNames.slice(0, 20)],
    criticalOutputs,
    overallRiskRating: formulaNodes.length || metadataNodes.length ? 'High' : 'Medium',
    estimatedDollarExposure: {},
    executiveBrief: {},
    reportSections,
    items,
    relationships,
    artifacts: [],
    backlog,
    evidenceIndex: [
      {
        id: evidenceId,
        type: 'source_extraction',
        location: sourceName,
        description: `Extracted ${extractedText.length} characters from uploaded source evidence.`,
        relatedObject: 'SRC-001'
      }
    ],
    lineageNodes: items.map((item) => ({
      node_id: item.id,
      node_type: item.type,
      name: item.name,
      criticality: item.criticality,
      owner: item.owner,
      confidence: item.confidence,
      status: item.status
    })),
    lineageEdges: relationships.map((edge) => ({
      edge_id: edge.id,
      from_id: edge.fromId,
      to_id: edge.toId,
      edge_type: edge.type,
      confidence: edge.confidence,
      evidence_id: edge.evidenceId
    })),
    excelObjects: [
      ...sheetNodes.map((node, index) => ({
        object_id: node.id,
        workbook_id: 'SRC-001',
        sheet: node.name,
        object_type: 'sheet',
        object_name: node.name,
        evidence_id: evidenceId,
        hidden_flag: null
      })),
      ...formulaNodes.map((node, index) => ({
        object_id: node.id,
        workbook_id: 'SRC-001',
        sheet: worksheetMatches[index]?.path || '',
        object_type: 'formula_block',
        object_name: node.name,
        formula_ref: worksheetMatches[index]?.formulas?.slice(0, 500) || '',
        evidence_id: evidenceId
      })),
      ...connectionNodes.map((node) => ({
        object_id: node.id,
        workbook_id: 'SRC-001',
        object_type: 'external_connection',
        object_name: node.name,
        evidence_id: evidenceId
      })),
      ...externalLinkNodes.map((node) => ({
        object_id: node.id,
        workbook_id: 'SRC-001',
        object_type: 'external_link',
        object_name: node.name,
        evidence_id: evidenceId
      })),
      ...powerQueryNodes.map((node) => ({
        object_id: node.id,
        workbook_id: 'SRC-001',
        object_type: 'power_query',
        object_name: node.name,
        evidence_id: evidenceId
      })),
      ...queryTableNodes.map((node) => ({
        object_id: node.id,
        workbook_id: 'SRC-001',
        object_type: 'query_table',
        object_name: node.name,
        evidence_id: evidenceId
      })),
      ...vbaNodes.map((node) => ({
        object_id: node.id,
        workbook_id: 'SRC-001',
        object_type: node.type,
        object_name: node.name,
        evidence_id: evidenceId
      })),
      ...definedNameNodes.map((node) => ({
        object_id: node.id,
        workbook_id: 'SRC-001',
        object_type: 'named_range',
        object_name: node.name,
        evidence_id: evidenceId
      }))
    ],
    transformationsRules: [
      ...formulaNodes.map((node, index) => ({
      transform_id: `TR-${String(index + 1).padStart(3, '0')}`,
      location: worksheetMatches[index]?.path || node.name,
      logic_type: 'excel_formula',
      code_ref: worksheetMatches[index]?.formulas?.slice(0, 500) || '',
      description: `Formula logic discovered in ${node.name}.`,
      business_meaning: 'Requires analyst interpretation.',
      rebuild_recommendation: 'Translate to dbt SQL or Snowpark with source-controlled tests after formulas and expected outputs are confirmed.'
      })),
      ...vbaModuleEvidence.slice(0, 20).map((module, index) => ({
        transform_id: `TR-VBA-${String(index + 1).padStart(3, '0')}`,
        location: module.name,
        logic_type: 'vba_module',
        code_ref: module.code.slice(0, 2200),
        description: `VBA module ${module.name} contains ${module.procedures.length} procedures, ${module.markers.length} automation/transformation markers, ${module.sqlStrings.length} SQL candidates, and ${module.fileRefs.length} file references.`,
        business_meaning: 'Macro code may orchestrate refresh, mutate workbook state, call external sources, write outputs, or encode business rules.',
        rebuild_recommendation: 'Build call graph, classify each procedure, translate data transformations into dbt/Snowpark logic, and retain event/control behavior as tests or orchestration.'
      })),
      ...vbaModuleEvidence.flatMap((module, moduleIndex) => module.sqlStrings.slice(0, 8).map((sql, sqlIndex) => ({
        transform_id: `TR-VBA-SQL-${String(moduleIndex + 1).padStart(3, '0')}-${String(sqlIndex + 1).padStart(3, '0')}`,
        location: module.name,
        logic_type: 'vba_sql_string',
        code_ref: sql.slice(0, 2200),
        description: `SQL candidate embedded in VBA module ${module.name}.`,
        business_meaning: 'Embedded SQL can define source extraction, filtering, joins, deletes, updates, or output generation.',
        rebuild_recommendation: 'Classify SQL command, map source/target objects, and convert to governed SQL/dbt/Snowpark implementation.'
      })))
    ],
    failureRisks: [
      {
        id: 'RISK-WORKBOOK-LOGIC',
        scenario: `${sourceName} workbook logic is migrated incompletely`,
        trigger: 'Hidden formulas, macros, Power Query, named ranges, or manual zones are not exported.',
        effect: 'Migrated outputs silently diverge from current state.',
        detection: 'Reconcile sample outputs, formula inventory, and owner signoff.',
        recovery: 'Export complete workbook metadata and rebuild with tests.',
        impactedOutput: criticalOutputs[0] || sourceName,
        confidence: 78
      }
    ],
    openQuestions: [
      {
        id: 'Q-SOURCE-OWNER',
        question: `Who owns ${sourceName}, its refresh SLA, and its official business outputs?`,
        owner: 'Source owner',
        impactIfUnanswered: 'Lineage, financial exposure, and acceptance criteria remain provisional.'
      }
    ]
  };
}

function canonicalItem({ id, type, name, businessPurpose, evidenceId, confidence, criticality, upstream = [], downstream = [], failureImpact, action }) {
  return {
    id,
    type,
    name: cleanName(name),
    businessPurpose,
    owner: 'Source owner',
    evidence: [
      {
        id: evidenceId,
        type: 'source_extraction',
        location: name,
        description: `Evidence extracted for ${cleanName(name)}.`
      }
    ],
    confidence,
    criticality,
    upstream,
    downstream,
    failureImpact,
    dollarExposure: {
      low: 0,
      base: 0,
      high: 0,
      assumptions: 'Dollar exposure requires business volume, unit value, labor recovery, SLA, customer, and compliance inputs.'
    },
    recommendedAction: {
      mode: 'document',
      summary: action,
      owner: 'Source owner',
      priority: criticality === 'high' ? 'P0' : 'P1',
      acceptanceCriteria: 'Evidence, owner, confidence, upstream/downstream relationships, failure impact, and migration action are validated.'
    },
    status: 'inferred_from_extracted_evidence'
  };
}

function relationship(id, fromId, toId, type, evidenceId, confidence) {
  return {
    id,
    fromId,
    toId,
    type,
    automated: null,
    cadence: null,
    transformId: null,
    evidenceId,
    confidence,
    status: 'inferred_from_extracted_evidence'
  };
}

function reportSection(title, body, confidence, evidenceIds) {
  return { title, body, confidence, evidenceIds };
}

function firstSourceName(text) {
  return text.match(/name:\s*([^\n\r]+)/i)?.[1]?.trim() || text.match(/--- SOURCE \d+:\s*([^\n\r-]+)/i)?.[1]?.trim() || '';
}

function inferSourceKindFromEvidence(text, sourceName) {
  const combined = `${sourceName}\n${text}`.toLowerCase();
  if (/\.(xlsx|xlsm|xls)\b|workbook sheets|xl\/worksheets/.test(combined)) return 'excel';
  if (/\.(accdb|mdb)\b|access binary|string recovery/.test(combined)) return 'access';
  if (/\.(docx|doc)\b|word\/document\.xml|document text/.test(combined)) return 'word';
  if (/\.(sql|csv|tsv)\b/.test(combined)) return 'database';
  return 'mixed';
}

function sourceKindLabel(kind) {
  const labels = {
    excel: 'Excel workbook',
    access: 'Access database',
    word: 'process document',
    database: 'database extract',
    mixed: 'multi-source'
  };
  return labels[kind] || `${kind || 'source'} artifact`;
}

function deriveCriticalOutputs(payload, sourceName, sheetNames, recoveredTerms) {
  const targets = safeArray(payload.targetOutputs).filter(Boolean).map((value) => String(value));
  if (targets.length) {
    return targets.filter((value) => !isGeneratedArtifactName(value)).slice(0, 12);
  }
  const outputTerms = [...sheetNames, ...recoveredTerms]
    .filter((value) => /report|output|dashboard|summary|calendar|plan|scorecard|invoice|claim|schedule|forecast/i.test(value))
    .filter((value) => !isGeneratedArtifactName(value))
    .slice(0, 8);
  return outputTerms.length ? outputTerms : [`${sourceName} business output`];
}

function isGeneratedArtifactName(value) {
  const normalized = String(value || '').replace(/\s+/g, '_').toLowerCase();
  return /executive_decision_brief|current_state_architecture_report|technical_discovery_workbook|auto_documentation_pack|diagram_pack|financial_impact_model|action_backlog|evidence_archive|metadata_manifest|discovery_action_pack/i.test(normalized);
}

function extractListAfter(text, regex, limit) {
  const match = text.match(regex);
  if (!match) {
    return [];
  }
  return match[1]
    .split(/\s*\|\s*|\s*,\s*/)
    .map(cleanName)
    .filter(Boolean)
    .slice(0, limit);
}

function parseVbaModuleEvidence(text, limit) {
  const modules = [];
  const excerptRegex = /^VBA source excerpt\s+([^:\n]+):\n([\s\S]*?)^VBA source excerpt end\s+[^\n]*$/gmi;
  for (const match of text.matchAll(excerptRegex)) {
    const rawName = cleanName(match[1]);
    const code = cleanVbaCode(match[2] || '');
    if (!code || !/\b(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+[A-Za-z_]|Attribute\s+VB_Name|Option\s+Explicit/i.test(code)) {
      continue;
    }
    const name = cleanName(code.match(/Attribute\s+VB_Name\s*=\s*"([^"]+)"/i)?.[1] || rawName || `VBA module ${modules.length + 1}`);
    modules.push({
      name,
      code: code.slice(0, 18000),
      procedures: uniqueStrings([...code.matchAll(/^\s*(?:Private\s+|Public\s+|Friend\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)/gmi)].map((procedure) => procedure[1]), 80),
      calls: uniqueStrings([
        ...[...code.matchAll(/\bCall\s+([A-Za-z_][A-Za-z0-9_.]*)/gi)].map((call) => `Call ${call[1]}`),
        ...[...code.matchAll(/\bApplication\.Run\s+["']?([^"'\r\n,)]+)/gi)].map((call) => `Application.Run ${call[1]}`),
        ...[...code.matchAll(/\b(?:Run|OnAction)\s*=\s*["']([^"']+)/gi)].map((call) => call[1])
      ], 80),
      markers: uniqueStrings(code.split(/\r?\n/)
        .map((line) => line.trim())
        .filter((line) => /Workbook_Open|Worksheet_Change|Worksheet_Calculate|Auto_Open|RefreshAll|QueryTables?|ListObjects?|WorkbookConnection|Connections?\(|PivotTables?|Recordset|ADODB|DAO|CurrentDb|DoCmd|TransferSpreadsheet|OpenDatabase|FileSystemObject|Shell|CreateObject|GetObject|Open\s+.*\s+For\s+|Print\s+#|Write\s+#|SaveAs|ExportAsFixedFormat|CopyFromRecordset|Range\(|Cells\(|\.Formula|\.Value|Replace\(|Split\(|Trim\(|CDate\(|CDbl\(|CLng\(|DateSerial|On Error/i.test(line)),
        160),
      sqlStrings: uniqueStrings([...code.matchAll(/"([^"\r\n]*(?:SELECT|INSERT|UPDATE|DELETE|MERGE|FROM|WHERE|JOIN)[^"\r\n]*)"/gi)].map((sql) => sql[1]), 80),
      fileRefs: uniqueStrings([
        ...[...code.matchAll(/\bhttps?:\/\/[^\s"'<>|)]+/gi)].map((url) => url[0].replace(/[.,;]+$/g, '')),
        ...[...code.matchAll(/["']([A-Za-z]:\\[^"']+|\\\\[^"']+|[^"']+\.(?:csv|txt|xlsx|xlsm|xls|accdb|mdb|pdf|xml|json|sql))["']/gi)].map((ref) => ref[1])
      ], 80)
    });
    if (modules.length >= limit) {
      break;
    }
  }
  return modules;
}

function cleanVbaCode(value) {
  return String(value || '')
    .replace(/\r\n?/g, '\n')
    .replace(/[^\x09\x0A\x20-\x7E]/g, ' ')
    .replace(/[ \t]+$/gm, '')
    .replace(/\n{4,}/g, '\n\n\n')
    .trim();
}

function uniqueStrings(values, limit) {
  const seen = new Set();
  const results = [];
  for (const value of values) {
    const cleaned = cleanName(value);
    const key = cleaned.toLowerCase();
    if (cleaned && !seen.has(key)) {
      seen.add(key);
      results.push(cleaned);
    }
    if (results.length >= limit) {
      break;
    }
  }
  return results;
}

function uniqueMatches(text, regex, limit) {
  const values = [];
  const seen = new Set();
  for (const match of text.matchAll(regex)) {
    const value = cleanName(match[1]);
    const key = value.toLowerCase();
    if (value && !seen.has(key)) {
      seen.add(key);
      values.push(value);
    }
    if (values.length >= limit) {
      break;
    }
  }
  return values;
}

function recoveredEvidenceTerms(text, limit) {
  const candidates = text
    .split(/\r?\n/)
    .map(cleanName)
    .filter((line) => line.length >= 3 && line.length <= 80)
    .filter((line) => !/^(name|extension|size|status|formulas|none detected|recovered strings)$/i.test(line))
    .filter((line) => !/^microsoft\.com:|^_xlfn\./i.test(line))
    .filter((line) => /[A-Za-z]/.test(line));
  return mergePrimitiveArrays([candidates]).slice(0, limit);
}

function metadataType(path) {
  if (/connection|external/i.test(path)) return 'connection';
  if (/query/i.test(path)) return 'query';
  if (/pivot/i.test(path)) return 'pivot';
  if (/customXml/i.test(path)) return 'custom_xml';
  return 'workbook_metadata';
}

function cleanName(value) {
  return String(value || '')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/_x([0-9A-F]{4})_/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 120);
}

function isLikelyWorkbookSheetName(value) {
  return Boolean(value)
    && !/^microsoft\.com:/i.test(value)
    && !/^_xlfn\./i.test(value)
    && !/^https?:\/\//i.test(value)
    && !/schemas\.openxmlformats\.org/i.test(value);
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

function chooseSourceAwareText(deltas, key, canonicalValue, sourceValue, fallback) {
  const best = chooseBestText(deltas, key) || stringValue(canonicalValue);
  if (!best || isInternalDiscoveryText(best)) {
    return sourceValue || fallback || '';
  }
  return best;
}

function stringValue(value) {
  return String(value || '').trim();
}

function isInternalDiscoveryRecord(record) {
  const id = stringValue(record.id || record.item_id || record.node_id).toLowerCase();
  const name = stringValue(record.name || record.object_name || record.title).toLowerCase();
  return /^(graph|graph-001|pack|pack-001)$/i.test(id)
    || /^microsoft\.com:|^_xlfn\./i.test(name)
    || isInternalDiscoveryText(name);
}

function isInternalDiscoveryText(value) {
  return /distillery discovery run|proof-grade discovery graph|canonical discovery graph|canonical proof graph|discovery action pack|mash bill source set|raw source set|evidence extraction|auto documentation|diagram pack|engineering backlog/i.test(stringValue(value));
}

function isRunDiagnosticAction(record) {
  const text = `${record.actionId || record.action_id || record.id || ''} ${record.title || ''} ${record.summary || ''} ${record.linkedItemId || ''}`;
  return /ACT-[A-Z_]+-RERUN|Rerun .*Architect|Rerun .*Lead|Rerun .*Principal|Rerun .*Examiner|Rerun .*Strategist|max_output_tokens|ended incomplete|returned valid JSON/i.test(text);
}

function isRunDiagnosticRisk(record) {
  const text = `${record.id || ''} ${record.scenario || ''} ${record.trigger || ''} ${record.effect || ''}`;
  return /RISK-[A-Z_]+-INCOMPLETE|did not complete|max_output_tokens|Distillery response status incomplete|merged action pack may be shallower/i.test(text);
}

function isRunDiagnosticQuestion(record) {
  const text = `${record.id || ''} ${record.question || ''} ${record.impactIfUnanswered || ''}`;
  return /Q-[A-Z_]+-INCOMPLETE|Should .* be rerun|native metadata exports\?|contains a documented blocker for this specialist pass/i.test(text);
}

function isRunDiagnosticEvidence(record) {
  const text = `${record.id || record.evidenceId || record.evidence_id || ''} ${record.type || ''} ${record.description || ''} ${record.relatedObject || ''}`;
  return /EV-[A-Z_]+-INCOMPLETE|distillery_run_status|ended incomplete|max_output_tokens|returned valid JSON/i.test(text);
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
