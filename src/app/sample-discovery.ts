import { DISCOVERY_AGENT_CONTRACT, OPENAI_DISCOVERY_MODEL } from './ai-orchestration';
import { DiscoveryModel } from './discovery-model';

export const discoveryModel: DiscoveryModel = {
  packageName: 'Discovery_Action_Pack',
  processName: 'Finance Close Revenue Distillation',
  businessFunction: 'Finance operations, billing assurance, executive reporting',
  recommendation: 'Stabilize the current close path, migrate authoritative flows to Snowflake, rebuild transformations in dbt, and orchestrate recurring extracts through Fivetran with Snowpark validation hooks.',
  decisionRequired: 'Approve migration wave 1 for billing, cash application, and close reporting dependencies; assign finance data owner for unresolved manual overrides.',
  systemsInScope: ['Access close mart', 'Excel revenue model', 'Shared drive extracts', 'ERP billing export', 'Collections workbook', 'Finance close memo'],
  criticalOutputs: ['Daily revenue flash', 'Month-end close package', 'Cash timing dashboard', 'Executive margin bridge'],
  overallRiskRating: 'high',
  estimatedDollarExposure: {
    low: 420000,
    base: 1850000,
    high: 6700000,
    assumptions: 'Exposure combines delayed billing, margin restatement risk, manual recovery labor, and SLA credits across five priced failure modes.'
  },
  items: [
    {
      id: 'SYS-001',
      type: 'system',
      name: 'ERP Billing',
      businessPurpose: 'Authoritative invoice, order, customer, and billing event source feeding revenue recognition and cash timing.',
      owner: 'Finance Systems',
      evidence: [
        {
          id: 'EV-SQL-001',
          type: 'SQL',
          location: '08_Evidence_Archive/SQL/erp_billing_extract.sql',
          description: 'Source extract SQL with invoice, order, and customer keys.'
        }
      ],
      confidence: 94,
      criticality: 'critical',
      upstream: [],
      downstream: ['DB-001', 'FILE-001'],
      failureImpact: 'No authoritative billing events reach close reporting, causing delayed revenue flash and cash application ambiguity.',
      dollarExposure: {
        low: 310000,
        base: 1400000,
        high: 5200000,
        assumptions: 'Three-day close window with affected invoices priced against average daily billed revenue and margin.'
      },
      recommendedAction: {
        mode: 'migrate',
        summary: 'Use Fivetran to land ERP billing entities directly into Snowflake with source freshness checks.',
        owner: 'Data Platform',
        priority: 'P0',
        acceptanceCriteria: 'Billing entities land in Snowflake with row-count parity, key constraints, and freshness SLA alerts.'
      },
      status: 'confirmed',
      tags: ['source-of-record', 'billing', 'fivetran']
    },
    {
      id: 'FILE-001',
      type: 'file',
      name: 'BillingExport_Daily.csv',
      businessPurpose: 'Daily flat-file bridge from ERP billing into Access when direct integration is unavailable.',
      owner: 'Revenue Analyst',
      evidence: [
        {
          id: 'EV-SCREEN-001',
          type: 'Screenshot',
          location: '08_Evidence_Archive/Screenshots/billing_export_folder.png',
          description: 'Shared drive export folder with daily CSV files.'
        }
      ],
      confidence: 82,
      criticality: 'high',
      upstream: ['SYS-001'],
      downstream: ['DB-001'],
      failureImpact: 'Late or malformed CSV blocks the Access import macro and shifts all downstream close reporting.',
      dollarExposure: {
        low: 120000,
        base: 620000,
        high: 1900000,
        assumptions: 'Late extract delays close package by one business day and triggers manual reconciliation.'
      },
      recommendedAction: {
        mode: 'retire',
        summary: 'Retire shared-drive CSV handoff once Fivetran source ingestion is live.',
        owner: 'Finance Systems',
        priority: 'P1',
        acceptanceCriteria: 'No production control depends on the shared-drive CSV for two consecutive close cycles.'
      },
      status: 'confirmed',
      tags: ['flat-file', 'manual-handoff', 'shadow-source']
    },
    {
      id: 'DB-001',
      type: 'database',
      name: 'CloseMart.accdb',
      businessPurpose: 'Access staging and reporting mart that joins billing exports, customer adjustments, and revenue rules.',
      owner: 'Revenue Analyst',
      evidence: [
        {
          id: 'EV-VBA-001',
          type: 'VBA',
          location: '08_Evidence_Archive/VBA/CloseMart_ModuleRefresh.bas',
          description: 'Refresh_All macro imports CSV, runs saved queries, and exports month-end tables.'
        }
      ],
      confidence: 89,
      criticality: 'critical',
      upstream: ['FILE-001', 'DOC-001'],
      downstream: ['QRY-001', 'CTRL-001'],
      failureImpact: 'Access refresh failure stops revenue flash production and hides transformation errors behind local desktop execution.',
      dollarExposure: {
        low: 220000,
        base: 1100000,
        high: 4100000,
        assumptions: 'Exposure based on one failed close run, finance recovery hours, delayed billing review, and reissue risk.'
      },
      recommendedAction: {
        mode: 'rebuild',
        summary: 'Rebuild saved queries as dbt models with explicit tests, lineage, and Snowpark reconciliation checks.',
        owner: 'Analytics Engineering',
        priority: 'P0',
        acceptanceCriteria: 'All critical Access queries have equivalent dbt models, passing row parity and metric reconciliation.'
      },
      status: 'confirmed',
      tags: ['access', 'desktop-risk', 'dbt']
    },
    {
      id: 'QRY-001',
      type: 'query',
      name: 'qry_RevenueFlash_Final',
      businessPurpose: 'Final revenue flash transformation joining invoices, adjustments, customer tier, and recognition exclusions.',
      owner: 'Revenue Analyst',
      evidence: [
        {
          id: 'EV-SQL-002',
          type: 'SQL',
          location: '08_Evidence_Archive/SQL/qry_RevenueFlash_Final.sql',
          description: 'Saved Access SQL with joins, hardcoded exclusions, and output fields.'
        }
      ],
      confidence: 91,
      criticality: 'critical',
      upstream: ['DB-001', 'WB-001'],
      downstream: ['OUT-001', 'OUT-004'],
      failureImpact: 'Incorrect transformation changes daily revenue, margin bridge, and executive reporting decisions.',
      dollarExposure: {
        low: 180000,
        base: 920000,
        high: 2800000,
        assumptions: 'Wrong-data scenario priced against impacted flash decisions and margin sensitivity.'
      },
      recommendedAction: {
        mode: 'govern',
        summary: 'Convert hardcoded exclusion logic into governed dbt seed and add metric tests for flash totals.',
        owner: 'Analytics Engineering',
        priority: 'P0',
        acceptanceCriteria: 'Exclusions are version controlled, owner-approved, and covered by automated dbt tests.'
      },
      status: 'confirmed',
      tags: ['business-logic', 'transformation', 'metric-risk']
    },
    {
      id: 'WB-001',
      type: 'workbook',
      name: 'Revenue_Close_Model.xlsm',
      businessPurpose: 'Excel model used for manual overrides, Power Query refresh, margin bridge, and finance review package.',
      owner: 'Finance Planning',
      evidence: [
        {
          id: 'EV-PQ-001',
          type: 'PowerQuery_M',
          location: '08_Evidence_Archive/PowerQuery_M/Revenue_Close_Model_Queries.m',
          description: 'Power Query M scripts, external links, and refresh sequencing.'
        },
        {
          id: 'EV-VBA-002',
          type: 'VBA',
          location: '08_Evidence_Archive/VBA/Workbook_Open.bas',
          description: 'Workbook open event triggers refresh and export button registration.'
        }
      ],
      confidence: 86,
      criticality: 'critical',
      upstream: ['DB-001', 'DOC-001'],
      downstream: ['QRY-001', 'OUT-002', 'OUT-004'],
      failureImpact: 'Manual overrides or stale Power Query results can misstate margin bridge and close package.',
      dollarExposure: {
        low: 90000,
        base: 530000,
        high: 1600000,
        assumptions: 'Wrong-data and unauditable scenarios priced against rework and executive reporting dependency.'
      },
      recommendedAction: {
        mode: 'automate',
        summary: 'Move manual override capture into governed Snowflake table with approvals and audit columns.',
        owner: 'Finance Planning',
        priority: 'P1',
        acceptanceCriteria: 'Overrides are submitted through governed workflow and no longer edited directly in protected workbook cells.'
      },
      status: 'partial',
      tags: ['excel', 'power-query', 'manual-overrides']
    },
    {
      id: 'DOC-001',
      type: 'document',
      name: 'Finance Close Memo.docx',
      businessPurpose: 'Process documentation describing close cadence, approvals, control points, exception paths, and SLA windows.',
      owner: 'Controller',
      evidence: [
        {
          id: 'EV-DOC-001',
          type: 'Document_Extract',
          location: '08_Evidence_Archive/Document_Extracts/Finance_Close_Memo_sections.json',
          description: 'Extracted headings, actors, controls, exceptions, and deadlines.'
        }
      ],
      confidence: 78,
      criticality: 'high',
      upstream: [],
      downstream: ['DB-001', 'WB-001', 'CTRL-001'],
      failureImpact: 'Outdated controls and undocumented exception paths slow migration design decisions.',
      dollarExposure: {
        low: 25000,
        base: 110000,
        high: 360000,
        assumptions: 'Rework exposure based on unresolved requirements and control redesign effort.'
      },
      recommendedAction: {
        mode: 'govern',
        summary: 'Generate current-state documentation from canonical graph and route controller signoff.',
        owner: 'Controller',
        priority: 'P2',
        acceptanceCriteria: 'Controller approves generated process narrative, controls, exceptions, and open questions.'
      },
      status: 'inferred',
      tags: ['word', 'controls', 'documentation']
    },
    {
      id: 'CTRL-001',
      type: 'control',
      name: 'Revenue Flash Variance Check',
      businessPurpose: 'Detects daily revenue variance beyond threshold before executive distribution.',
      owner: 'Controller',
      evidence: [
        {
          id: 'EV-CTRL-001',
          type: 'Control_Log',
          location: '08_Evidence_Archive/Document_Extracts/variance_control_log.csv',
          description: 'Sample control log with reviewer initials and variance threshold outcomes.'
        }
      ],
      confidence: 73,
      criticality: 'high',
      upstream: ['DB-001', 'DOC-001'],
      downstream: ['OUT-001'],
      failureImpact: 'Flash report can distribute with unreviewed anomalies and no auditable signoff.',
      dollarExposure: {
        low: 60000,
        base: 240000,
        high: 980000,
        assumptions: 'Control failure priced against correction cycle, executive decision delay, and customer credit exposure.'
      },
      recommendedAction: {
        mode: 'stabilize',
        summary: 'Implement automated variance threshold checks in Snowpark and publish signoff evidence.',
        owner: 'Data Platform',
        priority: 'P1',
        acceptanceCriteria: 'Variance checks run on every refresh and attach pass/fail evidence to the generated pack.'
      },
      status: 'partial',
      tags: ['control', 'snowpark', 'audit']
    },
    {
      id: 'OUT-001',
      type: 'output',
      name: 'Daily Revenue Flash',
      businessPurpose: 'Executive daily view of recognized revenue, anomalies, and billing health.',
      owner: 'Controller',
      evidence: [
        {
          id: 'EV-SCREEN-002',
          type: 'Screenshot',
          location: '08_Evidence_Archive/Screenshots/revenue_flash_output.png',
          description: 'Final flash report distributed to executive finance list.'
        }
      ],
      confidence: 88,
      criticality: 'critical',
      upstream: ['QRY-001', 'CTRL-001'],
      downstream: [],
      failureImpact: 'Leadership operates on stale or wrong revenue position during close and billing exception windows.',
      dollarExposure: {
        low: 260000,
        base: 1250000,
        high: 3900000,
        assumptions: 'Output-level exposure aggregates billing, margin, rework, and decision-delay risk.'
      },
      recommendedAction: {
        mode: 'migrate',
        summary: 'Certify this as a Snowflake governed output with dbt lineage, tests, and signed release evidence.',
        owner: 'Controller',
        priority: 'P0',
        acceptanceCriteria: 'Certified output has full upstream lineage, test history, business owner signoff, and distribution audit.'
      },
      status: 'confirmed',
      tags: ['critical-output', 'executive', 'certified']
    },
    {
      id: 'OUT-002',
      type: 'output',
      name: 'Month-End Close Package',
      businessPurpose: 'Finance close workbook package used for controller review and accounting handoff.',
      owner: 'Finance Planning',
      evidence: [
        {
          id: 'EV-SCREEN-003',
          type: 'Screenshot',
          location: '08_Evidence_Archive/Screenshots/close_package_tabs.png',
          description: 'Workbook tabs, hidden sheets, and final close package output.'
        }
      ],
      confidence: 81,
      criticality: 'critical',
      upstream: ['WB-001'],
      downstream: [],
      failureImpact: 'Close signoff is delayed or based on manually altered workbook logic.',
      dollarExposure: {
        low: 140000,
        base: 690000,
        high: 2100000,
        assumptions: 'Close delay and rework exposure using finance labor and reporting dependency proxy.'
      },
      recommendedAction: {
        mode: 'rebuild',
        summary: 'Rebuild package generation as governed Snowflake output and preserve Excel only as review layer.',
        owner: 'Finance Planning',
        priority: 'P1',
        acceptanceCriteria: 'Close package data is generated from Snowflake-certified tables with versioned extracts.'
      },
      status: 'partial',
      tags: ['critical-output', 'excel-output', 'close']
    },
    {
      id: 'OUT-004',
      type: 'output',
      name: 'Executive Margin Bridge',
      businessPurpose: 'Explains movement in margin across customer tier, product mix, exclusions, and manual adjustments.',
      owner: 'Finance Planning',
      evidence: [
        {
          id: 'EV-PROFILE-001',
          type: 'Profile',
          location: '08_Evidence_Archive/Profile/margin_bridge_profile.json',
          description: 'Column profile, formula ranges, and output reconciliation sample.'
        }
      ],
      confidence: 76,
      criticality: 'high',
      upstream: ['QRY-001', 'WB-001'],
      downstream: [],
      failureImpact: 'Margin story can be wrong or unauditable during executive review.',
      dollarExposure: {
        low: 70000,
        base: 390000,
        high: 1350000,
        assumptions: 'Decision-support proxy based on margin influenced and executive reporting dependency.'
      },
      recommendedAction: {
        mode: 'govern',
        summary: 'Convert formula blocks into transparent dbt metrics and reconcile bridge totals in Snowpark.',
        owner: 'Analytics Engineering',
        priority: 'P2',
        acceptanceCriteria: 'Each bridge component maps to a governed metric with formula parity evidence.'
      },
      status: 'partial',
      tags: ['margin', 'metric', 'decision-support']
    }
  ],
  relationships: [
    {
      id: 'EDGE-001',
      fromId: 'SYS-001',
      toId: 'FILE-001',
      type: 'exports_to',
      automated: false,
      cadence: 'Daily 5:30 AM',
      confidence: 82,
      evidenceId: 'EV-SCREEN-001'
    },
    {
      id: 'EDGE-002',
      fromId: 'FILE-001',
      toId: 'DB-001',
      type: 'imports_from',
      automated: true,
      cadence: 'Daily 5:45 AM',
      confidence: 87,
      transformId: 'TR-001',
      evidenceId: 'EV-VBA-001'
    },
    {
      id: 'EDGE-003',
      fromId: 'DB-001',
      toId: 'QRY-001',
      type: 'transforms',
      automated: true,
      cadence: 'Daily 6:00 AM',
      confidence: 91,
      transformId: 'TR-002',
      evidenceId: 'EV-SQL-002'
    },
    {
      id: 'EDGE-004',
      fromId: 'WB-001',
      toId: 'QRY-001',
      type: 'reads_from',
      automated: false,
      cadence: 'Close days',
      confidence: 74,
      transformId: 'TR-003',
      evidenceId: 'EV-PQ-001'
    },
    {
      id: 'EDGE-005',
      fromId: 'QRY-001',
      toId: 'OUT-001',
      type: 'writes_to',
      automated: true,
      cadence: 'Daily 6:15 AM',
      confidence: 88,
      evidenceId: 'EV-SCREEN-002'
    },
    {
      id: 'EDGE-006',
      fromId: 'CTRL-001',
      toId: 'OUT-001',
      type: 'approves',
      automated: false,
      cadence: 'Daily before distribution',
      confidence: 73,
      evidenceId: 'EV-CTRL-001'
    },
    {
      id: 'EDGE-007',
      fromId: 'DOC-001',
      toId: 'CTRL-001',
      type: 'documented_by',
      automated: false,
      cadence: 'Quarterly refresh',
      confidence: 78,
      evidenceId: 'EV-DOC-001'
    },
    {
      id: 'EDGE-008',
      fromId: 'WB-001',
      toId: 'OUT-002',
      type: 'writes_to',
      automated: false,
      cadence: 'Monthly close',
      confidence: 81,
      evidenceId: 'EV-SCREEN-003'
    },
    {
      id: 'EDGE-009',
      fromId: 'QRY-001',
      toId: 'OUT-004',
      type: 'writes_to',
      automated: true,
      cadence: 'Monthly close',
      confidence: 76,
      evidenceId: 'EV-PROFILE-001'
    },
    {
      id: 'EDGE-010',
      fromId: 'WB-001',
      toId: 'OUT-004',
      type: 'transforms',
      automated: false,
      cadence: 'Monthly close',
      confidence: 72,
      transformId: 'TR-004',
      evidenceId: 'EV-PROFILE-001'
    }
  ],
  artifacts: [
    {
      id: '01',
      name: '01_Executive_Decision_Brief.pdf',
      audience: 'Leadership',
      purpose: 'What this process does, why it matters, risk, dollars, and decision needed.',
      progress: 92,
      sourceModel: 'canonical graph'
    },
    {
      id: '02',
      name: '02_Current_State_Architecture_Report.pdf',
      audience: 'Business + IT',
      purpose: 'Concise narrative of how the process and data actually work today.',
      progress: 78,
      sourceModel: 'canonical graph'
    },
    {
      id: '03',
      name: '03_Technical_Discovery_Workbook.xlsx',
      audience: 'Engineers / analysts',
      purpose: 'Structured inventory of objects, logic, lineage, risks, and actions.',
      progress: 86,
      sourceModel: 'canonical graph'
    },
    {
      id: '04',
      name: '04_Auto_Documentation_Pack',
      audience: 'Automation / governance',
      purpose: 'Machine-readable current-state documentation generated from discovery.',
      progress: 81,
      sourceModel: 'canonical graph'
    },
    {
      id: '05',
      name: '05_Diagram_Pack',
      audience: 'All audiences',
      purpose: 'Visual truth of process flow, data flow, lineage, dependencies, and failures.',
      progress: 74,
      sourceModel: 'canonical graph'
    },
    {
      id: '06',
      name: '06_Financial_Impact_Model.xlsx',
      audience: 'Leadership / finance',
      purpose: 'Dollar value at risk when the process fails, is late, or is wrong.',
      progress: 83,
      sourceModel: 'canonical graph'
    },
    {
      id: '07',
      name: '07_Action_Backlog.csv',
      audience: 'Delivery teams',
      purpose: 'Ready-to-work remediation and modernization actions.',
      progress: 88,
      sourceModel: 'canonical graph'
    },
    {
      id: '08',
      name: '08_Evidence_Archive',
      audience: 'Audit / project team',
      purpose: 'Proof behind every finding.',
      progress: 69,
      sourceModel: 'canonical graph'
    },
    {
      id: '09',
      name: '09_Metadata_Manifest.json',
      audience: 'Automation / governance',
      purpose: 'Traceable metadata manifest for the complete package.',
      progress: 90,
      sourceModel: 'canonical graph'
    }
  ],
  confidenceAreas: [
    {
      area: 'Access objects',
      reviewed: 'Tables, saved queries, macros, VBA, reports, linked sources',
      coverage: 89,
      confidence: 87,
      blocked: 'Startup form dependency needs direct desktop inspection.'
    },
    {
      area: 'Excel logic',
      reviewed: 'Hidden sheets, named ranges, formulas, Power Query, VBA, pivots',
      coverage: 76,
      confidence: 78,
      blocked: 'Very hidden override sheet requires workbook password owner.'
    },
    {
      area: 'Word process docs',
      reviewed: 'Headings, actors, rules, exceptions, controls, deadlines',
      coverage: 84,
      confidence: 80,
      blocked: 'One approval path conflicts with interview notes.'
    },
    {
      area: 'Financial exposure',
      reviewed: 'Revenue, margin, labor, SLA, and audit exposure assumptions',
      coverage: 71,
      confidence: 73,
      blocked: 'Customer credit proxy needs finance approval.'
    }
  ],
  failureRisks: [
    {
      id: 'RISK-001',
      scenario: 'Process does not run',
      impactedOutput: 'Daily Revenue Flash',
      detection: 'No report by 7:00 AM distribution window',
      recovery: 'Manual ERP extract, Access import rerun, controller review',
      exposure: {
        low: 260000,
        base: 1250000,
        high: 3900000,
        assumptions: 'Daily revenue and executive reporting dependency with one-day delay.'
      },
      confidence: 82
    },
    {
      id: 'RISK-002',
      scenario: 'Process runs with wrong data',
      impactedOutput: 'Executive Margin Bridge',
      detection: 'Variance threshold or executive review challenge',
      recovery: 'Workbook formula tracing, manual reissue, signoff recertification',
      exposure: {
        low: 180000,
        base: 920000,
        high: 2800000,
        assumptions: 'Margin sensitivity plus manual rework and reissue risk.'
      },
      confidence: 76
    },
    {
      id: 'RISK-003',
      scenario: 'Process runs but cannot be audited',
      impactedOutput: 'Month-End Close Package',
      detection: 'Controller or audit request cannot tie output to evidence',
      recovery: 'Reconstruct refresh sequence and override history manually',
      exposure: {
        low: 80000,
        base: 310000,
        high: 980000,
        assumptions: 'Auditability failure and rework proxy across close team.'
      },
      confidence: 69
    }
  ],
  backlog: [
    {
      actionId: 'ACT-001',
      title: 'Land ERP billing tables in Snowflake through Fivetran',
      owner: 'Data Platform',
      priority: 'P0',
      dependency: 'ERP service account approval',
      dueDate: '2026-05-15',
      acceptanceCriteria: 'Tables land with row parity, key coverage, and freshness checks.',
      linkedItemId: 'SYS-001',
      mode: 'migrate'
    },
    {
      actionId: 'ACT-002',
      title: 'Rebuild qry_RevenueFlash_Final as dbt model',
      owner: 'Analytics Engineering',
      priority: 'P0',
      dependency: 'ACT-001',
      dueDate: '2026-05-29',
      acceptanceCriteria: 'dbt model reconciles to Access output within approved tolerance for two close cycles.',
      linkedItemId: 'QRY-001',
      mode: 'rebuild'
    },
    {
      actionId: 'ACT-003',
      title: 'Automate variance control in Snowpark',
      owner: 'Data Platform',
      priority: 'P1',
      dependency: 'Threshold owner signoff',
      dueDate: '2026-06-05',
      acceptanceCriteria: 'Snowpark check emits pass/fail evidence and blocks uncertified distribution.',
      linkedItemId: 'CTRL-001',
      mode: 'stabilize'
    },
    {
      actionId: 'ACT-004',
      title: 'Move manual overrides into governed Snowflake workflow',
      owner: 'Finance Planning',
      priority: 'P1',
      dependency: 'Override taxonomy approval',
      dueDate: '2026-06-12',
      acceptanceCriteria: 'Overrides include requester, approver, reason code, effective date, and audit history.',
      linkedItemId: 'WB-001',
      mode: 'automate'
    },
    {
      actionId: 'ACT-005',
      title: 'Generate controller-approved current-state documentation',
      owner: 'Controller',
      priority: 'P2',
      dependency: 'Resolve conflicting approval path',
      dueDate: '2026-06-19',
      acceptanceCriteria: 'Report sections, controls, exceptions, and open questions are signed by process owner.',
      linkedItemId: 'DOC-001',
      mode: 'govern'
    }
  ],
  extractionCapabilities: [
    {
      source: 'Access',
      autoExtracts: ['database file metadata', 'linked tables', 'row counts', 'relationships', 'saved SQL', 'macros', 'VBA', 'forms', 'reports', 'hidden objects'],
      currentStateOutputs: ['object inventory', 'transform rules', 'dependency map', 'failure modes'],
      readiness: 86
    },
    {
      source: 'Excel',
      autoExtracts: ['hidden sheets', 'tables', 'named ranges', 'formula regions', 'Power Query M', 'external links', 'pivots', 'VBA', 'manual inputs', 'hardcoded overrides'],
      currentStateOutputs: ['formula register', 'refresh timeline', 'manual override map', 'generated outputs'],
      readiness: 79
    },
    {
      source: 'Word',
      autoExtracts: ['metadata', 'headings', 'process steps', 'actors', 'inputs', 'outputs', 'approvals', 'rules', 'exceptions', 'controls', 'SLAs'],
      currentStateOutputs: ['process narrative', 'control narrative', 'open questions', 'system references'],
      readiness: 83
    },
    {
      source: 'Database',
      autoExtracts: ['schemas', 'tables', 'columns', 'keys', 'indexes', 'row counts', 'profiles', 'dependency hints', 'refresh markers'],
      currentStateOutputs: ['data dictionary', 'lineage nodes', 'lineage edges', 'source-of-truth assessment'],
      readiness: 88
    },
    {
      source: 'Interview',
      autoExtracts: ['tribal rules', 'manual paths', 'owners', 'approval reality', 'recovery steps', 'business impact assumptions'],
      currentStateOutputs: ['confidence gaps', 'decision log', 'financial assumptions', 'blocked lineage branches'],
      readiness: 72
    }
  ]
};

export const aiReadiness = {
  model: OPENAI_DISCOVERY_MODEL,
  contract: DISCOVERY_AGENT_CONTRACT,
  serverBoundary: 'All LLM calls run through a backend discovery agent. The Angular client receives normalized graph deltas only.'
};
