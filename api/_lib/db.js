const { randomUUID } = require('node:crypto');

const { createNeonHttpSql } = require('./neon-http');

const DISTILLERY_SCHEMA = 'distillery';
const REQUIRED_TABLES = [
  'discovery_runs',
  'discovery_sources',
  'discovery_items',
  'discovery_relationships',
  'discovery_artifacts',
  'discovery_backlog',
  'discovery_evidence_index',
  'discovery_lineage_nodes',
  'discovery_lineage_edges',
  'discovery_package_manifest',
  'discovery_people_roles',
  'discovery_process_steps',
  'discovery_access_objects',
  'discovery_excel_objects',
  'discovery_word_extracts',
  'discovery_data_elements',
  'discovery_transform_rules',
  'discovery_controls_exceptions',
  'discovery_data_quality',
  'discovery_security_access',
  'discovery_schedule_sla',
  'discovery_failure_modes',
  'discovery_financial_model',
  'discovery_open_questions'
];

const SCHEMA_SQL = `
create schema if not exists distillery;

create table if not exists distillery.discovery_runs (
  id text primary key,
  created_at timestamptz not null default now(),
  source_kind text not null,
  source_name text not null,
  model text not null,
  status text not null,
  known_artifacts jsonb not null default '[]'::jsonb,
  target_outputs jsonb not null default '[]'::jsonb,
  evidence_chars integer not null default 0,
  output_text text,
  canonical_delta jsonb,
  confidence numeric,
  error text
);

create table if not exists distillery.discovery_sources (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  source_id text not null,
  source_kind text not null,
  source_name text not null,
  location text,
  extension text,
  size_bytes bigint,
  extraction_status text,
  evidence_chars integer not null default 0,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, source_id)
);

create table if not exists distillery.discovery_items (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  item_id text not null,
  item_type text not null,
  item_name text not null,
  business_purpose text,
  owner text,
  evidence jsonb not null default '[]'::jsonb,
  confidence numeric,
  criticality text,
  upstream jsonb not null default '[]'::jsonb,
  downstream jsonb not null default '[]'::jsonb,
  failure_impact text,
  dollar_exposure jsonb not null default '{}'::jsonb,
  recommended_action jsonb not null default '{}'::jsonb,
  lineage_status text,
  tags jsonb not null default '[]'::jsonb,
  payload jsonb not null,
  created_at timestamptz not null default now(),
  primary key (run_id, item_id)
);

create table if not exists distillery.discovery_relationships (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  relationship_id text not null,
  from_id text not null,
  to_id text not null,
  relationship_type text not null,
  automated boolean,
  cadence text,
  transform_id text,
  evidence_id text,
  confidence numeric,
  payload jsonb not null,
  created_at timestamptz not null default now(),
  primary key (run_id, relationship_id)
);

create table if not exists distillery.discovery_artifacts (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  artifact_id text not null,
  artifact_name text not null,
  artifact_type text not null,
  audience text,
  purpose text,
  status text not null,
  progress numeric,
  payload jsonb not null,
  created_at timestamptz not null default now(),
  primary key (run_id, artifact_id)
);

create table if not exists distillery.discovery_backlog (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  action_id text not null,
  title text not null,
  owner text,
  priority text,
  dependency text,
  due_date text,
  acceptance_criteria text,
  linked_item_id text,
  action_mode text,
  payload jsonb not null,
  created_at timestamptz not null default now(),
  primary key (run_id, action_id)
);

create table if not exists distillery.discovery_evidence_index (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  evidence_id text not null,
  evidence_type text,
  location text,
  description text,
  related_object text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, evidence_id)
);

create table if not exists distillery.discovery_lineage_nodes (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  node_id text not null,
  node_type text not null,
  node_name text not null,
  parent_id text,
  criticality text,
  owner text,
  source_of_truth_candidate text,
  confidence numeric,
  status text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, node_id)
);

create table if not exists distillery.discovery_lineage_edges (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  edge_id text not null,
  from_id text not null,
  to_id text not null,
  edge_type text not null,
  automated boolean,
  transform_id text,
  cadence text,
  confidence numeric,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, edge_id)
);

create table if not exists distillery.discovery_package_manifest (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  manifest_id text not null,
  package_name text,
  report_date timestamptz not null default now(),
  assessor text,
  scope text,
  package_version text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, manifest_id)
);

create table if not exists distillery.discovery_people_roles (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  person_id text not null,
  role text,
  responsibility text,
  decision_rights text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, person_id)
);

create table if not exists distillery.discovery_process_steps (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  step_id text not null,
  swimlane text,
  trigger text,
  input text,
  action text,
  output text,
  system text,
  sla text,
  exception text,
  control text,
  evidence_id text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, step_id)
);

create table if not exists distillery.discovery_access_objects (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  object_id text not null,
  db_id text,
  object_type text,
  object_name text,
  linked_source text,
  sql_ref text,
  vba_ref text,
  source_objects jsonb not null default '[]'::jsonb,
  target_objects jsonb not null default '[]'::jsonb,
  purpose text,
  evidence_id text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, object_id)
);

create table if not exists distillery.discovery_excel_objects (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  object_id text not null,
  workbook_id text,
  sheet text,
  object_type text,
  object_name text,
  formula_ref text,
  pq_ref text,
  vba_ref text,
  external_link text,
  refresh_order integer,
  hidden_flag boolean,
  output_area text,
  evidence_id text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, object_id)
);

create table if not exists distillery.discovery_word_extracts (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  extract_id text not null,
  doc_id text,
  section_id text,
  heading text,
  actor text,
  step_text text,
  rule text,
  exception text,
  control text,
  system_reference text,
  evidence_id text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, extract_id)
);

create table if not exists distillery.discovery_data_elements (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  element_id text not null,
  element_name text not null,
  definition text,
  datatype text,
  key_flag boolean,
  pii_flag boolean,
  source text,
  target text,
  quality_issue text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, element_id)
);

create table if not exists distillery.discovery_transform_rules (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  transform_id text not null,
  location text,
  logic_type text,
  code_ref text,
  description text,
  business_meaning text,
  rebuild_recommendation text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, transform_id)
);

create table if not exists distillery.discovery_controls_exceptions (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  control_id text not null,
  control_type text,
  owner text,
  detection_method text,
  failure_condition text,
  mitigation text,
  evidence_id text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, control_id)
);

create table if not exists distillery.discovery_data_quality (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  issue_id text not null,
  asset text,
  issue text,
  example text,
  severity text,
  impact text,
  fix text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, issue_id)
);

create table if not exists distillery.discovery_security_access (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  security_id text not null,
  asset text,
  access_group text,
  credential_method text,
  pii boolean,
  retention text,
  auditability text,
  issue text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, security_id)
);

create table if not exists distillery.discovery_schedule_sla (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  schedule_id text not null,
  process text,
  cadence text,
  cut_off text,
  refresh_window text,
  dependency text,
  critical_time text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, schedule_id)
);

create table if not exists distillery.discovery_failure_modes (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  scenario_id text not null,
  trigger text,
  effect text,
  detection text,
  recovery text,
  duration text,
  impacted_output text,
  confidence numeric,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, scenario_id)
);

create table if not exists distillery.discovery_financial_model (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  scenario_id text not null,
  process_output text,
  failure_scenario text,
  frequency numeric,
  units_affected numeric,
  dollar_per_unit numeric,
  revenue_at_risk numeric,
  margin_at_risk numeric,
  rework_hours numeric,
  labor_recovery_cost numeric,
  customer_sla_exposure numeric,
  compliance_exposure numeric,
  low_impact numeric,
  base_impact numeric,
  high_impact numeric,
  confidence text,
  assumptions text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, scenario_id)
);

create table if not exists distillery.discovery_open_questions (
  run_id text not null references distillery.discovery_runs(id) on delete cascade,
  question_id text not null,
  question text not null,
  owner text,
  due_by text,
  impact_if_unanswered text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  primary key (run_id, question_id)
);

create index if not exists discovery_runs_created_at_idx on distillery.discovery_runs (created_at desc);
create index if not exists discovery_sources_kind_idx on distillery.discovery_sources (source_kind);
create index if not exists discovery_items_type_idx on distillery.discovery_items (item_type);
create index if not exists discovery_relationships_from_idx on distillery.discovery_relationships (from_id);
create index if not exists discovery_relationships_to_idx on distillery.discovery_relationships (to_id);
create index if not exists discovery_backlog_priority_idx on distillery.discovery_backlog (priority);
create index if not exists discovery_evidence_type_idx on distillery.discovery_evidence_index (evidence_type);
create index if not exists discovery_lineage_nodes_type_idx on distillery.discovery_lineage_nodes (node_type);
create index if not exists discovery_failure_modes_output_idx on distillery.discovery_failure_modes (impacted_output);
`;

let cachedSql;
let schemaReady = false;

function databaseUrl() {
  return process.env.DATABASE_URL || process.env.POSTGRES_URL || process.env.NEON_DATABASE_URL || '';
}

function hasDatabase() {
  return Boolean(databaseUrl());
}

function schemaStatements() {
  return SCHEMA_SQL
    .split(/;\s*(?:\r?\n|$)/)
    .map((statement) => statement.trim())
    .filter(Boolean)
    .map((statement) => `${statement};`);
}

function getSql() {
  if (cachedSql) {
    return cachedSql;
  }

  const url = databaseUrl();
  try {
    const { neon } = require('@neondatabase/serverless');
    cachedSql = neon(url);
    cachedSql.driver = '@neondatabase/serverless';
  } catch (error) {
    if (error?.code !== 'MODULE_NOT_FOUND') {
      throw error;
    }
    cachedSql = createNeonHttpSql(url);
  }

  return cachedSql;
}

async function ensureSchema() {
  if (!hasDatabase()) {
    return false;
  }

  if (schemaReady) {
    return true;
  }

  const sql = getSql();
  for (const statement of schemaStatements()) {
    await sql.query(statement);
  }

  schemaReady = true;
  return true;
}

function requiredTableValuesSql() {
  return REQUIRED_TABLES.map((tableName) => `('${tableName}')`).join(', ');
}

async function databaseStatus() {
  if (!hasDatabase()) {
    return {
      configured: false,
      ready: false,
      schema: DISTILLERY_SCHEMA,
      requiredTables: REQUIRED_TABLES
    };
  }

  await ensureSchema();
  const sql = getSql();
  const rows = await sql.query(`
    select
      r.table_name,
      to_regclass('${DISTILLERY_SCHEMA}.' || r.table_name) is not null as exists
    from (values ${requiredTableValuesSql()}) as r(table_name)
    order by r.table_name;
  `);

  const missingTables = rows.filter((row) => !row.exists).map((row) => row.table_name);

  return {
    configured: true,
    ready: missingTables.length === 0,
    schema: DISTILLERY_SCHEMA,
    driver: getSql().driver || '@neondatabase/serverless',
    tableCount: rows.length - missingTables.length,
    requiredTableCount: REQUIRED_TABLES.length,
    requiredTables: REQUIRED_TABLES,
    missingTables
  };
}

function averageConfidence(items) {
  if (!Array.isArray(items) || items.length === 0) {
    return null;
  }

  const confidences = items
    .map((item) => Number(item.confidence))
    .filter((confidence) => Number.isFinite(confidence));

  if (!confidences.length) {
    return null;
  }

  return confidences.reduce((sum, confidence) => sum + confidence, 0) / confidences.length;
}

function safeArray(value) {
  return Array.isArray(value) ? value : [];
}

function json(value, fallback = null) {
  return JSON.stringify(value ?? fallback);
}

function numberOrNull(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function boolOrNull(value) {
  return typeof value === 'boolean' ? value : null;
}

function sourceRecords(requestPayload) {
  const artifacts = safeArray(requestPayload.knownArtifacts);
  if (artifacts.length) {
    return artifacts.map((artifact, index) => ({
      sourceId: `SRC-${String(index + 1).padStart(3, '0')}`,
      sourceName: String(artifact).split(/[\\/]/).pop() || String(artifact),
      location: String(artifact),
      sourceKind: requestPayload.sourceKind,
      payload: {
        knownArtifact: artifact
      }
    }));
  }

  return [{
    sourceId: 'SRC-001',
    sourceName: requestPayload.sourceName,
    location: requestPayload.sourceName,
    sourceKind: requestPayload.sourceKind,
    payload: {
      knownArtifact: requestPayload.sourceName
    }
  }];
}

function evidenceRecords(items, relationships) {
  const byId = new Map();

  for (const item of items) {
    for (const evidence of safeArray(item.evidence)) {
      const evidenceId = String(evidence.id || evidence.evidenceId || randomUUID());
      byId.set(evidenceId, {
        evidenceId,
        evidenceType: evidence.type || evidence.evidenceType || null,
        location: evidence.location || null,
        description: evidence.description || null,
        relatedObject: item.id || null,
        payload: evidence
      });
    }
  }

  for (const relationship of relationships) {
    const evidenceId = relationship.evidenceId || relationship.evidence_id;
    if (evidenceId && !byId.has(String(evidenceId))) {
      byId.set(String(evidenceId), {
        evidenceId: String(evidenceId),
        evidenceType: 'relationship',
        location: null,
        description: `Evidence referenced by relationship ${relationship.id || relationship.relationshipId || ''}`.trim(),
        relatedObject: relationship.id || relationship.relationshipId || null,
        payload: relationship
      });
    }
  }

  return [...byId.values()];
}

async function saveDiscoveryRun(requestPayload, synthesis) {
  const runId = randomUUID();
  const delta = synthesis.canonicalDelta || {};
  const items = safeArray(delta.items);
  const relationships = safeArray(delta.relationships);
  const artifacts = safeArray(delta.artifacts);
  const backlog = safeArray(delta.backlog);
  const counts = {
    items: items.length,
    relationships: relationships.length,
    artifacts: artifacts.length,
    backlog: backlog.length
  };

  let stored = false;
  try {
    stored = await ensureSchema();
  } catch (error) {
    return {
      runId,
      stored: false,
      counts,
      persistenceError: 'Workspace storage is unavailable.'
    };
  }

  if (!stored) {
    return {
      runId,
      stored: false,
      counts,
      persistenceError: 'Workspace storage is not configured.'
    };
  }

  const sql = getSql();

  await sql`
    insert into distillery.discovery_runs (
      id,
      source_kind,
      source_name,
      model,
      status,
      known_artifacts,
      target_outputs,
      evidence_chars,
      output_text,
      canonical_delta,
      confidence
    )
    values (
      ${runId},
      ${requestPayload.sourceKind},
      ${requestPayload.sourceName},
      ${synthesis.model},
      ${synthesis.canonicalDelta ? 'completed' : 'completed_unparsed'},
      ${json(requestPayload.knownArtifacts, [])}::jsonb,
      ${json(requestPayload.targetOutputs, [])}::jsonb,
      ${String(requestPayload.extractedText || '').length},
      ${synthesis.outputText || ''},
      ${json(delta, {})}::jsonb,
      ${averageConfidence(items)}
    )
  `;

  for (const source of sourceRecords(requestPayload).slice(0, 500)) {
    await sql`
      insert into distillery.discovery_sources (
        run_id,
        source_id,
        source_kind,
        source_name,
        location,
        evidence_chars,
        payload
      )
      values (
        ${runId},
        ${source.sourceId},
        ${String(source.sourceKind || requestPayload.sourceKind || 'mixed')},
        ${String(source.sourceName || 'Unknown source')},
        ${source.location || null},
        ${String(requestPayload.extractedText || '').length},
        ${json(source.payload, {})}::jsonb
      )
      on conflict (run_id, source_id) do update set
        source_kind = excluded.source_kind,
        source_name = excluded.source_name,
        location = excluded.location,
        evidence_chars = excluded.evidence_chars,
        payload = excluded.payload
    `;
  }

  for (const item of items.slice(0, 2000)) {
    const itemId = String(item.id || randomUUID());
    await sql`
      insert into distillery.discovery_items (
        run_id,
        item_id,
        item_type,
        item_name,
        business_purpose,
        owner,
        evidence,
        confidence,
        criticality,
        upstream,
        downstream,
        failure_impact,
        dollar_exposure,
        recommended_action,
        lineage_status,
        tags,
        payload
      )
      values (
        ${runId},
        ${itemId},
        ${String(item.type || 'unknown')},
        ${String(item.name || item.id || 'Untitled item')},
        ${item.businessPurpose || item.business_purpose || null},
        ${item.owner || null},
        ${json(item.evidence, [])}::jsonb,
        ${numberOrNull(item.confidence)},
        ${item.criticality || null},
        ${json(item.upstream, [])}::jsonb,
        ${json(item.downstream, [])}::jsonb,
        ${item.failureImpact || item.failure_impact || null},
        ${json(item.dollarExposure || item.dollar_exposure, {})}::jsonb,
        ${json(item.recommendedAction || item.recommended_action, {})}::jsonb,
        ${item.status || item.lineageStatus || item.lineage_status || null},
        ${json(item.tags, [])}::jsonb,
        ${json(item, {})}::jsonb
      )
      on conflict (run_id, item_id) do update set
        item_type = excluded.item_type,
        item_name = excluded.item_name,
        business_purpose = excluded.business_purpose,
        owner = excluded.owner,
        evidence = excluded.evidence,
        confidence = excluded.confidence,
        criticality = excluded.criticality,
        upstream = excluded.upstream,
        downstream = excluded.downstream,
        failure_impact = excluded.failure_impact,
        dollar_exposure = excluded.dollar_exposure,
        recommended_action = excluded.recommended_action,
        lineage_status = excluded.lineage_status,
        tags = excluded.tags,
        payload = excluded.payload
    `;

    await sql`
      insert into distillery.discovery_lineage_nodes (
        run_id,
        node_id,
        node_type,
        node_name,
        criticality,
        owner,
        confidence,
        status,
        payload
      )
      values (
        ${runId},
        ${itemId},
        ${String(item.type || 'unknown')},
        ${String(item.name || item.id || 'Untitled node')},
        ${item.criticality || null},
        ${item.owner || null},
        ${numberOrNull(item.confidence)},
        ${item.status || item.lineageStatus || item.lineage_status || null},
        ${json(item, {})}::jsonb
      )
      on conflict (run_id, node_id) do update set
        node_type = excluded.node_type,
        node_name = excluded.node_name,
        criticality = excluded.criticality,
        owner = excluded.owner,
        confidence = excluded.confidence,
        status = excluded.status,
        payload = excluded.payload
    `;
  }

  for (const relationship of relationships.slice(0, 3000)) {
    const relationshipId = String(relationship.id || relationship.relationshipId || randomUUID());
    const fromId = String(relationship.fromId || relationship.from_id || '');
    const toId = String(relationship.toId || relationship.to_id || '');
    const relationshipType = String(relationship.type || relationship.relationshipType || 'depends_on');
    await sql`
      insert into distillery.discovery_relationships (
        run_id,
        relationship_id,
        from_id,
        to_id,
        relationship_type,
        automated,
        cadence,
        transform_id,
        evidence_id,
        confidence,
        payload
      )
      values (
        ${runId},
        ${relationshipId},
        ${fromId},
        ${toId},
        ${relationshipType},
        ${boolOrNull(relationship.automated)},
        ${relationship.cadence || null},
        ${relationship.transformId || relationship.transform_id || null},
        ${relationship.evidenceId || relationship.evidence_id || null},
        ${numberOrNull(relationship.confidence)},
        ${json(relationship, {})}::jsonb
      )
      on conflict (run_id, relationship_id) do update set
        from_id = excluded.from_id,
        to_id = excluded.to_id,
        relationship_type = excluded.relationship_type,
        automated = excluded.automated,
        cadence = excluded.cadence,
        transform_id = excluded.transform_id,
        evidence_id = excluded.evidence_id,
        confidence = excluded.confidence,
        payload = excluded.payload
    `;

    await sql`
      insert into distillery.discovery_lineage_edges (
        run_id,
        edge_id,
        from_id,
        to_id,
        edge_type,
        automated,
        transform_id,
        cadence,
        confidence,
        payload
      )
      values (
        ${runId},
        ${relationshipId},
        ${fromId},
        ${toId},
        ${relationshipType},
        ${boolOrNull(relationship.automated)},
        ${relationship.transformId || relationship.transform_id || null},
        ${relationship.cadence || null},
        ${numberOrNull(relationship.confidence)},
        ${json(relationship, {})}::jsonb
      )
      on conflict (run_id, edge_id) do update set
        from_id = excluded.from_id,
        to_id = excluded.to_id,
        edge_type = excluded.edge_type,
        automated = excluded.automated,
        transform_id = excluded.transform_id,
        cadence = excluded.cadence,
        confidence = excluded.confidence,
        payload = excluded.payload
    `;
  }

  for (const evidence of evidenceRecords(items, relationships).slice(0, 3000)) {
    await sql`
      insert into distillery.discovery_evidence_index (
        run_id,
        evidence_id,
        evidence_type,
        location,
        description,
        related_object,
        payload
      )
      values (
        ${runId},
        ${String(evidence.evidenceId)},
        ${evidence.evidenceType || null},
        ${evidence.location || null},
        ${evidence.description || null},
        ${evidence.relatedObject || null},
        ${json(evidence.payload, {})}::jsonb
      )
      on conflict (run_id, evidence_id) do update set
        evidence_type = excluded.evidence_type,
        location = excluded.location,
        description = excluded.description,
        related_object = excluded.related_object,
        payload = excluded.payload
    `;
  }

  for (const artifact of artifacts.slice(0, 500)) {
    await sql`
      insert into distillery.discovery_artifacts (
        run_id,
        artifact_id,
        artifact_name,
        artifact_type,
        audience,
        purpose,
        status,
        progress,
        payload
      )
      values (
        ${runId},
        ${String(artifact.id || randomUUID())},
        ${String(artifact.name || 'Untitled artifact')},
        ${String(artifact.type || artifact.artifactType || 'report')},
        ${artifact.audience || null},
        ${artifact.purpose || null},
        ${String(artifact.status || 'draft')},
        ${numberOrNull(artifact.progress)},
        ${json(artifact, {})}::jsonb
      )
      on conflict (run_id, artifact_id) do update set
        artifact_name = excluded.artifact_name,
        artifact_type = excluded.artifact_type,
        audience = excluded.audience,
        purpose = excluded.purpose,
        status = excluded.status,
        progress = excluded.progress,
        payload = excluded.payload
    `;
  }

  for (const action of backlog.slice(0, 1000)) {
    await sql`
      insert into distillery.discovery_backlog (
        run_id,
        action_id,
        title,
        owner,
        priority,
        dependency,
        due_date,
        acceptance_criteria,
        linked_item_id,
        action_mode,
        payload
      )
      values (
        ${runId},
        ${String(action.actionId || action.action_id || randomUUID())},
        ${String(action.title || action.summary || 'Untitled action')},
        ${action.owner || null},
        ${action.priority || null},
        ${action.dependency || null},
        ${action.dueDate || action.due_date || null},
        ${action.acceptanceCriteria || action.acceptance_criteria || null},
        ${action.linkedItemId || action.linked_item_id || null},
        ${action.mode || action.actionMode || action.action_mode || null},
        ${json(action, {})}::jsonb
      )
      on conflict (run_id, action_id) do update set
        title = excluded.title,
        owner = excluded.owner,
        priority = excluded.priority,
        dependency = excluded.dependency,
        due_date = excluded.due_date,
        acceptance_criteria = excluded.acceptance_criteria,
        linked_item_id = excluded.linked_item_id,
        action_mode = excluded.action_mode,
        payload = excluded.payload
    `;
  }

  await saveSupplementalRecords(runId, delta);

  return {
    runId,
    stored: true,
    counts
  };
}

async function saveSupplementalRecords(runId, delta) {
  const sql = getSql();

  const manifest = {
    packageName: delta.packageName || 'Discovery_Action_Pack',
    processName: delta.processName || null,
    businessFunction: delta.businessFunction || null,
    recommendation: delta.recommendation || null,
    decisionRequired: delta.decisionRequired || null
  };
  await sql`
    insert into distillery.discovery_package_manifest (
      run_id,
      manifest_id,
      package_name,
      scope,
      package_version,
      payload
    )
    values (
      ${runId},
      ${'MANIFEST-001'},
      ${manifest.packageName},
      ${manifest.processName || manifest.businessFunction || null},
      ${'1.0'},
      ${json(manifest, {})}::jsonb
    )
    on conflict (run_id, manifest_id) do update set
      package_name = excluded.package_name,
      scope = excluded.scope,
      package_version = excluded.package_version,
      payload = excluded.payload
  `;

  await saveGenericRecords('discovery_failure_modes', runId, safeArray(delta.failureRisks || delta.failure_risks), {
    id: (record) => record.id || record.scenarioId || record.scenario_id,
    columns: (record) => ({
      scenario_id: record.id || record.scenarioId || record.scenario_id || randomUUID(),
      trigger: record.trigger || record.scenario || null,
      effect: record.effect || record.failureImpact || record.failure_impact || null,
      detection: record.detection || null,
      recovery: record.recovery || null,
      impacted_output: record.impactedOutput || record.impacted_output || null,
      confidence: numberOrNull(record.confidence)
    })
  });

  await saveGenericRecords('discovery_open_questions', runId, safeArray(delta.openQuestions || delta.open_questions), {
    id: (record) => record.id || record.questionId || record.question_id,
    columns: (record) => ({
      question_id: record.id || record.questionId || record.question_id || randomUUID(),
      question: record.question || record.title || 'Unresolved discovery question',
      owner: record.owner || null,
      due_by: record.dueBy || record.due_by || null,
      impact_if_unanswered: record.impactIfUnanswered || record.impact_if_unanswered || null
    })
  });
}

async function saveGenericRecords(tableName, runId, records, config) {
  if (!records.length) {
    return;
  }

  const sql = getSql();
  for (const record of records.slice(0, 1000)) {
    const columns = config.columns(record);
    const names = Object.keys(columns);
    const values = Object.values(columns);
    const assignments = names
      .filter((name) => name !== names[0])
      .map((name) => `${name} = excluded.${name}`)
      .concat(['payload = excluded.payload'])
      .join(', ');
    const query = `
      insert into distillery.${tableName} (
        run_id,
        ${names.join(', ')},
        payload
      )
      values (
        $1,
        ${names.map((_, index) => `$${index + 2}`).join(', ')},
        $${names.length + 2}::jsonb
      )
      on conflict (run_id, ${names[0]}) do update set
        ${assignments}
    `;
    await sql.query(query, [runId, ...values, json(record, {})]);
  }
}

async function listRuns(limit = 20) {
  await ensureSchema();
  const sql = getSql();
  return sql`
    select
      id,
      created_at,
      source_kind,
      source_name,
      model,
      status,
      confidence,
      evidence_chars,
      jsonb_array_length(coalesce(canonical_delta->'items', '[]'::jsonb)) as item_count,
      jsonb_array_length(coalesce(canonical_delta->'backlog', '[]'::jsonb)) as backlog_count
    from distillery.discovery_runs
    order by created_at desc
    limit ${limit}
  `;
}

async function getRun(runId) {
  await ensureSchema();
  const sql = getSql();
  const [run] = await sql`select * from distillery.discovery_runs where id = ${runId}`;
  return run || null;
}

module.exports = {
  DISTILLERY_SCHEMA,
  REQUIRED_TABLES,
  databaseStatus,
  databaseUrl,
  ensureSchema,
  getRun,
  getSql,
  hasDatabase,
  listRuns,
  saveDiscoveryRun,
  schemaStatements
};
