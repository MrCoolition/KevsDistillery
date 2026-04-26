const { randomUUID } = require('node:crypto');

let cachedSql;
let schemaReady = false;

function databaseUrl() {
  return process.env.DATABASE_URL || process.env.POSTGRES_URL || process.env.NEON_DATABASE_URL || '';
}

function hasDatabase() {
  return Boolean(databaseUrl());
}

function getSql() {
  if (!cachedSql) {
    const { neon } = require('@neondatabase/serverless');
    cachedSql = neon(databaseUrl());
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
  await sql`
    create table if not exists discovery_runs (
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
    )
  `;
  await sql`
    create table if not exists discovery_items (
      run_id text not null references discovery_runs(id) on delete cascade,
      item_id text not null,
      item_type text not null,
      item_name text not null,
      owner text,
      confidence numeric,
      criticality text,
      payload jsonb not null,
      created_at timestamptz not null default now(),
      primary key (run_id, item_id)
    )
  `;
  await sql`
    create table if not exists discovery_relationships (
      run_id text not null references discovery_runs(id) on delete cascade,
      relationship_id text not null,
      from_id text not null,
      to_id text not null,
      relationship_type text not null,
      automated boolean,
      confidence numeric,
      payload jsonb not null,
      created_at timestamptz not null default now(),
      primary key (run_id, relationship_id)
    )
  `;
  await sql`
    create table if not exists discovery_artifacts (
      run_id text not null references discovery_runs(id) on delete cascade,
      artifact_id text not null,
      artifact_name text not null,
      artifact_type text not null,
      audience text,
      status text not null,
      payload jsonb not null,
      created_at timestamptz not null default now(),
      primary key (run_id, artifact_id)
    )
  `;
  await sql`
    create table if not exists discovery_backlog (
      run_id text not null references discovery_runs(id) on delete cascade,
      action_id text not null,
      title text not null,
      owner text,
      priority text,
      due_date text,
      linked_item_id text,
      payload jsonb not null,
      created_at timestamptz not null default now(),
      primary key (run_id, action_id)
    )
  `;
  await sql`create index if not exists discovery_runs_created_at_idx on discovery_runs (created_at desc)`;
  await sql`create index if not exists discovery_items_type_idx on discovery_items (item_type)`;
  await sql`create index if not exists discovery_backlog_priority_idx on discovery_backlog (priority)`;
  schemaReady = true;
  return true;
}

async function databaseStatus() {
  if (!hasDatabase()) {
    return {
      configured: false,
      ready: false
    };
  }

  const sql = getSql();
  const [status] = await sql`
    select
      to_regclass('public.discovery_runs') is not null as has_runs,
      to_regclass('public.discovery_items') is not null as has_items,
      to_regclass('public.discovery_relationships') is not null as has_relationships,
      to_regclass('public.discovery_artifacts') is not null as has_artifacts,
      to_regclass('public.discovery_backlog') is not null as has_backlog
  `;

  const ready = Boolean(
    status?.has_runs &&
    status?.has_items &&
    status?.has_relationships &&
    status?.has_artifacts &&
    status?.has_backlog
  );

  return {
    configured: true,
    ready,
    tables: status
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

async function saveDiscoveryRun(requestPayload, synthesis) {
  const runId = randomUUID();
  const delta = synthesis.canonicalDelta || {};
  const items = Array.isArray(delta.items) ? delta.items : [];
  const relationships = Array.isArray(delta.relationships) ? delta.relationships : [];
  const artifacts = Array.isArray(delta.artifacts) ? delta.artifacts : [];
  const backlog = Array.isArray(delta.backlog) ? delta.backlog : [];
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
      persistenceError: error instanceof Error ? error.message : 'Database persistence is unavailable.'
    };
  }

  if (!stored) {
    return {
      runId,
      stored: false,
      counts
    };
  }

  const sql = getSql();

  await sql`
    insert into discovery_runs (
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
      ${JSON.stringify(requestPayload.knownArtifacts || [])}::jsonb,
      ${JSON.stringify(requestPayload.targetOutputs || [])}::jsonb,
      ${String(requestPayload.extractedText || '').length},
      ${synthesis.outputText || ''},
      ${JSON.stringify(delta)}::jsonb,
      ${averageConfidence(items)}
    )
  `;

  for (const item of items.slice(0, 1000)) {
    await sql`
      insert into discovery_items (
        run_id,
        item_id,
        item_type,
        item_name,
        owner,
        confidence,
        criticality,
        payload
      )
      values (
        ${runId},
        ${String(item.id || randomUUID())},
        ${String(item.type || 'unknown')},
        ${String(item.name || item.id || 'Untitled item')},
        ${item.owner || null},
        ${Number.isFinite(Number(item.confidence)) ? Number(item.confidence) : null},
        ${item.criticality || null},
        ${JSON.stringify(item)}::jsonb
      )
      on conflict (run_id, item_id) do update set
        item_type = excluded.item_type,
        item_name = excluded.item_name,
        owner = excluded.owner,
        confidence = excluded.confidence,
        criticality = excluded.criticality,
        payload = excluded.payload
    `;
  }

  for (const relationship of relationships.slice(0, 1500)) {
    await sql`
      insert into discovery_relationships (
        run_id,
        relationship_id,
        from_id,
        to_id,
        relationship_type,
        automated,
        confidence,
        payload
      )
      values (
        ${runId},
        ${String(relationship.id || randomUUID())},
        ${String(relationship.fromId || relationship.from_id || '')},
        ${String(relationship.toId || relationship.to_id || '')},
        ${String(relationship.type || relationship.relationshipType || 'depends_on')},
        ${typeof relationship.automated === 'boolean' ? relationship.automated : null},
        ${Number.isFinite(Number(relationship.confidence)) ? Number(relationship.confidence) : null},
        ${JSON.stringify(relationship)}::jsonb
      )
      on conflict (run_id, relationship_id) do update set
        from_id = excluded.from_id,
        to_id = excluded.to_id,
        relationship_type = excluded.relationship_type,
        automated = excluded.automated,
        confidence = excluded.confidence,
        payload = excluded.payload
    `;
  }

  for (const artifact of artifacts.slice(0, 200)) {
    await sql`
      insert into discovery_artifacts (
        run_id,
        artifact_id,
        artifact_name,
        artifact_type,
        audience,
        status,
        payload
      )
      values (
        ${runId},
        ${String(artifact.id || randomUUID())},
        ${String(artifact.name || 'Untitled artifact')},
        ${String(artifact.type || artifact.artifactType || 'report')},
        ${artifact.audience || null},
        ${String(artifact.status || 'draft')},
        ${JSON.stringify(artifact)}::jsonb
      )
      on conflict (run_id, artifact_id) do update set
        artifact_name = excluded.artifact_name,
        artifact_type = excluded.artifact_type,
        audience = excluded.audience,
        status = excluded.status,
        payload = excluded.payload
    `;
  }

  for (const action of backlog.slice(0, 500)) {
    await sql`
      insert into discovery_backlog (
        run_id,
        action_id,
        title,
        owner,
        priority,
        due_date,
        linked_item_id,
        payload
      )
      values (
        ${runId},
        ${String(action.actionId || action.action_id || randomUUID())},
        ${String(action.title || action.summary || 'Untitled action')},
        ${action.owner || null},
        ${action.priority || null},
        ${action.dueDate || action.due_date || null},
        ${action.linkedItemId || action.linked_item_id || null},
        ${JSON.stringify(action)}::jsonb
      )
      on conflict (run_id, action_id) do update set
        title = excluded.title,
        owner = excluded.owner,
        priority = excluded.priority,
        due_date = excluded.due_date,
        linked_item_id = excluded.linked_item_id,
        payload = excluded.payload
    `;
  }

  return {
    runId,
    stored: true,
    counts
  };
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
    from discovery_runs
    order by created_at desc
    limit ${limit}
  `;
}

async function getRun(runId) {
  await ensureSchema();
  const sql = getSql();
  const [run] = await sql`select * from discovery_runs where id = ${runId}`;
  return run || null;
}

module.exports = {
  databaseStatus,
  databaseUrl,
  ensureSchema,
  getRun,
  hasDatabase,
  listRuns,
  saveDiscoveryRun
};
