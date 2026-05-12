create extension if not exists pgcrypto with schema public;

create schema if not exists data_discovery;

set search_path = data_discovery, public;

create table if not exists schema_migrations (
  migration_id text primary key,
  description text not null,
  applied_at timestamptz not null default now()
);

create table if not exists discovery_jobs (
  job_id uuid primary key,
  package_name text not null,
  source_process_name text not null,
  status text not null default 'QUEUED',
  generated_date date not null,
  source_file_count integer not null default 0,
  object_count integer not null default 0,
  query_count integer not null default 0,
  macro_count integer not null default 0,
  linked_source_count integer not null default 0,
  lineage_blocker_count integer not null default 0,
  qa_status text not null default 'PENDING',
  canonical_model jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists discovery_source_files (
  source_id text primary key,
  job_id uuid not null references discovery_jobs(job_id) on delete cascade,
  file_name text not null,
  file_type text not null,
  extension text not null,
  file_size_bytes bigint not null,
  sha256 text not null,
  evidence_id text not null,
  created_at timestamptz not null default now()
);

create table if not exists discovery_evidence_items (
  evidence_id text not null,
  job_id uuid not null references discovery_jobs(job_id) on delete cascade,
  title text not null,
  category text not null,
  relative_path text not null,
  source_file text not null,
  confidence text not null,
  summary text not null,
  created_at timestamptz not null default now(),
  primary key (job_id, evidence_id)
);

create table if not exists lineage_nodes (
  node_id text not null,
  job_id uuid not null references discovery_jobs(job_id) on delete cascade,
  node_type text not null,
  name text not null,
  description text not null,
  source_file text not null,
  business_purpose text not null,
  owner_status text not null,
  criticality text not null,
  confidence text not null,
  evidence_id text not null,
  recommended_action text not null,
  failure_impact text not null,
  dollar_exposure text not null,
  created_at timestamptz not null default now(),
  primary key (job_id, node_id)
);

create table if not exists lineage_edges (
  edge_id text not null,
  job_id uuid not null references discovery_jobs(job_id) on delete cascade,
  from_node_id text not null,
  to_node_id text not null,
  edge_type text not null,
  description text not null,
  automated_flag text not null,
  transformation_id text,
  cadence text,
  confidence text not null,
  evidence_id text not null,
  created_at timestamptz not null default now(),
  primary key (job_id, edge_id)
);

create table if not exists action_items (
  action_id text not null,
  job_id uuid not null references discovery_jobs(job_id) on delete cascade,
  title text not null,
  description text not null,
  source_asset text not null,
  owner_role text not null,
  recommended_owner text not null,
  action_type text not null,
  priority text not null,
  severity text not null,
  dependency text not null,
  due_date_or_phase text not null,
  acceptance_criteria text not null,
  evidence_id text not null,
  related_risk text not null,
  expected_business_value text not null,
  status text not null default 'Not Started',
  created_at timestamptz not null default now(),
  primary key (job_id, action_id)
);

create table if not exists financial_exposures (
  exposure_id uuid primary key default gen_random_uuid(),
  job_id uuid not null references discovery_jobs(job_id) on delete cascade,
  process_or_output text not null,
  failure_scenario text not null,
  frequency text not null,
  units_affected numeric not null,
  dollar_per_unit numeric not null,
  revenue_at_risk numeric not null,
  margin_percent numeric not null,
  margin_at_risk numeric not null,
  rework_hours numeric not null,
  labor_rate numeric not null,
  labor_recovery_cost numeric not null,
  customer_sla_exposure numeric not null,
  compliance_exposure numeric not null,
  cash_timing_cost numeric not null,
  low_impact numeric not null,
  base_impact numeric not null,
  high_impact numeric not null,
  annualized_low numeric not null,
  annualized_base numeric not null,
  annualized_high numeric not null,
  confidence text not null,
  assumptions text not null,
  evidence_id text not null,
  finance_validation_needed text not null,
  created_at timestamptz not null default now()
);

create table if not exists qa_checks (
  qa_id text not null,
  job_id uuid not null references discovery_jobs(job_id) on delete cascade,
  check_text text not null,
  status text not null,
  evidence_id text not null,
  notes text not null,
  created_at timestamptz not null default now(),
  primary key (job_id, qa_id)
);

create table if not exists dossier_packages (
  package_id uuid primary key default gen_random_uuid(),
  job_id uuid not null references discovery_jobs(job_id) on delete cascade,
  package_name text not null,
  zip_bytes bytea not null,
  byte_size bigint not null,
  qa_status text not null,
  created_at timestamptz not null default now()
);

create table if not exists ai_usage_events (
  usage_id uuid primary key default gen_random_uuid(),
  job_id uuid not null references discovery_jobs(job_id) on delete cascade,
  model text not null,
  requests integer not null default 0,
  input_tokens integer not null default 0,
  cached_input_tokens integer not null default 0,
  billable_input_tokens integer not null default 0,
  output_tokens integer not null default 0,
  reasoning_tokens integer not null default 0,
  total_tokens integer not null default 0,
  input_cost_usd numeric not null default 0,
  cached_input_cost_usd numeric not null default 0,
  output_cost_usd numeric not null default 0,
  total_cost_usd numeric not null default 0,
  cache_savings_usd numeric not null default 0,
  cache_hit_rate numeric not null default 0,
  pricing_source text not null,
  optimization_note text not null,
  created_at timestamptz not null default now()
);

create index if not exists idx_discovery_jobs_created_at on discovery_jobs(created_at desc);
create index if not exists idx_discovery_jobs_package_name on discovery_jobs(package_name);
create index if not exists idx_source_files_job_id on discovery_source_files(job_id);
create index if not exists idx_lineage_nodes_job_type on lineage_nodes(job_id, node_type);
create index if not exists idx_lineage_edges_job_type on lineage_edges(job_id, edge_type);
create index if not exists idx_action_items_priority on action_items(job_id, priority, severity);
create index if not exists idx_packages_job_id on dossier_packages(job_id);
create index if not exists idx_ai_usage_events_job_id on ai_usage_events(job_id);

insert into schema_migrations (migration_id, description)
values ('2026-05-08_001_initial_data_discovery_schema', 'Initial Data Source Discovery Distillery production schema')
on conflict (migration_id) do nothing;

comment on table discovery_jobs is 'One dossier generation run and its canonical JSONB discovery model.';
comment on table discovery_source_files is 'Uploaded source file metadata and checksum records.';
comment on table discovery_evidence_items is 'Evidence index rows mapped to generated package evidence files.';
comment on table lineage_nodes is 'Canonical discovery graph nodes.';
comment on table lineage_edges is 'Canonical discovery graph edges.';
comment on table action_items is 'Execution-ready remediation, governance, modernization, validation, and blocker actions.';
comment on table financial_exposures is 'Directional low/base/high exposure rows requiring finance validation.';
comment on table qa_checks is 'Package QA contract checks and status.';
comment on table dossier_packages is 'Generated ZIP package bytes. For high-volume production, replace bytea with object storage and keep a signed URL/reference here.';
comment on table ai_usage_events is 'OpenAI usage, token, cache, and estimated cost telemetry by dossier run and model.';
