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
);

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
);

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
);

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
);

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
);

create index if not exists discovery_runs_created_at_idx on discovery_runs (created_at desc);
create index if not exists discovery_items_type_idx on discovery_items (item_type);
create index if not exists discovery_backlog_priority_idx on discovery_backlog (priority);
