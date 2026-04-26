-- Uncle Kev's Distillery Neon schema
-- Runtime migration source of truth: api/_lib/db.js
-- Apply with: npm run db:migrate
-- API health and synthesis also auto-apply this schema when DATABASE_URL is configured.

create schema if not exists distillery;

-- Required distillery tables:
-- distillery.discovery_runs
-- distillery.discovery_sources
-- distillery.discovery_items
-- distillery.discovery_relationships
-- distillery.discovery_artifacts
-- distillery.discovery_backlog
-- distillery.discovery_evidence_index
-- distillery.discovery_lineage_nodes
-- distillery.discovery_lineage_edges
-- distillery.discovery_package_manifest
-- distillery.discovery_people_roles
-- distillery.discovery_process_steps
-- distillery.discovery_access_objects
-- distillery.discovery_excel_objects
-- distillery.discovery_word_extracts
-- distillery.discovery_data_elements
-- distillery.discovery_transform_rules
-- distillery.discovery_controls_exceptions
-- distillery.discovery_data_quality
-- distillery.discovery_security_access
-- distillery.discovery_schedule_sla
-- distillery.discovery_failure_modes
-- distillery.discovery_financial_model
-- distillery.discovery_open_questions

-- The executable DDL is embedded in api/_lib/db.js so Vercel serverless
-- functions can create and repair the dedicated app schema without relying
-- on filesystem bundling of this reference file.
