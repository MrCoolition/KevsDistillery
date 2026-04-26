# Production Setup

Uncle Kev's Distillery is productionized as a Vercel Angular app with Vercel Functions and Neon Postgres persistence.

## Vercel Environment Variables

Set these in the Vercel project settings:

```text
OPENAI_API_KEY=...
OPENAI_MODEL=gpt-5.5
DATABASE_URL=...
```

`DATABASE_URL`, `POSTGRES_URL`, or `NEON_DATABASE_URL` are accepted. The Neon Vercel integration usually provisions one of these automatically.

## Neon Schema

The API owns a dedicated `distillery` schema in Neon. Health checks and synthesis requests auto-create the schema and tables when `DATABASE_URL`, `POSTGRES_URL`, or `NEON_DATABASE_URL` is present. You can also run the migration explicitly:

```powershell
npm run db:migrate
```

Or call the deployed admin endpoint:

```text
POST /api/admin/migrate
```

The canonical tables are `distillery.discovery_runs`, `discovery_sources`, `discovery_items`, `discovery_relationships`, `discovery_artifacts`, `discovery_backlog`, `discovery_evidence_index`, `discovery_lineage_nodes`, `discovery_lineage_edges`, `discovery_package_manifest`, `discovery_people_roles`, `discovery_process_steps`, `discovery_access_objects`, `discovery_excel_objects`, `discovery_word_extracts`, `discovery_data_elements`, `discovery_transform_rules`, `discovery_controls_exceptions`, `discovery_data_quality`, `discovery_security_access`, `discovery_schedule_sla`, `discovery_failure_modes`, `discovery_financial_model`, and `discovery_open_questions`.

## Deploy

Vercel build settings are encoded in `vercel.json`:

- Framework: Angular
- Build command: `npm run build`
- Output directory: `dist/the-distillery/browser`
- API functions: `api/**/*.js`
- Region: `iad1`

After linking the local repo to the Vercel project, deploy production with:

```powershell
vercel --prod
```
