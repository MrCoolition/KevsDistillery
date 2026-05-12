# Data Source Discovery Distillery

Angular 21.2.12 app for generating evidence-backed, graph-backed Data Source Discovery Dossier packages on Vercel with Neon Postgres persistence and OpenAI Responses API synthesis after deterministic extraction.

## Stack

- Angular 21 standalone frontend.
- Vercel static hosting plus Node serverless API in `api/dossiers.ts`.
- Neon Postgres schema in `db/schema.sql`.
- Deterministic extraction libraries before AI: `xlsx`, `jszip`, `mammoth`, `csv-parse`, `node-sql-parser`, `sql-formatter`, `exceljs`, and `pdfkit`.
- OpenAI SDK `responses.create` is optional and only runs when `OPENAI_API_KEY` is configured.

## Local Setup

```bash
npm install
npm start
npm run build
npm run typecheck:api
```

For this Codex desktop environment, I used the bundled Node runtime and a workspace-local npm cache because the system npm wrapper is not available.

`npm start` runs a combined local development setup: Angular serves the UI on `http://127.0.0.1:4200`, `scripts/dev-api.mjs` serves the Vercel-compatible dossier API on `http://127.0.0.1:3001`, and `proxy.conf.json` forwards `/api/*` from Angular to the local API.

## Environment

Copy `.env.example` and configure:

- `DATABASE_URL`: Neon pooled or direct Postgres connection string.
- `OPENAI_API_KEY`: Optional. If absent, packages are generated deterministically without AI narrative synthesis.
- `OPENAI_MODEL`: Defaults to `gpt-5.5`.

## Neon Schema

Apply the named `data_discovery` schema with:

```bash
npm run db:migrate
```

## Vercel

`vercel.json` sets the Angular output directory and gives the dossier API a larger function budget.

```bash
npm run build
```

Then deploy from the project folder with Vercel.

## Dossier Contract

The workflow enforces the required package structure:

- README at the root.
- Folders `01` through `09`.
- Manifest only in `07_Metadata_Manifest`.
- Action backlog in `08_Action_Backlog`.
- Financial model in `09_Financial_Impact_Model`.
- Nine diagram PDFs, each with title, context, legends, criticality, confidence, and evidence notes.
- Query/macro inventories are separate; Access internals are marked blocked unless exported/native metadata is provided.
- No `.dot` files are emitted in the final package.

## Important Limitation

Browser uploads and Vercel serverless cannot safely inspect native Microsoft Access system catalogs through ACE/OLEDB. Access files are registered as databases, but MSysObjects/MSysQueries/macros/VBA/relationships require a trusted desktop collector or exported metadata. The package records that as a lineage blocker and creates P0 action items instead of inventing inventory.
