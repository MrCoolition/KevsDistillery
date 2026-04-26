# THE DISTILLERY

THE DISTILLERY is an Angular 21 standalone app prototype for graph-backed data discovery. It models how a single canonical discovery model can drive executive reports, auto-documentation, diagrams, recursive lineage, financial exposure, and remediation backlog outputs for Snowflake migration teams using Fivetran, dbt, SQL, and Snowpark.

## Run

Install dependencies with a package manager, then start the Angular dev server:

```powershell
npm install
npm start
```

The app expects Angular 21.2.8 packages. In this Codex workspace, Node is installed but `npm` is not available on PATH, so the local smoke test is dependency-free:

```powershell
node tools/smoke-test.mjs
```

## LLM Setup

Store the OpenAI key in `.env`:

```powershell
OPENAI_API_KEY=...
```

The local server-side example loads `.env` automatically:

```powershell
node tools/discovery-api.example.mjs
```

The browser must call this internal API boundary, not OpenAI directly.

For Vercel and Neon production setup, see [docs/production.md](docs/production.md).

## Product Surface

- Canonical discovery model with required evidence, confidence, criticality, impact, exposure, and next action fields.
- Executive command surface for scope, coverage, confidence, risk, and decision readiness.
- Recursive lineage graph with selected-node traceability.
- Action Pack generation tracker covering briefs, reports, workbooks, auto-docs, diagrams, financial model, backlog, evidence archive, and metadata manifest.
- Financial exposure model with low, base, and high estimates.
- Remediation backlog with acceptance criteria.
- Server-side LLM orchestration contract for OpenAI `gpt-5.5` via the Responses API.
