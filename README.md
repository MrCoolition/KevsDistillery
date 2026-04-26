# Uncle Kev's Distillery

Uncle Kev's Distillery is an Angular 21 standalone app for graph-backed enterprise data discovery. It lets teams browse real source files and folders, extract evidence, build one canonical proof graph, and generate executive reports, auto-documentation, diagrams, recursive lineage, business-impact analysis, and remediation backlog outputs for Snowflake migration teams using Fivetran, dbt, SQL, and Snowpark.

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
- Action Pack generation tracker covering briefs, reports, workbooks, auto-docs, diagrams, impact model, backlog, evidence archive, and metadata manifest.
- Business-impact model for availability, freshness, quality, auditability, recovery effort, and modernization priority.
- Remediation backlog with acceptance criteria.
- Server-side LLM orchestration contract for OpenAI `gpt-5.5` via the Responses API.
