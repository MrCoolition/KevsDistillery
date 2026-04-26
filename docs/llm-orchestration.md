# LLM Orchestration

Uncle Kev's Distillery keeps LLM work behind a server boundary. The Angular client never receives an OpenAI API key. The browser sends extracted source text and metadata to an internal discovery endpoint; the endpoint calls the OpenAI Responses API with `gpt-5.5`; the client receives canonical proof-graph deltas after validation.

## Live Production Configuration

- Model: `gpt-5.5`
- API surface: Responses API
- Reasoning effort: `high` by default, overrideable with `OPENAI_REASONING_EFFORT`
- Verbosity: `high` by default, overrideable with `OPENAI_VERBOSITY`
- Core output budget: `14000` tokens by default
- Specialist output budget: `11000` tokens per specialist pass by default
- Transport: stored background Responses runs, polled through the app API
- Public label: `The Distillery`
- Output: one merged canonical discovery model delta

## Specialist Passes

Production synthesis runs six background passes and merges them deterministically:

- Canonical Discovery Architect
- Current-State Auto Documentation Lead
- Recursive Lineage Principal
- Logic, Controls, and Failure-Mode Examiner
- Business Exposure and Action Backlog Strategist
- Diagram and Output Package Architect

Each pass returns the same canonical JSON contract. The merge layer deduplicates nodes, edges,
evidence, actions, report sections, lineage, failure modes, technical registers, artifacts, and
open questions into one Discovery Action Pack model.

## Required LLM Contract

Every generated finding must include:

- `id`
- `type`
- `businessPurpose`
- `owner`
- `evidence`
- `confidence`
- `criticality`
- `upstream`
- `downstream`
- `failureImpact`
- `dollarExposure`
- `recommendedAction`

Findings without evidence, confidence, and a next action stay unfinished.

The prompt explicitly requires:

- one canonical node-edge model for every output
- the 12 standard report sections
- all 9 Discovery_Action_Pack artifacts
- Access, Excel, Word, SQL, VBA, Power Query, manual-rule, and blocker extraction
- recursive lineage to terminal nodes or documented blockers
- low/base/high business exposure with assumptions
- owner, priority, dependency, and acceptance criteria for backlog actions
- no arbitrary low item caps

## Example

```powershell
$env:OPENAI_API_KEY="..."
node tools/openai-discovery-agent.example.mjs extracted-access-query.txt
```

The example is intentionally dependency-free and uses Node's native `fetch`. It also loads `.env` automatically when present.

## Local API Boundary

For app integration, start the local server-side proxy:

```powershell
node tools/discovery-api.example.mjs
```

Then post extracted evidence to:

```text
POST http://127.0.0.1:8787/api/discovery/synthesize
```

Request body:

```json
{
  "sourceKind": "access",
  "sourceName": "Operations_Source.accdb",
  "knownArtifacts": ["qry_CurrentState_Output", "Refresh_All"],
  "extractedText": "..."
}
```

The Angular app should call this internal endpoint, never OpenAI directly.
