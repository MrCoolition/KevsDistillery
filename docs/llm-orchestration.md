# LLM Orchestration

THE DISTILLERY keeps LLM work behind a server boundary. The Angular client never receives an OpenAI API key. The browser sends extracted source text and metadata to an internal discovery endpoint; the endpoint calls the OpenAI Responses API with `gpt-5.5`; the client receives canonical graph deltas after validation.

## Default Model

- Model: `gpt-5.5`
- API surface: Responses API
- Reasoning effort: `high` for normal discovery synthesis, `xhigh` for difficult lineage and control reconstruction
- Output: canonical discovery model delta

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
  "sourceName": "CloseMart.accdb",
  "knownArtifacts": ["qry_RevenueFlash_Final", "Refresh_All"],
  "extractedText": "..."
}
```

The Angular app should call this internal endpoint, never OpenAI directly.
