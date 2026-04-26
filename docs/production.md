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

The API auto-creates tables during the first synthesis request. You can also run the migration explicitly:

```powershell
npm run db:migrate
```

Or call the deployed admin endpoint:

```text
POST /api/admin/migrate
```

The canonical tables are:

- `discovery_runs`
- `discovery_items`
- `discovery_relationships`
- `discovery_artifacts`
- `discovery_backlog`

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
