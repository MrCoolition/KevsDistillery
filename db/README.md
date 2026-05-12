# Neon Postgres Schema

This schema supports the Data Source Discovery Dossier workflow on Neon Postgres. Production objects live under the named Postgres schema `data_discovery`.

Apply it with:

```bash
npm run db:migrate
```

The API persists completed runs when `DATABASE_URL` is configured. It stores run metadata, canonical graph nodes/edges, evidence index rows, action backlog rows, financial exposure rows, QA checks, AI token/cost usage events, and package ZIP bytes.

Package ZIP bytes are stored in `data_discovery.dossier_packages` for a clean first deployment; for large production runs, move ZIP storage to Vercel Blob or another object store and keep only the object reference in Neon.
