# Migrations

`../schema.sql` is the baseline schema (treat it as migration `001`, applied when a fresh database
is created). Everything in this directory is an incremental, additive change applied on top of it,
in filename order. There's no migration-runner tool wired up yet (small team, small number of
environments) — apply manually:

```bash
psql -U postgres -h localhost -d m365_cleanup -f src/db/migrations/002_cloud_connections.sql
```

When this grows past a handful of files, switch to a real migration tool (e.g. `node-pg-migrate` or
`Kysely`'s migrator) rather than continuing to track "have I applied this one" by hand.
