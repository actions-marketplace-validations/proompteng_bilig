# Durable Multiplayer Replication RFC

## Status

Archived historical RFC. The active production design has already moved to the monolith + Zero architecture.

## Current source of truth

- [architecture.md](/Users/gregkonush/github.com/bilig/docs/architecture.md)
- [browser-runtime.md](/Users/gregkonush/github.com/bilig/docs/browser-runtime.md)
- [backend-sync-service.md](/Users/gregkonush/github.com/bilig/docs/backend-sync-service.md)

## Current production shape

- `apps/bilig` is the only backend runtime.
- `apps/web` is the only browser shell.
- Zero is the read-sync and mutation ingress plane.
- Postgres is the durable source of truth.
- Durable recovery uses workbook checkpoints plus ordered event replay.
- There is no standalone `apps/local-server` or `apps/sync-server` product topology.

## Why this file remains

This file is kept only as historical context for the earlier replication design discussion. It is not an implementation checklist and it must not be treated as the current runtime contract.
