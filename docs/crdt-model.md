# CRDT and Local-First Model

## Status

Archived historical note.

The current production architecture is not CRDT-authoritative. The active design is:

- server-authoritative ordering in `apps/bilig`
- Zero as the narrow relational sync plane
- `@bilig/core` as the owner of local replica bookkeeping needed for replay and snapshot restore

## Current source of truth

- [design.md](../docs/design.md)
- [architecture.md](../docs/architecture.md)
- [05-06-next-phase.md](../docs/05-06-next-phase.md)
- [backend-sync-service.md](../docs/backend-sync-service.md)

## Why this file remains

This file is kept only as historical context for earlier local-first replication language. It must not be read as the current runtime contract.
