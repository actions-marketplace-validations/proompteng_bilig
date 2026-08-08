# Backend Sync Service

## Status

Current summary for the monolith backend.

## Current backend runtime

- The only supported backend runtime is `apps/bilig`.
- It serves the browser shell, session bootstrap, agent ingress, Zero query/mutate endpoints, and the authoritative workbook runtime.
- The product path does not depend on the retired `apps/sync-server` package.

## Current proof points

- [apps/bilig/src/index.ts](../apps/bilig/src/index.ts)
- [apps/bilig/src/http/sync-server.ts](../apps/bilig/src/http/sync-server.ts)
- [apps/bilig/src/zero/service.ts](../apps/bilig/src/zero/service.ts)

## Production boundary configuration

- `BILIG_AUTH_MODE` is required in production. Use `demo` only for public demo
  deployments, or `signed-proxy` behind a trusted identity proxy.
- `BILIG_SESSION_SECRET` must contain at least 32 bytes. `signed-proxy` also
  requires a distinct `BILIG_AUTH_PROXY_SECRET` of at least 32 bytes.
- A signed proxy sends `x-bilig-auth-user`, `x-bilig-auth-roles`,
  `x-bilig-auth-timestamp`, and `x-bilig-auth-signature`. The signature is a
  base64url HMAC-SHA256 over `timestamp`, user ID, and the comma-separated role
  header, joined with newlines. Assertions expire after 60 seconds.
- `BILIG_AGENT_IMPORT_MAX_BYTES` controls decoded workbook upload size and may
  not exceed 64 MiB. CSV cell budgets and XLSX expanded-byte, materialized-cell,
  and formula-cell budgets are enforced after ingress.
- `BILIG_REMOTE_MCP_ALLOWED_ORIGINS`, when set, is the complete comma-separated
  MCP CORS allowlist. Entries must be HTTP(S) origins. Local origins are allowed
  outside production by default and can be controlled explicitly with
  `BILIG_REMOTE_MCP_ALLOW_LOCAL_ORIGINS=true|false`.
