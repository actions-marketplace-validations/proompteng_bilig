import type { FastifyReply, FastifyRequest } from 'fastify'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import type { ZeroSyncService } from '../zero/service.js'
import { resolveSessionIdentity, type RequestSessionResolver, type SessionIdentity } from './session.js'

interface WorkbookSessionAuthorityInput {
  readonly request: FastifyRequest
  readonly reply: FastifyReply
  readonly sessionResolver: RequestSessionResolver
  readonly zeroSyncService?: ZeroSyncService
}

export function resolveWorkbookSessionWithAuthority(input: WorkbookSessionAuthorityInput): SessionIdentity {
  const session = resolveSessionIdentity(input.request, input.reply, input.sessionResolver)
  if (input.sessionResolver.mode === 'signed-proxy' && (!input.zeroSyncService?.enabled || !input.zeroSyncService.assertWorkbookAccess)) {
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_AUTHORIZATION_UNAVAILABLE',
      message: 'Private workbook access requires the authoritative Zero service',
      statusCode: 503,
      retryable: true,
    })
  }
  return session
}

export async function resolveAuthorizedWorkbookSession(input: {
  readonly request: FastifyRequest
  readonly reply: FastifyReply
  readonly documentId: string
  readonly sessionResolver: RequestSessionResolver
  readonly zeroSyncService?: ZeroSyncService
  readonly createIfMissing?: boolean
}): Promise<SessionIdentity> {
  const session = resolveWorkbookSessionWithAuthority(input)
  if (input.sessionResolver.mode === 'demo') {
    return session
  }
  await input.zeroSyncService?.assertWorkbookAccess?.(input.documentId, session, input.sessionResolver.mode, {
    createIfMissing: input.createIfMissing ?? false,
  })
  return session
}
