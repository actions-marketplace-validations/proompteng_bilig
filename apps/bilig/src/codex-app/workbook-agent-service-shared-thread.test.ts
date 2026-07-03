import { SpreadsheetEngine } from '@bilig/core'
import { describe, expect, it, vi } from 'vitest'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { WorkbookAgentThreadStateRecord } from '../zero/workbook-chat-thread-store.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import type { CodexAppServerClientOptions, CodexAppServerTransport } from './codex-app-server-client.js'
import { createWorkbookAgentService } from './workbook-agent-service.js'

import {
  FakeCodexTransport,
  createPreviewSummary,
  createReviewQueueItem,
  createZeroSyncStub,
  getPrimaryReviewBundle,
} from './workbook-agent-service.test-helpers.js'
describe('workbook agent service shared threads', () => {
  it('allows collaborators to reuse a shared thread session while persisting the canonical owner row', async () => {
    const fakeCodex = new FakeCodexTransport()
    let durableThreadState: WorkbookAgentThreadStateRecord | null = {
      documentId: 'doc-1',
      threadId: 'thr-shared',
      actorUserId: 'alex@example.com',
      scope: 'shared',
      executionPolicy: 'ownerReview',
      context: {
        selection: {
          sheetName: 'Sheet1',
          address: 'A1',
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      entries: [],
      reviewQueueItems: [],
      updatedAtUnixMs: 100,
    }
    const saveWorkbookAgentThreadState = vi.fn(async (record: WorkbookAgentThreadStateRecord) => {
      durableThreadState = structuredClone(record)
    })
    const zeroSync = createZeroSyncStub({
      async loadWorkbookAgentThreadState() {
        return durableThreadState ? structuredClone(durableThreadState) : null
      },
      saveWorkbookAgentThreadState,
    })
    const service = createWorkbookAgentService(zeroSync, {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => fakeCodex,
    })

    try {
      const alexSnapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-shared',
        },
      })

      const caseySnapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'casey@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-shared',
          context: {
            selection: {
              sheetName: 'Sheet1',
              address: 'B2',
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
          },
        },
      })

      expect(alexSnapshot.scope).toBe('shared')
      expect(caseySnapshot.threadId).toBe('thr-shared')
      expect(caseySnapshot.scope).toBe('shared')
      expect(caseySnapshot.context).toEqual(
        expect.objectContaining({
          selection: expect.objectContaining({
            address: 'B2',
          }),
        }),
      )

      await service.startTurn({
        documentId: 'doc-1',
        threadId: caseySnapshot.threadId,
        session: {
          userID: 'casey@example.com',
          roles: ['editor'],
        },
        body: {
          prompt: 'Review this shared thread',
        },
      })

      expect(saveWorkbookAgentThreadState).toHaveBeenLastCalledWith(
        expect.objectContaining({
          actorUserId: 'alex@example.com',
          scope: 'shared',
        }),
      )
    } finally {
      await service.close()
    }
  })

  it('loads shared execution history when a collaborator resumes a shared thread', async () => {
    const executionRecord = {
      id: 'run-shared-1',
      bundleId: 'bundle-shared-1',
      documentId: 'doc-1',
      threadId: 'thr-shared',
      turnId: 'turn-1',
      actorUserId: 'alex@example.com',
      goalText: 'Normalize imported rows',
      planText: 'Apply the shared cleanup plan',
      summary: 'Write cells in Sheet1!B2',
      scope: 'sheet' as const,
      riskClass: 'medium' as const,
      acceptedScope: 'full' as const,
      appliedBy: 'user' as const,
      baseRevision: 3,
      appliedRevision: 4,
      createdAtUnixMs: 100,
      appliedAtUnixMs: 200,
      context: null,
      commands: [
        {
          kind: 'writeRange' as const,
          sheetName: 'Sheet1',
          startAddress: 'B2',
          values: [[42]],
        },
      ],
      preview: null,
    }
    const listWorkbookAgentThreadRuns = vi.fn(async () => [executionRecord])
    const listWorkbookAgentRuns = vi.fn(async () => [])
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        listWorkbookAgentRuns,
        listWorkbookAgentThreadRuns,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-shared',
            actorUserId: 'alex@example.com',
            scope: 'shared',
            executionPolicy: 'ownerReview',
            context: null,
            entries: [],
            reviewQueueItems: [],
            updatedAtUnixMs: 100,
          }
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => new FakeCodexTransport(),
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'casey@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-shared',
        },
      })

      expect(snapshot.scope).toBe('shared')
      expect(snapshot.executionRecords).toEqual([executionRecord])
      expect(listWorkbookAgentThreadRuns).toHaveBeenCalledWith('doc-1', 'casey@example.com', 'thr-shared')
      expect(listWorkbookAgentRuns).not.toHaveBeenCalled()
    } finally {
      await service.close()
    }
  })

  it('requires the shared thread owner to apply medium/high-risk bundles', async () => {
    const applyAgentCommandBundle = vi.fn(async () => ({
      revision: 5,
      preview: createPreviewSummary(),
    }))
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-shared',
            actorUserId: 'alex@example.com',
            scope: 'shared',
            executionPolicy: 'ownerReview',
            context: null,
            entries: [],
            reviewQueueItems: [
              createReviewQueueItem({
                id: 'bundle-shared-1',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-1',
                goalText: 'Build a workbook-wide summary',
                summary: 'Create summary sheet and rewrite rollups',
                scope: 'workbook',
                riskClass: 'high',
                baseRevision: 4,
                createdAtUnixMs: 100,
                context: null,
                commands: [
                  {
                    kind: 'createSheet',
                    name: 'Summary',
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 0,
                sharedReview: null,
              }),
            ],
            updatedAtUnixMs: 100,
          }
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => new FakeCodexTransport(),
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'casey@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-shared',
        },
      })

      await expect(
        service.applyReviewItem({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          reviewItemId: 'bundle-shared-1',
          session: {
            userID: 'casey@example.com',
            roles: ['editor'],
          },
          appliedBy: 'user',
        }),
      ).rejects.toThrow('Shared medium/high-risk workbook bundles must be applied by the thread owner.')
      expect(applyAgentCommandBundle).not.toHaveBeenCalled()
    } finally {
      await service.close()
    }
  })

  it('still allows collaborators to apply low-risk shared bundles manually', async () => {
    const applyAgentCommandBundle = vi.fn(async () => ({
      revision: 5,
      preview: createPreviewSummary(),
    }))
    const appendWorkbookAgentRun = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-shared',
            actorUserId: 'alex@example.com',
            scope: 'shared',
            executionPolicy: 'ownerReview',
            context: null,
            entries: [],
            reviewQueueItems: [
              createReviewQueueItem({
                id: 'bundle-shared-low',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-1',
                goalText: 'Fix one visible cell',
                summary: 'Write cells in Sheet1!B2',
                scope: 'selection',
                riskClass: 'low',
                baseRevision: 4,
                createdAtUnixMs: 100,
                context: null,
                commands: [
                  {
                    kind: 'writeRange',
                    sheetName: 'Sheet1',
                    startAddress: 'B2',
                    values: [[42]],
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'B2',
                    endAddress: 'B2',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 1,
                sharedReview: null,
              }),
            ],
            updatedAtUnixMs: 100,
          }
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => new FakeCodexTransport(),
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'casey@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-shared',
        },
      })

      const applied = await service.applyReviewItem({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        reviewItemId: 'bundle-shared-low',
        session: {
          userID: 'casey@example.com',
          roles: ['editor'],
        },
        appliedBy: 'user',
      })

      expect(applyAgentCommandBundle).toHaveBeenCalled()
      expect(applied.reviewQueueItems).toEqual([])
    } finally {
      await service.close()
    }
  })

  it('requires owner approval before applying shared medium/high-risk bundles', async () => {
    const applyAgentCommandBundle = vi.fn(async () => ({
      revision: 6,
      preview: createPreviewSummary(),
    }))
    const appendWorkbookAgentRun = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-shared',
            actorUserId: 'alex@example.com',
            scope: 'shared',
            executionPolicy: 'ownerReview',
            context: null,
            entries: [],
            reviewQueueItems: [
              createReviewQueueItem({
                id: 'bundle-shared-review',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-1',
                goalText: 'Normalize the workbook',
                summary: 'Normalize shared workbook structure',
                scope: 'workbook',
                riskClass: 'high',
                baseRevision: 4,
                createdAtUnixMs: 100,
                context: null,
                commands: [
                  {
                    kind: 'createSheet',
                    name: 'Summary',
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 0,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'pending',
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [],
                },
              }),
            ],
            updatedAtUnixMs: 100,
          }
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => new FakeCodexTransport(),
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-shared',
        },
      })

      await expect(
        service.applyReviewItem({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          reviewItemId: 'bundle-shared-review',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          appliedBy: 'user',
        }),
      ).rejects.toThrow('Shared medium/high-risk workbook bundles must be approved by the thread owner before apply.')

      const reviewed = await service.reviewReviewItem({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        reviewItemId: 'bundle-shared-review',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          decision: 'approved',
        },
      })

      expect(getPrimaryReviewBundle(reviewed)).toEqual(
        expect.objectContaining({
          sharedReview: expect.objectContaining({
            status: 'approved',
            decidedByUserId: 'alex@example.com',
            recommendations: [],
          }),
        }),
      )
      expect(service.getObservabilitySnapshot().counters.sharedReviewApprovedCount).toBe(1)

      const applied = await service.applyReviewItem({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        reviewItemId: 'bundle-shared-review',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        appliedBy: 'user',
      })

      expect(applyAgentCommandBundle).toHaveBeenCalled()
      expect(appendWorkbookAgentRun).toHaveBeenCalled()
      expect(applied.reviewQueueItems).toEqual([])
    } finally {
      await service.close()
    }
  })

  it('does not apply a review item dismissed while authoritative preview is building', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'doc-1',
      replicaId: 'server:test',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    let releasePreview!: () => void
    const previewBlocked = new Promise<void>((resolve) => {
      releasePreview = resolve
    })
    let resolvePreviewStarted!: () => void
    const previewStarted = new Promise<void>((resolve) => {
      resolvePreviewStarted = resolve
    })
    const inspectWorkbook = async <T>(_documentId: string, task: (runtime: WorkbookRuntime) => T | Promise<T>): Promise<T> => {
      resolvePreviewStarted()
      await previewBlocked
      const runtime: WorkbookRuntime = {
        documentId: 'doc-1',
        engine,
        projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
          revision: 4,
          calculatedRevision: 4,
          ownerUserId: 'alex@example.com',
          updatedBy: 'alex@example.com',
          updatedAt: '2026-04-10T00:00:00.000Z',
        }),
        headRevision: 4,
        calculatedRevision: 4,
        ownerUserId: 'alex@example.com',
      }
      return await task(runtime)
    }
    const applyAgentCommandBundle = vi.fn(async () => ({
      revision: 6,
      preview: createPreviewSummary(),
    }))
    const appendWorkbookAgentRun = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
        inspectWorkbook,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-shared',
            actorUserId: 'alex@example.com',
            scope: 'shared',
            executionPolicy: 'ownerReview',
            context: null,
            entries: [],
            reviewQueueItems: [
              createReviewQueueItem({
                id: 'bundle-shared-review',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-1',
                goalText: 'Normalize the workbook',
                summary: 'Normalize shared workbook structure',
                scope: 'workbook',
                riskClass: 'high',
                baseRevision: 4,
                createdAtUnixMs: 100,
                context: null,
                commands: [
                  {
                    kind: 'createSheet',
                    name: 'Summary',
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 0,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'approved',
                  decidedByUserId: 'alex@example.com',
                  decidedAtUnixMs: 200,
                  recommendations: [],
                },
              }),
            ],
            updatedAtUnixMs: 100,
          }
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => new FakeCodexTransport(),
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-shared',
        },
      })

      const applyPromise = service.applyReviewItem({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        reviewItemId: 'bundle-shared-review',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        appliedBy: 'user',
      })
      await previewStarted
      const dismissed = await service.dismissReviewItem({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        reviewItemId: 'bundle-shared-review',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
      })
      expect(dismissed.reviewQueueItems).toEqual([])
      releasePreview()

      await expect(applyPromise).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_REVIEW_ITEM_NOT_FOUND',
        statusCode: 404,
        retryable: false,
      })
      expect(applyAgentCommandBundle).not.toHaveBeenCalled()
      expect(appendWorkbookAgentRun).not.toHaveBeenCalled()
      expect(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
        }).reviewQueueItems,
      ).toEqual([])
    } finally {
      releasePreview()
      await service.close()
    }
  })

  it('records collaborator recommendations before owner approval', async () => {
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-shared',
            actorUserId: 'alex@example.com',
            scope: 'shared',
            executionPolicy: 'ownerReview',
            context: null,
            entries: [],
            reviewQueueItems: [
              createReviewQueueItem({
                id: 'bundle-shared-review',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-1',
                goalText: 'Normalize the workbook',
                summary: 'Normalize shared workbook structure',
                scope: 'workbook',
                riskClass: 'high',
                baseRevision: 4,
                createdAtUnixMs: 100,
                context: null,
                commands: [
                  {
                    kind: 'createSheet',
                    name: 'Summary',
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 0,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'pending',
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [],
                },
              }),
            ],
            updatedAtUnixMs: 100,
          }
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => new FakeCodexTransport(),
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'pat@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-shared',
        },
      })

      const reviewed = await service.reviewReviewItem({
        documentId: 'doc-1',
        threadId: snapshot.threadId,
        reviewItemId: 'bundle-shared-review',
        session: {
          userID: 'pat@example.com',
          roles: ['editor'],
        },
        body: {
          decision: 'approved',
        },
      })

      expect(getPrimaryReviewBundle(reviewed)).toEqual(
        expect.objectContaining({
          sharedReview: expect.objectContaining({
            status: 'pending',
            decidedByUserId: null,
            recommendations: [
              expect.objectContaining({
                userId: 'pat@example.com',
                decision: 'approved',
              }),
            ],
          }),
        }),
      )
      expect(service.getObservabilitySnapshot().counters.sharedRecommendationApprovedCount).toBe(1)
      expect(reviewed.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'system',
            text: expect.stringContaining('pat@example.com shared a ready-to-apply review recommendation'),
          }),
        ]),
      )
    } finally {
      await service.close()
    }
  })

  it('prevents review updates after shared owner approval', async () => {
    const saveWorkbookAgentThreadState = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        saveWorkbookAgentThreadState,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-shared',
            actorUserId: 'alex@example.com',
            scope: 'shared',
            executionPolicy: 'ownerReview',
            context: null,
            entries: [],
            reviewQueueItems: [
              createReviewQueueItem({
                id: 'bundle-shared-review',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-1',
                goalText: 'Normalize the workbook',
                summary: 'Normalize shared workbook structure',
                scope: 'workbook',
                riskClass: 'high',
                baseRevision: 4,
                createdAtUnixMs: 100,
                context: null,
                commands: [
                  {
                    kind: 'createSheet',
                    name: 'Summary',
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 0,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'approved',
                  decidedByUserId: 'alex@example.com',
                  decidedAtUnixMs: 200,
                  recommendations: [],
                },
              }),
            ],
            updatedAtUnixMs: 100,
          }
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => new FakeCodexTransport(),
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'pat@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-shared',
        },
      })
      saveWorkbookAgentThreadState.mockClear()

      await expect(
        service.reviewReviewItem({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          reviewItemId: 'bundle-shared-review',
          session: {
            userID: 'pat@example.com',
            roles: ['editor'],
          },
          body: {
            decision: 'rejected',
          },
        }),
      ).rejects.toMatchObject({
        code: 'WORKBOOK_AGENT_SHARED_REVIEW_ALREADY_DECIDED',
        statusCode: 409,
        retryable: false,
      })

      expect(saveWorkbookAgentThreadState).not.toHaveBeenCalled()
      expect(service.getObservabilitySnapshot().counters.sharedRecommendationRejectedCount).toBe(0)
      expect(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'pat@example.com',
            roles: ['editor'],
          },
        }).reviewQueueItems[0],
      ).toEqual(
        expect.objectContaining({
          status: 'approved',
          decidedByUserId: 'alex@example.com',
          recommendations: [],
        }),
      )
    } finally {
      await service.close()
    }
  })

  it('prevents collaborators from dismissing pending shared owner-review items', async () => {
    const saveWorkbookAgentThreadState = vi.fn(async () => undefined)
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        saveWorkbookAgentThreadState,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-shared',
            actorUserId: 'alex@example.com',
            scope: 'shared',
            executionPolicy: 'ownerReview',
            context: null,
            entries: [],
            reviewQueueItems: [
              createReviewQueueItem({
                id: 'bundle-shared-review',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-1',
                goalText: 'Normalize the workbook',
                summary: 'Normalize shared workbook structure',
                scope: 'workbook',
                riskClass: 'high',
                baseRevision: 4,
                createdAtUnixMs: 100,
                context: null,
                commands: [
                  {
                    kind: 'createSheet',
                    name: 'Summary',
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 0,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'pending',
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [],
                },
              }),
            ],
            updatedAtUnixMs: 100,
          }
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => new FakeCodexTransport(),
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'pat@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-shared',
        },
      })
      saveWorkbookAgentThreadState.mockClear()

      await expect(
        service.dismissReviewItem({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          reviewItemId: 'bundle-shared-review',
          session: {
            userID: 'pat@example.com',
            roles: ['editor'],
          },
        }),
      ).rejects.toThrow('Only the shared thread owner can clear a pending owner-review item.')

      expect(saveWorkbookAgentThreadState).not.toHaveBeenCalled()
      expect(
        service.getSnapshot({
          documentId: 'doc-1',
          threadId: snapshot.threadId,
          session: {
            userID: 'pat@example.com',
            roles: ['editor'],
          },
        }).reviewQueueItems,
      ).toHaveLength(1)
    } finally {
      await service.close()
    }
  })

  it('loads only thread-scoped execution history for private threads', async () => {
    const threadExecutionRecord = {
      id: 'run-private-1',
      bundleId: 'bundle-private-1',
      documentId: 'doc-1',
      threadId: 'thr-private',
      turnId: 'turn-1',
      actorUserId: 'alex@example.com',
      goalText: 'Fix the selected range',
      planText: 'Repair formulas in the active thread',
      summary: 'Write formulas in Sheet1!C2:C5',
      scope: 'selection' as const,
      riskClass: 'low' as const,
      acceptedScope: 'full' as const,
      appliedBy: 'auto' as const,
      baseRevision: 7,
      appliedRevision: 8,
      createdAtUnixMs: 100,
      appliedAtUnixMs: 110,
      context: null,
      commands: [
        {
          kind: 'writeRange' as const,
          sheetName: 'Sheet1',
          startAddress: 'C2',
          values: [[42]],
        },
      ],
      preview: null,
    }
    const foreignExecutionRecord = {
      ...threadExecutionRecord,
      id: 'run-foreign-1',
      bundleId: 'bundle-foreign-1',
      threadId: 'thr-other',
      goalText: 'Foreign thread run',
      summary: 'Should never hydrate into thr-private',
    }
    const listWorkbookAgentThreadRuns = vi.fn(async () => [threadExecutionRecord])
    const listWorkbookAgentRuns = vi.fn(async () => [foreignExecutionRecord])
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        listWorkbookAgentRuns,
        listWorkbookAgentThreadRuns,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: 'doc-1',
            threadId: 'thr-private',
            actorUserId: 'alex@example.com',
            scope: 'private',
            executionPolicy: 'autoApplyAll',
            context: null,
            entries: [],
            reviewQueueItems: [],
            updatedAtUnixMs: 100,
          }
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport => new FakeCodexTransport(),
      },
    )

    try {
      const snapshot = await service.createSession({
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        body: {
          threadId: 'thr-private',
        },
      })

      expect(snapshot.scope).toBe('private')
      expect(snapshot.executionRecords).toEqual([threadExecutionRecord])
      expect(listWorkbookAgentThreadRuns).toHaveBeenCalledWith('doc-1', 'alex@example.com', 'thr-private')
      expect(listWorkbookAgentRuns).not.toHaveBeenCalled()
    } finally {
      await service.close()
    }
  })
})
