import { describe, expect, it } from 'vitest'

import { rootDir } from '../bilig-dominance-scorecard-input.ts'
import { loadOperatorWorkflowEvidence, operatorWorkflowGaps } from '../bilig-dominance-operator-workflow.ts'

describe('bilig dominance operator workflow evidence', () => {
  it('loads repo wiring for demoted dominance checks and standalone claim policy', () => {
    const evidence = loadOperatorWorkflowEvidence(rootDir)

    expect(operatorWorkflowGaps(evidence)).toEqual([])
    expect(evidence).toMatchObject({
      dominancePackageScriptsAbsent: true,
      googleSheetsTenXClaimGateScriptAbsent: true,
      publicClaimsCheckScriptPresent: true,
      runCiDominanceChecksAbsent: true,
      runCiPublicClaimsCheckPresent: true,
      generatedSourceChecksSerialized: true,
      promptArtifactAuditCoupledToLiveStatus: true,
    })
  })
})
