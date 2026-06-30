import type {
  LiteralInput,
  WorkbookCalculationSettingsSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookDrawingArtifactsSnapshot,
  WorkbookExternalLinkArtifactsSnapshot,
  WorkbookMacroPayloadSnapshot,
  WorkbookProtectionSnapshot,
  WorkbookTableSnapshot,
  WorkbookVolatileContextSnapshot,
} from '@bilig/protocol'
import { runWorkbookMetadataEffect } from './workbook-metadata-service.js'
import type { WorkbookMetadataService } from './workbook-metadata-service-contract.js'
import type {
  WorkbookCalculationSettingsRecord,
  WorkbookDefinedNameRecord,
  WorkbookDrawingArtifactsRecord,
  WorkbookExternalLinkArtifactsRecord,
  WorkbookMacroPayloadRecord,
  WorkbookMetadataRecord,
  WorkbookPropertyRecord,
  WorkbookProtectionRecord,
  WorkbookTableRecord,
  WorkbookVolatileContextRecord,
} from './workbook-metadata-types.js'
import { WorkbookStoreCommentAccessors } from './workbook-store-comment-accessors.js'

export abstract class WorkbookStoreMetadataAccessors extends WorkbookStoreCommentAccessors {
  protected abstract readonly metadata: WorkbookMetadataRecord
  protected abstract override get metadataService(): WorkbookMetadataService

  setWorkbookProperty(key: string, value: LiteralInput): WorkbookPropertyRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.setWorkbookProperty(key, value))
  }

  getWorkbookProperty(key: string): WorkbookPropertyRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getWorkbookProperty(key))
  }

  listWorkbookProperties(): WorkbookPropertyRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listWorkbookProperties())
  }

  setWorkbookProtection(record: WorkbookProtectionSnapshot): WorkbookProtectionRecord {
    return runWorkbookMetadataEffect(this.metadataService.setWorkbookProtection(record))
  }

  getWorkbookProtection(): WorkbookProtectionRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getWorkbookProtection())
  }

  setMacroPayload(record: WorkbookMacroPayloadSnapshot): WorkbookMacroPayloadRecord {
    return runWorkbookMetadataEffect(this.metadataService.setMacroPayload(record))
  }

  listMacroPayloads(): WorkbookMacroPayloadRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listMacroPayloads())
  }

  setCalculationSettings(settings: WorkbookCalculationSettingsSnapshot): WorkbookCalculationSettingsRecord {
    return runWorkbookMetadataEffect(this.metadataService.setCalculationSettings(settings))
  }

  getCalculationSettings(): WorkbookCalculationSettingsRecord {
    return runWorkbookMetadataEffect(this.metadataService.getCalculationSettings())
  }

  setVolatileContext(context: WorkbookVolatileContextSnapshot): WorkbookVolatileContextRecord {
    return runWorkbookMetadataEffect(this.metadataService.setVolatileContext(context))
  }

  getVolatileContext(): WorkbookVolatileContextRecord {
    return runWorkbookMetadataEffect(this.metadataService.getVolatileContext())
  }

  setDrawingArtifacts(artifacts: WorkbookDrawingArtifactsSnapshot): WorkbookDrawingArtifactsRecord {
    return runWorkbookMetadataEffect(this.metadataService.setDrawingArtifacts(artifacts))
  }

  getDrawingArtifacts(): WorkbookDrawingArtifactsRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getDrawingArtifacts())
  }

  clearDrawingArtifacts(): boolean {
    return runWorkbookMetadataEffect(this.metadataService.clearDrawingArtifacts())
  }

  setExternalLinkArtifacts(artifacts: WorkbookExternalLinkArtifactsSnapshot): WorkbookExternalLinkArtifactsRecord {
    return runWorkbookMetadataEffect(this.metadataService.setExternalLinkArtifacts(artifacts))
  }

  getExternalLinkArtifacts(): WorkbookExternalLinkArtifactsRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getExternalLinkArtifacts())
  }

  clearExternalLinkArtifacts(): boolean {
    return runWorkbookMetadataEffect(this.metadataService.clearExternalLinkArtifacts())
  }

  setDefinedName(name: string, value: WorkbookDefinedNameValueSnapshot, scopeSheetName?: string): WorkbookDefinedNameRecord {
    return runWorkbookMetadataEffect(this.metadataService.setDefinedName(name, value, scopeSheetName))
  }

  getDefinedName(name: string, scopeSheetName?: string): WorkbookDefinedNameRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getDefinedName(name, scopeSheetName))
  }

  deleteDefinedName(name: string, scopeSheetName?: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteDefinedName(name, scopeSheetName))
  }

  listDefinedNames(): WorkbookDefinedNameRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listDefinedNames())
  }

  setTable(record: WorkbookTableSnapshot): WorkbookTableRecord {
    return runWorkbookMetadataEffect(this.metadataService.setTable(record))
  }

  getTable(name: string): WorkbookTableRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getTable(name))
  }

  deleteTable(name: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteTable(name))
  }

  hasTables(): boolean {
    return this.metadata.tables.size > 0
  }

  listTables(): WorkbookTableRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listTables())
  }
}
