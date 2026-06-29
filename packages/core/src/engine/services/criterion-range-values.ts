import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { EngineRuntimeColumnStoreService, RuntimeColumnView } from './runtime-column-store-service.js'

export function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

export function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

export function decodeValueTag(rawTag: number | undefined): ValueTag {
  if (rawTag === undefined) {
    return ValueTag.Empty
  }
  switch (rawTag) {
    case 1:
      return ValueTag.Number
    case 2:
      return ValueTag.Boolean
    case 3:
      return ValueTag.String
    case 4:
      return ValueTag.Error
    case 0:
    default:
      return ValueTag.Empty
  }
}

export function normalizeSliceString(runtimeColumnStore: EngineRuntimeColumnStoreService, stringId: number): string {
  return stringId === 0 ? '' : runtimeColumnStore.normalizeStringId(stringId)
}

export function materializeSliceValue(
  view: RuntimeColumnView,
  offset: number,
  runtimeColumnStore: EngineRuntimeColumnStoreService,
): CellValue {
  const tag = decodeValueTag(view.readTagAt(offset))
  switch (tag) {
    case ValueTag.Empty:
      return { tag: ValueTag.Empty }
    case ValueTag.Number:
      return { tag: ValueTag.Number, value: view.readNumberAt(offset) }
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: view.readNumberAt(offset) !== 0 }
    case ValueTag.String: {
      const stringId = view.readStringIdAt(offset)
      return {
        tag: ValueTag.String,
        value: stringId === 0 ? '' : runtimeColumnStore.normalizeStringId(stringId),
        stringId,
      }
    }
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: view.readErrorAt(offset) ?? ErrorCode.None }
  }
}
