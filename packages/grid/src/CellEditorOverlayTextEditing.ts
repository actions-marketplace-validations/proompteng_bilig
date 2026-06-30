export function normalizeNumpadKey(key: string, code: string): string | null {
  if (!code.startsWith('Numpad')) {
    return null
  }
  const suffix = code.slice('Numpad'.length)
  if (/^\d$/.test(suffix)) {
    return suffix
  }
  if (suffix === 'Decimal') {
    return '.'
  }
  if (suffix === 'Add') {
    return '+'
  }
  if (suffix === 'Subtract') {
    return '-'
  }
  if (suffix === 'Multiply') {
    return '*'
  }
  if (suffix === 'Divide') {
    return '/'
  }
  return key.length === 1 ? key : null
}

export interface EditorTextSelection {
  readonly direction: 'backward' | 'forward' | 'none'
  readonly end: number
  readonly start: number
}

export interface EditorHistoryEntry {
  readonly selection: EditorTextSelection
  readonly value: string
}

export interface EditorHistoryState {
  readonly current: EditorHistoryEntry
  readonly redoStack: readonly EditorHistoryEntry[]
  readonly undoStack: readonly EditorHistoryEntry[]
}

const MAX_EDITOR_HISTORY_ENTRIES = 100

export function createCaretEndSelection(value: string): EditorTextSelection {
  return {
    direction: 'none',
    end: value.length,
    start: value.length,
  }
}

function sameTextSelection(left: EditorTextSelection, right: EditorTextSelection): boolean {
  return left.start === right.start && left.end === right.end && left.direction === right.direction
}

export function sameEditorHistoryEntry(left: EditorHistoryEntry, right: EditorHistoryEntry): boolean {
  return left.value === right.value && sameTextSelection(left.selection, right.selection)
}

export function trimEditorHistoryStack(entries: readonly EditorHistoryEntry[]): readonly EditorHistoryEntry[] {
  return entries.length <= MAX_EDITOR_HISTORY_ENTRIES ? entries : entries.slice(entries.length - MAX_EDITOR_HISTORY_ENTRIES)
}

export function resolveEditorHistoryShortcut(
  event: Pick<KeyboardEvent, 'altKey' | 'ctrlKey' | 'key' | 'metaKey' | 'shiftKey'>,
): 'redo' | 'undo' | null {
  const isPrimaryModified = (event.ctrlKey || event.metaKey) && !event.altKey
  if (!isPrimaryModified) {
    return null
  }
  const lowerKey = event.key.toLowerCase()
  if (lowerKey === 'z') {
    return event.shiftKey ? 'redo' : 'undo'
  }
  if (!event.shiftKey && lowerKey === 'y') {
    return 'redo'
  }
  return null
}

export function isComposingEditorKey(event: Pick<KeyboardEvent, 'isComposing' | 'key' | 'keyCode'>): boolean {
  return event.isComposing || event.key === 'Process' || event.keyCode === 229
}

export function clampTextOffset(offset: number, textLength: number): number {
  return Math.min(Math.max(0, offset), textLength)
}

export function previousTextBoundary(value: string, offset: number): number {
  if (offset <= 0) {
    return 0
  }
  let cursor = 0
  for (const segment of value) {
    const next = cursor + segment.length
    if (next >= offset) {
      return cursor
    }
    cursor = next
  }
  return cursor
}

export function nextTextBoundary(value: string, offset: number): number {
  if (offset >= value.length) {
    return value.length
  }
  let cursor = 0
  for (const segment of value) {
    const next = cursor + segment.length
    if (cursor >= offset || next > offset) {
      return next
    }
    cursor = next
  }
  return value.length
}
