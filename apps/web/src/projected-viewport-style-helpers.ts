import {
  MAX_COLS,
  MAX_ROWS,
  type CellStyleField,
  type CellStylePatch,
  type CellStyleRecord,
  type WorkbookAxisEntrySnapshot,
} from '@bilig/protocol'

export const DEFAULT_STYLE_ID = 'style-0'

export function normalizeProjectedCellStyle(style: Omit<CellStyleRecord, 'id'>): Omit<CellStyleRecord, 'id'> {
  return {
    ...(style.fill?.backgroundColor ? { fill: { backgroundColor: style.fill.backgroundColor } } : {}),
    ...(style.font && Object.keys(style.font).length > 0 ? { font: { ...style.font } } : {}),
    ...(style.alignment && Object.keys(style.alignment).length > 0 ? { alignment: { ...style.alignment } } : {}),
    ...(style.borders && Object.keys(style.borders).length > 0
      ? {
          borders: {
            ...(style.borders.top ? { top: { ...style.borders.top } } : {}),
            ...(style.borders.right ? { right: { ...style.borders.right } } : {}),
            ...(style.borders.bottom ? { bottom: { ...style.borders.bottom } } : {}),
            ...(style.borders.left ? { left: { ...style.borders.left } } : {}),
          },
        }
      : {}),
    ...(style.protection ? { protection: { ...style.protection } } : {}),
  }
}

export function projectedCellStyleKey(style: Omit<CellStyleRecord, 'id'>): string {
  return JSON.stringify({
    alignment: style.alignment ?? null,
    borders: style.borders ?? null,
    fill: style.fill?.backgroundColor ?? null,
    font: style.font ?? null,
    protection: style.protection ?? null,
  })
}

export function projectedCellStyleIdForKey(key: string): string {
  let hash = 2166136261
  for (let index = 0; index < key.length; index += 1) {
    hash ^= key.charCodeAt(index)
    hash = Math.imul(hash, 16777619)
  }
  return `style-local-${(hash >>> 0).toString(16)}`
}

function cloneProjectedStyleWithoutId(style: CellStyleRecord): Omit<CellStyleRecord, 'id'> {
  return normalizeProjectedCellStyle(style)
}

export function applyProjectedStylePatch(baseStyle: CellStyleRecord, patch: CellStylePatch): Omit<CellStyleRecord, 'id'> {
  const next = cloneProjectedStyleWithoutId(baseStyle)
  if (patch.fill === null) {
    delete next.fill
  } else if (patch.fill !== undefined) {
    const backgroundColor = patch.fill.backgroundColor
    if (backgroundColor === null) {
      delete next.fill
    } else if (backgroundColor !== undefined) {
      next.fill = { backgroundColor }
    }
  }
  if (patch.font === null) {
    delete next.font
  } else if (patch.font) {
    const font = { ...next.font }
    applyOptionalProjectedField(font, 'family', patch.font.family)
    applyOptionalProjectedField(font, 'size', patch.font.size)
    applyOptionalProjectedField(font, 'bold', patch.font.bold)
    applyOptionalProjectedField(font, 'italic', patch.font.italic)
    applyOptionalProjectedField(font, 'underline', patch.font.underline)
    applyOptionalProjectedField(font, 'color', patch.font.color)
    if (Object.keys(font).length > 0) {
      next.font = font
    } else {
      delete next.font
    }
  }
  if (patch.alignment === null) {
    delete next.alignment
  } else if (patch.alignment) {
    const alignment = { ...next.alignment }
    applyOptionalProjectedField(alignment, 'horizontal', patch.alignment.horizontal)
    applyOptionalProjectedField(alignment, 'vertical', patch.alignment.vertical)
    applyOptionalProjectedField(alignment, 'wrap', patch.alignment.wrap)
    applyOptionalProjectedField(alignment, 'indent', patch.alignment.indent)
    applyOptionalProjectedField(alignment, 'shrinkToFit', patch.alignment.shrinkToFit)
    applyOptionalProjectedField(alignment, 'readingOrder', patch.alignment.readingOrder)
    applyOptionalProjectedField(alignment, 'textRotation', patch.alignment.textRotation)
    applyOptionalProjectedField(alignment, 'justifyLastLine', patch.alignment.justifyLastLine)
    if (Object.keys(alignment).length > 0) {
      next.alignment = alignment
    } else {
      delete next.alignment
    }
  }
  if (patch.borders === null) {
    delete next.borders
  } else if (patch.borders) {
    const borders = { ...next.borders }
    applyProjectedBorderSidePatch(borders, 'top', patch.borders.top)
    applyProjectedBorderSidePatch(borders, 'right', patch.borders.right)
    applyProjectedBorderSidePatch(borders, 'bottom', patch.borders.bottom)
    applyProjectedBorderSidePatch(borders, 'left', patch.borders.left)
    if (Object.keys(borders).length > 0) {
      next.borders = borders
    } else {
      delete next.borders
    }
  }
  return normalizeProjectedCellStyle(next)
}

function applyProjectedBorderSidePatch(
  borders: NonNullable<CellStyleRecord['borders']>,
  side: keyof NonNullable<CellStyleRecord['borders']>,
  patch: NonNullable<CellStylePatch['borders']>['top'] | null | undefined,
): void {
  if (patch === undefined) {
    return
  }
  if (patch === null) {
    delete borders[side]
    return
  }
  const nextSide: Partial<NonNullable<CellStyleRecord['borders']>['top']> = { ...borders[side] }
  applyOptionalProjectedField(nextSide, 'style', patch.style)
  applyOptionalProjectedField(nextSide, 'weight', patch.weight)
  applyOptionalProjectedField(nextSide, 'color', patch.color)
  if (nextSide.style && nextSide.weight && nextSide.color) {
    borders[side] = {
      color: nextSide.color,
      style: nextSide.style,
      weight: nextSide.weight,
    }
  } else {
    delete borders[side]
  }
}

export function clearProjectedStyleFields(
  baseStyle: CellStyleRecord,
  fields: readonly CellStyleField[] | undefined,
): Omit<CellStyleRecord, 'id'> {
  if (fields === undefined || fields.length === 0) {
    return {}
  }
  const next = cloneProjectedStyleWithoutId(baseStyle)
  const cleared = new Set(fields)
  if (cleared.has('backgroundColor')) {
    delete next.fill
  }
  const font = filterProjectedStyleSection(
    next.font,
    [
      ['fontFamily', 'family'],
      ['fontSize', 'size'],
      ['fontBold', 'bold'],
      ['fontItalic', 'italic'],
      ['fontUnderline', 'underline'],
      ['fontColor', 'color'],
    ],
    cleared,
  )
  if (font) {
    next.font = font
  } else {
    delete next.font
  }
  const alignment = filterProjectedStyleSection(
    next.alignment,
    [
      ['alignmentHorizontal', 'horizontal'],
      ['alignmentVertical', 'vertical'],
      ['alignmentWrap', 'wrap'],
      ['alignmentIndent', 'indent'],
      ['alignmentShrinkToFit', 'shrinkToFit'],
      ['alignmentReadingOrder', 'readingOrder'],
      ['alignmentTextRotation', 'textRotation'],
      ['alignmentJustifyLastLine', 'justifyLastLine'],
    ],
    cleared,
  )
  if (alignment) {
    next.alignment = alignment
  } else {
    delete next.alignment
  }
  const borders = filterProjectedStyleSection(
    next.borders,
    [
      ['borderTop', 'top'],
      ['borderRight', 'right'],
      ['borderBottom', 'bottom'],
      ['borderLeft', 'left'],
    ],
    cleared,
  )
  if (borders) {
    next.borders = borders
  } else {
    delete next.borders
  }
  return normalizeProjectedCellStyle(next)
}

function filterProjectedStyleSection<T extends object>(
  section: T | undefined,
  fields: ReadonlyArray<readonly [CellStyleField, keyof T]>,
  cleared: ReadonlySet<CellStyleField>,
): T | undefined {
  if (!section) {
    return undefined
  }
  const nextSection = { ...section }
  fields.forEach(([field, key]) => {
    if (cleared.has(field)) {
      delete nextSection[key]
    }
  })
  return Object.keys(nextSection).length > 0 ? nextSection : undefined
}

function applyOptionalProjectedField<T extends object, K extends keyof T>(target: T, key: K, value: T[K] | null | undefined): void {
  if (value === undefined) {
    return
  }
  if (value === null) {
    delete target[key]
    return
  }
  target[key] = value
}

export function assertValidProjectedAxisMutation(axis: 'column' | 'row', index: number, size: number | undefined): void {
  const axisLength = axis === 'column' ? MAX_COLS : MAX_ROWS
  if (!Number.isInteger(index) || index < 0 || index >= axisLength) {
    throw new Error(`Invalid projected ${axis} index: ${index}`)
  }
  if (size !== undefined && (!Number.isFinite(size) || size < 0)) {
    throw new Error(`Invalid projected ${axis} size: ${size}`)
  }
}

export function buildAxisEntries(
  sizes: Readonly<Record<number, number>>,
  hiddenAxes: Readonly<Record<number, true>>,
  idPrefix: 'col' | 'row',
): WorkbookAxisEntrySnapshot[] {
  const indexes = new Set<number>()
  for (const key of Object.keys(sizes)) {
    const index = Number(key)
    if (Number.isInteger(index) && index >= 0) {
      indexes.add(index)
    }
  }
  for (const key of Object.keys(hiddenAxes)) {
    const index = Number(key)
    if (Number.isInteger(index) && index >= 0) {
      indexes.add(index)
    }
  }
  return [...indexes]
    .toSorted((left, right) => left - right)
    .map((index) => ({
      id: `${idPrefix}-${index}`,
      index,
      size: sizes[index] ?? null,
      hidden: hiddenAxes[index] === true ? true : null,
    }))
}
