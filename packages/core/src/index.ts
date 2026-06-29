// Public API surface of @openjsxl/core. Re-exports grow as features land
// (see IMPLEMENTATION.md). For now this is the cell/sheet data model.

// Addressing + date helpers (F0.3 / F0.4).
export type { CellRef } from './ooxml/a1'
export { columnToIndex, formatRef, indexToColumn, parseRef } from './ooxml/a1'
export { serialToDate } from './ooxml/dates'
export type { Cell, CellType, SheetInfo } from './types'
