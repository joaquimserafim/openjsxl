import { describe, expect, it } from 'vitest'
import { columnToIndex, formatRef, indexToColumn, parseRef } from '../a1'

describe('columnToIndex', () => {
	it('maps the bijective base-26 boundaries', () => {
		expect(columnToIndex('A')).toBe(1)
		expect(columnToIndex('Z')).toBe(26)
		expect(columnToIndex('AA')).toBe(27)
		expect(columnToIndex('AZ')).toBe(52)
		expect(columnToIndex('BA')).toBe(53)
		expect(columnToIndex('XFD')).toBe(16384) // Excel's last column
	})

	it('is case-insensitive', () => {
		expect(columnToIndex('aa')).toBe(27)
	})

	it('rejects invalid input', () => {
		expect(() => columnToIndex('')).toThrow()
		expect(() => columnToIndex('A1')).toThrow()
		expect(() => columnToIndex('-')).toThrow()
	})
})

describe('indexToColumn', () => {
	it('round-trips every index from 1 to 20000', () => {
		for (let i = 1; i <= 20000; i++) {
			expect(columnToIndex(indexToColumn(i))).toBe(i)
		}
	})

	it('rejects non-positive indices', () => {
		expect(() => indexToColumn(0)).toThrow()
	})
})

describe('parseRef / formatRef', () => {
	it('parses and formats A1 references', () => {
		expect(parseRef('A1')).toEqual({ col: 1, row: 1 })
		expect(parseRef('XFD1048576')).toEqual({ col: 16384, row: 1048576 })
		expect(formatRef({ col: 1, row: 1 })).toBe('A1')
	})

	it('round-trips references', () => {
		for (const ref of ['A1', 'B2', 'AA10', 'XFD1048576']) {
			expect(formatRef(parseRef(ref))).toBe(ref)
		}
	})
})
