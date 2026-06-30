import { describe, expect, it } from 'vitest'
import { type DecodeContext, decodeCell, type RawCell } from '../cell'

const ctx: DecodeContext = { sharedStrings: ['hello', 'world'] }

const raw = (ref: string, type: string | undefined, value: string | undefined): RawCell => ({
	ref,
	type,
	value,
})

describe('decodeCell', () => {
	it('resolves a shared-string index to its text', () => {
		expect(decodeCell(raw('A1', 's', '0'), ctx)).toEqual({
			ref: 'A1',
			type: 'string',
			value: 'hello',
		})
		expect(decodeCell(raw('A2', 's', '1'), ctx)).toEqual({
			ref: 'A2',
			type: 'string',
			value: 'world',
		})
	})

	it('treats an out-of-range or non-integer shared index as empty', () => {
		expect(decodeCell(raw('A1', 's', '9'), ctx).type).toBe('empty')
		expect(decodeCell(raw('A1', 's', '-1'), ctx).type).toBe('empty')
		expect(decodeCell(raw('A1', 's', 'x'), ctx).type).toBe('empty')
		expect(decodeCell(raw('A1', 's', undefined), ctx).type).toBe('empty')
	})

	it('reads inline and cached-formula strings straight from the value', () => {
		expect(decodeCell(raw('A1', 'inlineStr', 'inline'), ctx)).toEqual({
			ref: 'A1',
			type: 'string',
			value: 'inline',
		})
		expect(decodeCell(raw('A1', 'str', 'cached'), ctx)).toEqual({
			ref: 'A1',
			type: 'string',
			value: 'cached',
		})
	})

	it('reads booleans from 1/0, not as numbers', () => {
		expect(decodeCell(raw('A1', 'b', '1'), ctx)).toEqual({
			ref: 'A1',
			type: 'boolean',
			value: true,
		})
		expect(decodeCell(raw('A1', 'b', '0'), ctx)).toEqual({
			ref: 'A1',
			type: 'boolean',
			value: false,
		})
	})

	it('reads error literals', () => {
		expect(decodeCell(raw('A1', 'e', '#DIV/0!'), ctx)).toEqual({
			ref: 'A1',
			type: 'error',
			value: '#DIV/0!',
		})
	})

	it('defaults absent/"n" types to a number', () => {
		expect(decodeCell(raw('A1', undefined, '42'), ctx)).toEqual({
			ref: 'A1',
			type: 'number',
			value: 42,
		})
		expect(decodeCell(raw('A1', 'n', '3.14159'), ctx)).toEqual({
			ref: 'A1',
			type: 'number',
			value: 3.14159,
		})
		expect(decodeCell(raw('A1', undefined, '1e3'), ctx).value).toBe(1000)
	})

	it('treats a missing or non-numeric number value as empty', () => {
		expect(decodeCell(raw('A1', undefined, undefined), ctx)).toEqual({
			ref: 'A1',
			type: 'empty',
			value: null,
		})
		expect(decodeCell(raw('A1', undefined, ''), ctx).type).toBe('empty')
		expect(decodeCell(raw('A1', 'n', 'abc'), ctx).type).toBe('empty')
	})
})
