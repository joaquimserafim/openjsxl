import { describe, expect, it } from 'vitest'
import { serialToDate } from '../dates'

describe('serialToDate', () => {
	it('maps the 1900-system serial for the Unix epoch', () => {
		expect(serialToDate(25569).getTime()).toBe(Date.UTC(1970, 0, 1))
	})

	it('maps a modern 1900-system date', () => {
		expect(serialToDate(43831).getTime()).toBe(Date.UTC(2020, 0, 1))
	})

	it('honours the 1904 date system', () => {
		expect(serialToDate(42369, true).getTime()).toBe(Date.UTC(2020, 0, 1))
	})

	it('decodes fractional serials as times', () => {
		expect(serialToDate(43831.5).getTime()).toBe(Date.UTC(2020, 0, 1, 12))
	})
})
