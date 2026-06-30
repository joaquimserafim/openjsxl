import { describe, expect, it } from 'vitest'
import { loadFixture } from '../index'

describe('basic.xlsx fixture', () => {
	it('exists and is a valid ZIP (local file header signature "PK\\x03\\x04")', async () => {
		const bytes = await loadFixture('basic.xlsx')
		expect(bytes.length).toBeGreaterThan(0)
		expect([bytes[0], bytes[1], bytes[2], bytes[3]]).toEqual([0x50, 0x4b, 0x03, 0x04])
	})
})
