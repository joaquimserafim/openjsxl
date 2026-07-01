import { loadFixture } from '@openjsxl/fixtures'
import { describe, expect, it } from 'vitest'
import { XlsxError } from '../../errors'
import { openZip } from '../central-directory'

const PARTS = [
	'[Content_Types].xml',
	'_rels/.rels',
	'xl/_rels/workbook.xml.rels',
	'xl/sharedStrings.xml',
	'xl/styles.xml',
	'xl/workbook.xml',
	'xl/worksheets/sheet1.xml',
]

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes)

describe('openZip — real fixture (stored entries)', () => {
	it('lists every part of basic.xlsx', async () => {
		const zip = openZip(await loadFixture('basic.xlsx'))
		expect([...zip.entries.keys()].sort()).toEqual([...PARTS].sort())
	})

	it('reads parts back to their exact bytes', async () => {
		const zip = openZip(await loadFixture('basic.xlsx'))

		const contentTypes = decode(await zip.read('[Content_Types].xml'))
		expect(contentTypes).toContain('<Types')
		expect(contentTypes).toContain('spreadsheetml.sheet.main+xml')

		const sheet = decode(await zip.read('xl/worksheets/sheet1.xml'))
		expect(sheet).toContain('<c r="A1" t="s"><v>0</v></c>')
		expect(sheet).toContain('<f>B1*2</f><v>84</v>')

		const shared = decode(await zip.read('xl/sharedStrings.xml'))
		expect(shared).toContain('<t>hello</t>')
		expect(shared).toContain('<t>world</t>')
	})

	it('reports membership and rejects unknown entries', async () => {
		const zip = openZip(await loadFixture('basic.xlsx'))
		expect(zip.has('xl/workbook.xml')).toBe(true)
		expect(zip.has('nope.xml')).toBe(false)
		await expect(zip.read('nope.xml')).rejects.toThrow()
	})

	it('throws on a buffer that is not a zip', () => {
		expect(() => openZip(new Uint8Array([1, 2, 3, 4]))).toThrow(/not a zip/)
	})
})

// --- a configurable single-entry zip builder for the deflate / edge-case tests ---

const CRC_TABLE = (() => {
	const table = new Uint32Array(256)
	for (let n = 0; n < 256; n++) {
		let c = n
		for (let k = 0; k < 8; k++) c = (c & 1) !== 0 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
		table[n] = c >>> 0
	}
	return table
})()

function crc32(bytes: Uint8Array): number {
	let crc = 0xffffffff
	for (let i = 0; i < bytes.length; i++) {
		crc = ((CRC_TABLE[(crc ^ (bytes[i] as number)) & 0xff] as number) ^ (crc >>> 8)) >>> 0
	}
	return (crc ^ 0xffffffff) >>> 0
}

const u16 = (n: number): Uint8Array => Uint8Array.from([n & 0xff, (n >>> 8) & 0xff])
const u32 = (n: number): Uint8Array =>
	Uint8Array.from([n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff])

function concat(parts: Uint8Array[]): Uint8Array {
	const total = parts.reduce((sum, p) => sum + p.length, 0)
	const out = new Uint8Array(total)
	let offset = 0
	for (const part of parts) {
		out.set(part, offset)
		offset += part.length
	}
	return out
}

async function deflateRaw(data: Uint8Array): Promise<Uint8Array> {
	const stream = new Blob([data as unknown as BlobPart])
		.stream()
		.pipeThrough(new CompressionStream('deflate-raw'))
	return new Uint8Array(await new Response(stream).arrayBuffer())
}

interface ZipOptions {
	name: string
	content: string
	method?: number
	/** general-purpose bit 3: zero the local sizes and append a trailing data descriptor */
	dataDescriptor?: boolean
	descriptorSignature?: boolean
	/** bytes of (zeroed) extra field in the local header */
	localExtra?: number
	/** bytes of (zeroed) extra field in the central header */
	centralExtra?: number
	centralCompressedSize?: number
	centralUncompressedSize?: number
	centralLocalOffset?: number
	/** replace the payload with non-deflate garbage */
	corruptPayload?: boolean
}

async function buildZip(o: ZipOptions): Promise<Uint8Array> {
	const enc = new TextEncoder()
	const raw = enc.encode(o.content)
	const method = o.method ?? 8
	let payload = method === 8 ? await deflateRaw(raw) : raw
	if (o.corruptPayload) payload = Uint8Array.from([0xff, 0xff, 0xff, 0xff])
	const nameBytes = enc.encode(o.name)
	const crc = crc32(raw)
	const localExtra = new Uint8Array(o.localExtra ?? 0)
	const centralExtra = new Uint8Array(o.centralExtra ?? 0)
	const dd = o.dataDescriptor === true
	const flags = dd ? 0x08 : 0

	const local = concat([
		u32(0x04034b50),
		u16(20),
		u16(flags),
		u16(method),
		u16(0),
		u16(0x21),
		u32(dd ? 0 : crc),
		u32(dd ? 0 : payload.length),
		u32(dd ? 0 : raw.length),
		u16(nameBytes.length),
		u16(localExtra.length),
		nameBytes,
		localExtra,
	])
	const blockParts = [local, payload]
	if (dd) {
		blockParts.push(
			o.descriptorSignature
				? concat([u32(0x08074b50), u32(crc), u32(payload.length), u32(raw.length)])
				: concat([u32(crc), u32(payload.length), u32(raw.length)]),
		)
	}
	const localBlock = concat(blockParts)

	const central = concat([
		u32(0x02014b50),
		u16(20),
		u16(20),
		u16(flags),
		u16(method),
		u16(0),
		u16(0x21),
		u32(crc),
		u32(o.centralCompressedSize ?? payload.length),
		u32(o.centralUncompressedSize ?? raw.length),
		u16(nameBytes.length),
		u16(centralExtra.length),
		u16(0),
		u16(0),
		u16(0),
		u32(0),
		u32(o.centralLocalOffset ?? 0),
		nameBytes,
		centralExtra,
	])
	const eocd = concat([
		u32(0x06054b50),
		u16(0),
		u16(0),
		u16(1),
		u16(1),
		u32(central.length),
		u32(localBlock.length),
		u16(0),
	])
	return concat([localBlock, central, eocd])
}

describe('openZip — deflate and real-world layouts', () => {
	it('inflates a plain deflate (method 8) entry', async () => {
		const content = 'deflate me '.repeat(500)
		const zip = openZip(await buildZip({ name: 'big.xml', content }))
		expect(zip.entries.get('big.xml')?.method).toBe(8)
		expect(decode(await zip.read('big.xml'))).toBe(content)
	})

	it('reads a data-descriptor entry (GP bit 3, zeroed local sizes, real central sizes)', async () => {
		const content = 'rows '.repeat(400)
		for (const descriptorSignature of [false, true]) {
			const zip = openZip(
				await buildZip({
					name: 'sheet.xml',
					content,
					dataDescriptor: true,
					descriptorSignature,
				}),
			)
			expect(decode(await zip.read('sheet.xml'))).toBe(content)
		}
	})

	it('uses the local (not central) extra-field length to locate the data', async () => {
		const content = 'extra fields differ '.repeat(50)
		const zip = openZip(
			await buildZip({ name: 'p.xml', content, localExtra: 12, centralExtra: 0 }),
		)
		expect(decode(await zip.read('p.xml'))).toBe(content)
	})

	it('caps inflate at the declared uncompressed size (decompression-bomb guard)', async () => {
		const zip = openZip(
			await buildZip({
				name: 'bomb.xml',
				content: 'x'.repeat(10000),
				centralUncompressedSize: 16,
			}),
		)
		await expect(zip.read('bomb.xml')).rejects.toThrow(/failed to inflate/)
	})

	it('wraps inflate failures with the part name', async () => {
		const zip = openZip(
			await buildZip({ name: 'bad.xml', content: 'whatever', corruptPayload: true }),
		)
		await expect(zip.read('bad.xml')).rejects.toThrow(/corrupt zip: failed to inflate bad\.xml/)
		await expect(zip.read('bad.xml')).rejects.toMatchObject({ code: 'corrupt-zip' })
	})

	it('returns empty bytes for a zero-length deflate entry', async () => {
		const zip = openZip(
			await buildZip({
				name: 'empty.xml',
				content: '',
				centralCompressedSize: 0,
				centralUncompressedSize: 0,
			}),
		)
		expect((await zip.read('empty.xml')).length).toBe(0)
	})

	it('rejects ZIP64 archives with an explicit error', async () => {
		const zip = await buildZip({ name: 'p.xml', content: 'x', centralLocalOffset: 0xffffffff })
		expect(() => openZip(zip)).toThrow(/ZIP64/)
		expect(() => openZip(zip)).toThrow(XlsxError)
		try {
			openZip(zip)
		} catch (e) {
			expect((e as XlsxError).code).toBe('unsupported')
		}
	})

	it('rejects a part whose declared size exceeds maxPartBytes', async () => {
		const bytes = await buildZip({
			name: 'big.xml',
			content: 'x',
			centralUncompressedSize: 10000,
		})
		const zip = openZip(bytes, { maxPartBytes: 100 })
		await expect(zip.read('big.xml')).rejects.toMatchObject({ code: 'part-too-large' })
	})
})

describe('openZip — entry policy', () => {
	it('rejects a zip with duplicate entry names', async () => {
		const bytes = await loadFixture('edge-duplicate-entry.xlsx')
		expect(() => openZip(bytes)).toThrow(/duplicate entry name/)
		try {
			openZip(bytes)
		} catch (e) {
			expect((e as XlsxError).code).toBe('corrupt-zip')
		}
	})

	it('skips directory placeholder entries, keeping real parts', async () => {
		const zip = openZip(await loadFixture('edge-with-directory.xlsx'))
		expect(zip.entries.has('sub/')).toBe(false)
		expect(zip.has('keep.xml')).toBe(true)
	})
})

describe('readStream', () => {
	async function collect(chunks: AsyncIterable<Uint8Array>): Promise<Uint8Array> {
		const parts: Uint8Array[] = []
		let total = 0
		for await (const chunk of chunks) {
			parts.push(chunk)
			total += chunk.byteLength
		}
		const out = new Uint8Array(total)
		let offset = 0
		for (const part of parts) {
			out.set(part, offset)
			offset += part.byteLength
		}
		return out
	}

	it('streams the same bytes as read() for every part', async () => {
		const zip = openZip(await loadFixture('basic.xlsx'))
		for (const part of PARTS) {
			expect(decode(await collect(zip.readStream(part)))).toBe(decode(await zip.read(part)))
		}
	})
})
