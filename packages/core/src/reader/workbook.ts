import { parseRels, parseSharedStrings, parseWorkbook, resolveTarget } from '../ooxml'
import type { Cell, SheetInfo } from '../types'
import { openZip, type ZipArchive } from '../zip'
import { type Row, readRows } from './worksheet'

// The reader's public entry point. `openXlsx` follows the OPC relationship graph — never
// guessed filenames — from the package root to the workbook, then to each worksheet and the
// shared string table, and returns a Workbook of typed cells.
//
// Worksheet XML is decompressed up front (so cell access is synchronous) but only parsed
// into cells on first use; a sheet you never touch costs a decompression, not a parse.

const decoder = new TextDecoder()

// Relationship type URIs end in these segments; matching the suffix avoids hard-coding the
// 2006 namespace and tolerates the strict/transitional variants.
const REL_OFFICE_DOCUMENT = '/officeDocument'
const REL_SHARED_STRINGS = '/sharedStrings'

function directoryOf(path: string): string {
	const slash = path.lastIndexOf('/')
	return slash === -1 ? '' : path.slice(0, slash)
}

// The relationships for a part live in `<dir>/_rels/<file>.rels`.
function relsPathFor(path: string): string {
	const slash = path.lastIndexOf('/')
	const dir = slash === -1 ? '' : path.slice(0, slash)
	const file = slash === -1 ? path : path.slice(slash + 1)
	return dir === '' ? `_rels/${file}.rels` : `${dir}/_rels/${file}.rels`
}

async function readText(zip: ZipArchive, path: string): Promise<string> {
	if (!zip.has(path)) throw new Error(`xlsx is missing a required part: ${path}`)
	return decoder.decode(await zip.read(path))
}

export class Worksheet {
	/** Sheet name as shown on Excel's tab. */
	readonly name: string
	readonly #info: SheetInfo
	readonly #xml: string
	readonly #sharedStrings: string[]
	#cells: Map<string, Cell> | undefined

	constructor(info: SheetInfo, xml: string, sharedStrings: string[]) {
		this.name = info.name
		this.#info = info
		this.#xml = xml
		this.#sharedStrings = sharedStrings
	}

	/** Workbook-relative part path, e.g. `xl/worksheets/sheet1.xml`. */
	get path(): string {
		return this.#info.path
	}

	/** false for hidden or very-hidden sheets. */
	get visible(): boolean {
		return this.#info.visible
	}

	/** The cell at an A1 reference. Absent cells read as `empty` (Excel treats them blank). */
	cell(ref: string): Cell {
		return this.#index().get(ref) ?? { ref, type: 'empty', value: null }
	}

	/** Stream the populated rows in document order. Sparse: empty rows/cells are absent. */
	async *rows(): AsyncGenerator<Row> {
		for (const row of readRows(this.#xml, { sharedStrings: this.#sharedStrings })) {
			yield row
		}
	}

	#index(): Map<string, Cell> {
		if (this.#cells === undefined) {
			const cells = new Map<string, Cell>()
			for (const row of readRows(this.#xml, { sharedStrings: this.#sharedStrings })) {
				for (const cell of row.cells) cells.set(cell.ref, cell)
			}
			this.#cells = cells
		}
		return this.#cells
	}
}

export class Workbook {
	/** Sheets in tab order. */
	readonly sheets: readonly SheetInfo[]
	readonly #byName: Map<string, Worksheet>

	constructor(sheets: SheetInfo[], byName: Map<string, Worksheet>) {
		this.sheets = sheets
		this.#byName = byName
	}

	/** The worksheet with this tab name. Throws if there is none. */
	sheet(name: string): Worksheet {
		const worksheet = this.#byName.get(name)
		if (worksheet === undefined) {
			const available = this.sheets.map((s) => s.name).join(', ')
			throw new Error(`no sheet named ${JSON.stringify(name)}; available: ${available}`)
		}
		return worksheet
	}
}

export async function openXlsx(source: Uint8Array | ArrayBuffer): Promise<Workbook> {
	const bytes = source instanceof Uint8Array ? source : new Uint8Array(source)
	const zip = openZip(bytes)

	// Package relationships → the workbook part.
	const packageRels = parseRels(await readText(zip, '_rels/.rels'))
	const office = [...packageRels.values()].find((r) => r.type.endsWith(REL_OFFICE_DOCUMENT))
	if (office === undefined) throw new Error('not an xlsx: no officeDocument relationship')
	const workbookPath = resolveTarget('', office.target)
	const workbookDir = directoryOf(workbookPath)

	// Workbook sheet list + the workbook's own relationships.
	const workbookSheets = parseWorkbook(await readText(zip, workbookPath))
	const workbookRels = parseRels(await readText(zip, relsPathFor(workbookPath)))

	// Shared string table (optional — a workbook may use only inline strings).
	let sharedStrings: string[] = []
	const sst = [...workbookRels.values()].find((r) => r.type.endsWith(REL_SHARED_STRINGS))
	if (sst !== undefined && sst.targetMode !== 'External') {
		const sstPath = resolveTarget(workbookDir, sst.target)
		if (zip.has(sstPath)) {
			sharedStrings = parseSharedStrings(decoder.decode(await zip.read(sstPath)))
		}
	}

	// Resolve each sheet's r:id to a part, decompress it, and build the Worksheet.
	const infos: SheetInfo[] = []
	const byName = new Map<string, Worksheet>()
	for (const entry of workbookSheets) {
		const rel = workbookRels.get(entry.rid)
		if (rel === undefined || rel.targetMode === 'External') continue
		const path = resolveTarget(workbookDir, rel.target)
		if (!zip.has(path)) continue
		const xml = decoder.decode(await zip.read(path))
		const info: SheetInfo = { name: entry.name, path, visible: entry.visible }
		infos.push(info)
		// First definition wins if two sheets somehow share a name.
		if (!byName.has(entry.name)) byName.set(entry.name, new Worksheet(info, xml, sharedStrings))
	}

	return new Workbook(infos, byName)
}
