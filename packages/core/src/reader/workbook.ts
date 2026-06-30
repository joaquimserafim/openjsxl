import {
	type DecodeContext,
	parseRels,
	parseSharedStrings,
	parseStyles,
	parseWorkbook,
	resolveTarget,
} from '../ooxml'
import type { Cell, SheetInfo } from '../types'
import { openZip, type ZipArchive } from '../zip'
import { parseMergedCells, type Row, readRows, streamRows } from './worksheet'

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
const REL_STYLES = '/styles'

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
	readonly #context: DecodeContext

	#cells: Map<string, Cell> | undefined
	#merged: readonly string[] | undefined

	constructor(info: SheetInfo, xml: string, context: DecodeContext) {
		this.name = info.name
		this.#info = info
		this.#xml = xml
		this.#context = context
	}

	/** Workbook-relative part path, e.g. `xl/worksheets/sheet1.xml`. */
	get path(): string {
		return this.#info.path
	}

	/** false for hidden or very-hidden sheets. */
	get visible(): boolean {
		return this.#info.visible
	}

	/**
	 * Merged-cell ranges in A1 notation (e.g. `['A1:B1', 'A2:A4']`), in document order. Only the
	 * top-left cell of a merge holds a value; the rest read as `empty`. Empty when none.
	 */
	get mergedCells(): readonly string[] {
		if (this.#merged === undefined) this.#merged = parseMergedCells(this.#xml)
		return this.#merged
	}

	/** The cell at an A1 reference. Absent cells read as `empty` (Excel treats them blank). */
	cell(ref: string): Cell {
		return this.#index().get(ref) ?? { ref, type: 'empty', value: null }
	}

	/** Stream the populated rows in document order. Sparse: empty rows/cells are absent. */
	async *rows(): AsyncGenerator<Row> {
		for (const row of readRows(this.#xml, this.#context)) {
			yield row
		}
	}

	#index(): Map<string, Cell> {
		if (this.#cells === undefined) {
			const cells = new Map<string, Cell>()
			for (const row of readRows(this.#xml, this.#context)) {
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

interface LoadedWorkbook {
	zip: ZipArchive
	/** Shared decode context (shared strings, styles, date system) for every sheet. */
	context: DecodeContext
	/** Sheets in tab order, each with its resolved part path. */
	sheets: Array<{ info: SheetInfo; path: string }>
}

// Read the small parts every sheet depends on — relationships, the workbook, shared strings,
// styles — and resolve each sheet's part path through the relationship graph. Worksheets
// themselves are NOT read here, so this stays cheap whether the caller wants random access
// (openXlsx) or a constant-memory stream (streamSheetRows).
async function loadWorkbook(source: Uint8Array | ArrayBuffer): Promise<LoadedWorkbook> {
	const bytes = source instanceof Uint8Array ? source : new Uint8Array(source)
	const zip = openZip(bytes)

	// Package relationships → the workbook part.
	const packageRels = parseRels(await readText(zip, '_rels/.rels'))
	const office = [...packageRels.values()].find((r) => r.type.endsWith(REL_OFFICE_DOCUMENT))
	if (office === undefined) throw new Error('not an xlsx: no officeDocument relationship')
	const workbookPath = resolveTarget('', office.target)
	const workbookDir = directoryOf(workbookPath)

	// Workbook sheet list + date system + the workbook's own relationships.
	const { sheets: workbookSheets, date1904 } = parseWorkbook(await readText(zip, workbookPath))
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

	// Style table (optional) — needed to tell date-styled numbers from plain ones.
	const context: DecodeContext = { sharedStrings, date1904 }
	const stylesRel = [...workbookRels.values()].find((r) => r.type.endsWith(REL_STYLES))
	if (stylesRel !== undefined && stylesRel.targetMode !== 'External') {
		const stylesPath = resolveTarget(workbookDir, stylesRel.target)
		if (zip.has(stylesPath)) {
			context.styles = parseStyles(decoder.decode(await zip.read(stylesPath)))
		}
	}

	// Resolve each sheet's r:id to a part path.
	const sheets: Array<{ info: SheetInfo; path: string }> = []
	for (const entry of workbookSheets) {
		const rel = workbookRels.get(entry.rid)
		if (rel === undefined || rel.targetMode === 'External') continue
		const path = resolveTarget(workbookDir, rel.target)
		if (!zip.has(path)) continue
		sheets.push({ info: { name: entry.name, path, visible: entry.visible }, path })
	}

	return { zip, context, sheets }
}

export async function openXlsx(source: Uint8Array | ArrayBuffer): Promise<Workbook> {
	const { zip, context, sheets } = await loadWorkbook(source)

	// Decompress each worksheet (so cell access is synchronous) and build the Worksheet.
	const infos: SheetInfo[] = []
	const byName = new Map<string, Worksheet>()
	for (const { info, path } of sheets) {
		const xml = decoder.decode(await zip.read(path))
		infos.push(info)
		// First definition wins if two sheets somehow share a name.
		if (!byName.has(info.name)) byName.set(info.name, new Worksheet(info, xml, context))
	}

	return new Workbook(infos, byName)
}

/**
 * Stream the rows of one sheet with roughly constant memory: the worksheet is never
 * materialized as a whole string — it is decompressed and tokenized chunk by chunk, and each
 * row is yielded then discarded (F2.2). Use this for large sheets; use `openXlsx` when you
 * need random `cell()` access. `sheetName` defaults to the first sheet in tab order.
 */
export async function* streamSheetRows(
	source: Uint8Array | ArrayBuffer,
	sheetName?: string,
): AsyncGenerator<Row> {
	const { zip, context, sheets } = await loadWorkbook(source)
	const first = sheets[0]
	if (first === undefined) throw new Error('xlsx has no readable worksheets')

	let path = first.path
	if (sheetName !== undefined) {
		const match = sheets.find((s) => s.info.name === sheetName)
		if (match === undefined) {
			const available = sheets.map((s) => s.info.name).join(', ')
			throw new Error(`no sheet named ${JSON.stringify(sheetName)}; available: ${available}`)
		}
		path = match.path
	}

	yield* streamRows(zip.readStream(path), context)
}
