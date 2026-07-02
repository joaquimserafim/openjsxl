import { XlsxError } from '../errors'
import {
	type DecodeContext,
	parseRels,
	parseSharedStrings,
	parseStyles,
	parseWorkbook,
	type Relationship,
	resolveTarget,
	type StyleTable,
} from '../ooxml'
import type { Cell, CellStyle, Comment, Hyperlink, SheetInfo } from '../types'
import { openZip, type ZipArchive } from '../zip'
import {
	parseCellStyles,
	parseComments,
	parseDimension,
	parseHyperlinks,
	parseMergedCells,
	type Row,
	readRows,
	streamRows,
} from './worksheet'

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
const REL_COMMENTS = '/comments'

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
	if (!zip.has(path))
		throw new XlsxError('missing-part', `xlsx is missing a required part: ${path}`)
	return decoder.decode(await zip.read(path))
}

export class Worksheet {
	/** Sheet name as shown on Excel's tab. */
	readonly name: string
	readonly #info: SheetInfo
	readonly #xml: string
	readonly #context: DecodeContext
	readonly #rels: Map<string, Relationship> | undefined
	readonly #commentsXml: string | undefined

	#cells: Map<string, Cell> | undefined
	#merged: readonly string[] | undefined
	#hyperlinks: readonly Hyperlink[] | undefined
	#cellStyles: Map<string, number> | undefined
	#dimension: string | undefined
	#dimensionRead = false
	#comments: readonly Comment[] | undefined

	constructor(
		info: SheetInfo,
		xml: string,
		context: DecodeContext,
		rels?: Map<string, Relationship>,
		commentsXml?: string,
	) {
		this.name = info.name
		this.#info = info
		this.#xml = xml
		this.#context = context
		this.#rels = rels
		this.#commentsXml = commentsXml
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

	/**
	 * Hyperlinks declared on this sheet, in document order. Each carries the covered `ref` and,
	 * where present, a resolved external `target`, an in-workbook `location`, a `tooltip`, and a
	 * `display` override. Empty when none.
	 */
	get hyperlinks(): readonly Hyperlink[] {
		if (this.#hyperlinks === undefined) {
			this.#hyperlinks = parseHyperlinks(this.#xml, this.#rels)
		}
		return this.#hyperlinks
	}

	/**
	 * The number-format code applied to the cell at `ref` — a custom code like `"yyyy-mm-dd"`
	 * or `"0.00%"`, or a built-in one. `undefined` when the workbook has no style table or the
	 * id has no portable code. An unstyled or absent cell resolves to the default format (style
	 * 0, usually `"General"`), mirroring how date detection defaults.
	 */
	numberFormat(ref: string): string | undefined {
		return this.#context.styles?.formatCode(this.#cellStyleMap().get(ref))
	}

	/**
	 * The resolved style of the cell at `ref` — number format code, font, fill, border, and
	 * alignment (F4.1). Resolution shares the same effective-style map as {@link numberFormat}
	 * (cell `s` → row `customFormat` default → column default), so the two always agree.
	 * `undefined` for an unstyled cell, an absent cell, or a workbook with no style table —
	 * "no style" and "the default style" are deliberately the same answer. Objects are cached
	 * per distinct format record: two cells sharing a format return the same object.
	 */
	style(ref: string): CellStyle | undefined {
		return this.#context.styles?.cellStyle(this.#cellStyleMap().get(ref))
	}

	/**
	 * The sheet's declared used range in A1 notation (e.g. `"A1:E10"`, or a single cell), from
	 * the worksheet's `<dimension>`. `undefined` when the producer omits it — it is an optional
	 * hint, not authoritative, so treat a present value as advisory.
	 */
	get dimension(): string | undefined {
		if (!this.#dimensionRead) {
			this.#dimension = parseDimension(this.#xml)
			this.#dimensionRead = true
		}
		return this.#dimension
	}

	/**
	 * The comments anchored to cells on this sheet, in document order — each with its `ref`,
	 * resolved `author`, and plain `text`. Empty when the sheet has no comments part.
	 */
	get comments(): readonly Comment[] {
		if (this.#comments === undefined) {
			this.#comments = this.#commentsXml === undefined ? [] : parseComments(this.#commentsXml)
		}
		return this.#comments
	}

	#cellStyleMap(): Map<string, number> {
		if (this.#cellStyles === undefined) this.#cellStyles = parseCellStyles(this.#xml)
		return this.#cellStyles
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
			throw new XlsxError(
				'no-such-sheet',
				`no sheet named ${JSON.stringify(name)}; available: ${available}`,
			)
		}
		return worksheet
	}
}

interface LoadedWorkbook {
	readonly zip: ZipArchive
	/** Shared decode context (shared strings, styles, date system) for every sheet. */
	readonly context: DecodeContext
	/** Sheets in tab order, each with its resolved part path. */
	readonly sheets: ReadonlyArray<{ readonly info: SheetInfo; readonly path: string }>
}

// Read the small parts every sheet depends on — relationships, the workbook, shared strings,
// styles — and resolve each sheet's part path through the relationship graph. Worksheets
// themselves are NOT read here, so this stays cheap whether the caller wants random access
// (openXlsx) or a constant-memory stream (streamSheetRows).
async function loadWorkbook(
	source: Uint8Array | ArrayBuffer,
	options?: ReadOptions,
): Promise<LoadedWorkbook> {
	const bytes = source instanceof Uint8Array ? source : new Uint8Array(source)
	const zip = openZip(bytes, options)

	// Package relationships → the workbook part.
	const packageRels = parseRels(await readText(zip, '_rels/.rels'))
	const office = [...packageRels.values()].find((r) => r.type.endsWith(REL_OFFICE_DOCUMENT))
	if (office === undefined) {
		throw new XlsxError('not-xlsx', 'not an xlsx: no officeDocument relationship')
	}
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
	let styles: StyleTable | undefined
	const stylesRel = [...workbookRels.values()].find((r) => r.type.endsWith(REL_STYLES))
	if (stylesRel !== undefined && stylesRel.targetMode !== 'External') {
		const stylesPath = resolveTarget(workbookDir, stylesRel.target)
		if (zip.has(stylesPath)) {
			styles = parseStyles(decoder.decode(await zip.read(stylesPath)))
		}
	}
	const context: DecodeContext =
		styles !== undefined ? { sharedStrings, date1904, styles } : { sharedStrings, date1904 }

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

/**
 * Reader options. `maxPartBytes` caps the declared decompressed size of any single part — a
 * zip-bomb guard independent of the archive's own (untrusted) size fields. Omit for no ceiling.
 */
export interface ReadOptions {
	readonly maxPartBytes?: number
}

export async function openXlsx(
	source: Uint8Array | ArrayBuffer,
	options?: ReadOptions,
): Promise<Workbook> {
	const { zip, context, sheets } = await loadWorkbook(source, options)

	// Decompress each worksheet (so cell access is synchronous) and build the Worksheet.
	const infos: SheetInfo[] = []
	const byName = new Map<string, Worksheet>()
	for (const { info, path } of sheets) {
		const xml = decoder.decode(await zip.read(path))
		// The sheet's own relationships (xl/worksheets/_rels/sheetN.xml.rels) resolve hyperlink
		// r:ids and locate the comments part. Optional — a plain sheet has no rels part.
		const relsPath = relsPathFor(path)
		const rels = zip.has(relsPath)
			? parseRels(decoder.decode(await zip.read(relsPath)))
			: undefined

		// Comments live in a separate part linked from the worksheet rels.
		let commentsXml: string | undefined
		const commentsRel = rels && [...rels.values()].find((r) => r.type.endsWith(REL_COMMENTS))
		if (commentsRel !== undefined && commentsRel.targetMode !== 'External') {
			const commentsPath = resolveTarget(directoryOf(path), commentsRel.target)
			if (zip.has(commentsPath)) commentsXml = decoder.decode(await zip.read(commentsPath))
		}

		infos.push(info)
		// First definition wins if two sheets somehow share a name.
		if (!byName.has(info.name)) {
			byName.set(info.name, new Worksheet(info, xml, context, rels, commentsXml))
		}
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
	options?: ReadOptions,
): AsyncGenerator<Row> {
	const { zip, context, sheets } = await loadWorkbook(source, options)
	const first = sheets[0]
	if (first === undefined) throw new XlsxError('not-xlsx', 'xlsx has no readable worksheets')

	let path = first.path
	if (sheetName !== undefined) {
		const match = sheets.find((s) => s.info.name === sheetName)
		if (match === undefined) {
			const available = sheets.map((s) => s.info.name).join(', ')
			throw new XlsxError(
				'no-such-sheet',
				`no sheet named ${JSON.stringify(sheetName)}; available: ${available}`,
			)
		}
		path = match.path
	}

	yield* streamRows(zip.readStream(path), context)
}
