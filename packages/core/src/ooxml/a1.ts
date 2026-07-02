// A1 reference helpers. Excel columns use bijective base-26 (A=1 … Z=26, AA=27),
// which is not ordinary base-26 because there is no zero digit.

export interface CellRef {
	/** 1-based column index (A = 1). */
	readonly col: number
	/** 1-based row index. */
	readonly row: number
}

const CODE_UPPER_A = 65
const CODE_UPPER_Z = 90
const CODE_LOWER_A = 97
const CODE_LOWER_Z = 122

export function columnToIndex(letters: string): number {
	if (letters.length === 0) throw new Error("empty column reference")
	let index = 0
	for (let i = 0; i < letters.length; i++) {
		const code = letters.charCodeAt(i)
		let value = 0
		if (code >= CODE_UPPER_A && code <= CODE_UPPER_Z) value = code - CODE_UPPER_A + 1
		else if (code >= CODE_LOWER_A && code <= CODE_LOWER_Z) value = code - CODE_LOWER_A + 1
		else throw new Error(`invalid column reference: ${letters}`)
		index = index * 26 + value
		// An absurdly long ref (far beyond Excel's XFD/16384 limit) overflows past exact
		// integer precision — bail before it silently becomes a lossy float or Infinity, which
		// would poison downstream column arithmetic and formatRef. Bailing here also caps work
		// on a megabyte-long attacker-supplied ref. Callers that tolerate bad refs (the reader's
		// safeColumn) catch this and fall back to positional addressing.
		if (index > Number.MAX_SAFE_INTEGER)
			throw new Error(`column reference too large: ${letters}`)
	}
	return index
}

export function indexToColumn(index: number): string {
	if (!Number.isInteger(index) || index < 1) throw new Error(`invalid column index: ${index}`)
	let remaining = index
	let letters = ""
	while (remaining > 0) {
		const digit = (remaining - 1) % 26
		letters = String.fromCharCode(CODE_UPPER_A + digit) + letters
		remaining = Math.floor((remaining - 1) / 26)
	}
	return letters
}

const A1_PATTERN = /^([A-Za-z]+)([1-9][0-9]*)$/

export function parseRef(ref: string): CellRef {
	const match = A1_PATTERN.exec(ref)
	if (match === null) throw new Error(`invalid A1 reference: ${ref}`)
	return {
		col: columnToIndex(match[1] as string),
		row: Number.parseInt(match[2] as string, 10),
	}
}

export function formatRef(ref: CellRef): string {
	if (!Number.isInteger(ref.row) || ref.row < 1) throw new Error(`invalid row index: ${ref.row}`)
	return `${indexToColumn(ref.col)}${ref.row}`
}
