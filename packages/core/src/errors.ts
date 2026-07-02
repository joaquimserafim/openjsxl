// A single typed error for the ways opening or reading an .xlsx can fail, discriminated by
// `code` so callers can branch without matching on message strings. Messages stay
// human-readable; `cause` carries the underlying error where there is one (e.g. a failed
// inflate). All file-level failures thrown by the reader and zip layers are XlsxError.

export type XlsxErrorCode =
	| "not-a-zip" // the bytes are not a ZIP archive at all
	| "not-xlsx" // a valid ZIP, but not an OOXML spreadsheet
	| "missing-part" // a required part is absent from the package
	| "corrupt-zip" // the ZIP structure is malformed, or an entry failed to inflate
	| "unsupported" // a valid but unsupported feature (ZIP64, encryption, unknown method)
	| "no-such-sheet" // the caller asked for a sheet name the workbook does not have
	| "part-too-large" // a part's decompressed size exceeds the caller's maxPartBytes limit
	| "invalid-input" // (writer) a value or option passed to writeXlsx can't be represented in .xlsx

export class XlsxError extends Error {
	/** Machine-readable discriminant; branch on this rather than the message. */
	readonly code: XlsxErrorCode

	constructor(code: XlsxErrorCode, message: string, options?: { cause?: unknown }) {
		super(message, options)
		this.name = "XlsxError"
		this.code = code
	}
}
