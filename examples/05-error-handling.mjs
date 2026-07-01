// 05 — Malformed input throws a typed XlsxError with a discriminating `code`.
//
//   node 05-error-handling.mjs
//   pnpm --filter openjsxl-examples errors
//
// The reader never leaks a bare TypeError/RangeError from a corrupt file: every failure is an
// `XlsxError` whose `code` you can branch on ('not-a-zip' | 'not-xlsx' | 'missing-part' |
// 'corrupt-zip' | 'unsupported' | 'no-such-sheet' | 'part-too-large').

import { readFile } from 'node:fs/promises'
import { openXlsx, XlsxError } from 'openjsxl'

// 1) Garbage bytes — not a zip at all.
try {
	await openXlsx(new Uint8Array([0x50, 0x4b, 1, 2, 3, 4]))
} catch (err) {
	if (!(err instanceof XlsxError)) throw err
	console.log(`garbage input   → XlsxError code="${err.code}"`)
}

// 2) A valid workbook, but asking for a sheet that isn't there.
try {
	const wb = await openXlsx(await readFile(new URL('./data/sample.xlsx', import.meta.url)))
	wb.sheet('Does Not Exist')
} catch (err) {
	if (!(err instanceof XlsxError)) throw err
	console.log(`missing sheet   → XlsxError code="${err.code}"`)
}

// 3) The zip-bomb guard: cap any single decompressed part.
try {
	const bytes = await readFile(new URL('./data/sample.xlsx', import.meta.url))
	await openXlsx(bytes, { maxPartBytes: 32 })
} catch (err) {
	if (!(err instanceof XlsxError)) throw err
	console.log(`part over limit → XlsxError code="${err.code}"`)
}
