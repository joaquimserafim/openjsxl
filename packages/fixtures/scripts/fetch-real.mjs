// Download the real-producer .xlsx fixtures from their upstream sources into ../data.
//
// These are files exported by actual apps (Excel, LibreOffice, openpyxl), vendored from
// permissively-licensed projects so the reader is tested against genuine output — see
// ../THIRD_PARTY.md. Each entry is hash-pinned, so a re-run reproduces the committed bytes and
// any upstream drift (a changed or removed file) fails loudly rather than silently mutating a
// fixture. Run: pnpm fixtures:real  (needs network; not part of the test gate).
//
// To add a real fixture: drop it under ../data, add an entry here (file, url, license, sha256),
// and note it in ../THIRD_PARTY.md. Gathering fixtures that exercise each feature end-to-end is
// a great first contribution.

import { createHash } from "node:crypto"
import { writeFile } from "node:fs/promises"

const CALAMINE = "https://raw.githubusercontent.com/tafia/calamine/master/tests"
const CALAMINE_MIT = "MIT — Copyright (c) 2016 Johann Tuffe (tafia/calamine)"

/** @type {{ file: string, url: string, license: string, sha256: string }[]} */
const manifest = [
	{
		file: "merge_cells.xlsx",
		url: `${CALAMINE}/merge_cells.xlsx`,
		license: CALAMINE_MIT,
		sha256: "34f6b2779f49991859045791d8be0c5027e09b395f067c4363e38813d1e6431c",
	},
	{
		file: "merged_range.xlsx",
		url: `${CALAMINE}/merged_range.xlsx`,
		license: CALAMINE_MIT,
		sha256: "13553289594edda304fdd66e0176e8cb5dce365eff8f589b6d1290090f1bc575",
	},
	{
		file: "hyperlinks.xlsx",
		url: `${CALAMINE}/hyperlinks.xlsx`,
		license: CALAMINE_MIT,
		sha256: "ab148c6f8889d1a27d8a2dcc87156513a9ae54df9f2213a70e56c7710a108b2d",
	},
	{
		file: "date.xlsx",
		url: `${CALAMINE}/date.xlsx`,
		license: CALAMINE_MIT,
		sha256: "0440916ec1b76ee4dd54955d59c2f42b266e9f5e4fccafb8cc1ed875908b26d2",
	},
	{
		file: "date_1904.xlsx",
		url: `${CALAMINE}/date_1904.xlsx`,
		license: CALAMINE_MIT,
		sha256: "8b9f2b71a46833d7b9d2250d60b3c0948024f3d2b26f38ef990c0f2b2683fe08",
	},
	{
		file: "errors.xlsx",
		url: `${CALAMINE}/errors.xlsx`,
		license: CALAMINE_MIT,
		sha256: "8d2b17a3170a47647f18798dc3dbe7504aaeb256f9b73485c69d103b0e0021f8",
	},
	{
		file: "inventory-table.xlsx",
		url: `${CALAMINE}/inventory-table.xlsx`,
		license: CALAMINE_MIT,
		sha256: "d63fa89e938aad9635cd556f683b901cc1c118bf11b494b1156c4711e5329631",
	},
	{
		file: "any_sheets.xlsx",
		url: `${CALAMINE}/any_sheets.xlsx`,
		license: CALAMINE_MIT,
		sha256: "fba48e0141f847d80738a4417f0a336cba92e4ac692d8dd7516eac41e32c9e4f",
	},
]

const dataDir = new URL("../data/", import.meta.url)
let failures = 0

for (const { file, url, sha256 } of manifest) {
	try {
		const res = await fetch(url)
		if (!res.ok) throw new Error(`HTTP ${res.status}`)
		const bytes = new Uint8Array(await res.arrayBuffer())
		const got = createHash("sha256").update(bytes).digest("hex")
		if (got !== sha256) throw new Error(`sha256 mismatch: expected ${sha256}, got ${got}`)
		await writeFile(new URL(file, dataDir), bytes)
		console.log(`ok   ${file} (${bytes.length} bytes)`)
	} catch (err) {
		failures++
		console.error(`FAIL ${file}: ${err instanceof Error ? err.message : err}`)
	}
}

if (failures > 0) {
	console.error(`\n${failures} fixture(s) failed to fetch/verify.`)
	process.exitCode = 1
} else {
	console.log(`\n${manifest.length} real-producer fixtures verified.`)
}
