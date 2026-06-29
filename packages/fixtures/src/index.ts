import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'

// Test corpus access. Real .xlsx binaries live under ./data — both programmatic
// fixtures (see scripts/generate.mjs, F0.2) and files produced by Excel, LibreOffice,
// and Google Sheets.

const dataDir = fileURLToPath(new URL('../data/', import.meta.url))

/** Absolute path to a committed fixture file under packages/fixtures/data. */
export function fixturePath(name: string): string {
	return `${dataDir}${name}`
}

/** Read a fixture file as raw bytes. */
export async function loadFixture(name: string): Promise<Uint8Array> {
	return new Uint8Array(await readFile(fixturePath(name)))
}
