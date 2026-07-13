import { readdirSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

// Test corpus access. Real .xlsx binaries live under ./data — both programmatic
// fixtures (see scripts/generate.mjs, F0.2) and files produced by Excel, LibreOffice,
// and Google Sheets.

const dataDir = fileURLToPath(new URL("../data/", import.meta.url));
const localDir = fileURLToPath(new URL("../local/", import.meta.url));

/**
 * Every committed fixture file name under ./data (the `.md` docs excluded), sorted for determinism.
 * Used by the fuzz harness (F9.4) to enumerate seed files for byte-mutation replay.
 */
export function listFixtures(): readonly string[] {
	return readdirSync(dataDir)
		.filter((f) => !f.endsWith(".md"))
		.sort();
}

/** Absolute path to a committed fixture file under packages/fixtures/data. */
export function fixturePath(name: string): string {
	return `${dataDir}${name}`;
}

/** Read a fixture file as raw bytes. */
export async function loadFixture(name: string): Promise<Uint8Array> {
	return new Uint8Array(await readFile(fixturePath(name)));
}

/**
 * Absolute path to a git-ignored, local-only fixture under packages/fixtures/local (e.g. a
 * differently-licensed real file). May not exist on a fresh clone — tests that use it must
 * skip when it is absent. See packages/fixtures/local/README.md.
 */
export function localFixturePath(name: string): string {
	return `${localDir}${name}`;
}

/** Read a local-only fixture as raw bytes (throws if the file is absent — guard with exists). */
export async function loadLocalFixture(name: string): Promise<Uint8Array> {
	return new Uint8Array(await readFile(localFixturePath(name)));
}
