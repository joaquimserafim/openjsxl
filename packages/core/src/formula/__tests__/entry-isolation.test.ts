import { readdirSync, readFileSync } from "node:fs";
import { join, relative } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

// The opt-in `openjsxl/formula` entry (parser + evaluator + 90+ built-ins) must never be pulled into
// the main `openjsxl` bundle: a consumer who never touches formulas must not pay for it, and the
// core bundle must stay byte-for-byte unchanged whether or not the entry exists (F8.1 contract;
// re-asserted by F11.9's scope line "do not pull the optional formula entry point into the core
// reader bundle"). This pins the IMPORT GRAPH at the source level — no build needed — so a future
// "shared scanner" refactor cannot silently drag the evaluator graph into the reader.

const SRC = fileURLToPath(new URL("../..", import.meta.url));
const FORMULA_DIR = join(SRC, "formula");

function tsFiles(dir: string, out: string[] = []): string[] {
	for (const entry of readdirSync(dir, { withFileTypes: true })) {
		const path = join(dir, entry.name);
		if (entry.isDirectory()) {
			if (entry.name !== "__tests__" && entry.name !== "node_modules") tsFiles(path, out);
		} else if (entry.name.endsWith(".ts") && !entry.name.endsWith(".d.ts")) {
			out.push(path);
		}
	}
	return out;
}

// Every `from "…"` specifier in a file that resolves into src/formula/.
function importsIntoFormula(file: string): string[] {
	const text = readFileSync(file, "utf8");
	const hits: string[] = [];
	for (const m of text.matchAll(/from\s+["']([^"']+)["']/g)) {
		const spec = m[1] as string;
		if (!spec.startsWith(".")) continue;
		const target = join(file, "..", spec);
		if (target === FORMULA_DIR || target.startsWith(`${FORMULA_DIR}/`)) hits.push(spec);
	}
	return hits;
}

describe("openjsxl/formula stays isolated from the main entry", () => {
	it("no core module outside src/formula/ imports anything from src/formula/", () => {
		const offenders: Record<string, string[]> = {};
		for (const file of tsFiles(SRC)) {
			if (file === FORMULA_DIR || file.startsWith(`${FORMULA_DIR}/`)) continue;
			// The `openjsxl/formula` entry IS src/formula/index.ts (tsup entry `formula`), already
			// excluded above; nothing outside that directory may reach into it.
			const hits = importsIntoFormula(file);
			if (hits.length > 0) offenders[relative(SRC, file)] = hits;
		}
		expect(offenders).toEqual({});
	});
});
