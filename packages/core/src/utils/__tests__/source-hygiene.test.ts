import { readdir, readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

// A raw NUL (0x00) inside a source file makes git and grep treat the whole file as binary, which
// silently defeats the repo's grep-based audits (a search for a symbol in such a file returns
// "no match" indistinguishable from a real absence) and hides the file's diff from review. It is
// also never legal in an XML document (see `isXmlSafe`), so it has no business in our source. Raw
// NULs have crept into template literals used as Map keys before (F10.6); this pins that they stay
// escaped (`\u0000`), which is the identical string at runtime but keeps the file plain text.
async function tsFiles(dir: string): Promise<string[]> {
	const out: string[] = [];
	for (const entry of await readdir(dir, { withFileTypes: true })) {
		const path = `${dir}/${entry.name}`;
		if (entry.isDirectory()) out.push(...(await tsFiles(path)));
		else if (entry.name.endsWith(".ts")) out.push(path);
	}
	return out;
}

describe("source hygiene", () => {
	it("no source file under core/src carries a raw NUL byte", async () => {
		const srcDir = fileURLToPath(new URL("../../", import.meta.url));
		const files = await tsFiles(srcDir);
		expect(files.length).toBeGreaterThan(50); // sanity: the walk found the tree
		const offenders: string[] = [];
		for (const file of files) {
			if ((await readFile(file)).includes(0x00)) offenders.push(file);
		}
		expect(offenders).toEqual([]);
	});
});
