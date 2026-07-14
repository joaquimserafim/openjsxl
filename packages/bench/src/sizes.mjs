// Library size matrix. Measures the real cost of adding each spreadsheet library to a project — the
// number openjsxl's zero-dependency design targets — by doing a CLEAN production install of each in a
// throwaway directory and measuring what lands on disk. Every number comes from a measurement (npm
// install / du / npm view), never an estimate. Writes ../.cache/sizes.json; report.mjs renders it.
//
//   node src/sizes.mjs            # -> packages/bench/.cache/sizes.json
//   pnpm bench --render-only      # re-render docs with the size table filled in
//
// Method per library: `npm install --omit=dev` the published version into a fresh temp project, then
// `npm ls --all --omit=dev` for the exact transitive package count, `du -sk` for the on-disk
// footprint (the library + every runtime dependency), and `npm view … dist.unpackedSize` for the
// package's OWN unpacked size. openjsxl is measured at its last published version (0.6.0); M7 adds
// reader code but ZERO dependencies, so its dependency/footprint story is unchanged.

import { execSync } from "node:child_process";
import { existsSync, mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const CACHE = fileURLToPath(new URL("../.cache/", import.meta.url));

// Every library is measured the SAME way — a clean `npm install` of its published package — so the
// footprints compare like for like. openjsxl tracks `latest` (not a pinned tag), so the row stays
// current with each release automatically; the competitors are pinned to a specific version.
const LIBS = [
	{
		name: "openjsxl",
		spec: "openjsxl",
		note: "its one dependency is its own `@openjsxl/core` — zero third-party packages. Measured from the published `latest`, so this row moves with each release.",
	},
	{ name: "ExcelJS", spec: "exceljs@4.4.0" },
	{ name: "SheetJS", spec: "xlsx@0.18.5" },
];

function sh(cmd, cwd) {
	return execSync(cmd, { cwd, encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] }).trim();
}

function measure(lib) {
	const dir = mkdtempSync(join(tmpdir(), `openjsxl-size-${lib.name}-`));
	process.stdout.write(`▶ sizing ${lib.spec} … `);
	sh("npm init -y", dir);
	sh(`npm install --omit=dev --no-audit --no-fund --loglevel=error ${lib.spec}`, dir);
	// `npm ls --all --parseable` prints one line per installed package: the ROOT project, the library
	// ITSELF, then its transitive deps — so the dependency count excludes the first two (counting the
	// library as its own dependency inflated every row by one; post-M7 review).
	const lsLines = sh("npm ls --all --omit=dev --parseable", dir).split("\n").filter(Boolean);
	const totalDeps = Math.max(0, lsLines.length - 2);
	const directDeps = Object.keys(
		JSON.parse(sh(`npm view ${lib.spec} dependencies --json`) || "{}"),
	).length;
	const installKB = Number(sh("du -sk node_modules", dir).split(/\s+/)[0]);
	const ownBytes = Number(sh(`npm view ${lib.spec} dist.unpackedSize`));
	// Record the RESOLVED version, so a floating spec like `openjsxl` (= latest) shows its real
	// number in the table rather than the literal "latest"/bare name.
	const resolvedVersion = sh(`npm view ${lib.spec} version`);
	const pkgName = lib.spec.split("@")[0];
	const out = {
		name: lib.name,
		spec: `${pkgName}@${resolvedVersion}`,
		directDeps,
		totalDeps,
		installKB,
		ownBytes,
	};
	if (lib.note) out.note = lib.note;
	process.stdout.write(
		`${directDeps} direct / ${totalDeps} total deps · ${(installKB / 1024).toFixed(2)} MB installed\n`,
	);
	return out;
}

function main() {
	mkdirSync(CACHE, { recursive: true });
	const libs = LIBS.map(measure);
	const node = process.version;
	const out = { date: new Date().toISOString().slice(0, 10), node, libs };
	const path = join(CACHE, "sizes.json");
	writeFileSync(path, JSON.stringify(out, null, 2));
	process.stdout.write(`\nWrote ${path}\n`);
	if (!existsSync(join(CACHE, "js-results.json"))) {
		process.stdout.write("Run `pnpm bench` (or `--render-only`) to fold this into docs.\n");
	}
}

main();
