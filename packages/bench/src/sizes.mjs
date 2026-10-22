// Out-of-band size measurements of exact reviewed distributions. npm pack supplies the artifact
// metadata; npm installs that local tarball with lifecycle scripts disabled. No floating latest
// lookup, URL-to-package-name guessing, or registry-only fallback for the SheetJS distribution.
import { execFileSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { captureProvenance, PROVENANCE_VERSION, verifyPackedArtifact } from "./provenance.mjs";

const CACHE = fileURLToPath(new URL("../.cache/", import.meta.url));

function run(command, args, cwd) {
	return execFileSync(command, args, {
		cwd,
		encoding: "utf8",
		stdio: ["ignore", "pipe", "inherit"],
	}).trim();
}

function measure(library) {
	const source =
		library.id === "openjsxl" ? `${library.packageName}@${library.version}` : library.source;
	const spec = source.startsWith("https://")
		? source
		: `${library.packageName}@${library.version}`;
	const dir = mkdtempSync(join(tmpdir(), `openjsxl-size-${library.packageName}-`));
	process.stdout.write(`▶ sizing ${spec} … `);
	try {
		const [packed] = JSON.parse(
			run(
				"npm",
				["pack", spec, "--ignore-scripts", "--json", "--pack-destination", dir],
				dir,
			),
		);
		const integrity = verifyPackedArtifact(
			library,
			packed,
			readFileSync(join(dir, packed.filename)),
		);
		const overrides = library.id === "exceljs" ? { exceljs: library.overrides } : {};
		writeFileSync(
			join(dir, "package.json"),
			JSON.stringify({
				name: "openjsxl-size-probe",
				private: true,
				version: "0.0.0",
				dependencies: { [library.packageName]: `file:${join(dir, packed.filename)}` },
				overrides,
			}),
		);
		run(
			"npm",
			[
				"install",
				"--omit=dev",
				"--ignore-scripts",
				"--no-audit",
				"--no-fund",
				"--loglevel=error",
			],
			dir,
		);
		const installed = JSON.parse(
			readFileSync(join(dir, "node_modules", library.packageName, "package.json"), "utf8"),
		);
		if (installed.version !== packed.version)
			throw new Error(`Installed version differs from packed ${spec}`);
		const lines = run("npm", ["ls", "--all", "--omit=dev", "--parseable"], dir)
			.split("\n")
			.filter(Boolean);
		const out = {
			name: library.name,
			packageName: library.packageName,
			version: installed.version,
			source,
			integrity,
			expectedIntegrity: library.expectedIntegrity,
			overrides: library.overrides,
			directDeps: Object.keys(installed.dependencies ?? {}).length,
			totalDeps: Math.max(0, lines.length - 2),
			installKB: Number(run("du", ["-sk", "node_modules"], dir).split(/\s+/)[0]),
			ownBytes: packed.unpackedSize,
		};
		if (library.id === "openjsxl") {
			out.note =
				"Footprint of the exact published version; timing uses the workspace bundle identified above. Its sole runtime dependency is its own @openjsxl/core.";
		}
		if (library.id === "exceljs")
			out.note = `Measured with the benchmark's uuid ${library.overrides.uuid} override; other transitives resolve during this dated clean install.`;
		if (library.id === "xlsx") out.note = `Official distribution: ${source}.`;
		process.stdout.write(
			`${out.directDeps} direct / ${out.totalDeps} total deps · ${(out.installKB / 1024).toFixed(2)} MB installed\n`,
		);
		return out;
	} finally {
		rmSync(dir, { recursive: true, force: true });
	}
}

const provenance = captureProvenance();
const libs = provenance.libraries.map(measure);
mkdirSync(CACHE, { recursive: true });
const path = join(CACHE, "sizes.json");
writeFileSync(
	path,
	JSON.stringify(
		{
			schemaVersion: PROVENANCE_VERSION,
			date: new Date().toISOString().slice(0, 10),
			node: process.version,
			lockfileSha256: provenance.lockfileSha256,
			libs,
		},
		null,
		2,
	),
);
process.stdout.write(`\nWrote ${path}. Re-render with pnpm bench --render-only.\n`);
