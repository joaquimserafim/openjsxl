import { createHash } from "node:crypto";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { machineStamp } from "./machine.mjs";

const require = createRequire(import.meta.url);
const ROOT = fileURLToPath(new URL("../../../", import.meta.url));
const BENCH = fileURLToPath(new URL("../package.json", import.meta.url));
export const PROVENANCE_VERSION = 1;
export const FIXTURE_REVISION = 1;

export function sha256(bytes) {
	return createHash("sha256").update(bytes).digest("hex");
}

// Deliberately accept only pnpm's generated inline resolution form for our exact comparator
// keys. A lockfile format change must be reviewed instead of silently losing the artifact pin.
export function lockedIntegrity(lockfile, packageName, source) {
	// pnpm 11 also writes a separate bootstrap document for the package manager itself.
	const documents = lockfile
		.split(/^---\n/m)
		.filter((doc) => doc.includes("\n  packages/bench:\n"));
	const sections = documents.length === 1 ? documents[0].split("\npackages:\n") : [];
	const blocks =
		sections.length === 2
			? sections[1].split("\nsnapshots:\n")[0].split(`\n  ${packageName}@${source}:\n`)
			: [];
	const resolution = blocks.length === 2 ? blocks[1].split("\n")[0] : "";
	const match =
		/^ {4}resolution: \{integrity: (sha512-[A-Za-z0-9+/]{86}==)(?:, tarball: ([^}]+))?\}$/.exec(
			resolution,
		);
	if (!match || (source.startsWith("https://") && match[2] !== source)) {
		throw new Error(
			`Missing or unsupported locked artifact integrity for ${packageName}@${source}`,
		);
	}
	return match[1];
}

export function verifyPackedArtifact(library, packed, bytes) {
	if (packed?.name !== library.packageName || packed.version !== library.version) {
		throw new Error(`Unexpected artifact identity for ${library.packageName}`);
	}
	const integrity = `sha512-${createHash("sha512").update(bytes).digest("base64")}`;
	if (
		packed.integrity !== integrity ||
		(library.expectedIntegrity !== null && library.expectedIntegrity !== integrity)
	) {
		throw new Error(`Artifact integrity mismatch for ${library.packageName}`);
	}
	return integrity;
}

// Some distributions (including SheetJS 0.20.3) do not export package.json. Walk upward from
// their resolved entry instead of guessing a package directory from a URL or version string.
function installedPackage(name) {
	let dir = dirname(require.resolve(name));
	for (;;) {
		const path = join(dir, "package.json");
		if (existsSync(path)) {
			const pkg = JSON.parse(readFileSync(path, "utf8"));
			if (pkg.name === name) return { pkg, path };
		}
		const parent = dirname(dir);
		if (parent === dir) throw new Error(`Cannot locate installed ${name} package metadata`);
		dir = parent;
	}
}

export function librarySnapshot(lockfile = readFileSync(join(ROOT, "pnpm-lock.yaml"), "utf8")) {
	const manifest = JSON.parse(readFileSync(BENCH, "utf8"));
	const libraries = [
		{ id: "openjsxl", name: "openjsxl", packageName: "openjsxl", source: "workspace" },
		{
			id: "exceljs",
			name: "ExcelJS",
			packageName: "exceljs",
			source: manifest.devDependencies.exceljs,
		},
		{ id: "xlsx", name: "SheetJS", packageName: "xlsx", source: manifest.devDependencies.xlsx },
	];
	return libraries.map((library) => {
		const { pkg, path } = installedPackage(library.packageName);
		const overrides =
			library.id === "exceljs"
				? { uuid: createRequire(path)("uuid/package.json").version }
				: {};
		const expectedIntegrity =
			library.id === "openjsxl"
				? null
				: lockedIntegrity(lockfile, library.packageName, library.source);
		return { ...library, version: pkg.version, overrides, expectedIntegrity };
	});
}

export function captureProvenance() {
	const lockfile = readFileSync(join(ROOT, "pnpm-lock.yaml"), "utf8");
	return {
		schemaVersion: PROVENANCE_VERSION,
		measuredAt: new Date().toISOString(),
		machine: machineStamp(),
		libraries: librarySnapshot(lockfile),
		lockfileSha256: sha256(lockfile),
		bundles: {
			core: sha256(readFileSync(join(ROOT, "packages/core/dist/index.js"))),
			facade: sha256(readFileSync(join(ROOT, "packages/openjsxl/dist/index.js"))),
		},
	};
}

export function requireProvenance(provenance) {
	if (
		provenance?.schemaVersion !== PROVENANCE_VERSION ||
		typeof provenance.measuredAt !== "string" ||
		!Number.isFinite(Date.parse(provenance.measuredAt)) ||
		!provenance.machine?.node ||
		typeof provenance.lockfileSha256 !== "string" ||
		typeof provenance.bundles?.core !== "string" ||
		!Array.isArray(provenance.libraries) ||
		!["openjsxl", "exceljs", "xlsx"].every((id) =>
			provenance.libraries.some(
				(lib) =>
					lib.id === id &&
					lib.version &&
					lib.source &&
					(id === "openjsxl" ||
						/^sha512-[A-Za-z0-9+/]{86}==$/.test(lib.expectedIntegrity)),
			),
		)
	) {
		throw new Error("Benchmark cache lacks measurement provenance; run `pnpm bench` again.");
	}
	return provenance;
}

// Cache identity is for reproducibility, not an authentication boundary. Invalidate on changes
// to the producer artifact, dependency resolution, or controlled dataset/fixture recipe.
export function fixtureIdentity(provenance, producer, dataset) {
	return {
		revision: FIXTURE_REVISION,
		producer:
			producer === "csv"
				? { id: "csv", source: "workloads.mjs" }
				: provenance.libraries.find((lib) => lib.id === producer),
		lockfileSha256: provenance.lockfileSha256,
		dataset,
	};
}

export async function ensureCachedFixture(path, identity, generate) {
	const sidecar = `${path}.json`;
	if (existsSync(path) && existsSync(sidecar)) {
		try {
			const cached = JSON.parse(readFileSync(sidecar, "utf8"));
			if (
				JSON.stringify(cached.identity) === JSON.stringify(identity) &&
				cached.sha256 === sha256(readFileSync(path))
			)
				return path;
		} catch {
			// A legacy, partial, or damaged cache is regenerated from the controlled recipe.
		}
	}
	const bytes = await generate();
	writeFileSync(path, bytes);
	writeFileSync(sidecar, JSON.stringify({ identity, sha256: sha256(bytes) }, null, 2));
	return path;
}

export function matchingSizeSnapshot(data, provenance) {
	if (
		data?.schemaVersion !== PROVENANCE_VERSION ||
		!Array.isArray(data.libs) ||
		data.lockfileSha256 !== provenance.lockfileSha256
	)
		return false;
	return provenance.libraries.every((library) => {
		const measured = data.libs.find((lib) => lib.packageName === library.packageName);
		return (
			measured?.version === library.version &&
			measured.source ===
				(library.id === "openjsxl"
					? `${library.packageName}@${library.version}`
					: library.source) &&
			measured.expectedIntegrity === library.expectedIntegrity &&
			(library.id === "openjsxl" || measured.integrity === library.expectedIntegrity) &&
			JSON.stringify(measured.overrides) === JSON.stringify(library.overrides)
		);
	});
}

export function matchingPythonSnapshot(data, meta) {
	return (
		data?.schemaVersion === PROVENANCE_VERSION &&
		Array.isArray(data.fixtures) &&
		Array.isArray(meta.fixtures) &&
		typeof data.benchmarkMeasuredAt === "string" &&
		data.benchmarkMeasuredAt === meta.provenance.measuredAt &&
		JSON.stringify(data.fixtures) === JSON.stringify(meta.fixtures)
	);
}
