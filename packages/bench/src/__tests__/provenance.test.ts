import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ensureFormatFixture } from "../formats.mjs";
import {
	ensureCachedFixture,
	fixtureIdentity,
	librarySnapshot,
	lockedIntegrity,
	matchingPythonSnapshot,
	matchingSizeSnapshot,
	requireProvenance,
	verifyPackedArtifact,
} from "../provenance.mjs";
import { renderReport } from "../report.mjs";

const source = "https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz";
const artifact = Buffer.from("reviewed artifact bytes");
const integrity = `sha512-${createHash("sha512").update(artifact).digest("base64")}`;
const dataset = { workload: "numbers", rows: 1_000, columns: 10, format: "xlsx" };
const provenance = {
	schemaVersion: 1,
	measuredAt: "2024-02-03T04:05:06.000Z",
	machine: {
		cpu: "Recorded CPU",
		cores: 4,
		memoryGiB: 8,
		platform: "recorded-os",
		node: "v24.0.0",
	},
	libraries: [
		{
			id: "openjsxl",
			name: "openjsxl",
			packageName: "openjsxl",
			version: "1.0.0",
			source: "workspace",
			overrides: {},
			expectedIntegrity: null,
		},
		{
			id: "exceljs",
			name: "ExcelJS",
			packageName: "exceljs",
			version: "4.4.0",
			source: "4.4.0",
			overrides: { uuid: "11.1.1" },
			expectedIntegrity: integrity,
		},
		{
			id: "xlsx",
			name: "SheetJS",
			packageName: "xlsx",
			version: "0.20.3",
			source,
			overrides: {},
			expectedIntegrity: integrity,
		},
	],
	lockfileSha256: "recorded-lock",
	bundles: { core: "recorded-core", facade: "recorded-facade" },
};
const meta = {
	provenance,
	fixtures: [{ name: "read-numbers-10k.xlsx", sha256: "recorded-fixture" }],
	sizes: [{ key: "10k", rows: 1_000, cells: 10_000 }],
	workloads: ["numbers"],
	ops: ["read"],
};
const dirs: string[] = [];
const scratch = () => {
	const dir = mkdtempSync(join(tmpdir(), "openjsxl-bench-provenance-"));
	dirs.push(dir);
	return dir;
};
afterEach(() => {
	for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

describe("benchmark provenance", () => {
	it("records the installed official SheetJS version and reviewed ExcelJS override", () => {
		const libraries = librarySnapshot();
		expect(libraries.find((lib) => lib.id === "xlsx")).toMatchObject({
			version: "0.20.3",
			source,
		});
		expect(libraries.find((lib) => lib.id === "exceljs")).toMatchObject({
			version: "4.4.0",
			overrides: { uuid: "11.1.1" },
		});
	});

	it("rejects a legacy result instead of attaching today's version labels", () => {
		expect(() => requireProvenance(undefined)).toThrow("lacks measurement provenance");
		expect(() => renderReport([], { ...meta, provenance: undefined })).toThrow(
			"run `pnpm bench` again",
		);
	});

	it("renders the recorded date, machine, versions, source, and build even on another machine", () => {
		const text = renderReport([], meta);
		expect(text).toContain("on 2024-02-03");
		expect(text).toContain("Recorded CPU · 4 cores · 8 GiB · recorded-os · Node v24.0.0");
		expect(text).toContain("**SheetJS** `0.20.3`");
		expect(text).toContain(source);
		expect(text).toContain("override uuid 11.1.1");
		expect(text).toContain("recorded-core");
		expect(text).not.toContain("0.18.5");
	});

	it("revalidates cached reads, hides invalid timings, and lists every failed cell", () => {
		const text = renderReport(
			[
				{
					ok: true,
					op: "read",
					workload: "numbers",
					rows: 1_000,
					sink: 7_492_500,
					label: "openjsxl",
					sizeKey: "10k",
					format: "xlsb",
					section: "format-read",
					timeMs: 123,
					peakRss: 777 * 1024 ** 2,
				},
				{
					ok: false,
					op: "write",
					workload: "styled",
					label: "SheetJS",
					sizeKey: "10k",
					reason: "unsupported: SheetJS community build does not emit cell styles",
				},
			],
			meta,
		);
		expect(text).toContain("Failed or unsupported measurements");
		expect(text).toContain(
			"| openjsxl | read | numbers | xlsb | 10k | Read checksum mismatch:",
		);
		expect(text).toContain("received 7492500");
		expect(text).toContain("| SheetJS | write | styled | xlsx | 10k | unsupported:");
		expect(text).not.toContain("123 ms");
		expect(text).not.toContain("777 MB");
	});

	it("regenerates legacy caches, reuses matching bytes, and invalidates changed producers", async () => {
		const path = join(scratch(), "fixture.xlsx");
		writeFileSync(path, "old cache without metadata");
		const generate = vi.fn(async () => Buffer.from("controlled bytes"));
		const identity = fixtureIdentity(provenance, "xlsx", dataset);
		await ensureCachedFixture(path, identity, generate);
		expect(readFileSync(path, "utf8")).toBe("controlled bytes");
		await ensureCachedFixture(path, identity, generate);
		expect(generate).toHaveBeenCalledTimes(1);
		await ensureCachedFixture(path, { ...identity, lockfileSha256: "updated-lock" }, generate);
		expect(generate).toHaveBeenCalledTimes(2);
		await ensureCachedFixture(
			path,
			{ ...identity, producer: { source: "changed-distribution" } },
			generate,
		);
		expect(generate).toHaveBeenCalledTimes(3);
	});

	it("regenerates modified bytes and damaged metadata outside the measured worker", async () => {
		const path = join(scratch(), "fixture.ods");
		const generate = vi.fn(async () => Buffer.from("controlled"));
		const identity = fixtureIdentity(provenance, "xlsx", dataset);
		await ensureCachedFixture(path, identity, generate);
		writeFileSync(path, "substitute");
		await ensureCachedFixture(path, identity, generate);
		expect(generate).toHaveBeenCalledTimes(2);
		writeFileSync(`${path}.json`, "broken JSON");
		await ensureCachedFixture(path, identity, generate);
		expect(generate).toHaveBeenCalledTimes(3);
	});

	it("regenerates a real CSV when the row count changes under the same size key", async () => {
		const cache = `${scratch()}/`;
		const path = await ensureFormatFixture(
			cache,
			{ key: "10k", rows: 2 },
			"csv",
			"unused",
			provenance,
		);
		expect(readFileSync(path, "utf8").trim().split("\n")).toHaveLength(2);
		await ensureFormatFixture(cache, { key: "10k", rows: 3 }, "csv", "unused", provenance);
		expect(readFileSync(path, "utf8").trim().split("\n")).toHaveLength(3);
		expect(JSON.parse(readFileSync(`${path}.json`, "utf8")).identity.dataset).toEqual({
			workload: "numbers",
			rows: 3,
			columns: 10,
			format: "csv",
		});
	});

	it("extracts only the exact locked comparator resolution and rejects missing or changed forms", () => {
		const block = `\n  xlsx@${source}:\n    resolution: {integrity: ${integrity}, tarball: ${source}}\n    version: 0.20.3\n`;
		const lockfile = `lockfileVersion: '9.0'\n\nimporters:\n\n  packages/bench:\n    devDependencies: {}\n\npackages:\n${block}\nsnapshots:\n  xlsx@${source}: {}\n`;
		expect(lockedIntegrity(lockfile, "xlsx", source)).toBe(integrity);
		expect(
			lockedIntegrity(
				`---\nlockfileVersion: '9.0'\npackages:\n  pnpm@11.28.6: {}\nsnapshots:\n  pnpm@11.28.6: {}\n---\n${lockfile}`,
				"xlsx",
				source,
			),
		).toBe(integrity);
		expect(() => lockedIntegrity(lockfile, "xlsx", "0.20.3")).toThrow(
			"locked artifact integrity",
		);
		expect(() =>
			lockedIntegrity(
				lockfile.replace(`tarball: ${source}`, "tarball: https://other.invalid/file.tgz"),
				"xlsx",
				source,
			),
		).toThrow("locked artifact integrity");
		expect(() =>
			lockedIntegrity(
				lockfile.replace("\nsnapshots:", `${block}\nsnapshots:`),
				"xlsx",
				source,
			),
		).toThrow("locked artifact integrity");
		expect(() =>
			lockedIntegrity(
				lockfile.replace("resolution: {", "resolution:\n      {"),
				"xlsx",
				source,
			),
		).toThrow("locked artifact integrity");
	});

	it("checks the actual packed bytes and npm metadata against the lockfile artifact", () => {
		const library = { packageName: "xlsx", version: "0.20.3", expectedIntegrity: integrity };
		const packed = { name: "xlsx", version: "0.20.3", integrity };
		expect(verifyPackedArtifact(library, packed, artifact)).toBe(integrity);
		expect(() =>
			verifyPackedArtifact(library, packed, Buffer.from("same version, different bytes")),
		).toThrow("integrity mismatch");
		const changed = Buffer.from("repacked artifact");
		const changedIntegrity = `sha512-${createHash("sha512").update(changed).digest("base64")}`;
		expect(() =>
			verifyPackedArtifact(library, { ...packed, integrity: changedIntegrity }, changed),
		).toThrow("integrity mismatch");
		expect(() =>
			verifyPackedArtifact(library, { ...packed, integrity: changedIntegrity }, artifact),
		).toThrow("integrity mismatch");
	});

	it("excludes legacy Python and size measurements rather than publishing stale figures", () => {
		const text = renderReport([], meta, {
			python: { date: "2020-01-01", results: [] },
			sizes: { libs: [{ name: "SheetJS", spec: "xlsx@0.18.5" }] },
		});
		expect(text).toContain("No matching Python reference measurements");
		expect(text).toContain("No matching size measurements");
		expect(text).not.toContain("2020-01-01");
	});

	it("requires Python measurements to match both the JS run and exact fixture hashes", () => {
		const data = {
			schemaVersion: 1,
			benchmarkMeasuredAt: provenance.measuredAt,
			fixtures: meta.fixtures,
		};
		expect(matchingPythonSnapshot(data, meta)).toBe(true);
		expect(matchingPythonSnapshot({ ...data, benchmarkMeasuredAt: "another-run" }, meta)).toBe(
			false,
		);
		expect(
			matchingPythonSnapshot(
				{ ...data, fixtures: [{ name: "read-numbers-10k.xlsx", sha256: "changed" }] },
				meta,
			),
		).toBe(false);
		expect(matchingPythonSnapshot({ ...data, fixtures: undefined }, meta)).toBe(false);
		const text = renderReport([], meta, {
			python: {
				...data,
				results: [],
				libraries: { openpyxl: "3.1.5", "python-calamine": "0.6.2" },
			},
		});
		expect(text).toContain("Libraries: openpyxl 3.1.5; python-calamine 0.6.2.");
	});

	it("requires size version, source, override, and dependency resolution to match", () => {
		const libs = provenance.libraries.map((lib) => ({
			...lib,
			integrity: lib.expectedIntegrity,
			source: lib.id === "openjsxl" ? "openjsxl@1.0.0" : lib.source,
		}));
		const data = { schemaVersion: 1, lockfileSha256: provenance.lockfileSha256, libs };
		expect(matchingSizeSnapshot(data, provenance)).toBe(true);
		expect(matchingSizeSnapshot({ ...data, lockfileSha256: "old" }, provenance)).toBe(false);
		for (const change of [
			{ version: "0.18.5" },
			{ source: "0.20.3" },
			{ overrides: { uuid: "8.3.2" } },
			{ integrity: "different-artifact" },
			{ expectedIntegrity: undefined },
		]) {
			expect(
				matchingSizeSnapshot(
					{
						...data,
						libs: libs.map((lib) => (lib.id === "xlsx" ? { ...lib, ...change } : lib)),
					},
					provenance,
				),
			).toBe(false);
		}
	});
});

import { createHash } from "node:crypto";
