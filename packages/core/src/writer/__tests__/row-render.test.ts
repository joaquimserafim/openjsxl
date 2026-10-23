import { describe, expect, it } from "vitest";
import { bestDoublingRatio, LINEAR_RATIO_LIMIT } from "../../ooxml/__tests__/complexity";
import {
	forEachRowColumn,
	type RowListings,
	renderCell,
	renderStreamRow,
	SPARSE_ROW_SLACK,
} from "../sheet";
import { createStyleRegistry } from "../styles";
import type { CellInput } from "../types";
import { countingRow, farRightRows, type RowReads } from "./sparse-row-fixtures";

// F11.54: the row and cell renderers both writers share. A row visits only its populated columns.

const empty = { value: undefined, formula: undefined, style: undefined } as const;

describe("forEachRowColumn", () => {
	it("visits a dense row's columns, probing each index once and never listing keys", () => {
		const reads: RowReads = { probes: 0, keyListings: 0 };
		const visits: [number, boolean][] = [];
		forEachRowColumn(countingRow([1, 2, 3], reads), 3, (c, own) => visits.push([c, own]));
		expect(visits).toEqual([
			[0, true],
			[1, true],
			[2, true],
		]);
		expect(reads).toEqual({ probes: 3, keyListings: 0 });
	});

	it("skips holes and visits nothing for an empty row", () => {
		const row: CellInput[] = [];
		row[0] = 1;
		row[2] = 3;
		const visits: number[] = [];
		forEachRowColumn(row, 3, (c) => visits.push(c));
		forEachRowColumn([], 0, (c) => visits.push(c));
		expect(visits).toEqual([0, 2]);
	});

	it("probes each column just before its visit, so a cell filled during a visit is seen", () => {
		const row: CellInput[] = [];
		row[0] = "first";
		const visits: number[] = [];
		forEachRowColumn(row, 3, (c) => {
			visits.push(c);
			if (c === 0) row[2] = "added";
		});
		expect(visits).toEqual([0, 2]);
	});

	it("lists the row's keys once its holes outnumber its visits by more than the slack", () => {
		expect(SPARSE_ROW_SLACK).toBe(32);
		const reads: RowReads = { probes: 0, keyListings: 0 };
		const visits: number[] = [];
		forEachRowColumn(countingRow(farRightRows(1, 16_384)[0] ?? [], reads), 16_384, (c) =>
			visits.push(c),
		);
		expect(visits).toEqual([16_383]);
		// 33 probes before the switch, then one probe for the listed column.
		expect(reads).toEqual({ probes: SPARSE_ROW_SLACK + 2, keyListings: 1 });
	});

	it("probes every index while the holes stay within the slack", () => {
		const reads: RowReads = { probes: 0, keyListings: 0 };
		const row: CellInput[] = [];
		row[SPARSE_ROW_SLACK] = "x";
		const visits: number[] = [];
		forEachRowColumn(countingRow(row, reads), SPARSE_ROW_SLACK + 1, (c) => visits.push(c));
		expect(visits).toEqual([SPARSE_ROW_SLACK]);
		expect(reads).toEqual({ probes: SPARSE_ROW_SLACK + 1, keyListings: 0 });
	});

	it("lists a reused row object once per listings map", () => {
		const reads: RowReads = { probes: 0, keyListings: 0 };
		const row = countingRow(farRightRows(1, 16_384)[0] ?? [], reads);
		const listings: RowListings = new WeakMap();
		const visits: number[] = [];
		for (let r = 0; r < 5; r++)
			forEachRowColumn(row, 16_384, (c) => visits.push(c), undefined, listings);
		expect(visits).toEqual([16_383, 16_383, 16_383, 16_383, 16_383]);
		expect(reads.keyListings).toBe(1);
	});

	it("visits captured columns below the row length in both phases, without a probe saying own", () => {
		const short: CellInput[] = [];
		short[1] = "h";
		const shortVisits: [number, boolean][] = [];
		forEachRowColumn(short, 2, (c, own) => shortVisits.push([c, own]), new Set([0, 5]));
		expect(shortVisits).toEqual([
			[0, false],
			[1, true],
		]);
		const long: CellInput[] = [];
		long[200] = "h";
		const longVisits: [number, boolean][] = [];
		forEachRowColumn(
			long,
			300,
			(c, own) => longVisits.push([c, own]),
			new Set([100, 200, 299, 300]),
		);
		expect(longVisits).toEqual([
			[100, false],
			[200, true],
			[299, false],
		]);
	});

	it("sorts a Proxy's key listing and keeps only canonical indices below the length", () => {
		const row: CellInput[] = [];
		row[0] = 1;
		row[400] = 2;
		row[900] = 3;
		const named = Object.assign(row, { extra: "named", "01": "not an index" });
		const reads: RowReads = { probes: 0, keyListings: 0 };
		const visits: number[] = [];
		forEachRowColumn(countingRow(named, reads, true), 1000, (c) => visits.push(c));
		forEachRowColumn(row, 500, (c) => visits.push(c));
		expect(visits).toEqual([0, 400, 900, 0, 400]);
	});
});

describe("renderCell", () => {
	it("renders nothing for an empty, unstyled cell, before building its ref", () => {
		const styles = createStyleRegistry();
		expect(renderCell(1, 1, empty, false, styles)).toBeUndefined();
		expect(renderCell(1, 1, { ...empty, value: null }, false, styles)).toBeUndefined();
		// Column 0 has no A1 ref: building one would throw.
		expect(renderCell(0, 1, empty, false, styles)).toBeUndefined();
		expect(() => renderCell(0, 1, { ...empty, value: 1 }, false, styles)).toThrow(
			/invalid column index/,
		);
	});

	it("renders each value type", () => {
		const fresh = createStyleRegistry();
		expect([
			renderCell(
				2,
				1,
				{ ...empty, value: null, style: { font: { bold: true } } },
				false,
				fresh,
			),
			renderCell(1, 1, { ...empty, value: "a" }, false, fresh),
			renderCell(1, 1, { ...empty, value: true }, false, fresh),
			renderCell(1, 1, { ...empty, value: 2.5 }, false, fresh),
			renderCell(1, 1, { ...empty, formula: "1+1", value: 2 }, false, fresh),
		]).toEqual([
			'<c r="B1" s="1"/>',
			'<c r="A1" t="inlineStr"><is><t>a</t></is></c>',
			'<c r="A1" t="b"><v>1</v></c>',
			'<c r="A1"><v>2.5</v></c>',
			'<c r="A1"><f>1+1</f><v>2</v></c>',
		]);
		expect(
			renderCell(
				1,
				1,
				{ ...empty, value: new Date(Date.UTC(2020, 0, 1)) },
				false,
				createStyleRegistry(),
			),
		).toBe('<c r="A1" s="1"><v>43831</v></c>');
	});

	it.each([
		[Number.NaN, "cell A1: NaN is not a finite number"],
		[new Date(Number.NaN), "cell A1: invalid Date"],
		[1n, "cell A1: unsupported cell value type"],
	])("rejects %s, which it cannot represent", (value, message) => {
		expect(() =>
			// A JS caller can pass a value the types forbid.
			renderCell(1, 1, { ...empty, value: value as never }, false, createStyleRegistry()),
		).toThrow(expect.objectContaining({ code: "invalid-input", message }));
	});
});

describe("renderStreamRow", () => {
	const styles = createStyleRegistry();

	it("renders nothing for a missing row and a property-only row for row attributes", () => {
		expect(renderStreamRow("S", 3, undefined, new Map(), false, styles)).toBe("");
		expect(renderStreamRow("S", 3, [], new Map([[3, ' ht="20"']]), false, styles)).toBe(
			'<row r="3" ht="20"/>',
		);
	});

	it("renders only a row's populated columns", () => {
		const row: CellInput[] = [];
		row[1] = 7;
		row[16_383] = "z";
		expect(renderStreamRow("S", 3, row, new Map(), false, styles)).toBe(
			'<row r="3"><c r="B3"><v>7</v></c><c r="XFD3" t="inlineStr"><is><t>z</t></is></c></row>',
		);
	});

	it("lists a reused sparse row once across the rows that share a listings map", () => {
		const reads: RowReads = { probes: 0, keyListings: 0 };
		const row = countingRow(farRightRows(1, 16_384)[0] ?? [], reads);
		const listings: RowListings = new WeakMap();
		for (let r = 1; r <= 3; r++)
			renderStreamRow("S", r, row, new Map(), false, styles, listings);
		expect(reads.keyListings).toBe(1);
	});

	it("rejects a non-array row and an over-long row", () => {
		// A JS caller can pass a row the types forbid.
		expect(() => renderStreamRow("S", 3, "abc" as never, new Map(), false, styles)).toThrow(
			/a row must be an array of cell values/,
		);
		expect(() => renderStreamRow("S", 3, new Array(16_385), new Map(), false, styles)).toThrow(
			/more than 16384 cells/,
		);
	});

	it("is linear in populated cells on far-right rows", () => {
		const ratio = bestDoublingRatio((n) => {
			const rows = farRightRows(n, n);
			return () => {
				const attrs = new Map<number, string>();
				const registry = createStyleRegistry();
				const listings: RowListings = new WeakMap();
				let xml = "";
				for (let r = 0; r < rows.length; r++) {
					xml += renderStreamRow("S", r + 1, rows[r], attrs, false, registry, listings);
				}
				return xml;
			};
		}, 2_000);
		expect(ratio).toBeLessThan(LINEAR_RATIO_LIMIT);
	}, 120_000);
});
