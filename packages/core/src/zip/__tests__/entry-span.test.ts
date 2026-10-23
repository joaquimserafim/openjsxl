import { describe, expect, it } from "vitest";
import { writeZip } from "../../writer/zip";
import { type EntrySpan, findSpanConflict, localEntrySpan, openZip } from "../central-directory";

const span = (name: string, start: number, end: number): EntrySpan => ({
	name,
	start,
	dataStart: start + 30,
	end,
});

describe("localEntrySpan", () => {
	it("reads the span from the local header", async () => {
		const zip = await writeZip([{ name: "a", data: new TextEncoder().encode("first") }]);
		const entry = openZip(zip).entries.get("a");
		if (entry === undefined) throw new Error("expected entry a");
		expect(localEntrySpan(new DataView(zip.buffer), entry)).toEqual({
			name: "a",
			start: 0,
			dataStart: 31,
			end: 31 + entry.compressedSize,
		});
	});

	it("has no span without a local header at the entry's offset", async () => {
		const zip = await writeZip([{ name: "a", data: new TextEncoder().encode("first") }]);
		const entry = openZip(zip).entries.get("a");
		if (entry === undefined) throw new Error("expected entry a");
		const view = new DataView(zip.buffer);
		expect(localEntrySpan(view, { ...entry, localHeaderOffset: 1 })).toBeUndefined();
		expect(
			localEntrySpan(view, { ...entry, localHeaderOffset: zip.byteLength - 29 }),
		).toBeUndefined();
	});
});

describe("findSpanConflict", () => {
	it("accepts disjoint spans in any order, touching each other and the central directory", () => {
		expect(findSpanConflict([], 0)).toBeUndefined();
		expect(findSpanConflict([span("b", 40, 100), span("a", 0, 40)], 100)).toBeUndefined();
		expect(findSpanConflict([span("a", 0, 40), span("b", 60, 90)], 100)).toBeUndefined();
	});

	it("reports two spans that share bytes, in local-header order", () => {
		expect(findSpanConflict([span("b", 0, 40), span("a", 0, 40)], 100)).toMatchObject({
			kind: "shared-bytes",
		});
		expect(findSpanConflict([span("b", 40, 80), span("a", 0, 41)], 100)).toEqual({
			kind: "shared-bytes",
			first: "a",
			second: "b",
		});
		expect(findSpanConflict([span("outer", 0, 90), span("inner", 40, 50)], 100)).toEqual({
			kind: "shared-bytes",
			first: "outer",
			second: "inner",
		});
	});

	it("reports a span that reaches past the central directory's offset", () => {
		expect(findSpanConflict([span("a", 0, 40), span("b", 40, 101)], 100)).toEqual({
			kind: "into-central-directory",
			name: "b",
		});
		expect(findSpanConflict([span("a", 120, 160)], 100)).toEqual({
			kind: "into-central-directory",
			name: "a",
		});
	});
});
