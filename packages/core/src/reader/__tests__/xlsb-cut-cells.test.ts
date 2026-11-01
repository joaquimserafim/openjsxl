import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { type BiffRecord, readRecords } from "../../biff";
import type { Cell } from "../../types";
import { streamXlsx, workbookToInput, writeXlsx } from "../../writer";
import { writeZip } from "../../writer/zip";
import { openZip } from "../../zip";
import { openXlsx, type Workbook } from "../workbook";
import { openXlsb } from "../xlsb";

// F11.100 at the reader level. SheetJS's short-cell fixture is rebuilt with two cut records on its
// Mixed sheet: J2's BrtShortIsst keeps its complete style, set to 4 ("0.000"), but loses its string
// index, and F3's BrtShortBool keeps one byte, the first byte of style 4. Each used to invent a value
// (J2 the first shared string, F3 `true` from its style byte), which the bridge then wrote.

const FIXTURE = "sheetjs-short-cells-sst.xlsb";
const MIXED = "xl/worksheets/sheet3.bin";
const SHORT_BOOL = 0x0f;
const SHORT_ISST = 0x12;

// A record framed as the parser reads it: the id in one byte or two, then a varint length.
function frame({ id, data }: BiffRecord): number[] {
	const length: number[] = [];
	let remaining = data.length;
	do {
		const byte = remaining & 0x7f;
		remaining >>>= 7;
		length.push(byte | (remaining === 0 ? 0 : 0x80));
	} while (remaining !== 0);
	return [...(id < 0x80 ? [id] : [id & 0xff, id >>> 8]), ...length, ...data];
}

const sameBytes = (data: Uint8Array, expected: readonly number[]) =>
	data.length === expected.length && data.every((byte, i) => byte === expected[i]);

// The fixture with each record that `edit` returns new bytes for replaced.
async function rebuilt(
	edit: (record: BiffRecord, index: number, all: readonly BiffRecord[]) => Uint8Array | undefined,
): Promise<Uint8Array> {
	const zip = openZip(await loadFixture(FIXTURE));
	const records = [...readRecords(await zip.read(MIXED))];
	const sheet = records.flatMap((record, index) =>
		frame({ id: record.id, data: edit(record, index, records) ?? record.data }),
	);
	const parts: { name: string; data: Uint8Array }[] = [];
	for (const name of zip.entries.keys()) {
		parts.push({ name, data: name === MIXED ? Uint8Array.from(sheet) : await zip.read(name) });
	}
	return writeZip(parts);
}

// J2 is the only short shared-string record pointing at string 3; F3 is the sheet's last short Boolean.
const cut = () =>
	rebuilt((record, index, all) => {
		if (record.id === SHORT_ISST && sameBytes(record.data, [0, 0, 0, 0, 3, 0, 0, 0])) {
			return Uint8Array.from([4, 0, 0, 0]);
		}
		const lastBool = all.findLastIndex((r) => r.id === SHORT_BOOL);
		if (index === lastBool && sameBytes(record.data, [0, 0, 0, 0, 0]))
			return Uint8Array.from([4]);
		return undefined;
	});

async function cells(workbook: Workbook, name: string): Promise<Cell[]> {
	const out: Cell[] = [];
	for await (const row of workbook.sheet(name).rows()) out.push(...row.cells);
	return out;
}

const EMPTY = (ref: string): Cell => ({ ref, type: "empty", value: null });
const dropCut = (list: readonly Cell[]) =>
	list.filter((cell) => cell.ref !== "J2" && cell.ref !== "F3");

describe("XLSB cut cell records — rebuilt SheetJS fixture (F11.100)", () => {
	it("rebuilds the producer's sheet byte for byte, then cuts exactly the two records", async () => {
		const zip = openZip(await loadFixture(FIXTURE));
		const original = await zip.read(MIXED);
		const same = openZip(await rebuilt(() => undefined));
		expect(sameBytes(await same.read(MIXED), [...original])).toBe(true);
		const edited = [...readRecords(await openZip(await cut()).read(MIXED))];
		const before = [...readRecords(original)];
		expect(edited).toHaveLength(before.length);
		const changed = edited.filter(
			(record, i) => !sameBytes(record.data, [...(before[i]?.data ?? [])]),
		);
		expect(changed.map((record) => [record.id, [...record.data]])).toEqual([
			[SHORT_ISST, [4, 0, 0, 0]],
			[SHORT_BOOL, [4]],
		]);
	});

	it("reads both cut cells as empty with their record's format, and every other cell as before", async () => {
		const original = await openXlsb(await loadFixture(FIXTURE));
		const workbook = await openXlsb(await cut());
		const mixed = workbook.sheet("Mixed");
		expect(original.sheet("Mixed").cell("J2")?.value).toBe("tail");
		expect(original.sheet("Mixed").cell("F3")?.value).toBe(false);
		expect(mixed.cell("J2")).toEqual(EMPTY("J2"));
		expect(mixed.cell("F3")).toEqual(EMPTY("F3"));
		expect(mixed.numberFormat("J2")).toBe("0.000"); // its complete style, 4
		expect(mixed.numberFormat("F3")).toBe("General"); // its style is cut, so style 0
		expect(await cells(workbook, "Mixed")).toEqual(dropCut(await cells(original, "Mixed")));
		for (const name of ["Integers", "Decimals"]) {
			expect(await cells(workbook, name), name).toEqual(await cells(original, name));
		}
		expect(mixed.dimension).toBe(original.sheet("Mixed").dimension);
	});

	it.each(["buffered", "streaming"])(
		"exports no value for either cut cell through the %s bridge",
		async (mode) => {
			const write = async (book: Workbook) => {
				const input = await workbookToInput(book);
				return mode === "buffered"
					? writeXlsx(input)
					: new Uint8Array(await new Response(streamXlsx(input)).arrayBuffer());
			};
			const original = await openXlsx(
				await write(await openXlsb(await loadFixture(FIXTURE))),
			);
			const bridged = await openXlsx(await write(await openXlsb(await cut())));
			expect(bridged.sheet("Mixed").cell("J2")).toEqual(EMPTY("J2"));
			expect(bridged.sheet("Mixed").cell("F3")).toEqual(EMPTY("F3"));
			expect(await cells(bridged, "Mixed")).toEqual(dropCut(await cells(original, "Mixed")));
		},
	);
});
