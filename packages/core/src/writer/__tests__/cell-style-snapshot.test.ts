import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import type { CellStyle } from "../../types";
import { openZip } from "../../zip";
import { streamXlsx } from "../stream";
import type { CellInput, StreamSheetInput, StyledCell, WorkbookInput } from "../types";
import { writeXlsx } from "../workbook";

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
	const reader = stream.getReader();
	const chunks: Uint8Array[] = [];
	let length = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		chunks.push(value);
		length += value.length;
	}
	const bytes = new Uint8Array(length);
	let offset = 0;
	for (const chunk of chunks) {
		bytes.set(chunk, offset);
		offset += chunk.length;
	}
	return bytes;
}

const writers: readonly {
	name: string;
	write: (workbook: WorkbookInput) => Promise<Uint8Array>;
}[] = [
	{ name: "writeXlsx", write: writeXlsx },
	{ name: "streamXlsx", write: (workbook) => drain(streamXlsx(workbook)) },
];

for (const { name, write } of writers) {
	describe(`${name} — cell property snapshots`, () => {
		it("retains the first style when a getter later returns undefined", async () => {
			let reads = 0;
			const cell: StyledCell = { value: 7 };
			Object.defineProperty(cell, "style", {
				get() {
					reads++;
					return reads === 1 ? { font: { bold: true } } : undefined;
				},
			});

			const workbook = await openXlsx(
				await write({ sheets: [{ name: "S", rows: [[cell]] }] }),
			);

			expect(workbook.sheet("S").style("A1")).toEqual({ font: { bold: true } });
			expect(reads).toBe(1);
		});

		it("does not invoke a style getter that throws on its second read", async () => {
			let reads = 0;
			const cell: StyledCell = {
				value: 7,
				get style() {
					reads++;
					if (reads > 1) throw new Error("style was read twice");
					return { font: { bold: true } };
				},
			};

			const workbook = await openXlsx(
				await write({ sheets: [{ name: "S", rows: [[cell]] }] }),
			);

			expect(workbook.sheet("S").style("A1")).toEqual({ font: { bold: true } });
			expect(reads).toBe(1);
		});

		it("reads value, formula, and style once through a Proxy", async () => {
			const reads = { value: 0, formula: 0, style: 0 };
			const cell = new Proxy(
				{ value: 3, formula: "1+2", style: { font: { bold: true } } },
				{
					get(target, key, receiver) {
						if (key === "value" || key === "formula" || key === "style") {
							reads[key]++;
							if (reads[key] > 1) throw new Error(`${key} was read twice`);
						}
						return Reflect.get(target, key, receiver);
					},
				},
			);

			const workbook = await openXlsx(
				await write({ sheets: [{ name: "S", rows: [[cell]] }] }),
			);
			const sheet = workbook.sheet("S");

			expect(reads).toEqual({ value: 1, formula: 1, style: 1 });
			expect(sheet.cell("A1").value).toBe(3);
			expect(sheet.formula("A1")).toBe("1+2");
			expect(sheet.style("A1")).toEqual({ font: { bold: true } });
		});

		it("keeps an undefined first style without adding a stylesheet", async () => {
			let reads = 0;
			const cell: StyledCell = { value: 7 };
			Object.defineProperty(cell, "style", {
				get() {
					reads++;
					if (reads > 1) throw new Error("style was read twice");
					return undefined;
				},
			});

			const bytes = await write({ sheets: [{ name: "S", rows: [[cell]] }] });

			expect(reads).toBe(1);
			expect(openZip(bytes).has("xl/styles.xml")).toBe(false);
		});

		it.each([
			{ label: "null", style: null },
			{ label: "a string", style: "bold" },
			{ label: "an array", style: [] },
			{ label: "an unknown key", style: { typo: true } },
			{ label: "an invalid font field", style: { font: { bold: "yes" } } },
		])(
			"rejects $label as the first style even if a later style is valid",
			async ({ style }) => {
				let reads = 0;
				const cell: StyledCell = { value: 7 };
				Object.defineProperty(cell, "style", {
					get() {
						reads++;
						return reads === 1 ? style : { font: { bold: true } };
					},
				});

				const result = write({ sheets: [{ name: "S", rows: [[cell]] }] });

				await expect(result).rejects.toBeInstanceOf(XlsxError);
				await expect(result).rejects.toMatchObject({ code: "invalid-input" });
				expect(reads).toBe(1);
			},
		);

		it("preserves a styled blank and an uncached formula cell", async () => {
			const reads = [0, 0];
			const cells: StyledCell[] = [{ value: null }, { formula: "1+2" }];
			for (const [index, cell] of cells.entries()) {
				Object.defineProperty(cell, "style", {
					get() {
						reads[index] = (reads[index] ?? 0) + 1;
						return reads[index] === 1 ? { font: { bold: true } } : undefined;
					},
				});
			}

			const workbook = await openXlsx(
				await write({ sheets: [{ name: "S", rows: [cells] }] }),
			);
			const sheet = workbook.sheet("S");

			expect(reads).toEqual([1, 1]);
			expect(sheet.cell("A1").type).toBe("empty");
			expect(sheet.cell("A1").value).toBeNull();
			expect(sheet.style("A1")).toEqual({ font: { bold: true } });
			expect(sheet.formula("B1")).toBe("1+2");
			expect(sheet.cell("B1").value).toBeNull();
			expect(sheet.style("B1")).toEqual({ font: { bold: true } });
		});

		const dateStyles: readonly {
			label: string;
			style: CellStyle | undefined;
			expected: CellStyle;
		}[] = [
			{
				label: "an absent style",
				style: undefined,
				expected: { numberFormat: "mm-dd-yy" },
			},
			{
				label: "a font-only style",
				style: { font: { bold: true } },
				expected: { numberFormat: "mm-dd-yy", font: { bold: true } },
			},
			{
				label: "an explicit date format",
				style: { numberFormat: "yyyy-mm-dd", font: { bold: true } },
				expected: { numberFormat: "yyyy-mm-dd", font: { bold: true } },
			},
		];
		it.each(dateStyles)("preserves date formatting for $label", async ({ style, expected }) => {
			const date = new Date("2024-02-29T12:00:00Z");
			const reads = { value: 0, style: 0 };
			const cell: StyledCell = {
				get value() {
					reads.value++;
					if (reads.value > 1) throw new Error("value was read twice");
					return date;
				},
			};
			Object.defineProperty(cell, "style", {
				get() {
					reads.style++;
					if (reads.style > 1) throw new Error("style was read twice");
					return style;
				},
			});

			const workbook = await openXlsx(
				await write({ sheets: [{ name: "S", rows: [[cell]] }] }),
			);

			expect(workbook.sheet("S").cell("A1")).toMatchObject({ type: "date", value: date });
			expect(workbook.sheet("S").style("A1")).toEqual(expected);
			expect(reads).toEqual({ value: 1, style: 1 });
		});

		it("snapshots a reused object once per cell coordinate", async () => {
			const reads = { value: 0, style: 0 };
			const cell: StyledCell = {
				get value() {
					return ++reads.value;
				},
				get style() {
					reads.style++;
					return { font: { bold: reads.style % 2 === 1, italic: reads.style % 2 === 0 } };
				},
			};

			const workbook = await openXlsx(
				await write({ sheets: [{ name: "S", rows: [[cell, cell]] }] }),
			);
			const sheet = workbook.sheet("S");

			expect(reads).toEqual({ value: 2, style: 2 });
			expect(sheet.cell("A1").value).toBe(1);
			expect(sheet.cell("B1").value).toBe(2);
			expect(sheet.style("A1")).toEqual({ font: { bold: true } });
			expect(sheet.style("B1")).toEqual({ font: { italic: true } });
		});
	});
}

describe("streamXlsx — lazy cell style snapshots", () => {
	it("captures each async row before advancing a source that reuses its cell", async () => {
		const count = 128;
		let current = 0;
		let started = 0;
		let closed = 0;
		const reads = { value: 0, style: 0 };
		const cell: StyledCell = {
			get value() {
				reads.value++;
				return current;
			},
			get style() {
				reads.style++;
				return current % 2 === 1 ? { font: { bold: true } } : { font: { italic: true } };
			},
		};
		async function* rows(): AsyncGenerator<readonly CellInput[]> {
			started++;
			try {
				for (current = 1; current <= count; current++) {
					yield [cell];
					expect(reads).toEqual({ value: current, style: current });
				}
			} finally {
				closed++;
			}
		}
		let preparations = 0;
		const trailing: StreamSheetInput = {
			name: "Trailing",
			rows: [],
			get columns() {
				preparations++;
				expect(started).toBe(0);
				expect(reads).toEqual({ value: 0, style: 0 });
				return [];
			},
		};

		const workbook = await openXlsx(
			await drain(streamXlsx({ sheets: [{ name: "S", rows: rows() }, trailing] })),
		);
		const sheet = workbook.sheet("S");

		expect(preparations).toBe(1);
		expect(started).toBe(1);
		expect(closed).toBe(1);
		expect(reads).toEqual({ value: count, style: count });
		expect(sheet.cell("A1").value).toBe(1);
		expect(sheet.cell(`A${count}`).value).toBe(count);
		expect(sheet.style("A1")).toEqual({ font: { bold: true } });
		expect(sheet.style(`A${count}`)).toEqual({ font: { italic: true } });
	});

	it("rejects a captured invalid style and closes the source before reading future cells", async () => {
		let styleReads = 0;
		let futureReads = 0;
		let closed = 0;
		const bad: StyledCell = { value: 7 };
		Object.defineProperty(bad, "style", {
			get() {
				styleReads++;
				return styleReads === 1 ? { font: { bold: "invalid" } } : undefined;
			},
		});
		const future: StyledCell = {
			value: 8,
			get style(): never {
				futureReads++;
				throw new Error("future cell should remain untouched");
			},
		};
		async function* rows(): AsyncGenerator<readonly CellInput[]> {
			try {
				yield [bad];
				yield [future];
			} finally {
				closed++;
			}
		}

		const result = drain(streamXlsx({ sheets: [{ name: "S", rows: rows() }] }));

		await expect(result).rejects.toBeInstanceOf(XlsxError);
		await expect(result).rejects.toMatchObject({ code: "invalid-input" });
		expect(styleReads).toBe(1);
		expect(futureReads).toBe(0);
		expect(closed).toBe(1);
	});
});
