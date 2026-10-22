import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { openXlsx } from "../../reader/workbook";
import { streamXlsx } from "../stream";
import type { StreamRows, StreamSheetInput, StreamWorkbookInput } from "../types";

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

function sheetWithRows(getRows: () => unknown): StreamSheetInput {
	const sheet: StreamSheetInput = { name: "S", rows: [] };
	Object.defineProperty(sheet, "rows", { get: getRows });
	return sheet;
}

describe("streamXlsx — sheet.rows snapshot", () => {
	it("emits the first iterable when the rows getter changes", async () => {
		let reads = 0;
		const sheet: StreamSheetInput = {
			name: "S",
			get rows() {
				reads++;
				return reads === 1 ? [[42]] : [[99]];
			},
		};

		const workbook = await openXlsx(await drain(streamXlsx({ sheets: [sheet] })));

		expect(workbook.sheet("S").cell("A1")?.value).toBe(42);
		expect(reads).toBe(1);
	});

	it.each(["null", "throw"])("never reads a second %s rows value", async (second) => {
		let reads = 0;
		const sheet = sheetWithRows(() => {
			reads++;
			if (reads === 1) return [[42]];
			if (second === "null") return null;
			throw new Error("rows was read twice");
		});

		const workbook = await openXlsx(await drain(streamXlsx({ sheets: [sheet] })));

		expect(workbook.sheet("S").cell("A1")?.value).toBe(42);
		expect(reads).toBe(1);
	});

	it("performs exactly one Proxy rows get per sheet", async () => {
		const reads = [0, 0];
		const sheets = [42, 84].map(
			(value, index) =>
				new Proxy(
					{ name: `S${index}`, rows: [[value]] },
					{
						get(target, key, receiver) {
							if (key === "rows") reads[index] = (reads[index] ?? 0) + 1;
							return Reflect.get(target, key, receiver);
						},
					},
				),
		);

		const workbook = await openXlsx(await drain(streamXlsx({ sheets })));

		expect(reads).toEqual([1, 1]);
		expect(workbook.sheet("S0").cell("A1")?.value).toBe(42);
		expect(workbook.sheet("S1").cell("A1")?.value).toBe(84);
	});

	it.each([null, undefined, 42, "rows", {}])(
		"rejects first rows value %j with a typed error without a second read",
		async (first) => {
			let reads = 0;
			const sheet = sheetWithRows(() => {
				reads++;
				return reads === 1 ? first : [[42]];
			});
			const result = drain(streamXlsx({ sheets: [sheet] }));

			await expect(result).rejects.toBeInstanceOf(XlsxError);
			await expect(result).rejects.toMatchObject({ code: "invalid-input" });
			expect(reads).toBe(1);
		},
	);

	for (const mode of ["sync", "async"]) {
		it(`keeps a one-shot ${mode} generator untouched during preparation`, async () => {
			let started = 0;
			let yielded = 0;
			let closed = 0;
			let reads = 0;
			function* syncRows(): Generator<readonly number[] | undefined> {
				started++;
				try {
					for (const row of [[42], undefined, [84]]) {
						yielded++;
						yield row;
					}
				} finally {
					closed++;
				}
			}
			async function* asyncRows(): AsyncGenerator<readonly number[] | undefined> {
				yield* syncRows();
			}
			const source = mode === "sync" ? syncRows() : asyncRows();
			const sheet = sheetWithRows(() => {
				reads++;
				return source;
			});
			let preparations = 0;
			const trailing: StreamSheetInput = {
				name: "Trailing",
				rows: [],
				get columns() {
					preparations++;
					expect(started).toBe(0);
					expect(yielded).toBe(0);
					return [];
				},
			};

			const workbook = await openXlsx(await drain(streamXlsx({ sheets: [sheet, trailing] })));

			expect(preparations).toBe(1);
			expect(reads).toBe(1);
			expect(started).toBe(1);
			expect(yielded).toBe(3);
			expect(closed).toBe(1);
			expect(workbook.sheet("S").cell("A1")?.value).toBe(42);
			expect(workbook.sheet("S").cell("A2")?.value).toBeNull();
			expect(workbook.sheet("S").cell("A3")?.value).toBe(84);
		});

		it(`does not start the ${mode} source if later sheet validation fails`, async () => {
			let started = false;
			let closed = false;
			function* syncRows(): Generator<readonly number[]> {
				started = true;
				try {
					yield [42];
				} finally {
					closed = true;
				}
			}
			async function* asyncRows(): AsyncGenerator<readonly number[]> {
				yield* syncRows();
			}
			const source = mode === "sync" ? syncRows() : asyncRows();
			const result = drain(
				streamXlsx({
					sheets: [
						{ name: "S", rows: source },
						{ name: "Bad", rows: [], columns: [{ min: 1, max: 1, width: -1 }] },
					],
				}),
			);

			await expect(result).rejects.toMatchObject({ code: "invalid-input" });
			expect(started).toBe(false);
			expect(closed).toBe(false);
		});

		it(`preserves ${mode} producer errors and closes the captured source`, async () => {
			const failure = new Error("row source failed");
			let closed = 0;
			let reads = 0;
			function* syncRows(): Generator<readonly number[]> {
				try {
					yield [42];
					throw failure;
				} finally {
					closed++;
				}
			}
			async function* asyncRows(): AsyncGenerator<readonly number[]> {
				yield* syncRows();
			}
			const source = mode === "sync" ? syncRows() : asyncRows();
			const sheet = sheetWithRows(() => {
				reads++;
				return source;
			});

			await expect(drain(streamXlsx({ sheets: [sheet] }))).rejects.toBe(failure);
			expect(closed).toBe(1);
			expect(reads).toBe(1);
		});

		it(`rejects malformed ${mode} yielded cells typed and closes the source`, async () => {
			let closed = 0;
			let advancedPastInvalid = false;
			function* syncRows(): Generator<readonly number[]> {
				try {
					yield [42];
					yield [Number.NaN];
					advancedPastInvalid = true;
					yield [84];
				} finally {
					closed++;
				}
			}
			async function* asyncRows(): AsyncGenerator<readonly number[]> {
				yield* syncRows();
			}
			const source = mode === "sync" ? syncRows() : asyncRows();
			const result = drain(streamXlsx({ sheets: [{ name: "S", rows: source }] }));

			await expect(result).rejects.toBeInstanceOf(XlsxError);
			await expect(result).rejects.toMatchObject({ code: "invalid-input" });
			expect(advancedPastInvalid).toBe(false);
			expect(closed).toBe(1);
		});

		it(`closes the captured ${mode} source on cancellation`, async () => {
			let yielded = 0;
			let closed = 0;
			let reads = 0;
			function* syncRows(): Generator<readonly number[]> {
				try {
					for (let i = 1; ; i++) {
						yielded++;
						yield [i];
					}
				} finally {
					closed++;
				}
			}
			async function* asyncRows(): AsyncGenerator<readonly number[]> {
				yield* syncRows();
			}
			const source = mode === "sync" ? syncRows() : asyncRows();
			const sheet = sheetWithRows(() => {
				reads++;
				return source;
			});
			const reader = streamXlsx({ sheets: [sheet] }).getReader();
			try {
				while (yielded === 0) {
					const { done } = await reader.read();
					expect(done).toBe(false);
				}
			} finally {
				await reader.cancel();
			}

			expect(closed).toBe(1);
			expect(reads).toBe(1);
		});
	}

	it("still completes cancellation when source cleanup throws", async () => {
		let yielded = 0;
		let closed = 0;
		function closeCursor(): never {
			closed++;
			throw new Error("cursor cleanup failed");
		}
		async function* rows(): AsyncGenerator<readonly number[]> {
			try {
				for (let i = 1; ; i++) {
					yielded++;
					yield [i];
				}
			} finally {
				closeCursor();
			}
		}
		const reader = streamXlsx({ sheets: [{ name: "S", rows: rows() }] }).getReader();
		try {
			while (yielded === 0) {
				const { done } = await reader.read();
				expect(done).toBe(false);
			}
		} finally {
			await expect(reader.cancel()).resolves.toBeUndefined();
		}
		expect(closed).toBe(1);
	});

	it("retains identical stream bytes for ordinary rows and getter-backed equivalents", async () => {
		const rows: readonly (readonly number[] | undefined)[] = [[42, 7], undefined, [84]];
		const input = (source: StreamRows): StreamWorkbookInput => ({
			date1904: true,
			sheets: [
				{
					name: "S",
					rows: source,
					columns: [{ min: 1, max: 1, width: 18 }],
					freeze: { rows: 1 },
					comments: [{ ref: "A1", author: "A", text: "note" }],
				},
				{ name: "Empty", rows: [] },
			],
		});
		const control = await drain(streamXlsx(input(rows)));
		function* syncRows(): Generator<readonly number[] | undefined> {
			yield* rows;
		}
		async function* asyncRows(): AsyncGenerator<readonly number[] | undefined> {
			yield* rows;
		}

		for (const source of [rows, syncRows(), asyncRows()]) {
			const workbook = input(source);
			const sheet = workbook.sheets[0];
			expect(sheet).toBeDefined();
			if (sheet === undefined) throw new Error("missing test sheet");
			let reads = 0;
			Object.defineProperty(sheet, "rows", {
				get() {
					reads++;
					return source;
				},
			});

			expect(await drain(streamXlsx(workbook))).toEqual(control);
			expect(reads).toBe(1);
		}
	});
});
