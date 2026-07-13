import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import {
	MAX_DV_TEXT_LEN,
	MAX_DV_TITLE_LEN,
	MAX_SQREF_RANGES,
	parseDataValidations,
} from "../../ooxml/data-validation";
import { openXlsx } from "../../reader/workbook";
import type { DataValidation } from "../../types";
import { openZip } from "../../zip";
import { streamXlsx } from "../stream";
import type { SheetInput } from "../types";
import { writeXlsx } from "../workbook";

const decoder = new TextDecoder();

// The worksheet part text (holds the inline <dataValidations> block).
const sheetXml = async (sheet: SheetInput): Promise<string> =>
	decoder.decode(
		await openZip(await writeXlsx({ sheets: [sheet] })).read("xl/worksheets/sheet1.xml"),
	);

async function roundTrip(dvs: readonly DataValidation[]) {
	const sheet: SheetInput = { name: "S", rows: [["h"]], dataValidations: dvs };
	const book = await openXlsx(await writeXlsx({ sheets: [sheet] }));
	return book.sheet("S").dataValidations;
}

const rejects = (dvs: readonly unknown[]) =>
	expect(
		writeXlsx({
			// biome-ignore lint/suspicious/noExplicitAny: exercising hostile JS callers past the types
			sheets: [{ name: "S", rows: [["h"]], dataValidations: dvs as any }],
		}),
	).rejects.toThrow(XlsxError);

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
	const reader = stream.getReader();
	const chunks: Uint8Array[] = [];
	let total = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		chunks.push(value);
		total += value.length;
	}
	const out = new Uint8Array(total);
	let at = 0;
	for (const c of chunks) {
		out.set(c, at);
		at += c.length;
	}
	return out;
}

describe("writeXlsx — data validation round-trip", () => {
	it("round-trips every type + operator + multi-range sqref deep-equal", async () => {
		const dvs: DataValidation[] = [
			{
				sqref: ["A2:A10"],
				type: "whole",
				operator: "between",
				formula1: "1",
				formula2: "10",
			},
			{
				sqref: ["B2:B10", "D2:D10"],
				type: "decimal",
				operator: "greaterThanOrEqual",
				formula1: "0",
			},
			{ sqref: ["C2:C10"], type: "list", formula1: '"a,b,c"' },
			{ sqref: ["E2:E10"], type: "list", formula1: "Lists!$A$1:$A$3", showDropDown: false },
			{
				sqref: ["F2:F10"],
				type: "date",
				operator: "notBetween",
				formula1: "1",
				formula2: "2",
			},
			{ sqref: ["G2:G10"], type: "time", operator: "greaterThan", formula1: "0.5" },
			{ sqref: ["H2:H10"], type: "textLength", operator: "lessThanOrEqual", formula1: "20" },
			{ sqref: ["I2:I10"], type: "custom", formula1: "ISNUMBER(I2)" },
			{
				sqref: ["J2:J10"],
				type: "none",
				showInputMessage: true,
				promptTitle: "Note",
				prompt: "free-form",
			},
		];
		expect(await roundTrip(dvs)).toEqual(dvs);
	});

	it("carries prompts, errors, errorStyle, and allowBlank", async () => {
		const dv: DataValidation = {
			sqref: ["A1"],
			type: "whole",
			operator: "between",
			formula1: "1",
			formula2: "9",
			allowBlank: true,
			showInputMessage: true,
			showErrorMessage: true,
			errorStyle: "warning",
			promptTitle: "Q",
			prompt: "enter",
			errorTitle: "Bad",
			error: "no",
		};
		expect((await roundTrip([dv]))[0]).toEqual(dv);
	});

	it('emits showDropDown INVERTED: intuitive true → "0", false → "1"', async () => {
		const xml = await sheetXml({
			name: "S",
			rows: [["h"]],
			dataValidations: [
				{ sqref: ["A1"], type: "list", formula1: '"x"', showDropDown: true },
				{ sqref: ["B1"], type: "list", formula1: '"y"', showDropDown: false },
			],
		});
		expect(xml).toContain('showDropDown="0" sqref="A1"'); // intuitive true → file "0"
		expect(xml).toContain('showDropDown="1" sqref="B1"'); // intuitive false → file "1"
	});

	it("strips a leading = from formula operands (stored form)", async () => {
		const [dv] = await roundTrip([{ sqref: ["A1"], type: "custom", formula1: "=ISBLANK(A1)" }]);
		expect(dv?.formula1).toBe("ISBLANK(A1)");
	});

	it("does not emit <dataValidations> when the sheet has none (byte-identity path)", async () => {
		expect(await sheetXml({ name: "S", rows: [["a", 1]] })).not.toContain("dataValidation");
		// An explicitly empty array is also invisible.
		expect(await sheetXml({ name: "S", rows: [["a", 1]], dataValidations: [] })).not.toContain(
			"dataValidation",
		);
	});

	it("orders <dataValidations> between </sheetData> and <hyperlinks> (schema order)", async () => {
		const xml = await sheetXml({
			name: "S",
			rows: [["h"]],
			dataValidations: [{ sqref: ["A1"], type: "custom", formula1: "TRUE" }],
			hyperlinks: [{ ref: "A1", location: "S!A1" }],
		});
		const dv = xml.indexOf("<dataValidations");
		const hl = xml.indexOf("<hyperlinks");
		const sd = xml.indexOf("</sheetData>");
		expect(sd).toBeLessThan(dv);
		expect(dv).toBeLessThan(hl);
	});
});

describe("writeXlsx — data validation rejects malformed input (typed)", () => {
	const base = (over: Record<string, unknown>): DataValidation[] => [
		{ sqref: ["A1"], type: "whole", ...over } as DataValidation,
	];

	it("rejects an unknown property", async () => {
		await rejects(base({ bogus: 1 }));
	});
	it("rejects an unknown type / operator / errorStyle", async () => {
		await rejects(base({ type: "listy" }));
		await rejects(base({ operator: "approx" }));
		await rejects(base({ errorStyle: "shout" }));
	});
	it("rejects an empty or non-array sqref, and a non-canonical range", async () => {
		await rejects(base({ sqref: [] }));
		await rejects(base({ sqref: "A1" }));
		await rejects(base({ sqref: ["A1", 5] }));
		await rejects(base({ sqref: ["$A$1"] })); // sqref refs are relative — no $
		await rejects(base({ sqref: ["B2:A1"] })); // reversed range
		await rejects(base({ sqref: ["ZZZZ9"] })); // outside the grid
	});
	it("rejects a too-long prompt/error title (>32) or body (>255)", async () => {
		await rejects(base({ promptTitle: "x".repeat(MAX_DV_TITLE_LEN + 1) }));
		await rejects(base({ errorTitle: "x".repeat(MAX_DV_TITLE_LEN + 1) }));
		await rejects(base({ prompt: "x".repeat(MAX_DV_TEXT_LEN + 1) }));
		await rejects(base({ error: "x".repeat(MAX_DV_TEXT_LEN + 1) }));
	});
	it("rejects an over-long inline list literal for a list type", async () => {
		await rejects([
			{ sqref: ["A1"], type: "list", formula1: `"${"a,".repeat(MAX_DV_TEXT_LEN)}"` },
		]);
	});
	it("rejects a sqref with more ranges than the cap", async () => {
		const many = Array.from({ length: MAX_SQREF_RANGES + 1 }, (_, i) => `A${(i % 1000) + 1}`);
		await rejects([{ sqref: many, type: "whole" }]);
	});
	it("rejects non-boolean flags and non-string formulas/text", async () => {
		await rejects(base({ allowBlank: 1 }));
		await rejects(base({ showDropDown: "yes" }));
		await rejects(base({ formula1: 5 }));
		await rejects(base({ prompt: 42 }));
	});
	it("rejects a control character in prompt text or a formula", async () => {
		await rejects(base({ prompt: "a\u0000b" }));
		await rejects(base({ type: "custom", formula1: "A1\u0007" }));
	});
});

describe("streamXlsx — data validation lands on the streaming writer too", () => {
	it("emits the same rules through the streamed footer", async () => {
		const bytes = await drain(
			streamXlsx({
				sheets: [
					{
						name: "S",
						rows: [["h"], ["x"]],
						dataValidations: [
							{
								sqref: ["A2:A10"],
								type: "list",
								formula1: '"a,b"',
								showDropDown: false,
							},
						],
					},
				],
			}),
		);
		const ws = (await openXlsx(bytes)).sheet("S");
		expect(ws.dataValidations).toEqual([
			{ sqref: ["A2:A10"], type: "list", formula1: '"a,b"', showDropDown: false },
		]);
		expect(ws.cell("A2").value).toBe("x"); // body streamed correctly around the footer block
	});

	it("rejects malformed validations on the streaming writer with the same typed error", async () => {
		expect(
			drain(
				streamXlsx({
					// biome-ignore lint/suspicious/noExplicitAny: hostile input past the types
					sheets: [{ name: "S", rows: [["h"]], dataValidations: [{ sqref: [] }] as any }],
				}),
			),
		).rejects.toThrow(XlsxError);
	});
});

// F9.2 adversarial-review regression (the core invariant the review found broken): whatever the
// tolerant reader RETURNS, the strict writer must ACCEPT. A hostile worksheet — non-canonical sqref,
// control-char prompt/formula, leading-`=` formula — is parsed to the reader's DEGRADED model, and
// that model is fed straight back to writeXlsx. It must NOT throw, and it must round-trip deep-equal.
describe("data validation — reader output is always writer-acceptable (bridge never aborts)", () => {
	const wsWith = (body: string): string =>
		`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><dataValidations>${body}</dataValidations></worksheet>`;

	it("re-writes a workbook whose file used forms the writer would reject verbatim", async () => {
		const model = parseDataValidations(
			wsWith(
				// A:A / $A$1 / A1:A2000000 dropped, B2:C3 kept; a control-char promptTitle dropped;
				// a leading-`=` custom formula normalized to stored form.
				'<dataValidation type="list" sqref="A:A B2:C3 $A$1 A1:A2000000" promptTitle="bad&#1;title" prompt="ok">' +
					'<formula1>"Red,Green"</formula1></dataValidation>' +
					'<dataValidation type="custom" sqref="D1"><formula1>=ISBLANK(D1)</formula1></dataValidation>',
			),
		);
		// The reader already degraded everything into the writable set.
		expect(model[0]?.sqref).toEqual(["B2:C3"]);
		expect(model[0]?.promptTitle).toBeUndefined();
		expect(model[1]?.formula1).toBe("ISBLANK(D1)");

		// Feeding that model straight back to the writer must succeed AND round-trip deep-equal.
		const book = await openXlsx(
			await writeXlsx({ sheets: [{ name: "S", rows: [["h"]], dataValidations: model }] }),
		);
		expect(book.sheet("S").dataValidations).toEqual(model);
	});
});
