import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { escapeAttr } from "../../writer/xml";
import { writeZip } from "../../writer/zip";
import { openOds } from "../ods";

const encoder = new TextEncoder();

async function readDurations(values: readonly (string | undefined)[]) {
	const rows = values
		.map((value) => {
			const attr = value === undefined ? "" : ` office:time-value="${escapeAttr(value)}"`;
			return `<table:table-row><table:table-cell office:value-type="time"${attr}><text:p>display text</text:p></table:table-cell></table:table-row>`;
		})
		.join("");
	const content =
		'<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
		'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
		'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">' +
		`<office:body><office:spreadsheet><table:table table:name="Durations">${rows}</table:table>` +
		"</office:spreadsheet></office:body></office:document-content>";
	const bytes = await writeZip([
		{
			name: "mimetype",
			data: encoder.encode("application/vnd.oasis.opendocument.spreadsheet"),
		},
		{ name: "content.xml", data: encoder.encode(content) },
	]);
	return (await openOds(bytes)).sheet("Durations");
}

describe("openOds — signed time durations", () => {
	it("preserves the minus sign of -PT1H", async () => {
		const sheet = await readDurations(["-PT1H"]);

		expect(sheet.cell("A1")).toEqual({ ref: "A1", type: "number", value: -1 / 24 });
	});

	it.each([
		{ duration: "PT1H", days: 1 / 24 },
		{ duration: "PT30M", days: 1 / 48 },
		{ duration: "PT45S", days: 45 / 86_400 },
		{ duration: "PT0.125S", days: 0.125 / 86_400 },
		{ duration: "PT1H30M", days: 1 / 16 },
		{ duration: "P2DT3H4M5.25S", days: 2 + 11_045.25 / 86_400 },
		{ duration: "P10DT0S", days: 10 },
		{ duration: "PT49H", days: 49 / 24 },
		{ duration: "PT90M", days: 1 / 16 },
		{ duration: "PT120.5S", days: 120.5 / 86_400 },
	])(
		"applies the sign to all of $duration and leaves the positive value unchanged",
		async ({ duration, days }) => {
			const sheet = await readDurations([`-${duration}`, duration]);

			expect(sheet.cell("A1").type).toBe("number");
			expect(sheet.cell("A1").value).toBeCloseTo(-days, 12);
			expect(sheet.cell("A2").type).toBe("number");
			expect(sheet.cell("A2").value).toBeCloseTo(days, 12);
			expect(sheet.dimension).toBe("A1:A2");
			const values = [];
			for await (const row of sheet.rows()) values.push(row.cells[0]?.value);
			expect(values).toEqual([sheet.cell("A1").value, sheet.cell("A2").value]);
		},
	);

	it("normalizes positive and negative zero durations to ordinary zero", async () => {
		const values = [
			"PT0H",
			"PT0M",
			"PT0S",
			"PT0.000S",
			"P0DT0S",
			"P0DT0H0M0.0S",
			"PT000.000S",
			"P00DT00H00M00S",
			`PT0.${"0".repeat(323)}1S`,
			`PT0.${"0".repeat(319)}1S`,
		].flatMap((value) => [value, `-${value}`]);
		const sheet = await readDurations(values);

		for (let row = 1; row <= values.length; row++) {
			const cell = sheet.cell(`A${row}`);
			expect(cell.type).toBe("number");
			expect(cell.value).toBe(0);
			expect(Object.is(cell.value, -0)).toBe(false);
		}
	});

	it("keeps the existing nonzero-day form without time components", async () => {
		// The existing parser accepts PnDT for nonzero n, but not P0DT. This task only fixes signs.
		const sheet = await readDurations(["P2DT", "-P2DT"]);

		expect(sheet.cell("A1").value).toBe(2);
		expect(sheet.cell("A2").value).toBe(-2);
	});

	it("drops malformed values without using their display text as a fallback", async () => {
		const values = [
			undefined,
			"",
			"PT",
			"-PT",
			"P0DT",
			"-P0DT",
			"PT1",
			"PT1Hgarbage",
			"PT1H2H",
			"PT1M1H",
			"PT-1H",
			"PT+1H",
			"--PT1H",
			"+PT1H",
			"pt1h",
			"PT.5S",
			"PT1.S",
			"PT1,5S",
			"PT1e2S",
			"PTNaNS",
			"PTInfinityS",
			" PT1H",
			"PT1H ",
		];
		const sheet = await readDurations(values);

		for (let row = 1; row <= values.length; row++) {
			expect(sheet.cell(`A${row}`)).toEqual({ ref: `A${row}`, type: "empty", value: null });
		}
		expect(sheet.dimension).toBeUndefined();
		const rows = [];
		for await (const row of sheet.rows()) rows.push(row);
		expect(rows).toEqual([]);
	});

	it("does not add support for duration forms outside the existing grammar", async () => {
		// Day-only, year/month, week, and fractional-hour/day forms remain outside this parser.
		const forms = ["P1D", "P1Y", "P1M", "P1W", "P1.5DT0S", "PT1.5H", "PT1.5M"];
		const values = forms.flatMap((value) => [value, `-${value}`]);
		const sheet = await readDurations(values);

		for (let row = 1; row <= values.length; row++) {
			expect(sheet.cell(`A${row}`).type).toBe("empty");
		}
		expect(sheet.dimension).toBeUndefined();
	});

	it("keeps the finite-value guard for either sign and every component", async () => {
		const huge = "9".repeat(400);
		const values = [`P${huge}DT0S`, `PT${huge}H`, `PT${huge}M`, `PT${huge}S`].flatMap(
			(value) => [value, `-${value}`],
		);
		const sheet = await readDurations(values);

		for (let row = 1; row <= values.length; row++) {
			expect(sheet.cell(`A${row}`).type).toBe("empty");
		}
		expect(sheet.dimension).toBeUndefined();
	});

	it("retains large finite days while rejecting a long malformed component", async () => {
		const largeDays = `1${"0".repeat(307)}`;
		const sheet = await readDurations([
			`P${largeDays}DT0S`,
			`-P${largeDays}DT0S`,
			`-PT${"9".repeat(4_096)}X`,
		]);

		expect(sheet.cell("A1").value).toBe(1e307);
		expect(sheet.cell("A2").value).toBe(-1e307);
		expect(sheet.cell("A3").type).toBe("empty");
		expect(sheet.dimension).toBe("A1:A2");
	});

	it("reads signed duration values from the real LibreOffice fixture verbatim", async () => {
		const sheet = (
			await openOds(await loadFixture("libreoffice-negative-durations.ods"))
		).sheet("Durations");
		// LibreOffice writes elapsed days as hours: the compound is ±PT51H04M05.25S.
		const values = [
			-1 / 24,
			-(2 + 11_045.25 / 86_400),
			-0.125 / 86_400,
			1 / 24,
			0,
			-1,
			2 + 11_045.25 / 86_400,
		];

		for (const [index, value] of values.entries()) {
			const cell = sheet.cell(`B${index + 2}`);
			expect(cell.type).toBe("number");
			expect(cell.value).toBeCloseTo(value, 12);
		}
		expect(Object.is(sheet.cell("B6").value, -0)).toBe(false);
		expect(sheet.dimension).toBe("A1:B8");
	});
});
