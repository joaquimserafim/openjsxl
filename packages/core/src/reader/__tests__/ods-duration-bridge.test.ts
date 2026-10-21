import { loadFixture } from "@openjsxl/fixtures";
import { describe, expect, it } from "vitest";
import { streamXlsx, workbookToInput, writeXlsx } from "../../writer";
import { openOds } from "../ods";
import { openXlsx } from "../workbook";

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

describe("ODS negative durations — XLSX bridge (F11.15)", () => {
	it.each([
		"buffered",
		"streaming",
	])("preserves signed numeric values through %s output", async (mode) => {
		// Read the real LibreOffice producer bytes verbatim. ODS time values are numeric days in
		// the shared model; the bridge must not turn a negative duration into a date or its magnitude.
		const source = await openOds(await loadFixture("libreoffice-negative-durations.ods"));
		const input = await workbookToInput(source);
		const bytes = mode === "buffered" ? await writeXlsx(input) : await drain(streamXlsx(input));
		const rewritten = await openXlsx(bytes);
		expect(rewritten.sheets.map((sheet) => sheet.name)).toEqual(["Durations"]);
		const sheet = rewritten.sheet("Durations");
		const expected = [
			{ ref: "B2", value: -1 / 24 },
			{ ref: "B3", value: -(2 + 3 / 24 + 4 / 1_440 + 5.25 / 86_400) },
			{ ref: "B4", value: -0.125 / 86_400 },
			{ ref: "B5", value: 1 / 24 },
			{ ref: "B6", value: 0 },
			{ ref: "B7", value: -1 },
			{ ref: "B8", value: 2 + 3 / 24 + 4 / 1_440 + 5.25 / 86_400 },
		];
		for (const { ref, value } of expected) {
			const cell = sheet.cell(ref);
			expect(cell.type, ref).toBe("number");
			expect(cell.value, ref).toBeCloseTo(value, 14);
			expect(cell.value, ref).toBe(source.sheet("Durations").cell(ref).value);
		}
		expect(Object.is(sheet.cell("B6").value, -0)).toBe(false);
	});
});
