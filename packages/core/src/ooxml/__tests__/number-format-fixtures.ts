import { listFixtures, loadFixture } from "@openjsxl/fixtures";
import { RecordData, readRecords } from "../../biff";
import { detectSpreadsheetFormat } from "../../reader/detect";
import { writeXlsx, writeZip } from "../../writer";
import type { SheetInput } from "../../writer/types";
import { FMT } from "../../xlsb/records";
import { tokenize } from "../../xml";
import { openZip } from "../../zip";

// F11.50 fixtures: number-format codes at and past Excel's 255-character limit, styles parts and
// workbooks that carry them, and the 1.0.0 date scan kept as an equivalence oracle.

/** A 255-character date code (the longest Excel accepts). */
export const DATE_255 = `yyyy${"0".repeat(251)}`;
/** A 256-character date code (one past the limit). */
export const DATE_256 = `${DATE_255}0`;
/** A 255-character non-date code. */
export const CODE_255 = `0.${"0".repeat(253)}`;
/** A 256-character non-date code. */
export const CODE_256 = `${CODE_255}0`;

/**
 * The 1.0.0 bracket pass, quadratic on unclosed brackets, kept only as an oracle for short inputs.
 *
 * @example legacyStrip("[Red]0.00") // "0.00"
 */
export function legacyStrip(code: string): string {
	return code.replace(/\[[^\]]*\]/g, "");
}

/**
 * The 1.0.0 date decision, kept only as an oracle for short inputs.
 *
 * @example legacyIsDate("yyyy-mm-dd") // true
 * @example legacyIsDate("[h]") // true (elapsed time)
 * @example legacyIsDate("0.00") // false
 */
export function legacyIsDate(code: string): boolean {
	const withoutLiterals = code.replace(/"[^"]*"|\\.|[_*]./g, "");
	if (/\[(?:h+|m+|s+)\]/i.test(withoutLiterals)) return true;
	return /[dmyhs]/i.test(legacyStrip(withoutLiterals));
}

/**
 * `count` deterministic pseudo-random codes (length 0-31) over the characters the date scan treats
 * specially.
 *
 * @example [...randomCodes(2)].length // 2
 */
export function* randomCodes(count: number): Generator<string> {
	const alphabet = '[]"\\_*hdmsy0a ;';
	let seed = 0x2f6e2b1;
	const next = (): number => {
		seed = (Math.imul(seed, 1103515245) + 12345) >>> 0;
		return seed;
	};
	for (let i = 0; i < count; i++) {
		const length = next() % 32;
		let code = "";
		for (let j = 0; j < length; j++) code += alphabet.charAt(next() % alphabet.length);
		yield code;
	}
}

/**
 * Every custom number-format code in the committed xlsx/xlsm/xlsb corpus. Deliberately malformed
 * archives (detection answers `undefined`) have no style table to sample and are passed over.
 *
 * @example (await corpusFormatCodes()).length > 0 // true
 */
export async function corpusFormatCodes(): Promise<string[]> {
	const codes: string[] = [];
	for (const name of listFixtures()) {
		const bytes = await loadFixture(name);
		const format = await detectSpreadsheetFormat(bytes);
		if (format !== "xlsx" && format !== "xlsb") continue;
		const isBinary = format === "xlsb";
		const zip = openZip(bytes);
		const suffix = isBinary ? "styles.bin" : "styles.xml";
		const part = [...zip.entries.keys()].find((entry) => entry.endsWith(suffix));
		if (part === undefined) continue;
		const styles = await zip.read(part);
		if (isBinary) {
			for (const record of readRecords(styles)) {
				if (record.id !== FMT) continue;
				const data = new RecordData(record.data);
				data.u16();
				const code = data.wideString();
				if (code !== undefined) codes.push(code);
			}
			continue;
		}
		for (const token of tokenize(new TextDecoder().decode(styles))) {
			if (token.kind !== "open" || !token.name.endsWith("numFmt")) continue;
			const code = token.attrs.formatCode;
			if (code !== undefined) codes.push(code);
		}
	}
	return codes;
}

/**
 * A styles part with the given custom formats ([numFmtId, code], codes must not need escaping) and
 * one cellXfs entry per id in `xfFormatIds`.
 *
 * @example stylesXml([[164, "0.0"]], [0, 164]) // xf 1 uses custom format 164
 */
export function stylesXml(
	formats: readonly (readonly [number, string])[],
	xfFormatIds: readonly number[],
): string {
	const numFmts = formats.map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${code}"/>`);
	const xfs = xfFormatIds.map(
		(id) => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0"/>`,
	);
	return (
		'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
		`<numFmts count="${formats.length}">${numFmts.join("")}</numFmts>` +
		'<fonts count="1"><font/></fonts>' +
		'<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
		'<borders count="1"><border/></borders>' +
		`<cellXfs count="${xfFormatIds.length}">${xfs.join("")}</cellXfs>` +
		"</styleSheet>"
	);
}

/**
 * A one-cell workbook (`S!A1` = 45000) styled with `DATE_255`, whose styles.xml code is then
 * replaced by `code` — the way to get a code the writer would refuse into a readable file.
 *
 * @example await workbookWithFormat(DATE_256) // xlsx bytes whose A1 format is 256 characters
 */
export async function workbookWithFormat(code: string): Promise<Uint8Array> {
	const written = await writeXlsx({
		sheets: [{ name: "S", rows: [[{ value: 45000, style: { numberFormat: DATE_255 } }]] }],
	});
	const zip = openZip(written);
	const parts: { name: string; data: Uint8Array }[] = [];
	for (const name of zip.entries.keys()) {
		const data = await zip.read(name);
		if (name !== "xl/styles.xml") {
			parts.push({ name, data });
			continue;
		}
		const xml = new TextDecoder().decode(data);
		if (!xml.includes(DATE_255)) throw new Error("styles.xml does not carry the seed code");
		parts.push({ name, data: new TextEncoder().encode(xml.replace(DATE_255, code)) });
	}
	return writeZip(parts);
}

/**
 * A sheet whose A1 carries `numberFormat` as a cell style.
 *
 * @example styledSheet("0.00").rows // [[{ value: 1.5, style: { numberFormat: "0.00" } }]]
 */
export function styledSheet(numberFormat: string): SheetInput {
	return { name: "S", rows: [[{ value: 1.5, style: { numberFormat } }]] };
}

/**
 * A sheet with one conditional-formatting rule whose dxf carries `numberFormat`.
 *
 * @example cfSheet("0.0%").conditionalFormatting?.[0]?.rules[0] // a cellIs rule with that dxf code
 */
export function cfSheet(numberFormat: string): SheetInput {
	return {
		name: "S",
		rows: [[1], [2]],
		conditionalFormatting: [
			{
				sqref: ["A1:A2"],
				rules: [
					{
						type: "cellIs",
						priority: 1,
						operator: "greaterThan",
						formulas: ["1"],
						dxf: { numberFormat },
					},
				],
			},
		],
	};
}
