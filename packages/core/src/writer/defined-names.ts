import { XlsxError } from "../errors";
import { type CellRef, indexToColumn, parseCanonicalRange } from "../ooxml/a1";
import { MAX_FORMULA_LEN } from "../ooxml/formula";
import {
	type DefinedNameProblem,
	definedNameProblem,
	FILTER_DATABASE_NAME,
	isFilterDatabaseName,
	MAX_NAME_LEN,
} from "../ooxml/name";
import type { DefinedName } from "../ooxml/workbook";
import { encodeXstring } from "../ooxml/xstring";
import { escapeAttr, escapeText, isPlainRecord, isXmlSafe } from "./xml";

/**
 * Serialize trusted defined names for workbook.xml.
 * Input `[]` produces `""`; logical `_x0041_` is protected as wire text `_x005F_x0041_`.
 */
export function definedNamesXml(names: readonly DefinedName[]): string {
	if (names.length === 0) return "";
	const items = names
		.map((dn) => {
			// CT_DefinedName attribute order: name, …, localSheetId, hidden. Both @name and the
			// complete logical formula use ST_Xstring before their XML context is escaped.
			const scope = dn.localSheetId !== undefined ? ` localSheetId="${dn.localSheetId}"` : "";
			const hidden = dn.hidden ? ' hidden="1"' : "";
			return `<definedName name="${escapeAttr(encodeXstring(dn.name))}"${scope}${hidden}>${escapeText(encodeXstring(dn.refersTo))}</definedName>`;
		})
		.join("");
	return `<definedNames>${items}</definedNames>`;
}

/** Map an invalid-name code to its writer message. Input `"cell-ref"` explains the A1 collision. */
export function definedNameProblemMessage(name: string, problem: DefinedNameProblem): string {
	switch (problem) {
		case "empty":
			return "must not be empty";
		case "too-long":
			return `"${name}" exceeds ${MAX_NAME_LEN} characters`;
		case "not-xml-safe":
			return `"${name}" contains a character not allowed in XML`;
		case "whitespace":
			return `"${name}" must not contain whitespace`;
		case "bad-start":
			return `"${name}" must start with a letter, underscore, or backslash`;
		case "cell-ref":
			return `"${name}" must not look like a cell reference`;
		case "bad-builtin":
			return `"${name}" uses the reserved "_xlnm." prefix but is not a spec built-in name`;
	}
}

/**
 * Validate caller input into the trusted logical objects emitted by {@link definedNamesXml}.
 * Input `undefined` produces `[]`; invalid shapes and values throw a typed `invalid-input` error.
 * Each property is read once so a getter cannot diverge validation from later emission.
 */
export function validateDefinedNames(raw: unknown, sheetCount: number): DefinedName[] {
	if (raw === undefined) return [];
	if (!Array.isArray(raw)) {
		throw new XlsxError("invalid-input", "definedNames must be an array");
	}
	const out: DefinedName[] = [];
	// Duplicate detection per scope: key = `${scopeKey}\u0000${NAME_UPPERCASE}`. A legal name is
	// XML-safe, so the NUL separator cannot merge distinct scope/name pairs.
	const seen = new Set<string>();
	for (let i = 0; i < raw.length; i++) {
		const entry = raw[i];
		if (!isPlainRecord(entry)) {
			throw new XlsxError("invalid-input", `definedNames[${i}] must be an object`);
		}
		for (const key of Object.keys(entry)) {
			if (
				key !== "name" &&
				key !== "refersTo" &&
				key !== "localSheetId" &&
				key !== "hidden"
			) {
				throw new XlsxError(
					"invalid-input",
					`definedNames[${i}] has an unknown property "${key}"`,
				);
			}
		}
		const name = entry.name;
		const refersTo = entry.refersTo;
		const localSheetId = entry.localSheetId;
		const hidden = entry.hidden;
		if (typeof name !== "string") {
			throw new XlsxError("invalid-input", `definedNames[${i}].name must be a string`);
		}
		const problem = definedNameProblem(name);
		if (problem !== undefined) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}].name ${definedNameProblemMessage(name, problem)}`,
			);
		}
		// `_xlnm._FilterDatabase` belongs to SheetInput.autoFilter; accepting it here would duplicate it.
		if (isFilterDatabaseName(name)) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}]: "${name}" is managed by SheetInput.autoFilter — set the sheet's autoFilter instead of a defined name`,
			);
		}
		if (typeof refersTo !== "string") {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}] ("${name}"): refersTo must be a string`,
			);
		}
		if (refersTo.length === 0) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}] ("${name}"): refersTo must not be empty`,
			);
		}
		if (refersTo.length > MAX_FORMULA_LEN) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}] ("${name}"): refersTo exceeds Excel's ${MAX_FORMULA_LEN}-character limit`,
			);
		}
		if (refersTo.startsWith("=")) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}] ("${name}"): refersTo must be in stored form, without a leading "="`,
			);
		}
		if (!isXmlSafe(refersTo)) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}] ("${name}"): refersTo contains a character not allowed in XML`,
			);
		}
		let scopeKey = "*";
		if (localSheetId !== undefined) {
			if (
				typeof localSheetId !== "number" ||
				!Number.isInteger(localSheetId) ||
				localSheetId < 0 ||
				localSheetId >= sheetCount
			) {
				throw new XlsxError(
					"invalid-input",
					`definedNames[${i}] ("${name}"): localSheetId must be an integer index of an existing sheet (0..${sheetCount - 1})`,
				);
			}
			scopeKey = String(localSheetId);
		}
		if (hidden !== undefined && typeof hidden !== "boolean") {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}] ("${name}"): hidden must be a boolean`,
			);
		}
		const dupKey = `${scopeKey}\u0000${name.toUpperCase()}`;
		if (seen.has(dupKey)) {
			throw new XlsxError(
				"invalid-input",
				`definedNames[${i}]: duplicate name "${name}" in the same scope (names are case-insensitive)`,
			);
		}
		seen.add(dupKey);
		out.push({
			name,
			refersTo,
			...(localSheetId !== undefined ? { localSheetId } : {}),
			...(hidden === true ? { hidden: true } : {}),
		});
	}
	return out;
}

/** Input `("Bob's", "A1:C3")`; output `'Bob''s'!$A$1:$C$3`. */
export function qualifiedAbsoluteRange(sheetName: string, ref: string): string {
	const quoted = `'${sheetName.replace(/'/g, "''")}'`;
	const range = parseCanonicalRange(ref);
	if (range === undefined) return `${quoted}!${ref}`;
	const abs = (cell: CellRef): string => `$${indexToColumn(cell.col)}$${cell.row}`;
	const cells = ref.includes(":") ? `${abs(range.from)}:${abs(range.to)}` : abs(range.from);
	return `${quoted}!${cells}`;
}

/**
 * Build the hidden sheet-scoped name paired with an auto-filter.
 * Input `("Data", 0, "A1:C3")`; output refers to `'Data'!$A$1:$C$3` in sheet scope 0.
 */
export function filterDatabaseName(
	sheetName: string,
	sheetIndex: number,
	ref: string,
): DefinedName {
	return {
		name: FILTER_DATABASE_NAME,
		refersTo: qualifiedAbsoluteRange(sheetName, ref),
		localSheetId: sheetIndex,
		hidden: true,
	};
}
