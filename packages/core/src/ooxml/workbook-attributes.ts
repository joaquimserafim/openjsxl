import type { WorkbookProtection } from "../types";
import { MAX_SPIN_COUNT } from "./styles";

// Pure conversions from workbook XML attributes to reader model values. This module has no
// tokenizer state and is internal to the OOXML reader.

/** Input `"1"`/`"true"` gives `true`; `"0"`/`"false"` gives `false`; anything else is dropped. */
export function booleanAttribute(value: string | undefined): boolean | undefined {
	if (value === "1" || value === "true") return true;
	if (value === "0" || value === "false") return false;
	return undefined;
}

/**
 * Convert one `<workbookProtection>` attribute map into the reader model.
 * Input: `{ lockStructure: "1", workbookSpinCount: "2" }`.
 * Output: `{ lockStructure: true, workbookSpinCount: 2 }`; invalid values are omitted.
 */
export function workbookProtectionFromAttributes(
	attrs: Readonly<Record<string, string>>,
): WorkbookProtection | undefined {
	const out: {
		lockStructure?: boolean;
		lockWindows?: boolean;
		workbookPassword?: string;
		workbookAlgorithmName?: string;
		workbookHashValue?: string;
		workbookSaltValue?: string;
		workbookSpinCount?: number;
	} = {};
	const lockStructure = booleanAttribute(attrs.lockStructure);
	if (lockStructure !== undefined) out.lockStructure = lockStructure;
	const lockWindows = booleanAttribute(attrs.lockWindows);
	if (lockWindows !== undefined) out.lockWindows = lockWindows;
	if (attrs.workbookPassword !== undefined) out.workbookPassword = attrs.workbookPassword;
	if (attrs.workbookAlgorithmName !== undefined) {
		out.workbookAlgorithmName = attrs.workbookAlgorithmName;
	}
	if (attrs.workbookHashValue !== undefined) out.workbookHashValue = attrs.workbookHashValue;
	if (attrs.workbookSaltValue !== undefined) out.workbookSaltValue = attrs.workbookSaltValue;

	// Input `"0".."4294967295"` is an xsd:unsignedInt; malformed or larger values are omitted.
	const spin = attrs.workbookSpinCount;
	if (spin !== undefined && /^[0-9]+$/.test(spin)) {
		const number = Number(spin);
		if (Number.isInteger(number) && number <= MAX_SPIN_COUNT) {
			out.workbookSpinCount = number;
		}
	}
	return Object.keys(out).length > 0 ? out : undefined;
}
