import { XlsxError } from "../errors";

/**
 * Throw a writer error scoped to one worksheet.
 *
 * Input: sheet `Data`, message `comments must be an array`.
 * Output: `XlsxError("invalid-input", 'sheet "Data": comments must be an array')`.
 */
export function sheetInvalid(sheetName: string, message: string): never {
	throw new XlsxError("invalid-input", `sheet "${sheetName}": ${message}`);
}

/**
 * Reject object properties outside a feature's allowlist.
 *
 * Input: `{ ref: "A1", typo: true }`, allowed `["ref"]`.
 * Output: an invalid-input error naming `typo`.
 */
export function checkKeys(
	sheetName: string,
	what: string,
	obj: Record<string, unknown>,
	allowed: readonly string[],
): void {
	for (const key of Object.keys(obj)) {
		if (!allowed.includes(key)) {
			sheetInvalid(sheetName, `${what} has an unknown property "${key}"`);
		}
	}
}

/**
 * Keep invalid caller values short enough for useful error messages.
 *
 * Input: a 30-character reference. Output: its first 24 characters plus `…`.
 */
export function shortened(value: string): string {
	return value.length > 24 ? `${value.slice(0, 24)}…` : value;
}
