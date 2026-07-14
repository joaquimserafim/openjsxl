import { XlsxError } from "../errors";

// One shared UTF-8 decoder for whole-part reads (stateless in one-shot mode, so reuse is safe).
const decoder = new TextDecoder();

/**
 * Decode a whole part's bytes to a string, as a TYPED failure when it can't be materialized. A
 * part whose decoded text would exceed the JS engine's maximum string length (V8 caps at
 * ~0x1fffffe8 ≈ 512 MiB chars) makes `TextDecoder.decode` throw a bare `Error`/`RangeError` — a
 * hostile ~KB deflated part can land in the 512 MiB–2 GiB band that passes the zip-bomb guards yet
 * still overflows the string. The tolerant-reader contract is "resolve OR typed `XlsxError`, never
 * a bare throw", so surface it as `part-too-large` (F9.7 review follow-up). Only the decode itself
 * is shielded; a decode of representable bytes never throws.
 */
export function decodeText(bytes: Uint8Array): string {
	try {
		return decoder.decode(bytes);
	} catch (cause) {
		throw new XlsxError(
			"part-too-large",
			"a part's decoded text exceeds the maximum string length this runtime supports",
			{ cause },
		);
	}
}
