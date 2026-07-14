import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { decodeText } from "../decode";

// F9.7 review follow-up — a whole-part decode must never bare-throw. A part whose decoded text
// would exceed the JS engine's maximum string length (V8 ~0x1fffffe8 chars) can slip past the
// zip-bomb guards (its output is under the 2 GiB ceiling and its ratio under 300x) yet still
// overflow the string; the tolerant-reader contract requires a typed XlsxError, not the engine's
// bare Error.

describe("decodeText", () => {
	it("decodes representable bytes normally", () => {
		expect(decodeText(new TextEncoder().encode("<x>héllo 😀</x>"))).toBe("<x>héllo 😀</x>");
		expect(decodeText(new Uint8Array(0))).toBe("");
	});

	it("surfaces an over-the-string-limit part as typed part-too-large, not a bare throw", () => {
		// V8 caps a string at 0x1fffffe8 chars; decoding one more single-byte char overflows. The
		// failure is a synchronous length check (the string is never allocated), so this only holds
		// the input array in memory. Non-V8 engines with a higher limit simply decode it — the guard
		// is still correct there (nothing to catch).
		const overLimit = 0x1fffffe8 + 8;
		let bytes: Uint8Array;
		try {
			bytes = new Uint8Array(overLimit).fill(0x41); // 'A'
		} catch {
			return; // a runtime that can't even allocate the array — nothing to assert
		}
		try {
			decodeText(bytes);
			// Some engines allow a longer string; if so there is no error to type — accept that.
		} catch (e) {
			expect(e).toBeInstanceOf(XlsxError);
			expect((e as XlsxError).code).toBe("part-too-large");
		}
	});
});
