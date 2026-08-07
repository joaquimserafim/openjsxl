import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import { crc32 } from "../../utils/crc32";
import {
	createEntryIntegrityVerifier,
	type VerifiableZipEntry,
	verifyEntryCrc32,
	verifyEntryOutput,
	verifyEntrySize,
} from "../entry-integrity";

const encode = (value: string): Uint8Array => new TextEncoder().encode(value);

function entry(overrides: Partial<VerifiableZipEntry> = {}): VerifiableZipEntry {
	return {
		name: "xl/worksheets/sheet1.xml",
		crc32: 0,
		uncompressedSize: 0,
		...overrides,
	};
}

describe("verifyEntrySize", () => {
	it("accepts the exact central-directory size", () => {
		expect(() => verifyEntrySize(entry({ uncompressedSize: 7 }), 7)).not.toThrow();
	});

	it.each([6, 8])("rejects an actual size of %i when 7 bytes were declared", (actualSize) => {
		const verify = () => verifyEntrySize(entry({ uncompressedSize: 7 }), actualSize);

		expect(verify).toThrowError(XlsxError);
		expect(verify).toThrow(/produced \d+ bytes; expected 7/);
	});
});

describe("verifyEntryCrc32", () => {
	it("accepts an exact CRC, including the valid zero value", () => {
		expect(() => verifyEntryCrc32(entry({ crc32: 0 }), 0)).not.toThrow();
	});

	it("rejects a mismatch and formats both CRC values as eight hexadecimal digits", () => {
		const verify = () => verifyEntryCrc32(entry({ crc32: 0x00000001 }), 0x00000002);

		expect(verify).toThrowError(XlsxError);
		expect(verify).toThrow(/got 00000002, expected 00000001/);
	});
});

describe("createEntryIntegrityVerifier", () => {
	it("applies one checksum consistently across multiple output chunks", () => {
		const first = encode("123");
		const second = encode("456789");
		const expected = entry({
			crc32: crc32(encode("123456789")),
			uncompressedSize: first.byteLength + second.byteLength,
		});
		const verifier = createEntryIntegrityVerifier(expected, true);

		verifier.update(first);
		verifier.update(second);

		expect(() => verifier.finish()).not.toThrow();
	});

	it("always checks size when checksum verification is disabled", () => {
		const verifier = createEntryIntegrityVerifier(entry({ uncompressedSize: 2 }), false);
		verifier.update(encode("one"));

		expect(() => verifier.finish()).toThrow(/produced 3 bytes; expected 2/);
	});

	it("skips only CRC comparison when checksum verification is disabled", () => {
		const verifier = createEntryIntegrityVerifier(
			entry({ crc32: 0xdeadbeef, uncompressedSize: 3 }),
			false,
		);
		verifier.update(encode("one"));

		expect(() => verifier.finish()).not.toThrow();
	});

	it("validates a naturally empty entry", () => {
		const verifier = createEntryIntegrityVerifier(entry(), true);

		expect(() => verifier.finish()).not.toThrow();
	});
});

describe("verifyEntryOutput", () => {
	it("gives buffered output the same size and checksum checks as streamed output", () => {
		const output = encode("buffered");
		const expected = entry({ crc32: crc32(output), uncompressedSize: output.byteLength });

		expect(() => verifyEntryOutput(expected, output, true)).not.toThrow();
		expect(() => verifyEntryOutput({ ...expected, crc32: 1 }, output, true)).toThrow(
			/CRC-32 mismatch/,
		);
	});
});
