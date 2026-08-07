import { XlsxError } from "../errors";
import { CRC32_INIT, crc32Final, crc32Update } from "../utils/crc32";

/** The central-directory fields needed to validate a fully consumed ZIP entry. */
export interface VerifiableZipEntry {
	readonly name: string;
	readonly crc32: number;
	readonly uncompressedSize: number;
}

export interface EntryIntegrityVerifier {
	/** Include one more uncompressed output chunk in the final integrity checks. */
	update(chunk: Uint8Array): void;
	/** Validate the accumulated output after the entry reaches its natural end. */
	finish(): void;
}

/** Require the actual output length to match the central-directory declaration exactly. */
export function verifyEntrySize(entry: VerifiableZipEntry, actualSize: number): void {
	if (actualSize !== entry.uncompressedSize) {
		throw new XlsxError(
			"corrupt-zip",
			`corrupt zip: entry ${entry.name} produced ${actualSize} bytes; expected ${entry.uncompressedSize}`,
		);
	}
}

/** Require the actual output CRC-32 to match the central-directory declaration exactly. */
export function verifyEntryCrc32(entry: VerifiableZipEntry, actualCrc32: number): void {
	if (actualCrc32 !== entry.crc32) {
		const actual = actualCrc32.toString(16).padStart(8, "0");
		const expected = entry.crc32.toString(16).padStart(8, "0");
		throw new XlsxError(
			"corrupt-zip",
			`corrupt zip: CRC-32 mismatch for ${entry.name} (got ${actual}, expected ${expected})`,
		);
	}
}

/**
 * Build one verifier shared by buffered and streaming ZIP reads.
 *
 * Call `finish` only after natural EOF. A cancelled stream intentionally makes no
 * whole-entry integrity claim.
 */
export function createEntryIntegrityVerifier(
	entry: VerifiableZipEntry,
	verifyChecksum: boolean,
): EntryIntegrityVerifier {
	let actualSize = 0;
	let runningCrc32 = CRC32_INIT;

	return {
		update(chunk): void {
			actualSize += chunk.byteLength;
			if (verifyChecksum) runningCrc32 = crc32Update(runningCrc32, chunk);
		},
		finish(): void {
			verifyEntrySize(entry, actualSize);
			if (verifyChecksum) verifyEntryCrc32(entry, crc32Final(runningCrc32));
		},
	};
}

/** Validate one already-buffered entry through the same verifier used by streamed reads. */
export function verifyEntryOutput(
	entry: VerifiableZipEntry,
	output: Uint8Array,
	verifyChecksum: boolean,
): void {
	const verifier = createEntryIntegrityVerifier(entry, verifyChecksum);
	verifier.update(output);
	verifier.finish();
}
