import { openZip, type ZipArchive } from "../zip";
import type { ReadOptions } from "./workbook";

// F7.4 — format detection. `detectSpreadsheetFormat` sniffs a file's container so a caller can route
// "a user uploaded a spreadsheet" to the right opener (openXlsx / openXlsb / openOds / openCsv)
// without a try-each cascade. It is a best-effort CLASSIFIER, not a validator: a positive result
// means "this looks like format X"; the matching `open*` function is still what fully parses — and,
// on the tolerant-reader contract, rejects — the bytes. Detection reuses the hardened zip reader, so
// it inherits the same bounds (a corrupt / ZIP64 / bomb archive classifies as unknown, never hangs).

/**
 * The spreadsheet formats openjsxl can read. {@link detectSpreadsheetFormat} returns one of these,
 * or `undefined` when the bytes match no known container and are not decodable as text.
 */
export type SpreadsheetFormat = "xlsx" | "xlsb" | "ods" | "csv";

const decoder = new TextDecoder();

// The ODF spreadsheet media type; a spreadsheet TEMPLATE (`…spreadsheet-template`) shares the prefix.
const ODS_SPREADSHEET = "application/vnd.oasis.opendocument.spreadsheet";
// The xlsb workbook part's content type — BIFF12 binary, distinct from the XML spreadsheetml types.
// (All `.xlsb`, macro-enabled or not, use this exact type; its presence uniquely marks a binary book.)
const XLSB_MAIN = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";

// How much of a non-zip input to sample when deciding "is this text?": enough to catch binary noise
// cheaply, bounded so a huge file isn't fully scanned just to classify it.
const TEXT_SAMPLE_BYTES = 8192;

// A ceiling on how many DECOMPRESSED bytes to read when classifying a zip part (mimetype /
// [Content_Types].xml / content.xml): real ones are a few KB and the format markers sit near the
// top, so 1 MiB peeks the whole of any legitimate part. Crucially it bounds a decompression bomb —
// detection STREAMS each part and stops at this cap rather than inflating it whole, so a part
// declaring gigabytes can't force an unbounded allocation on the default (no-maxPartBytes) call.
const MAX_SNIFF_BYTES = 1 << 20;

/**
 * Best-effort spreadsheet-format detection from a file's bytes. Returns `'xlsx'` (also for `.xlsm` /
 * `.xltx` / `.xltm` — all read by {@link openXlsx}), `'xlsb'`, `'ods'`, `'csv'`, or `undefined`.
 *
 * ZIP-based formats (xlsx / xlsb / ods) are told apart by peeking the package: an ODF `mimetype`
 * entry names a spreadsheet → `'ods'` (a `mimetype`-less ODS is still recognized via its
 * `content.xml` spreadsheet body, matching `openOds`'s tolerance); otherwise `[Content_Types].xml`
 * names a binary workbook → `'xlsb'` or an XML one → `'xlsx'`. A non-zip input that decodes as UTF-8
 * text is `'csv'` — CSV/TSV has no magic bytes, so this is a documented heuristic (any
 * delimited-or-not text reads as `'csv'`), never as certain as a container sniff. Corrupt /
 * unsupported archives and binary noise are `undefined`; the corresponding `open*` function remains
 * the authority that parses and rejects.
 */
export async function detectSpreadsheetFormat(
	source: Uint8Array | ArrayBuffer,
	options?: ReadOptions,
): Promise<SpreadsheetFormat | undefined> {
	const bytes = source instanceof Uint8Array ? source : new Uint8Array(source);
	if (bytes.byteLength === 0) return undefined;

	// xlsx / xlsb / ods all share the ZIP (OPC / ODF) container. A leading PK signature marks a zip;
	// peek the entries that name the format. Any zip error (not-a-zip, corrupt, ZIP64) → unknown.
	if (looksLikeZip(bytes)) {
		try {
			const zip = openZip(bytes, options);
			// ODS names its media type in a `mimetype` entry (ODF requires it first + stored).
			if (zip.has("mimetype")) {
				const mimetype = (await sniff(zip, "mimetype")).trim();
				if (mimetype.startsWith(ODS_SPREADSHEET)) return "ods";
			}
			// OOXML distinguishes a binary (xlsb) from an XML (xlsx / xlsm / xltx) workbook in the
			// package content types.
			if (zip.has("[Content_Types].xml")) {
				const ct = await sniff(zip, "[Content_Types].xml");
				if (ct.includes(XLSB_MAIN)) return "xlsb";
				if (isOoxmlSpreadsheet(ct)) return "xlsx";
			} else if (zip.has("content.xml")) {
				// A `mimetype`-less ODS (a real-producer quirk `openOds` deliberately tolerates): the
				// spreadsheet body lives in content.xml, whose `office:spreadsheet` root also tells an
				// .ods from an .odt / .odp — so detect's verdict is never stricter than its opener.
				const content = await sniff(zip, "content.xml");
				if (content.includes("office:spreadsheet")) return "ods";
			}
		} catch {
			// A corrupt / ZIP64 / otherwise-unreadable zip is a container we can't classify.
		}
		return undefined; // a zip, but not a spreadsheet package (e.g. a .docx, or a plain .zip)
	}

	// Not a zip. CSV/TSV has no magic bytes, so decodable UTF-8 text is classified `csv` (best-effort);
	// binary noise is `undefined`.
	return looksLikeText(bytes) ? "csv" : undefined;
}

// A ZIP starts with a PK signature: a local file header (`03 04`), or the empty-archive EOCD
// (`05 06`) / spanned marker (`07 08`). Checking the record bytes — not merely "PK" — keeps a CSV
// that happens to begin with "PK" from being mistaken for a zip.
function looksLikeZip(b: Uint8Array): boolean {
	if (b.length < 4 || b[0] !== 0x50 || b[1] !== 0x4b) return false;
	return (
		(b[2] === 0x03 && b[3] === 0x04) ||
		(b[2] === 0x05 && b[3] === 0x06) ||
		(b[2] === 0x07 && b[3] === 0x08)
	);
}

// The non-binary OOXML spreadsheet workbook part comes in four content types: `.xlsx` / `.xltx`
// (`spreadsheetml.(sheet|template).main+xml`) and macro-enabled `.xlsm` / `.xltm`
// (`ms-excel.(sheet|template).macroEnabled.main+xml`). Any of them ⇒ `openXlsx` reads it.
function isOoxmlSpreadsheet(ct: string): boolean {
	return (
		ct.includes("spreadsheetml.sheet.main") ||
		ct.includes("spreadsheetml.template.main") ||
		ct.includes("ms-excel.sheet.macroEnabled.main") ||
		ct.includes("ms-excel.template.macroEnabled.main")
	);
}

// Is a non-zip input decodable text (→ csv), or binary (→ undefined)? Sample a bounded prefix, reject
// C0 control bytes other than tab / newline / CR (NUL and friends mark binary, though they are valid
// UTF-8), and require the prefix to be valid UTF-8. A multi-byte sequence split by the sample
// boundary is tolerated (decoded in stream mode), not treated as invalid.
function looksLikeText(b: Uint8Array): boolean {
	const sample = b.byteLength > TEXT_SAMPLE_BYTES ? b.subarray(0, TEXT_SAMPLE_BYTES) : b;
	for (const c of sample) {
		if (c < 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) return false;
	}
	try {
		new TextDecoder("utf-8", { fatal: true }).decode(sample, { stream: true });
	} catch {
		return false;
	}
	return true;
}

// Read up to MAX_SNIFF_BYTES DECOMPRESSED bytes of a zip part, streaming and stopping early. The
// early break cancels the underlying inflate (inflateRawStream's `finally`), so a bomb part inflates
// at most ~one chunk past the cap instead of its full declared size — this is what keeps detection
// bounded on the default (no-maxPartBytes) call. Enough to peek any real part's format markers.
async function sniff(zip: ZipArchive, name: string): Promise<string> {
	const chunks: Uint8Array[] = [];
	let total = 0;
	for await (const chunk of zip.readStream(name)) {
		const room = MAX_SNIFF_BYTES - total;
		const take = chunk.byteLength <= room ? chunk : chunk.subarray(0, room);
		chunks.push(take);
		total += take.byteLength;
		if (total >= MAX_SNIFF_BYTES) break;
	}
	const joined = new Uint8Array(total);
	let offset = 0;
	for (const c of chunks) {
		joined.set(c, offset);
		offset += c.byteLength;
	}
	return decoder.decode(joined);
}
