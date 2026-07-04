// The workbook-level media registry for picture write (F6.3). Media parts are shared across the
// whole workbook: identical image bytes are written ONCE as `xl/media/imageK.<ext>` and every
// picture that uses them points at that one part. Numbering is first-occurrence order (deterministic
// given the sheet + image order), so the output stays byte-deterministic.
//
// Dedup keys on a deterministic content hash (FNV-1a) with an exact byte-compare on collision, NOT a
// bare length-grouped byte-compare: the bridge (F6.4) can hand this images read from a hostile file,
// so an attacker controls the image COUNT — a pure pairwise compare would be O(n²) on same-length
// images (a CLAUDE.md "no O(n²) on attacker-controlled counts" violation). FNV-1a is deterministic
// (no randomness) and linear in total media bytes.

export interface MediaRegistry {
	/** Intern image bytes (already validated) with their extension; returns the 1-based media number. */
	intern(bytes: Uint8Array, ext: string): number;
	/** The media parts to emit, in number order: `xl/media/imageK.<ext>` → the bytes. */
	parts(): readonly { readonly name: string; readonly data: Uint8Array }[];
	/** The distinct extensions used, for the content-type `Default` entries. */
	extensions(): readonly string[];
}

// Boil the image bytes down to one small number — a quick "fingerprint" — so we can tell in one
// step whether we've probably seen this exact image before (a full compare then confirms it).
function fnv1a(bytes: Uint8Array): number {
	let h = 0x811c9dc5;
	for (let i = 0; i < bytes.length; i++) {
		h ^= bytes[i] as number;
		h = Math.imul(h, 0x01000193);
	}
	return h >>> 0;
}

// True only when two buffers hold exactly the same bytes.
function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
	if (a.length !== b.length) return false;
	for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
	return true;
}

// Keeps track of the images in a workbook. Hand it image bytes and it hands back a number; the same
// image handed in twice gets the same number, so its bytes are written to the file only once.
export function createMediaRegistry(): MediaRegistry {
	const entries: { bytes: Uint8Array; ext: string; number: number }[] = [];
	// content-hash key → indices into `entries` (a list, so a hash collision stays exact).
	const byHash = new Map<string, number[]>();

	return {
		// Store these bytes and give back their 1-based number. If the exact same image is already
		// stored, return its existing number (and don't store a second copy). The extension is part of
		// the identity: the media part is named `imageN.<ext>`, so the same bytes under two different
		// extensions are two different files — never merge them, or a picture would point at a file
		// that was never written.
		intern(bytes, ext) {
			const key = `${bytes.length}:${fnv1a(bytes)}:${ext}`;
			const bucket = byHash.get(key);
			if (bucket !== undefined) {
				for (const idx of bucket) {
					const entry = entries[idx];
					if (
						entry !== undefined &&
						entry.ext === ext &&
						bytesEqual(entry.bytes, bytes)
					) {
						return entry.number;
					}
				}
			}
			const number = entries.length + 1;
			const idx = entries.length;
			entries.push({ bytes, ext, number });
			if (bucket !== undefined) bucket.push(idx);
			else byHash.set(key, [idx]);
			return number;
		},
		// The image files to put in the .xlsx, one per unique image: `xl/media/imageN.png` → its bytes.
		parts() {
			return entries.map((e) => ({
				name: `xl/media/image${e.number}.${e.ext}`,
				data: e.bytes,
			}));
		},
		// The set of file extensions used (e.g. png, jpeg) — the package needs to list each one.
		extensions() {
			return [...new Set(entries.map((e) => e.ext))];
		},
	};
}
