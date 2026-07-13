// Half B — a seeded, dependency-free byte-mutation engine. Given a real fixture and a {@link Prng},
// it produces a corrupted variant to replay against every reader. Mutations span three layers:
//   • generic bytes  — bit/byte flips, zero runs, truncation, slice duplication (deflate-stream damage
//                       and length lies);
//   • zip structure  — perturb the fields the EOCD/central-directory/local-header walk trusts (entry
//                       counts, offsets, compressed/uncompressed sizes, CRCs, name lengths);
//   • XML-ish text   — where structural markers sit in the CLEAR (CSV, an ODS `mimetype`, an
//                       uncompressed part), blow up a `count="…"`/`ref="…"` number or inject nesting.
// Every mutator degrades gracefully (a no-op if its target isn't present), so the same set applies to
// xlsx / xlsb / ods / csv seeds alike. The engine never allocates from an attacker-controlled length;
// growth mutators are capped so a campaign can't OOM itself.

import type { Prng } from "./prng";

// ZIP signatures (little-endian on disk): local file header, central-directory header, EOCD.
const SIG_LOCAL = [0x50, 0x4b, 0x03, 0x04] as const;
const SIG_CENTRAL = [0x50, 0x4b, 0x01, 0x02] as const;
const SIG_EOCD = [0x50, 0x4b, 0x05, 0x06] as const;

const MAX_GROWTH = 4096; // cap a single duplication so growth can't run away

function findSig(bytes: Uint8Array, sig: readonly number[], prng: Prng): number {
	// Collect every offset of the signature, then pick one (so repeated calls hit different records).
	const hits: number[] = [];
	for (let i = 0; i + sig.length <= bytes.length; i++) {
		let ok = true;
		for (let j = 0; j < sig.length; j++) {
			if (bytes[i + j] !== sig[j]) {
				ok = false;
				break;
			}
		}
		if (ok) hits.push(i);
	}
	if (hits.length === 0) return -1;
	return hits[prng.int(hits.length)] as number;
}

// Overwrite a little-endian uint at `off` (width 2 or 4) with a hostile value.
function pokeUint(bytes: Uint8Array, off: number, width: number, value: number): void {
	for (let i = 0; i < width; i++) {
		if (off + i < bytes.length) bytes[off + i] = (value >>> (8 * i)) & 0xff;
	}
}

const HOSTILE_UINTS: readonly number[] = [
	0, 1, 0xffff, 0xffff_ffff, 0x7fff_ffff, 0x8000_0000, 0xdead_beef,
];

// ── individual mutators (each returns the possibly-resized array) ───────────────────────────────

function bitFlip(b: Uint8Array, prng: Prng): Uint8Array {
	if (b.length === 0) return b;
	const i = prng.int(b.length);
	b[i] = (b[i] as number) ^ (1 << prng.int(8));
	return b;
}

function byteSet(b: Uint8Array, prng: Prng): Uint8Array {
	if (b.length === 0) return b;
	const n = 1 + prng.int(4);
	for (let k = 0; k < n; k++) b[prng.int(b.length)] = prng.byte();
	return b;
}

function zeroRun(b: Uint8Array, prng: Prng): Uint8Array {
	if (b.length === 0) return b;
	const start = prng.int(b.length);
	const len = 1 + prng.int(Math.min(64, b.length - start));
	b.fill(0, start, start + len);
	return b;
}

function truncate(b: Uint8Array, prng: Prng): Uint8Array {
	if (b.length === 0) return b;
	return b.subarray(0, prng.int(b.length));
}

function duplicateSlice(b: Uint8Array, prng: Prng): Uint8Array {
	if (b.length === 0) return b;
	const start = prng.int(b.length);
	const len = 1 + prng.int(Math.min(MAX_GROWTH, b.length - start));
	const out = new Uint8Array(b.length + len);
	out.set(b.subarray(0, start), 0);
	out.set(b.subarray(start, start + len), start);
	out.set(b.subarray(start), start + len);
	return out;
}

function zipStructure(b: Uint8Array, prng: Prng): Uint8Array {
	const sig = prng.pick([SIG_EOCD, SIG_CENTRAL, SIG_LOCAL]);
	const at = findSig(b, sig, prng);
	if (at < 0) return byteSet(b, prng); // no such record — fall back to a generic poke
	const value = prng.pick(HOSTILE_UINTS);
	// Poke one of a few structurally-load-bearing fields relative to the signature. Widths/offsets are
	// the ZIP-spec field layouts; an out-of-range poke is harmless (pokeUint bounds-checks).
	if (sig === SIG_EOCD) {
		// total-entries (off+10, u16), CD size (off+12, u32), CD offset (off+16, u32), comment len (off+20, u16)
		const [o, w] = prng.pick([
			[10, 2],
			[12, 4],
			[16, 4],
			[20, 2],
		] as const);
		pokeUint(b, at + o, w, value);
	} else if (sig === SIG_CENTRAL) {
		// crc (16,u32), compressed (20,u32), uncompressed (24,u32), name-len (28,u16), local-hdr-offset (42,u32)
		const [o, w] = prng.pick([
			[16, 4],
			[20, 4],
			[24, 4],
			[28, 2],
			[42, 4],
		] as const);
		pokeUint(b, at + o, w, value);
	} else {
		// local header: crc (14,u32), compressed (18,u32), uncompressed (22,u32), name-len (26,u16)
		const [o, w] = prng.pick([
			[14, 4],
			[18, 4],
			[22, 4],
			[26, 2],
		] as const);
		pokeUint(b, at + o, w, value);
	}
	return b;
}

// XML-ish text mutations — only bite where markers sit in the clear (CSV / ODS mimetype / an
// uncompressed part). On a fully-deflated xlsx these patterns aren't found and the mutator no-ops.
const TEXT_TARGETS: readonly RegExp[] = [
	/count="(\d+)"/,
	/ref="([^"]*)"/,
	/ span="[^"]*"/,
	/ r="([A-Z]+\d+)"/,
];

function xmlText(b: Uint8Array, prng: Prng): Uint8Array {
	// Decode latin1 (byte-preserving), rewrite, re-encode — safe for the ASCII markers we target.
	let s = "";
	for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i] as number);
	const rx = prng.pick(TEXT_TARGETS);
	const m = rx.exec(s);
	if (m === null) return zeroRun(b, prng); // no clear-text marker — fall back
	const inject = prng.pick([
		"4294967295",
		"999999999999",
		"-1",
		"A1:XFD1048576",
		"".padEnd(2000, "A"),
	]);
	const patched = `${s.slice(0, m.index)}${m[0].replace(m[1] ?? "", inject)}${s.slice(m.index + m[0].length)}`;
	const out = new Uint8Array(patched.length);
	for (let i = 0; i < patched.length; i++) out[i] = patched.charCodeAt(i) & 0xff;
	return out;
}

const MUTATORS: readonly ((b: Uint8Array, prng: Prng) => Uint8Array)[] = [
	bitFlip,
	byteSet,
	zeroRun,
	truncate,
	duplicateSlice,
	zipStructure,
	zipStructure, // weight the structural mutator higher — it targets the reader's trust points
	xmlText,
];

/**
 * Produce one corrupted variant of `seed` driven entirely by `prng` (so a crasher is reproducible
 * from its seed). Applies 1–4 mutators over a fresh copy; the seed bytes are never modified.
 */
export function mutate(seed: Uint8Array, prng: Prng): Uint8Array {
	let b: Uint8Array = seed.slice(); // own copy — never mutate the caller's fixture bytes
	const rounds = 1 + prng.int(4);
	for (let k = 0; k < rounds; k++) {
		const mutator = prng.pick(MUTATORS);
		b = mutator(b, prng);
	}
	return b;
}
