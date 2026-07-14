import { XlsxError } from "../errors";
import { isWhitespace } from "../utils";
import { tokenize, type XmlToken } from "./tokenizer";

// A chunk-fed wrapper around the streaming tokenizer (./tokenizer). Decompressed worksheet
// text arrives in arbitrary chunks, and a single token — a tag, a comment, a CDATA section,
// an entity reference — can straddle a chunk boundary. `createXmlStream` buffers the input
// and only ever hands the underlying `tokenize` a prefix it is *certain* is complete, so
// peak memory tracks the largest single unfinished construct (one tag, or one run of text
// up to the next `<`), not the whole part.
//
// F9.7 hostile-input hardening: an unfinished construct's bytes are HELD as an array of
// chunks and only the NEW chunk (plus a terminator-sized overlap) is searched per push —
// never the whole construct again. The pre-F9.7 code rescanned (and re-flattened) one
// growing string from index 0 on every push, which made a single giant straddling construct
// O(n²) CPU on attacker-controlled input. And a single construct can no longer pin the
// buffer without bound: past MAX_UNFINISHED_CONSTRUCT the stream fails typed.
//
// Consumers must still concatenate consecutive `text` tokens: a text run split across two
// chunks emits as two `text` events, exactly as CDATA/entities already split runs.

export interface XmlStream {
	/** Feed the next text chunk; returns the tokens that became complete. */
	push(text: string): XmlToken[];
	/** Tokenize whatever remains after the last chunk (call once at end of input). */
	flush(): XmlToken[];
}

const CDATA_OPEN = "<![CDATA[";
const COMMENT_OPEN = "<!--";

// A real entity reference is short (`&#x10FFFF;` is 10 chars). If a `&` in trailing text is
// further than this from the end with no `;`, it cannot be a split entity — it is a literal
// `&` (only valid in escaped form, but tolerated) and is emitted rather than pinning the
// buffer, so a stray `&` can't defeat the constant-memory guarantee.
const MAX_ENTITY_TAIL = 16;

/**
 * A single unterminated markup construct — one tag, comment, CDATA section, or PI — larger than
 * this fails typed (`part-too-large`) instead of buffering without bound. Cell TEXT is drained
 * progressively and never hits this; only pathological/hostile markup pins this much unclosed
 * construct, and 64 MiB is orders of magnitude above anything a real producer writes (F9.7).
 */
export const MAX_UNFINISHED_CONSTRUCT = 64 * 1024 * 1024;

// The delimited (non-tag) constructs: opener length + terminator (searched with a plain indexOf,
// quote-UNAWARE — exactly as the tokenizer scans each). An OPEN tag is handled separately because
// its terminator (`>`) must respect quoted attribute values. `gt` covers the two constructs the
// tokenizer ends at the first plain `>`: a `<!…>` declaration (e.g. `<!DOCTYPE …>`) and a `</…>`
// close tag.
const DELIMS = {
	comment: { open: COMMENT_OPEN.length, term: "-->" },
	cdata: { open: CDATA_OPEN.length, term: "]]>" },
	pi: { open: 2, term: "?>" },
	gt: { open: 2, term: ">" },
} as const;

type PendingKind = keyof typeof DELIMS | "tag";

// Is `buf.slice(at)` a non-empty proper prefix of `marker` — i.e. an incomplete marker that
// might complete once more text arrives? (e.g. "<!-" is a prefix of "<!--".)
function isPartialMarkerOf(buf: string, at: number, marker: string): boolean {
	const tail = buf.length - at;
	return tail < marker.length && marker.startsWith(buf.slice(at));
}

// Where the tag scanner is within a tag, so a chunk boundary can resume mid-tag. This MIRRORS the
// tokenizer's tag parse exactly (tokenizer.ts): a NAME comes first (a `>` ends it; `'`/`"`/`=` are
// literal name chars), then ATTRS, and only a value AFTER `=` protects a `>` inside quotes. Getting
// this wrong (e.g. treating a `'` in the name as a quote) makes safeBoundary end a tag at a
// different `>` than the tokenizer, splitting the stream where one-shot would not.
type TagPhase = "name" | "attrs" | "vstart" | "vunq" | "vq";
interface TagState {
	readonly phase: TagPhase;
	readonly quote: string; // the open value-quote in phase "vq"
}
const TAG_START: TagState = { phase: "name", quote: "" };

// Scan a tag from `from` with the carried `state`, returning the index of the '>' that ends it
// (-1 if still open) and the state to resume with. One char per step, no re-scan — linear.
function scanTag(buf: string, from: number, state: TagState): { close: number; state: TagState } {
	const len = buf.length;
	let phase = state.phase;
	let quote = state.quote;
	for (let k = from; k < len; k++) {
		const ch = buf[k];
		switch (phase) {
			case "name":
				if (ch === ">") return { close: k, state: TAG_START };
				if (isWhitespace(ch) || ch === "/") phase = "attrs";
				break;
			case "attrs":
				if (ch === ">") return { close: k, state: TAG_START };
				if (ch === "=") phase = "vstart"; // a value follows; ws/'/'/attr-name chars stay in attrs
				break;
			case "vstart": // after '=', skipping whitespace to the value start
				if (ch === '"' || ch === "'") {
					phase = "vq";
					quote = ch;
				} else if (ch === ">") {
					return { close: k, state: TAG_START }; // empty unquoted value → '>' ends the tag
				} else if (ch === "/") {
					phase = "attrs";
				} else if (!isWhitespace(ch)) {
					phase = "vunq";
				}
				break;
			case "vunq": // an unquoted value: ends at whitespace / '/' / '>'
				if (ch === ">") return { close: k, state: TAG_START };
				if (isWhitespace(ch) || ch === "/") phase = "attrs";
				break;
			case "vq": // inside a quoted value: only the matching quote closes it — '>' is protected
				if (ch === quote) {
					phase = "attrs";
					quote = "";
				}
				break;
		}
	}
	return { close: -1, state: { phase, quote } };
}

// Classify `buf`: `cut` is the end of its longest complete-constructs-only prefix, and `open`
// describes the unterminated construct starting at `cut` (or null when the remainder is just a
// short ambiguity — a partial `<!-`/`<![CD` marker or a possibly-split entity — that is cheap
// to rescan whole once more input arrives).
function safeBoundary(buf: string): {
	cut: number;
	open: { kind: PendingKind; tagState: TagState } | null;
} {
	const len = buf.length;
	let i = 0; // scan cursor
	let safe = 0; // end of the last complete construct (the cut point)

	while (i < len) {
		const lt = buf.indexOf("<", i);

		if (lt === -1) {
			// Trailing text with no following markup. It is safe except for a split entity
			// (e.g. "&amp" awaiting its ";"): keep from an unterminated "&" onward — but only
			// while it could still be a short entity, so a stray literal "&" doesn't pin it.
			const amp = buf.lastIndexOf("&");
			const splitEntity =
				amp >= i && buf.indexOf(";", amp) === -1 && len - amp <= MAX_ENTITY_TAIL;
			return { cut: splitEntity ? amp : len, open: null };
		}

		// An incomplete comment/CDATA marker (e.g. "<!-", "<![CD") must wait for more input,
		// otherwise it would be misread as a declaration. The ambiguity is < 9 chars, so
		// leaving it unclassified (whole-rescan next push) cannot grow unbounded.
		if (isPartialMarkerOf(buf, lt, COMMENT_OPEN) || isPartialMarkerOf(buf, lt, CDATA_OPEN)) {
			return { cut: lt, open: null };
		}

		if (buf.startsWith(COMMENT_OPEN, lt)) {
			const end = buf.indexOf("-->", lt + COMMENT_OPEN.length);
			if (end === -1) return { cut: lt, open: { kind: "comment", tagState: TAG_START } };
			i = end + 3;
		} else if (buf.startsWith(CDATA_OPEN, lt)) {
			const end = buf.indexOf("]]>", lt + CDATA_OPEN.length);
			if (end === -1) return { cut: lt, open: { kind: "cdata", tagState: TAG_START } };
			i = end + 3;
		} else if (buf[lt + 1] === "!" || buf[lt + 1] === "/") {
			// A `<!…>` declaration that is NOT `<!--`/`<![CDATA[` (both handled above, their
			// partials by the marker check), OR a `</…>` close tag: the tokenizer ends BOTH at the
			// first plain `>` — quote-UNAWARE (tokenizer.ts:75 / :82), unlike an open tag's value
			// quotes. Routing a close tag through the quote-aware scanTag would end it at a later
			// `>`, splitting the stream where one-shot would not (the F9.7 review's close-tag bug).
			const end = buf.indexOf(">", lt + 2);
			if (end === -1) return { cut: lt, open: { kind: "gt", tagState: TAG_START } };
			i = end + 1;
		} else if (buf[lt + 1] === "?") {
			const end = buf.indexOf("?>", lt + 2);
			if (end === -1) return { cut: lt, open: { kind: "pi", tagState: TAG_START } };
			i = end + 2;
		} else {
			const next = buf[lt + 1];
			// A '<' NOT followed by a name-start / '!' / '?' / '/' is a LITERAL '<' in text, not a
			// tag — the tokenizer's phantom-element guard (tokenizer.ts). safeBoundary must agree,
			// or it would classify e.g. `<=` as an unterminated tag and later "complete" it at a
			// stray '>' (inside `]]>`), splitting the stream where one-shot tokenize would not.
			if (next === ">" || next === "=" || isWhitespace(next)) {
				i = lt + 1; // the '<' drains as text; keep scanning (safe is NOT advanced past text)
				continue;
			}
			const { close, state } = scanTag(buf, lt + 1, TAG_START);
			if (close === -1) return { cut: lt, open: { kind: "tag", tagState: state } };
			i = close + 1;
		}
		safe = i;
	}

	return { cut: safe, open: null };
}

export function createXmlStream(): XmlStream {
	// Unclassified input with NO known-open construct. Stays small: complete constructs drain
	// every push, and the only unclassified remainders are a partial `<!-`/`<![CD` marker
	// (< 9 chars) or a split-entity tail (≤ MAX_ENTITY_TAIL).
	let buffer = "";
	// The unterminated construct currently open, held as PIECES — joined exactly once, when it
	// completes. Only the newest chunk is ever searched, so total work stays linear.
	let pending: PendingKind | null = null;
	let held: string[] = [];
	let heldLen = 0;
	let tagState: TagState = TAG_START; // tag resume state (phase + open value-quote) across chunks
	let tail = ""; // delimited resume state: last term.length-1 SEARCHED chars (never opener bytes)

	const capCheck = () => {
		if (heldLen + buffer.length > MAX_UNFINISHED_CONSTRUCT) {
			throw new XlsxError(
				"part-too-large",
				`malformed xml: a single unterminated markup construct exceeds ${MAX_UNFINISHED_CONSTRUCT} characters`,
			);
		}
	};

	// Drain `buffer`: tokenize its complete prefix, and move an unterminated construct (if one
	// starts) into `held` with its resume state, leaving `buffer` with only the short ambiguous
	// remainder (if any).
	const classify = (tokens: XmlToken[]): void => {
		const { cut, open } = safeBoundary(buffer);
		if (cut > 0) {
			tokens.push(...tokenize(buffer.slice(0, cut)));
		}
		if (open !== null) {
			const rest = buffer.slice(cut);
			pending = open.kind;
			held = [rest];
			heldLen = rest.length;
			tagState = open.tagState;
			// The searched region excludes the construct's opener — a terminator may not borrow
			// opener bytes (`<!-->` is NOT a complete comment), so neither may the overlap window.
			tail =
				open.kind === "tag"
					? ""
					: rest.slice(
							Math.max(
								DELIMS[open.kind].open,
								rest.length - (DELIMS[open.kind].term.length - 1),
							),
						);
			buffer = "";
		} else if (cut > 0) {
			buffer = buffer.slice(cut);
		}
		capCheck();
	};

	// The held construct is complete: `endInText` is where it ends within the newest chunk.
	// Join it once, tokenize it, reset, and classify whatever follows it in the same chunk.
	const complete = (tokens: XmlToken[], text: string, endInText: number): void => {
		const construct = held.join("") + text.slice(0, endInText);
		pending = null;
		held = [];
		heldLen = 0;
		tagState = TAG_START;
		tail = "";
		tokens.push(...tokenize(construct));
		buffer = text.slice(endInText);
		classify(tokens);
	};

	return {
		push(text: string): XmlToken[] {
			const tokens: XmlToken[] = [];
			if (pending === null) {
				buffer += text;
				classify(tokens);
				return tokens;
			}
			if (pending === "tag") {
				const scan = scanTag(text, 0, tagState);
				if (scan.close === -1) {
					held.push(text);
					heldLen += text.length;
					tagState = scan.state;
					capCheck();
					return tokens;
				}
				complete(tokens, text, scan.close + 1);
				return tokens;
			}
			const { term } = DELIMS[pending];
			// Search the new chunk plus the last term.length-1 already-held chars, so a
			// terminator split across the boundary is still found — and nothing older is ever
			// re-touched (the linearity guarantee).
			const window = tail + text;
			const at = window.indexOf(term);
			if (at === -1) {
				held.push(text);
				heldLen += text.length;
				// Keep only the last term.length-1 chars as the next overlap. A 1-char terminator
				// (`gt`) can't straddle a boundary, so its tail is EMPTY — `slice(-0)` would keep the
				// WHOLE window and re-search it every push (the F9.7 review's decl O(n²) bug).
				tail = term.length > 1 ? window.slice(-(term.length - 1)) : "";
				capCheck();
				return tokens;
			}
			// ≥ 1: a terminator wholly inside `tail` would have been found on a previous push.
			complete(tokens, text, at + term.length - tail.length);
			return tokens;
		},
		flush(): XmlToken[] {
			const rest = held.join("") + buffer;
			pending = null;
			held = [];
			heldLen = 0;
			tagState = TAG_START;
			tail = "";
			buffer = "";
			if (rest === "") return [];
			return [...tokenize(rest)];
		},
	};
}
