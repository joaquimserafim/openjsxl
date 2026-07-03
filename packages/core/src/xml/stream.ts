import { tokenize, type XmlToken } from "./tokenizer";

// A chunk-fed wrapper around the streaming tokenizer (./tokenizer). Decompressed worksheet
// text arrives in arbitrary chunks, and a single token — a tag, a comment, a CDATA section,
// an entity reference — can straddle a chunk boundary. `createXmlStream` buffers the input
// and only ever hands the underlying `tokenize` a prefix it is *certain* is complete, so
// peak memory tracks the largest single unfinished construct (one tag, or one run of text
// up to the next `<`), not the whole part.
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

// Is `buf.slice(at)` a non-empty proper prefix of `marker` — i.e. an incomplete marker that
// might complete once more text arrives? (e.g. "<!-" is a prefix of "<!--".)
function isPartialMarkerOf(buf: string, at: number, marker: string): boolean {
	const tail = buf.length - at;
	return tail < marker.length && marker.startsWith(buf.slice(at));
}

// The length of the longest prefix of `buf` that contains only complete constructs, so it
// can be tokenized now with no construct straddling the cut. The remainder is kept buffered.
function safeBoundary(buf: string): number {
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
			return splitEntity ? amp : len;
		}

		// An incomplete comment/CDATA marker (e.g. "<!-", "<![CD") must wait for more input,
		// otherwise it would be misread as a declaration.
		if (isPartialMarkerOf(buf, lt, COMMENT_OPEN) || isPartialMarkerOf(buf, lt, CDATA_OPEN)) {
			return lt;
		}

		let end: number;
		if (buf.startsWith(COMMENT_OPEN, lt)) {
			end = buf.indexOf("-->", lt + COMMENT_OPEN.length);
			if (end === -1) return lt;
			i = end + 3;
		} else if (buf.startsWith(CDATA_OPEN, lt)) {
			end = buf.indexOf("]]>", lt + CDATA_OPEN.length);
			if (end === -1) return lt;
			i = end + 3;
		} else if (buf[lt + 1] === "?") {
			end = buf.indexOf("?>", lt + 2);
			if (end === -1) return lt;
			i = end + 2;
		} else {
			// A tag (open/close/declaration). Scan for its '>', skipping quoted attribute
			// values so a '>' inside a value can't end it early.
			let k = lt + 1;
			let quote = "";
			let close = -1;
			while (k < len) {
				const ch = buf[k];
				if (quote !== "") {
					if (ch === quote) quote = "";
				} else if (ch === '"' || ch === "'") {
					quote = ch;
				} else if (ch === ">") {
					close = k;
					break;
				}
				k++;
			}
			if (close === -1) return lt; // tag (or its quoted value) not closed yet
			i = close + 1;
		}
		safe = i;
	}

	return safe;
}

export function createXmlStream(): XmlStream {
	let buffer = "";

	return {
		push(text: string): XmlToken[] {
			buffer += text;
			const cut = safeBoundary(buffer);
			if (cut === 0) return [];
			const tokens = [...tokenize(buffer.slice(0, cut))];
			buffer = buffer.slice(cut);
			return tokens;
		},
		flush(): XmlToken[] {
			if (buffer === "") return [];
			const tokens = [...tokenize(buffer)];
			buffer = "";
			return tokens;
		},
	};
}
