// A tiny, zero-dependency, declarative .xlsx builder for test fixtures.
//
// Describe a workbook as data — sheets, cells, merges, hyperlinks, comments — and buildWorkbook
// emits only the OOXML parts that description needs, wires up content-types and relationships,
// and packs it all into a STORED (uncompressed) ZIP with a tiny CRC32 (no zip/deflate dep).
// Output is deterministic (fixed DOS timestamps), so re-running produces byte-identical files.
//
// This exists so synthetic fixtures are generated in code, never hand-written. It is NOT a
// general writer — it covers exactly what the reader's fixtures exercise. Note that testing the
// reader against files this builds is circular (both share assumptions); real-producer files in
// ../data are what actually catch mistakes. Use this for controlled edge-case shapes.
//
// Cell spec (one field selects the kind): { ref, text } | { ref, number } | { ref, bool } |
//   { ref, serial, numFmtId? | numFmt? } | { ref, formula, number? }. Any cell may carry a
//   numFmtId (built-in) or numFmt (custom code string) to style it.
//
// Column / row default styles (exercise style inheritance): a sheet may carry
//   columns: [{ min, max, numFmtId? | numFmt? }]  → emits <col min max style> (column default)
//   rowStyles: { <rowNumber>: { numFmtId? | numFmt? } } → emits <row s customFormat="1"> for that
//     row, so its cells that omit their own `s` inherit the format.

const encoder = new TextEncoder();

const CRC_TABLE = (() => {
	const table = new Uint32Array(256);
	for (let n = 0; n < 256; n++) {
		let c = n;
		for (let k = 0; k < 8; k++) c = (c & 1) !== 0 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
		table[n] = c >>> 0;
	}
	return table;
})();

function crc32(bytes) {
	let crc = 0xffffffff;
	for (let i = 0; i < bytes.length; i++) crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
	return (crc ^ 0xffffffff) >>> 0;
}

const u16 = (n) => Uint8Array.from([n & 0xff, (n >>> 8) & 0xff]);
const u32 = (n) =>
	Uint8Array.from([n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff]);

function concat(parts) {
	let total = 0;
	for (const part of parts) total += part.length;
	const out = new Uint8Array(total);
	let offset = 0;
	for (const part of parts) {
		out.set(part, offset);
		offset += part.length;
	}
	return out;
}

// Fixed DOS date (1980-01-01) and time (00:00) for deterministic output.
const DOS_TIME = 0;
const DOS_DATE = 0x0021;

function zipStore(files) {
	const local = [];
	const central = [];
	let offset = 0;

	for (const file of files) {
		const name = encoder.encode(file.name);
		const { data } = file;
		const crc = crc32(data);
		const size = data.length;

		const header = concat([
			u32(0x04034b50),
			u16(20),
			u16(0),
			u16(0), // method: 0 = stored
			u16(DOS_TIME),
			u16(DOS_DATE),
			u32(crc),
			u32(size),
			u32(size),
			u16(name.length),
			u16(0),
			name,
		]);
		local.push(header, data);

		central.push(
			concat([
				u32(0x02014b50),
				u16(20),
				u16(20),
				u16(0),
				u16(0),
				u16(DOS_TIME),
				u16(DOS_DATE),
				u32(crc),
				u32(size),
				u32(size),
				u16(name.length),
				u16(0),
				u16(0),
				u16(0),
				u16(0),
				u32(0),
				u32(offset),
				name,
			]),
		);
		offset += header.length + data.length;
	}

	const directory = concat(central);
	const eocd = concat([
		u32(0x06054b50),
		u16(0),
		u16(0),
		u16(files.length),
		u16(files.length),
		u32(directory.length),
		u32(offset),
		u16(0),
	]);
	return concat([...local, directory, eocd]);
}

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT = "application/vnd.openxmlformats-officedocument.spreadsheetml";

const escapeText = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const escapeAttr = (s) => escapeText(s).replace(/"/g, "&quot;");
const needsPreserve = (s) => s !== s.trim();
const rowOf = (ref) => Number(/\d+$/.exec(ref)?.[0]);

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function relsPart(rels) {
	const items = rels
		.map(
			(r) =>
				`<Relationship Id="${r.id}" Type="${r.type}" Target="${escapeAttr(r.target)}"${
					r.mode === "External" ? ' TargetMode="External"' : ""
				}/>`,
		)
		.join("");
	return `${XML_DECL}\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
}

// Accumulates the shared string table across every sheet, deduping and indexing.
function makeSharedStrings() {
	const list = [];
	const index = new Map();
	let total = 0;
	return {
		has: () => list.length > 0,
		ref(text) {
			total++;
			let i = index.get(text);
			if (i === undefined) {
				i = list.length;
				list.push(text);
				index.set(text, i);
			}
			return i;
		},
		xml() {
			const items = list
				.map(
					(s) =>
						`<si><t${needsPreserve(s) ? ' xml:space="preserve"' : ""}>${escapeText(s)}</t></si>`,
				)
				.join("");
			return `${XML_DECL}\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${total}" uniqueCount="${list.length}">${items}</sst>`;
		},
	};
}

// Accumulates cell formats. Index 0 is always the default (General); each distinct numFmt gets
// its own cellXf. Custom codes are written into <numFmts> with ids >= 164.
function makeStyles() {
	const custom = new Map(); // code -> numFmtId
	const xfs = [0]; // cellXf index -> numFmtId; 0 = General
	const byId = new Map([[0, 0]]); // numFmtId -> cellXf index
	let nextCustom = 164;
	return {
		has: () => xfs.length > 1 || custom.size > 0,
		indexFor(cell) {
			let numFmtId;
			if (typeof cell.numFmtId === "number") numFmtId = cell.numFmtId;
			else if (typeof cell.numFmt === "string") {
				numFmtId = custom.get(cell.numFmt);
				if (numFmtId === undefined) {
					numFmtId = nextCustom++;
					custom.set(cell.numFmt, numFmtId);
				}
			} else return undefined;
			let idx = byId.get(numFmtId);
			if (idx === undefined) {
				idx = xfs.length;
				xfs.push(numFmtId);
				byId.set(numFmtId, idx);
			}
			return idx;
		},
		xml() {
			const numFmts = [...custom];
			const numFmtsBlock = numFmts.length
				? `<numFmts count="${numFmts.length}">${numFmts
						.map(
							([code, id]) =>
								`<numFmt numFmtId="${id}" formatCode="${escapeAttr(code)}"/>`,
						)
						.join("")}</numFmts>`
				: "";
			const cellXfs = xfs
				.map(
					(id) =>
						`<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0"${
							id !== 0 ? ' applyNumberFormat="1"' : ""
						}/>`,
				)
				.join("");
			return `${XML_DECL}
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmtsBlock}<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${cellXfs}</cellXfs></styleSheet>`;
		},
	};
}

function cellXml(cell, sst, styles) {
	const s = styles.indexFor(cell);
	const sAttr = s !== undefined ? ` s="${s}"` : "";
	if (cell.text !== undefined)
		return `<c r="${cell.ref}"${sAttr} t="s"><v>${sst.ref(cell.text)}</v></c>`;
	if (cell.bool !== undefined)
		return `<c r="${cell.ref}"${sAttr} t="b"><v>${cell.bool ? 1 : 0}</v></c>`;
	if (cell.formula !== undefined) {
		const v = cell.number !== undefined ? `<v>${cell.number}</v>` : "";
		return `<c r="${cell.ref}"${sAttr}><f>${escapeText(cell.formula)}</f>${v}</c>`;
	}
	if (cell.serial !== undefined) return `<c r="${cell.ref}"${sAttr}><v>${cell.serial}</v></c>`;
	if (cell.number !== undefined) return `<c r="${cell.ref}"${sAttr}><v>${cell.number}</v></c>`;
	return `<c r="${cell.ref}"${sAttr}/>`;
}

function sheetDataXml(cells, sst, styles, rowStyles) {
	const rows = new Map();
	for (const cell of cells) {
		const r = rowOf(cell.ref);
		if (!rows.has(r)) rows.set(r, []);
		rows.get(r).push(cellXml(cell, sst, styles));
	}
	return [...rows.keys()]
		.sort((a, b) => a - b)
		.map((r) => {
			// A row default style becomes `s` + customFormat="1" (the reader honors row `s` only
			// under customFormat). Cells that set their own `s` still override it.
			const rowStyle = rowStyles?.[r];
			const idx = rowStyle !== undefined ? styles.indexFor(rowStyle) : undefined;
			const attrs = idx !== undefined ? ` s="${idx}" customFormat="1"` : "";
			return `<row r="${r}"${attrs}>${rows.get(r).join("")}</row>`;
		})
		.join("");
}

function commentsXml(comments) {
	const authors = [];
	const authorIndex = new Map();
	const idFor = (name) => {
		let i = authorIndex.get(name);
		if (i === undefined) {
			i = authors.length;
			authors.push(name);
			authorIndex.set(name, i);
		}
		return i;
	};
	const list = comments
		.map(
			(c) =>
				`<comment ref="${c.ref}" authorId="${idFor(c.author ?? "")}"><text><t xml:space="preserve">${escapeText(
					c.text,
				)}</t></text></comment>`,
		)
		.join("");
	const authorsXml = authors.map((a) => `<author>${escapeText(a)}</author>`).join("");
	return `${XML_DECL}
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>${authorsXml}</authors><commentList>${list}</commentList></comments>`;
}

/** Build a valid .xlsx from a declarative spec. Returns the archive bytes (Uint8Array). */
export function buildWorkbook(spec) {
	const sheets = spec.sheets;
	const sst = makeSharedStrings();
	const styles = makeStyles();
	const parts = []; // worksheet + comments + worksheet-rels parts, in emit order
	const overrides = []; // Content_Types overrides for the above
	let commentsCount = 0;

	sheets.forEach((sheet, i) => {
		const num = i + 1;
		const wsRels = [];

		let hyperlinksXml = "";
		if (sheet.hyperlinks?.length) {
			const links = sheet.hyperlinks.map((h) => {
				let rid = "";
				if (h.target !== undefined) {
					const id = `rId${wsRels.length + 1}`;
					wsRels.push({
						id,
						type: `${REL}/hyperlink`,
						target: h.target,
						mode: "External",
					});
					rid = ` r:id="${id}"`;
				}
				const loc = h.location !== undefined ? ` location="${escapeAttr(h.location)}"` : "";
				const tip = h.tooltip !== undefined ? ` tooltip="${escapeAttr(h.tooltip)}"` : "";
				const disp = h.display !== undefined ? ` display="${escapeAttr(h.display)}"` : "";
				return `<hyperlink ref="${h.ref}"${loc}${tip}${disp}${rid}/>`;
			});
			hyperlinksXml = `<hyperlinks xmlns:r="${REL}">${links.join("")}</hyperlinks>`;
		}

		if (sheet.comments?.length) {
			commentsCount++;
			const path = `xl/comments${commentsCount}.xml`;
			parts.push({ name: path, xml: commentsXml(sheet.comments) });
			overrides.push({ part: `/${path}`, type: `${CT}.comments+xml` });
			wsRels.push({
				id: `rId${wsRels.length + 1}`,
				type: `${REL}/comments`,
				target: `../comments${commentsCount}.xml`,
			});
		}

		// OOXML order within <worksheet>: dimension, cols, sheetData, …, mergeCells, …, hyperlinks.
		const merges = sheet.merges?.length
			? `<mergeCells count="${sheet.merges.length}">${sheet.merges
					.map((r) => `<mergeCell ref="${r}"/>`)
					.join("")}</mergeCells>`
			: "";
		const dim = sheet.dimension ? `<dimension ref="${sheet.dimension}"/>` : "";
		const cols = sheet.columns?.length
			? `<cols>${sheet.columns
					.map((c) => {
						const idx = styles.indexFor(c);
						const style = idx !== undefined ? ` style="${idx}"` : "";
						return `<col min="${c.min}" max="${c.max}"${style}/>`;
					})
					.join("")}</cols>`
			: "";
		const data = sheetDataXml(sheet.cells ?? [], sst, styles, sheet.rowStyles);
		parts.push({
			name: `xl/worksheets/sheet${num}.xml`,
			xml: `${XML_DECL}
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${dim}${cols}<sheetData>${data}</sheetData>${merges}${hyperlinksXml}</worksheet>`,
		});
		overrides.push({ part: `/xl/worksheets/sheet${num}.xml`, type: `${CT}.worksheet+xml` });

		if (wsRels.length) {
			parts.push({ name: `xl/worksheets/_rels/sheet${num}.xml.rels`, xml: relsPart(wsRels) });
		}
	});

	const wbRels = sheets.map((_, i) => ({
		id: `rId${i + 1}`,
		type: `${REL}/worksheet`,
		target: `worksheets/sheet${i + 1}.xml`,
	}));

	// styles.xml and sharedStrings.xml are optional parts — emit each (with its rel and
	// content-type override) only when the workbook actually has styles / strings, so a minimal
	// workbook produces a minimal package.
	const optionalFiles = [];
	const optionalOverrides = [];
	let nextRid = sheets.length + 1;
	if (styles.has()) {
		wbRels.push({ id: `rId${nextRid++}`, type: `${REL}/styles`, target: "styles.xml" });
		optionalFiles.push({ name: "xl/styles.xml", xml: styles.xml() });
		optionalOverrides.push({ part: "/xl/styles.xml", type: `${CT}.styles+xml` });
	}
	if (sst.has()) {
		wbRels.push({
			id: `rId${nextRid++}`,
			type: `${REL}/sharedStrings`,
			target: "sharedStrings.xml",
		});
		optionalFiles.push({ name: "xl/sharedStrings.xml", xml: sst.xml() });
		optionalOverrides.push({ part: "/xl/sharedStrings.xml", type: `${CT}.sharedStrings+xml` });
	}

	const workbookPr = spec.date1904 ? '<workbookPr date1904="1"/>' : "";
	const sheetsXml = sheets
		.map((s, i) => {
			const state =
				s.visible === false || s.visible === "hidden"
					? ' state="hidden"'
					: s.visible === "veryHidden"
						? ' state="veryHidden"'
						: "";
			return `<sheet name="${escapeAttr(s.name)}" sheetId="${i + 1}"${state} r:id="rId${i + 1}"/>`;
		})
		.join("");

	const allOverrides = [
		{ part: "/xl/workbook.xml", type: `${CT}.sheet.main+xml` },
		...overrides,
		...optionalOverrides,
	];

	const files = [
		{
			name: "[Content_Types].xml",
			xml: `${XML_DECL}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${allOverrides
				.map((o) => `<Override PartName="${o.part}" ContentType="${o.type}"/>`)
				.join("")}</Types>`,
		},
		{
			name: "_rels/.rels",
			xml: relsPart([
				{ id: "rId1", type: `${REL}/officeDocument`, target: "xl/workbook.xml" },
			]),
		},
		{
			name: "xl/workbook.xml",
			xml: `${XML_DECL}
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}">${workbookPr}<sheets>${sheetsXml}</sheets></workbook>`,
		},
		{ name: "xl/_rels/workbook.xml.rels", xml: relsPart(wbRels) },
		...parts,
		...optionalFiles,
	];

	return zipStore(files.map((f) => ({ name: f.name, data: encoder.encode(f.xml) })));
}

/**
 * Pack arbitrary parts into a STORED .xlsx-shaped ZIP, without any workbook wiring. For crafting
 * deliberately-broken or otherwise hand-built packages in tests (a missing/malformed part, or a
 * drawing + binary media). Each part is `{ name, xml }` for a text part or `{ name, data }` with a
 * Uint8Array for a binary one (e.g. image media).
 */
export function packParts(parts) {
	return zipStore(parts.map((p) => ({ name: p.name, data: p.data ?? encoder.encode(p.xml) })));
}
