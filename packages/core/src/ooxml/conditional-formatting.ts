import type {
	CfHighlightType,
	Cfvo,
	Color,
	ConditionalFormatting,
	ConditionalFormattingRule,
	DxfStyle,
} from "../types";
import { isXmlSafe, localName } from "../utils";
import { tokenize } from "../xml";
import { isCanonicalSqrefToken, MAX_SQREF_RANGES } from "./data-validation";
import { parseColor } from "./styles";

// Conditional-formatting parser (F9.3). Reads the worksheet's MAIN-part `<conditionalFormatting>`
// blocks into the shared model, resolving each `<cfRule @dxfId>` to an inline {@link DxfStyle} against
// the workbook `<dxfs>` table (decision 3 — the numeric id never surfaces). TOLERANT, and every value
// it returns is one the writer accepts (shared bounds, the F9.2 lesson):
//   - sqref tokens are filtered to canonical A1 (shared `isCanonicalSqrefToken`), count-capped;
//   - `<formula>` text is taken in stored form (leading `=` stripped) and dropped if empty/XML-unsafe;
//   - a `cfvo`'s `val` and a rule's `text` are dropped if XML-unsafe;
//   - an unknown `type`, or a rule with no priority, is handled without throwing.
//
// x14 lives under `<extLst>` — worksheet-level `x14:conditionalFormattings` AND the cfRule-level
// `<extLst>` x14:id twin — and is SKIPPED entirely (decision 4): a dangling GUID must not round-trip.

const HIGHLIGHT_TYPES: ReadonlySet<string> = new Set<CfHighlightType>([
	"cellIs",
	"expression",
	"top10",
	"aboveAverage",
	"uniqueValues",
	"duplicateValues",
	"containsText",
	"notContainsText",
	"beginsWith",
	"endsWith",
	"containsBlanks",
	"notContainsBlanks",
	"containsErrors",
	"notContainsErrors",
	"timePeriod",
]);
// Type-predicate guard (no `as`): the set holds exactly the CfHighlightType members.
function isHighlightType(v: string): v is CfHighlightType {
	return HIGHLIGHT_TYPES.has(v);
}

const bool01 = (v: string | undefined): boolean | undefined => {
	if (v === "1" || v === "true") return true;
	if (v === "0" || v === "false") return false;
	return undefined;
};

// A formula operand in stored form (leading `=` stripped), or undefined when empty / XML-unsafe.
const storedFormula = (text: string): string | undefined => {
	const s = text.startsWith("=") ? text.slice(1) : text;
	return s !== "" && isXmlSafe(s) ? s : undefined;
};

const xmlSafeAttr = (v: string | undefined): string | undefined =>
	v !== undefined && isXmlSafe(v) ? v : undefined;

// Split @sqref into canonical tokens, capped (shared with data validation).
function parseSqref(raw: string | undefined): string[] {
	if (raw === undefined) return [];
	const out: string[] = [];
	let processed = 0;
	for (const token of raw.split(/\s+/)) {
		if (token === "") continue;
		if (++processed > MAX_SQREF_RANGES) break;
		if (isCanonicalSqrefToken(token)) out.push(token);
	}
	return out;
}

// The mutable rule under construction. `sub` collects a colorScale/dataBar/iconSet child's cfvo/colors.
interface RuleBuilder {
	type: string;
	priority?: number;
	dxfId?: number;
	stopIfTrue?: boolean;
	operator?: string;
	text?: string;
	timePeriod?: string;
	rank?: number;
	percent?: boolean;
	bottom?: boolean;
	aboveAverage?: boolean;
	equalAverage?: boolean;
	stdDev?: number;
	formulas: string[];
	iconSetName?: string;
	cfvo: Cfvo[];
	colors: Color[];
}

function intAttr(v: string | undefined): number | undefined {
	// Canonical (optionally-signed) decimal only — `Number("")`/`Number("1e2")`/`Number("0x1")` would
	// otherwise coerce an empty/odd attribute into a phantom priority/rank/dxfId (F9.3 review).
	if (v === undefined || !/^-?[0-9]+$/.test(v)) return undefined;
	const n = Number(v);
	return Number.isInteger(n) ? n : undefined;
}

function startRule(attrs: Readonly<Record<string, string>>): RuleBuilder {
	const b: RuleBuilder = { type: attrs.type ?? "", formulas: [], cfvo: [], colors: [] };
	const priority = intAttr(attrs.priority);
	if (priority !== undefined && priority >= 1) b.priority = priority;
	const dxfId = intAttr(attrs.dxfId);
	if (dxfId !== undefined && dxfId >= 0) b.dxfId = dxfId;
	const stop = bool01(attrs.stopIfTrue);
	if (stop !== undefined) b.stopIfTrue = stop;
	const op = xmlSafeAttr(attrs.operator);
	if (op !== undefined) b.operator = op;
	const text = xmlSafeAttr(attrs.text);
	if (text !== undefined) b.text = text;
	const tp = xmlSafeAttr(attrs.timePeriod);
	if (tp !== undefined) b.timePeriod = tp;
	const rank = intAttr(attrs.rank);
	if (rank !== undefined) b.rank = rank;
	const percent = bool01(attrs.percent);
	if (percent !== undefined) b.percent = percent;
	const bottom = bool01(attrs.bottom);
	if (bottom !== undefined) b.bottom = bottom;
	const above = bool01(attrs.aboveAverage);
	if (above !== undefined) b.aboveAverage = above;
	const equal = bool01(attrs.equalAverage);
	if (equal !== undefined) b.equalAverage = equal;
	const stdDev = intAttr(attrs.stdDev);
	if (stdDev !== undefined) b.stdDev = stdDev;
	return b;
}

function cfvoFrom(attrs: Readonly<Record<string, string>>): Cfvo {
	// `type` is gated for XML-safety like every other reader-returned CF string — an unsafe (or
	// absent) value degrades to "num", so the reader never yields a cfvo the writer would reject.
	const cfvo: { type: string; val?: string; gte?: boolean } = {
		type: xmlSafeAttr(attrs.type) ?? "num",
	};
	const val = xmlSafeAttr(attrs.val);
	if (val !== undefined) cfvo.val = val;
	const gte = bool01(attrs.gte);
	if (gte !== undefined) cfvo.gte = gte;
	return cfvo;
}

/**
 * The most `<formula>` operands one cfRule can carry (CT_CfRule: `maxOccurs="3"` — e.g. `between`
 * needs 2). Single-sourced (F9.6): the tolerant reader IGNORES the 4th+ formula of a malformed rule,
 * the strict writer REJECTS one, so neither side emits a schema-invalid rule Excel would repair.
 */
export const MAX_CF_FORMULAS = 3;

// Shared cfvo/color count bounds (decision 5) — the reader DROPS a graphical rule that violates them,
// the writer REJECTS one, so neither side can produce a schema-invalid <colorScale>/<dataBar>/<iconSet>
// that Excel would repair-prompt on.
/** Icon count for a built-in ST_IconSetType (3TrafficLights1 → 3); default set is 3; unknown → undefined. */
export function iconSetCount(name: string | undefined): number | undefined {
	if (name === undefined) return 3;
	const m = /^([345])/.exec(name);
	return m ? Number(m[1]) : undefined;
}
/** A color scale needs 2–3 cfvos and an equal number of colors. */
export function colorScaleCountsOk(cfvoLen: number, colorLen: number): boolean {
	return (cfvoLen === 2 || cfvoLen === 3) && cfvoLen === colorLen;
}
/** A data bar needs exactly 2 cfvos. */
export function dataBarCountsOk(cfvoLen: number): boolean {
	return cfvoLen === 2;
}
/** An icon set needs one cfvo per icon (from its name; an unknown name accepts 3–5). */
export function iconSetCountsOk(name: string | undefined, cfvoLen: number): boolean {
	const expected = iconSetCount(name);
	return expected !== undefined ? cfvoLen === expected : cfvoLen >= 3 && cfvoLen <= 5;
}

// Build the discriminated rule, resolving the dxf. Returns undefined for an unrepresentable rule
// (unknown type, or a scale/bar/icon set missing its child data).
function buildRule(
	b: RuleBuilder,
	dxfs: readonly DxfStyle[] | undefined,
	fallbackPriority: number,
): ConditionalFormattingRule | undefined {
	const priority = b.priority ?? fallbackPriority;
	const base = { priority, ...(b.stopIfTrue !== undefined ? { stopIfTrue: b.stopIfTrue } : {}) };

	if (b.type === "colorScale") {
		// Out-of-count (schema-invalid) scales are DROPPED — the writer rejects them (shared bound).
		if (!colorScaleCountsOk(b.cfvo.length, b.colors.length)) return undefined;
		return { type: "colorScale", ...base, cfvo: b.cfvo, colors: b.colors };
	}
	if (b.type === "dataBar") {
		const color = b.colors[0];
		if (!dataBarCountsOk(b.cfvo.length) || color === undefined) return undefined;
		return { type: "dataBar", ...base, cfvo: b.cfvo, color };
	}
	if (b.type === "iconSet") {
		if (!iconSetCountsOk(b.iconSetName, b.cfvo.length)) return undefined;
		return {
			type: "iconSet",
			...base,
			...(b.iconSetName !== undefined ? { iconSet: b.iconSetName } : {}),
			cfvo: b.cfvo,
		};
	}
	const t = b.type;
	if (isHighlightType(t)) {
		const dxf = b.dxfId !== undefined ? dxfs?.[b.dxfId] : undefined;
		return {
			type: t,
			...base,
			...(dxf !== undefined ? { dxf } : {}),
			...(b.operator !== undefined ? { operator: b.operator } : {}),
			...(b.text !== undefined ? { text: b.text } : {}),
			...(b.timePeriod !== undefined ? { timePeriod: b.timePeriod } : {}),
			...(b.rank !== undefined ? { rank: b.rank } : {}),
			...(b.percent !== undefined ? { percent: b.percent } : {}),
			...(b.bottom !== undefined ? { bottom: b.bottom } : {}),
			...(b.aboveAverage !== undefined ? { aboveAverage: b.aboveAverage } : {}),
			...(b.equalAverage !== undefined ? { equalAverage: b.equalAverage } : {}),
			...(b.stdDev !== undefined ? { stdDev: b.stdDev } : {}),
			...(b.formulas.length > 0 ? { formulas: b.formulas } : {}),
		};
	}
	return undefined; // unknown ST_CfType — degrade (drop)
}

/**
 * Parse a worksheet's main `<conditionalFormatting>` blocks, resolving `dxfId`s against `dxfs`. x14
 * extensions under `<extLst>` are skipped. Never throws.
 */
export function parseConditionalFormatting(
	xml: string,
	dxfs: readonly DxfStyle[] | undefined,
): ConditionalFormatting[] {
	const out: ConditionalFormatting[] = [];
	let extLstDepth = 0;
	let block: { sqref: string[]; rules: ConditionalFormattingRule[] } | undefined;
	let ruleCount = 0; // for a fallback priority within the sheet
	let rule: RuleBuilder | undefined;
	let sub: "colorScale" | "dataBar" | "iconSet" | undefined; // open scale/bar/icon child
	let inFormula = false;
	let formulaText = "";

	for (const token of tokenize(xml)) {
		if (token.kind === "text") {
			if (inFormula) formulaText += token.value;
			continue;
		}
		const name = localName(token.name);

		if (token.kind === "open") {
			if (name === "extLst") {
				if (!token.selfClosing) extLstDepth++;
				continue;
			}
			if (extLstDepth > 0) continue;
			if (name === "conditionalFormatting") {
				block = { sqref: parseSqref(token.attrs.sqref), rules: [] };
				continue;
			}
			if (block === undefined) continue;
			if (name === "cfRule") {
				rule = startRule(token.attrs);
				sub = undefined;
				inFormula = false;
				if (token.selfClosing) {
					const built = buildRule(rule, dxfs, ++ruleCount);
					if (built !== undefined) block.rules.push(built);
					rule = undefined;
				}
				continue;
			}
			if (rule === undefined) continue;
			if (name === "colorScale" || name === "dataBar" || name === "iconSet") {
				sub = name;
				if (name === "iconSet") {
					const iset = xmlSafeAttr(token.attrs.iconSet);
					if (iset !== undefined) rule.iconSetName = iset;
				}
			} else if (name === "cfvo" && sub !== undefined) {
				rule.cfvo.push(cfvoFrom(token.attrs));
			} else if (name === "color" && sub !== undefined) {
				const color = parseColor(token.attrs);
				if (color !== undefined) rule.colors.push(color);
			} else if (name === "formula") {
				inFormula = !token.selfClosing;
				formulaText = "";
			}
		} else if (token.kind === "close") {
			if (name === "extLst") {
				if (extLstDepth > 0) extLstDepth--;
				continue;
			}
			if (extLstDepth > 0) continue;
			if (name === "conditionalFormatting") {
				// A block with no covered range or no rule is meaningless — drop it.
				if (block !== undefined && block.sqref.length > 0 && block.rules.length > 0) {
					out.push({ sqref: block.sqref, rules: block.rules });
				}
				block = undefined;
				continue;
			}
			if (rule === undefined) continue;
			if (name === "formula") {
				if (inFormula) {
					const f = storedFormula(formulaText);
					// The 4th+ formula of a malformed rule is ignored (shared MAX_CF_FORMULAS bound,
					// F9.6) — the writer rejects a rule carrying more, so a tolerated file still writes.
					if (f !== undefined && rule.formulas.length < MAX_CF_FORMULAS)
						rule.formulas.push(f);
				}
				inFormula = false;
			} else if (name === "colorScale" || name === "dataBar" || name === "iconSet") {
				sub = undefined;
			} else if (name === "cfRule" && block !== undefined) {
				const built = buildRule(rule, dxfs, ++ruleCount);
				if (built !== undefined) block.rules.push(built);
				rule = undefined;
			}
		}
	}
	return out;
}
