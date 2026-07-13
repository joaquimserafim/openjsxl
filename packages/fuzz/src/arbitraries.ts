// fast-check generators for writer inputs (Half A). Two families:
//   • TYPED, stable, mostly-valid trees (`plainScalarWorkbookArb`, `plainWorkbookArb`) used by the
//     determinism and round-trip properties — no casts, no adversarial shapes. These still carry
//     runtime-invalid-but-type-valid values (bad refs, empty names, malformed merges) so the reject
//     path is exercised deterministically; only ENUM fields are kept to valid members (an invalid
//     enum can't be typed, so those are fuzzed in the untyped hostile corpus below).
//   • `hostileWorkbookArb: Arbitrary<unknown>` — the same trees with poisoned VALUES and hostile
//     SHAPES (non-plain prototypes, unknown keys, transparent Proxies, wrong-typed / invalid-enum
//     fields) mixed in, used only by the resolve-or-typed-error property, which casts once.
//
// Deliberately NOT generated: getters that THROW. The writer reads caller data; a caller whose getter
// throws gets that throw back — that is outside the "reject invalid DATA with a typed error" contract,
// so asserting XlsxError on it would be an unsound false failure. Poisoned VALUES (returned by plain
// props or by transparent Proxies) ARE in contract and are covered.

import type {
	CellInput,
	CellStyle,
	Color,
	ConditionalFormatting,
	ConditionalFormattingRule,
	DataValidation,
	DataValidationErrorStyle,
	DataValidationOperator,
	DataValidationType,
	HorizontalAlignment,
	PatternType,
	SheetInput,
	SheetState,
	TableInfo,
	UnderlineStyle,
	VerticalAlignment,
	WorkbookInput,
} from "@openjsxl/core";
import fc from "fast-check";
import { colToA1 } from "./a1";

// Typed enum pools — `fc.constantFrom(...pool)` keeps the literal union (a bare `constantFrom("a")`
// widens to `Arbitrary<string>`, which won't satisfy the model's unions under strict types).
const UNDERLINES: readonly UnderlineStyle[] = ["single", "double"];
const PATTERNS: readonly PatternType[] = ["solid", "none", "gray125"];
const HALIGN: readonly HorizontalAlignment[] = ["left", "center", "right"];
const VALIGN: readonly VerticalAlignment[] = ["top", "center", "bottom"];
const DV_TYPES: readonly DataValidationType[] = [
	"none",
	"whole",
	"decimal",
	"list",
	"date",
	"time",
	"textLength",
	"custom",
];
const DV_OPS: readonly DataValidationOperator[] = [
	"between",
	"notBetween",
	"equal",
	"greaterThan",
	"lessThan",
];
const DV_ERR: readonly DataValidationErrorStyle[] = ["stop", "warning", "information"];
const STATES: readonly SheetState[] = ["visible", "hidden", "veryHidden"];
const CF_HL_TYPES = [
	"cellIs",
	"expression",
	"containsText",
	"top10",
	"duplicateValues",
	"timePeriod",
] as const;
const AUTO_COLOR: Color = { auto: true };

// ── Colors + styles ──────────────────────────────────────────────────────────────────────────

const colorArb: fc.Arbitrary<Color> = fc.oneof(
	fc
		.tuple(fc.constantFrom("FF", "00", ""), fc.integer({ min: 0, max: 0xffffff }))
		.map(([a, n]) => ({ rgb: a + n.toString(16).padStart(6, "0").toUpperCase() })),
	fc.record({
		theme: fc.integer({ min: 0, max: 9 }),
		tint: fc.double({ min: -1, max: 1, noNaN: true }),
	}),
	fc.record({ indexed: fc.integer({ min: 0, max: 65 }) }),
	fc.constant(AUTO_COLOR),
);

const cellStyleArb: fc.Arbitrary<CellStyle> = fc.record(
	{
		numberFormat: fc.constantFrom("0.00", "yyyy-mm-dd", "0%", "@", "#,##0"),
		font: fc.record(
			{
				name: fc.constantFrom("Calibri", "Arial"),
				size: fc.integer({ min: 1, max: 409 }),
				bold: fc.boolean(),
				italic: fc.boolean(),
				underline: fc.constantFrom(...UNDERLINES),
				strike: fc.boolean(),
				color: colorArb,
			},
			{ requiredKeys: [] },
		),
		fill: fc.record(
			{ patternType: fc.constantFrom(...PATTERNS), fgColor: colorArb, bgColor: colorArb },
			{ requiredKeys: ["patternType"] },
		),
		alignment: fc.record(
			{
				horizontal: fc.constantFrom(...HALIGN),
				vertical: fc.constantFrom(...VALIGN),
				wrapText: fc.boolean(),
				indent: fc.integer({ min: 0, max: 250 }),
				textRotation: fc.integer({ min: 0, max: 180 }),
			},
			{ requiredKeys: [] },
		),
	},
	{ requiredKeys: [] },
);

// ── Cell values ──────────────────────────────────────────────────────────────────────────────

// Strings including the nasty stuff the emitter must escape or reject: XML metacharacters, control
// chars (only \t \n \r are XML-safe), lone surrogates, a leading `=`.
const nastyStringArb: fc.Arbitrary<string> = fc.oneof(
	fc.string({ maxLength: 24 }),
	fc.constantFrom("<&>\"'", " ", "", "\x1b[0m", "\uD800", "\uDFFF", "\t\n\r", "="),
	fc.string({ unit: "binary", maxLength: 16 }),
);

// Numbers including the ones .xlsx cannot represent (NaN, ±Infinity) plus boundary magnitudes.
const numberArb: fc.Arbitrary<number> = fc.oneof(
	fc.double(),
	fc.constantFrom(
		0,
		-0,
		1,
		-1,
		Number.NaN,
		Number.POSITIVE_INFINITY,
		Number.NEGATIVE_INFINITY,
		1e308,
		-1e308,
		Number.MAX_SAFE_INTEGER,
	),
	fc.integer(),
);

const scalarArb: fc.Arbitrary<string | number | boolean | Date | null> = fc.oneof(
	nastyStringArb,
	numberArb,
	fc.boolean(),
	fc.oneof(fc.date(), fc.constant(new Date(Number.NaN))),
	fc.constant(null),
);

const styledCellArb: fc.Arbitrary<CellInput> = fc.record(
	{
		value: scalarArb,
		style: cellStyleArb,
		formula: fc.oneof(fc.constantFrom("SUM(A1:A2)", "1+1", "=A1", ""), nastyStringArb),
	},
	{ requiredKeys: [] },
);

const cellInputArb: fc.Arbitrary<CellInput> = fc.oneof(
	{ weight: 3, arbitrary: scalarArb },
	{ weight: 1, arbitrary: styledCellArb },
	{ weight: 1, arbitrary: fc.constant(undefined) },
);

// Scalar cell for the strict round-trip property: printable, non-space ASCII strings (exact
// round-trip), bounded numbers, booleans, null — no styles, formulas, holes, or empties to reason
// about. (Empty string and null both write an empty cell that reads back as `empty`; excluded here.)
const roundTripScalarArb: fc.Arbitrary<string | number | boolean | null> = fc.oneof(
	fc
		.string({ maxLength: 24 })
		.map((s) => s.replace(/[^\x21-\x7e]/g, ""))
		.filter((s) => s.length > 0),
	fc.integer({ min: -1_000_000, max: 1_000_000 }),
	fc.double({ min: -1e6, max: 1e6, noNaN: true, noDefaultInfinity: true }),
	fc.boolean(),
	fc.constant(null),
);

// ── M9 features: data validation, conditional formatting, tables ───────────────────────────────

const sqrefTokenArb = fc.constantFrom(
	"A1",
	"A1:A10",
	"B2:C3",
	"A1:B2",
	"D5",
	"A:A",
	"$A$1",
	"z9",
	"",
);

const dataValidationArb: fc.Arbitrary<DataValidation> = fc.record(
	{
		sqref: fc.array(sqrefTokenArb, { minLength: 0, maxLength: 3 }),
		type: fc.constantFrom(...DV_TYPES),
		operator: fc.constantFrom(...DV_OPS),
		formula1: fc.oneof(
			fc.constantFrom('"a,b,c"', "$A$1:$A$3", "10", "=B1", "TODAY()"),
			nastyStringArb,
		),
		formula2: fc.constantFrom("20", "100", ""),
		allowBlank: fc.boolean(),
		showDropDown: fc.boolean(),
		showInputMessage: fc.boolean(),
		showErrorMessage: fc.boolean(),
		errorStyle: fc.constantFrom(...DV_ERR),
		promptTitle: fc.oneof(fc.string({ maxLength: 40 }), nastyStringArb),
		prompt: fc.oneof(fc.string({ maxLength: 60 }), nastyStringArb),
		errorTitle: fc.string({ maxLength: 40 }),
		error: fc.string({ maxLength: 60 }),
	},
	{ requiredKeys: ["sqref", "type"] },
);

const cfvoArb = fc.record(
	{
		type: fc.constantFrom("num", "percent", "max", "min", "percentile", "formula", "<bad>"),
		val: fc.constantFrom("0", "50", "100", "=A1"),
		gte: fc.boolean(),
	},
	{ requiredKeys: ["type"] },
);

const cfRuleArb: fc.Arbitrary<ConditionalFormattingRule> = fc.oneof(
	fc.record(
		{
			type: fc.constantFrom(...CF_HL_TYPES),
			priority: fc.integer({ min: -2, max: 5 }),
			dxf: fc.record(
				{
					font: fc.record({ bold: fc.boolean(), color: colorArb }, { requiredKeys: [] }),
					fill: fc.record({ bgColor: colorArb }, { requiredKeys: [] }),
				},
				{ requiredKeys: [] },
			),
			operator: fc.constantFrom("greaterThan", "between", "containsText"),
			text: fc.string({ maxLength: 12 }),
			formulas: fc.array(fc.constantFrom("A1>0", "10", '"x"'), { maxLength: 3 }),
		},
		{ requiredKeys: ["type", "priority"] },
	),
	fc.record({
		type: fc.constant("colorScale" as const),
		priority: fc.integer({ min: 0, max: 5 }),
		cfvo: fc.array(cfvoArb, { minLength: 0, maxLength: 4 }),
		colors: fc.array(colorArb, { minLength: 0, maxLength: 4 }),
	}),
	fc.record({
		type: fc.constant("dataBar" as const),
		priority: fc.integer({ min: 0, max: 5 }),
		cfvo: fc.array(cfvoArb, { minLength: 0, maxLength: 3 }),
		color: colorArb,
	}),
	fc.record({
		type: fc.constant("iconSet" as const),
		priority: fc.integer({ min: 0, max: 5 }),
		iconSet: fc.constantFrom("3TrafficLights1", "5Rating", "bogusSet"),
		cfvo: fc.array(cfvoArb, { minLength: 0, maxLength: 5 }),
	}),
);

const conditionalFormattingArb: fc.Arbitrary<ConditionalFormatting> = fc.record({
	sqref: fc.array(sqrefTokenArb, { minLength: 0, maxLength: 2 }),
	rules: fc.array(cfRuleArb, { minLength: 0, maxLength: 3 }),
});

const tableArb: fc.Arbitrary<TableInfo> = fc.record(
	{
		name: fc.oneof(
			fc.constantFrom("Table1", "Sales", "my table", "A1", ""),
			fc.string({ maxLength: 20 }),
		),
		ref: fc.constantFrom("A1:C5", "A1:A1", "B2:D4", "ZZ1:ZZ9", "notaref", "A1"),
		columns: fc.array(
			fc.record({ name: fc.string({ maxLength: 12 }) }, { requiredKeys: ["name"] }),
			{ minLength: 0, maxLength: 4 },
		),
		headerRow: fc.boolean(),
		totalsRow: fc.boolean(),
	},
	{ requiredKeys: ["name", "ref", "columns", "headerRow", "totalsRow"] },
);

// ── Sheets + workbooks ─────────────────────────────────────────────────────────────────────────

const validNameArb: fc.Arbitrary<string> = fc.integer({ min: 1, max: 9999 }).map((n) => `S${n}`);
const hostileNameArb: fc.Arbitrary<string> = fc.oneof(
	validNameArb,
	fc.constantFrom("", "a".repeat(40), "bad/name", "x[1]", "a:b", "*"),
);

const rowArb: fc.Arbitrary<readonly CellInput[]> = fc.array(cellInputArb, {
	minLength: 0,
	maxLength: 5,
});

function sheetArb(nameArb: fc.Arbitrary<string>): fc.Arbitrary<SheetInput> {
	return fc.record(
		{
			name: nameArb,
			rows: fc.array(
				fc.oneof(
					{ weight: 5, arbitrary: rowArb },
					{ weight: 1, arbitrary: fc.constant(undefined) },
				),
				{ minLength: 0, maxLength: 6 },
			),
			columns: fc.array(
				fc.record(
					{
						min: fc.integer({ min: 1, max: 20 }),
						max: fc.integer({ min: 1, max: 20 }),
						width: fc.double({ min: 0, max: 300, noNaN: true }),
						hidden: fc.boolean(),
					},
					{ requiredKeys: ["min", "max"] },
				),
				{ maxLength: 3 },
			),
			freeze: fc.record(
				{ rows: fc.integer({ min: 0, max: 5 }), cols: fc.integer({ min: 0, max: 5 }) },
				{ requiredKeys: [] },
			),
			merges: fc.array(
				fc.constantFrom("A1:B2", "A1:A3", "B2:C3", "A1", "A1:A1", "bad", "C3:B2"),
				{ maxLength: 3 },
			),
			hyperlinks: fc.array(
				fc.record(
					{
						ref: fc.constantFrom("A1", "B2:C3"),
						target: fc.webUrl(),
						location: fc.constantFrom("Sheet1!A1", ""),
						tooltip: fc.string({ maxLength: 10 }),
					},
					{ requiredKeys: ["ref"] },
				),
				{ maxLength: 2 },
			),
			state: fc.constantFrom(...STATES),
			dataValidations: fc.array(dataValidationArb, { maxLength: 2 }),
			conditionalFormatting: fc.array(conditionalFormattingArb, { maxLength: 2 }),
			tables: fc.array(tableArb, { maxLength: 2 }),
		},
		{ requiredKeys: ["name", "rows"] },
	);
}

/** Strict round-trip corpus: one sheet of scalar cells, valid name — writeXlsx must round-trip values. */
export const plainScalarWorkbookArb: fc.Arbitrary<WorkbookInput> = fc
	.record({
		name: validNameArb,
		rows: fc.array(fc.array(roundTripScalarArb, { maxLength: 5 }), {
			minLength: 1,
			maxLength: 6,
		}),
	})
	.map((sheet) => ({ sheets: [sheet] }));

// ── Guaranteed-valid corpus (reliably WRITES) ───────────────────────────────────────────────────
// The hostile/plain corpora pack a runtime-invalid value into ~99.8% of workbooks, so writeXlsx
// rejects them BEFORE producing bytes — which makes the determinism and resolve/re-read arms nearly
// vacuous (a finding from the F9.4 harness soundness review). This corpus is constructed to ALWAYS
// write: unique index-based names, finite bounded numbers, XML-safe strings, writer-legal styles, and
// VALID M9 blocks (a derive-from-header table, a whole-number DV, a cellIs CF). It powers the
// determinism + resolve/re-read properties so their headline claims are actually exercised — and, as a
// bonus, asserts the writer never REJECTS legitimately-valid input.

const safeStringArb: fc.Arbitrary<string> = fc.string({ maxLength: 12 }).map((s) => {
	const t = s.replace(/[^\x20-\x7e]/g, "");
	return t.length > 0 ? t : "x";
});
const safeNumberArb: fc.Arbitrary<number> = fc.oneof(
	fc.integer({ min: -1_000_000, max: 1_000_000 }),
	fc.double({ min: -1e6, max: 1e6, noNaN: true, noDefaultInfinity: true }),
);
const safeScalarArb: fc.Arbitrary<string | number | boolean | null> = fc.oneof(
	safeStringArb,
	safeNumberArb,
	fc.boolean(),
	fc.constant(null),
);
const safeColorArb: fc.Arbitrary<Color> = fc.oneof(
	fc
		.integer({ min: 0, max: 0xffffff })
		.map((n) => ({ rgb: `FF${n.toString(16).padStart(6, "0").toUpperCase()}` })),
	fc.record({ theme: fc.integer({ min: 0, max: 9 }) }),
	fc.record({ indexed: fc.integer({ min: 0, max: 63 }) }),
	fc.constant(AUTO_COLOR),
);
const safeStyleArb: fc.Arbitrary<CellStyle> = fc.record(
	{
		numberFormat: fc.constantFrom("0.00", "0%", "@"),
		font: fc.record(
			{
				bold: fc.boolean(),
				italic: fc.boolean(),
				size: fc.integer({ min: 6, max: 72 }),
				color: safeColorArb,
			},
			{ requiredKeys: [] },
		),
		fill: fc.record(
			{ patternType: fc.constant<PatternType>("solid"), fgColor: safeColorArb },
			{ requiredKeys: ["patternType"] },
		),
		alignment: fc.record(
			{
				horizontal: fc.constantFrom(...HALIGN),
				wrapText: fc.boolean(),
				indent: fc.integer({ min: 0, max: 10 }),
			},
			{ requiredKeys: [] },
		),
	},
	{ requiredKeys: [] },
);
const safeCellArb: fc.Arbitrary<CellInput> = fc.oneof(
	{ weight: 3, arbitrary: safeScalarArb },
	{
		weight: 1,
		arbitrary: fc.record(
			{ value: safeScalarArb, style: safeStyleArb },
			{ requiredKeys: ["value"] },
		),
	},
);

interface ValidSpec {
	readonly width: number;
	readonly data: readonly (readonly CellInput[])[];
	readonly table: boolean;
	readonly dv: boolean;
	readonly cf: boolean;
}
const validSpecArb: fc.Arbitrary<ValidSpec> = fc.record({
	width: fc.integer({ min: 1, max: 4 }),
	data: fc.array(fc.array(safeCellArb, { minLength: 0, maxLength: 4 }), {
		minLength: 1,
		maxLength: 4,
	}),
	table: fc.boolean(),
	dv: fc.boolean(),
	cf: fc.boolean(),
});

function toValidSheet(spec: ValidSpec, i: number): SheetInput {
	// Row 0 is a header of DISTINCT non-empty text (H0…H{width-1}) so a derive-from-header table has
	// unique, non-empty column names.
	const header: string[] = [];
	for (let c = 0; c < spec.width; c++) header.push(`H${c}`);
	const rows: readonly (readonly CellInput[])[] = [header, ...spec.data];
	const sheet: {
		name: string;
		rows: readonly (readonly CellInput[])[];
		tables?: readonly TableInfo[];
		dataValidations?: readonly DataValidation[];
		conditionalFormatting?: readonly ConditionalFormatting[];
	} = { name: `S${i}`, rows };
	if (spec.table) {
		// `Table_${i}` is a legal identifier (not cell-ref-shaped, workbook-unique); empty `columns`
		// means "derive from the header row"; ref spans header + all data rows.
		sheet.tables = [
			{
				name: `Table_${i}`,
				ref: `A1:${colToA1(spec.width)}${rows.length}`,
				columns: [],
				headerRow: true,
				totalsRow: false,
			},
		];
	}
	if (spec.dv)
		sheet.dataValidations = [
			{ sqref: ["A2"], type: "whole", operator: "greaterThan", formula1: "0" },
		];
	if (spec.cf) {
		sheet.conditionalFormatting = [
			{
				sqref: ["A2:A5"],
				rules: [
					{
						type: "cellIs",
						priority: 1,
						operator: "greaterThan",
						formulas: ["0"],
						dxf: { fill: { bgColor: { rgb: "FFFFC7CE" } } },
					},
				],
			},
		];
	}
	return sheet;
}

/** Guaranteed-valid corpus: reliably writes (see above). Names/table-names are index-unique. */
export const validWorkbookArb: fc.Arbitrary<WorkbookInput> = fc
	.array(validSpecArb, { minLength: 1, maxLength: 3 })
	.map((specs) => ({ sheets: specs.map(toValidSheet) }));

// ── Hostile shape injectors (applied only in the resolve-or-throw corpus) ──────────────────────

function withUnknownKey(obj: object): Record<string, unknown> {
	return { ...obj, __openjsxl_evil__: 1 };
}
function withBadProto(obj: object): object {
	return Object.assign(Object.create({ poisoned: true }), obj);
}
function asTransparentProxy(obj: object): object {
	return new Proxy(obj, {});
}
const WRAPS: readonly ((o: object) => unknown)[] = [
	(o) => o,
	withUnknownKey,
	withBadProto,
	asTransparentProxy,
];

// Randomly wrap a generated value in one hostile shape (or leave it be).
function poison(arb: fc.Arbitrary<object>): fc.Arbitrary<unknown> {
	return arb.chain((value) => fc.constantFrom(...WRAPS).map((wrap) => wrap(value)));
}

// A data-validation / CF / table block with an INVALID enum member — typed `unknown` so the malformed
// literal is legal here (it can't live in the strict corpus). Attached onto an otherwise-valid sheet.
const invalidEnumSheetArb: fc.Arbitrary<unknown> = sheetArb(hostileNameArb).map((s) => ({
	...s,
	dataValidations: [{ sqref: ["A1"], type: "bogus", operator: "??" }],
	conditionalFormatting: [{ sqref: ["A1"], rules: [{ type: "notAType", priority: 1 }] }],
}));

/**
 * Resolve-or-typed-error corpus (`Arbitrary<unknown>`): hostile names, poisoned scalar values (NaN,
 * ∞, control chars, invalid Dates), malformed M9 blocks (incl. invalid enums), plus a chance the
 * workbook / a sheet / a cell carries a non-plain prototype, an unknown key, or a Proxy. The property
 * casts once at its boundary.
 */
export const hostileWorkbookArb: fc.Arbitrary<unknown> = fc.oneof(
	poison(
		fc
			.array(sheetArb(hostileNameArb), { minLength: 0, maxLength: 3 })
			.map((sheets) => ({ sheets })),
	),
	poison(sheetArb(hostileNameArb)).map((sheet) => ({ sheets: [sheet] })),
	invalidEnumSheetArb.map((sheet) => ({ sheets: [sheet] })),
	fc.record({
		sheets: fc.oneof(
			fc.constant(null),
			fc.integer(),
			fc.string(),
			fc.constant([]),
			fc.array(fc.oneof(fc.constant(null), fc.integer(), fc.string()), { maxLength: 3 }),
		),
	}),
	fc.oneof(fc.constant(null), fc.constant(undefined), fc.integer(), fc.string(), fc.boolean()),
);
