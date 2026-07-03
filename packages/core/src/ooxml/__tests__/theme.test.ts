import { describe, expect, it } from "vitest";
import { parseTheme, resolveColor, resolveTint } from "../theme";

// F5.3 — theme parse + Excel tint resolution. The tint algorithm is Excel's own Win32 integer HLS
// (HLSMAX=240); the vectors below were produced by an INDEPENDENT reference implementation of that
// algorithm, and the grayscale rows plus several accent rows (558ED5, 96B4D8, DA9694, A9D08E,
// F4B084) are the exact swatch values Excel shows for those theme colors — so this pins both the
// implementation-agreement and the real-Excel-match.

// [base 6-hex, tint, expected 6-hex]
const VECTORS: [string, number, string][] = [
	["FFFFFF", 0.0, "FFFFFF"],
	["FFFFFF", 0.3999755851924192, "FFFFFF"],
	["FFFFFF", 0.5999938962981048, "FFFFFF"],
	["FFFFFF", 0.7999816888943144, "FFFFFF"],
	["FFFFFF", -0.249977111117893, "BFBFBF"],
	["FFFFFF", -0.4999237898604048, "808080"],
	["FFFFFF", 0.1, "FFFFFF"],
	["FFFFFF", -0.1, "E6E6E6"],
	["000000", 0.0, "000000"],
	["000000", 0.3999755851924192, "666666"],
	["000000", 0.5999938962981048, "999999"],
	["000000", 0.7999816888943144, "CCCCCC"],
	["000000", -0.249977111117893, "000000"],
	["000000", -0.4999237898604048, "000000"],
	["000000", 0.1, "1A1A1A"],
	["000000", -0.1, "000000"],
	["4F81BD", 0.0, "4F81BD"],
	["4F81BD", 0.3999755851924192, "96B4D8"],
	["4F81BD", 0.5999938962981048, "B8CCE4"],
	["4F81BD", 0.7999816888943144, "DCE6F1"],
	["4F81BD", -0.249977111117893, "376193"],
	["4F81BD", -0.4999237898604048, "244062"],
	["4F81BD", 0.1, "608CC4"],
	["4F81BD", -0.1, "4273AE"],
	["C0504D", 0.0, "C0504D"],
	["C0504D", 0.3999755851924192, "DA9694"],
	["C0504D", 0.5999938962981048, "E7B9B8"],
	["C0504D", 0.7999816888943144, "F2DCDB"],
	["C0504D", -0.249977111117893, "963634"],
	["C0504D", -0.4999237898604048, "652523"],
	["C0504D", 0.1, "C7615F"],
	["C0504D", -0.1, "B4423F"],
	["1F497D", 0.0, "1F497D"],
	["1F497D", 0.3999755851924192, "558ED5"],
	["1F497D", 0.5999938962981048, "8DB4E2"],
	["1F497D", 0.7999816888943144, "C7DAF1"],
	["1F497D", -0.249977111117893, "17375E"],
	["1F497D", -0.4999237898604048, "10253F"],
	["1F497D", 0.1, "265A99"],
	["1F497D", -0.1, "1C4271"],
	["EEECE1", 0.0, "EEECE1"],
	["EEECE1", 0.3999755851924192, "F5F4ED"],
	["EEECE1", 0.5999938962981048, "F8F7F3"],
	["EEECE1", 0.7999816888943144, "FCFCFA"],
	["EEECE1", -0.249977111117893, "C5BE98"],
	["EEECE1", -0.4999237898604048, "948A54"],
	["EEECE1", 0.1, "EFEEE4"],
	["EEECE1", -0.1, "DDD9C4"],
	["FF0000", 0.0, "FF0000"],
	["FF0000", 0.3999755851924192, "FF6666"],
	["FF0000", 0.5999938962981048, "FF9999"],
	["FF0000", 0.7999816888943144, "FFCCCC"],
	["FF0000", -0.249977111117893, "BF0000"],
	["FF0000", -0.4999237898604048, "800000"],
	["FF0000", 0.1, "FF1A1A"],
	["FF0000", -0.1, "E60000"],
	["00B050", 0.0, "00B050"],
	["00B050", 0.3999755851924192, "37FF92"],
	["00B050", 0.5999938962981048, "79FFB6"],
	["00B050", 0.7999816888943144, "BDFFDB"],
	["00B050", -0.249977111117893, "00843C"],
	["00B050", -0.4999237898604048, "005928"],
	["00B050", 0.1, "00D25F"],
	["00B050", -0.1, "009F48"],
	["44546A", 0.0, "44546A"],
	["44546A", 0.3999755851924192, "8497B0"],
	["44546A", 0.5999938962981048, "ADBACB"],
	["44546A", 0.7999816888943144, "D6DCE4"],
	["44546A", -0.249977111117893, "344050"],
	["44546A", -0.4999237898604048, "222B35"],
	["44546A", 0.1, "52657E"],
	["44546A", -0.1, "3E4D60"],
	["ED7D31", 0.0, "ED7D31"],
	["ED7D31", 0.3999755851924192, "F4B084"],
	["ED7D31", 0.5999938962981048, "F8CBAD"],
	["ED7D31", 0.7999816888943144, "FCE4D6"],
	["ED7D31", -0.249977111117893, "C65911"],
	["ED7D31", -0.4999237898604048, "853D0C"],
	["ED7D31", 0.1, "EF8A47"],
	["ED7D31", -0.1, "EB6C18"],
	["4472C4", 0.0, "4472C4"],
	["4472C4", 0.3999755851924192, "8EA9DB"],
	["4472C4", 0.5999938962981048, "B5C7E8"],
	["4472C4", 0.7999816888943144, "DAE2F3"],
	["4472C4", -0.249977111117893, "305496"],
	["4472C4", -0.4999237898604048, "203764"],
	["4472C4", 0.1, "5780CA"],
	["4472C4", -0.1, "3965B5"],
	["70AD47", 0.0, "70AD47"],
	["70AD47", 0.3999755851924192, "A9D08E"],
	["70AD47", 0.5999938962981048, "C6E0B4"],
	["70AD47", 0.7999816888943144, "E2EFDA"],
	["70AD47", -0.249977111117893, "548235"],
	["70AD47", -0.4999237898604048, "385724"],
	["70AD47", 0.1, "7EBA56"],
	["70AD47", -0.1, "659C41"],
];

describe("resolveTint — Excel integer-HLS tint algorithm", () => {
	it("matches every independent reference vector", () => {
		for (const [base, tint, expected] of VECTORS) {
			expect(resolveTint(base, tint), `${base} @ ${tint}`).toBe(expected);
		}
	});

	it("is the identity for tint 0 even where the HLS round trip is lossy", () => {
		// 4472C4 does not survive an integer HLS round trip, so tint 0 must short-circuit.
		expect(resolveTint("4472C4", 0)).toBe("4472C4");
		expect(resolveTint("abcdef", 0)).toBe("ABCDEF"); // also uppercases
	});

	it("clamps an out-of-range or non-finite tint instead of going out of gamut", () => {
		expect(resolveTint("4F81BD", 5)).toBe(resolveTint("4F81BD", 1)); // >1 clamps to +1
		expect(resolveTint("4F81BD", -5)).toBe(resolveTint("4F81BD", -1)); // <-1 clamps to -1
		expect(resolveTint("4F81BD", Number.NaN)).toBe("4F81BD"); // non-finite → identity
		expect(resolveTint("4F81BD", Number.POSITIVE_INFINITY)).toBe("4F81BD");
	});

	it("matches Excel's documented grayscale swatches exactly", () => {
		// Black lighter / white darker — unambiguous middle grays only the integer algorithm nails.
		expect(resolveTint("000000", 0.499984740745262)).toBe("808080");
		expect(resolveTint("000000", 0.3499862666707358)).toBe("595959");
		expect(resolveTint("FFFFFF", -0.049989318502332)).toBe("F2F2F2");
		expect(resolveTint("FFFFFF", -0.4999237898604048)).toBe("808080");
	});
});

describe("parseTheme — <a:clrScheme>", () => {
	// Distinct colors in the four dark/light slots so the index swap is observable.
	const scheme = (dk1sys = false) =>
		`<a:theme xmlns:a="x"><a:themeElements><a:clrScheme name="t">` +
		(dk1sys
			? '<a:dk1><a:sysClr val="windowText" lastClr="0A0B0C"/></a:dk1>'
			: '<a:dk1><a:srgbClr val="111111"/></a:dk1>') +
		'<a:lt1><a:srgbClr val="222222"/></a:lt1>' +
		'<a:dk2><a:srgbClr val="333333"/></a:dk2>' +
		'<a:lt2><a:srgbClr val="444444"/></a:lt2>' +
		'<a:accent1><a:srgbClr val="aa0001"/></a:accent1>' +
		'<a:accent2><a:srgbClr val="aa0002"/></a:accent2>' +
		'<a:accent3><a:srgbClr val="aa0003"/></a:accent3>' +
		'<a:accent4><a:srgbClr val="aa0004"/></a:accent4>' +
		'<a:accent5><a:srgbClr val="aa0005"/></a:accent5>' +
		'<a:accent6><a:srgbClr val="aa0006"/></a:accent6>' +
		'<a:hlink><a:srgbClr val="aa0007"/></a:hlink>' +
		'<a:folHlink><a:srgbClr val="aa0008"/></a:folHlink>' +
		// a fmtScheme srgbClr that must NOT be mistaken for a scheme color
		'</a:clrScheme><a:fmtScheme><a:solidFill><a:srgbClr val="deadbe"/></a:solidFill></a:fmtScheme>' +
		"</a:themeElements></a:theme>";

	it("applies the dark/light index swap (0→lt1, 1→dk1, 2→lt2, 3→dk2)", () => {
		const t = parseTheme(scheme());
		expect(t).toBeDefined();
		expect(t?.[0]).toBe("222222"); // theme 0 = Background 1 = lt1
		expect(t?.[1]).toBe("111111"); // theme 1 = Text 1 = dk1
		expect(t?.[2]).toBe("444444"); // theme 2 = Background 2 = lt2
		expect(t?.[3]).toBe("333333"); // theme 3 = Text 2 = dk2
		expect(t?.[4]).toBe("AA0001"); // theme 4 = accent1 (uppercased)
		expect(t?.[11]).toBe("AA0008"); // theme 11 = folHlink
	});

	it("reads <a:sysClr> via lastClr", () => {
		expect(parseTheme(scheme(true))?.[1]).toBe("0A0B0C"); // dk1 sysClr lastClr
	});

	it("falls back to the standard system color when sysClr has no lastClr", () => {
		const s =
			'<a:clrScheme><a:dk1><a:sysClr val="windowText"/></a:dk1>' +
			'<a:lt1><a:sysClr val="window"/></a:lt1>' +
			'<a:dk2><a:srgbClr val="333333"/></a:dk2><a:lt2><a:srgbClr val="444444"/></a:lt2>' +
			[1, 2, 3, 4, 5, 6]
				.map((n) => `<a:accent${n}><a:srgbClr val="00000${n}"/></a:accent${n}>`)
				.join("") +
			'<a:hlink><a:srgbClr val="000007"/></a:hlink><a:folHlink><a:srgbClr val="000008"/></a:folHlink></a:clrScheme>';
		const t = parseTheme(s);
		expect(t?.[1]).toBe("000000"); // dk1 = windowText → black
		expect(t?.[0]).toBe("FFFFFF"); // lt1 = window → white
	});

	it("returns undefined when a slot is missing or the scheme is absent", () => {
		expect(parseTheme("<a:theme><a:themeElements/></a:theme>")).toBeUndefined();
		const missingAccent6 = scheme().replace(
			'<a:accent6><a:srgbClr val="aa0006"/></a:accent6>',
			"",
		);
		expect(parseTheme(missingAccent6)).toBeUndefined();
	});
});

describe("resolveColor", () => {
	const theme = parseTheme(
		'<a:clrScheme><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
			'<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>' +
			'<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>' +
			[2, 3, 4, 5, 6]
				.map((n) => `<a:accent${n}><a:srgbClr val="00000${n}"/></a:accent${n}>`)
				.join("") +
			'<a:hlink><a:srgbClr val="000007"/></a:hlink><a:folHlink><a:srgbClr val="000008"/></a:folHlink></a:clrScheme>',
	);

	it("normalizes an rgb color to 8-digit uppercase ARGB", () => {
		expect(resolveColor({ rgb: "ff0000" }, theme)).toBe("FFFF0000"); // 6-digit gains FF alpha
		expect(resolveColor({ rgb: "8000FF00" }, theme)).toBe("8000FF00"); // 8-digit passthrough
	});

	it("resolves a theme color with and without tint", () => {
		expect(resolveColor({ theme: 4 }, theme)).toBe("FF4F81BD"); // accent1, no tint
		expect(resolveColor({ theme: 4, tint: 0.3999755851924192 }, theme)).toBe("FF96B4D8");
		expect(resolveColor({ theme: 0 }, theme)).toBe("FFFFFFFF"); // Background 1 = white (swap)
	});

	it("returns undefined when it cannot resolve", () => {
		expect(resolveColor({ theme: 4 }, undefined)).toBeUndefined(); // no theme part
		expect(resolveColor({ theme: 99 }, theme)).toBeUndefined(); // index past the 12 slots
		expect(resolveColor({ indexed: 5 }, theme)).toBeUndefined(); // palette out of F5.3 scope
		expect(resolveColor({ auto: true }, theme)).toBeUndefined(); // consumer decides
	});
});
