import { describe, expect, it } from "vitest";
import { XlsxError } from "../../errors";
import type { CellStyle } from "../../types";
import { createStyleRegistry } from "../styles";

const cyclic: { child?: object } = {};
cyclic.child = cyclic;

const invalidColors: readonly { label: string; value: unknown; description: string }[] = [
	{ label: "object", value: {}, description: "object" },
	{ label: "cyclic object", value: cyclic, description: "object" },
	{ label: "array", value: [], description: "object" },
	{ label: "null", value: null, description: "object" },
	{ label: "BigInt", value: 1n, description: "bigint" },
	{ label: "symbol", value: Symbol("invalid"), description: "symbol" },
	{ label: "number", value: 123, description: "number" },
	{ label: "undefined", value: undefined, description: "undefined" },
	{ label: "string", value: 'bad"\r', description: '"bad\\"\\r"' },
];

describe.each(["cell style", "differential style"])("%s invalid color diagnostics", (mode) => {
	it.each(invalidColors)(
		"rejects $label without inherited serialization",
		({ value, description }) => {
			const previous = Object.getOwnPropertyDescriptor(Object.prototype, "toJSON");
			let reads = 0;
			let error: unknown;
			try {
				Object.defineProperty(Object.prototype, "toJSON", {
					configurable: true,
					get() {
						reads++;
						throw new Error("inherited toJSON getter");
					},
				});
				const registry = createStyleRegistry();
				if (mode === "cell style") {
					const style: CellStyle = {};
					Object.defineProperty(style, "font", {
						enumerable: true,
						value: { color: { rgb: value } },
					});
					registry.xfIndexFor(style, false, "A1");
				} else {
					registry.internDxf({ font: { color: { rgb: value } } }, "rule 1");
				}
			} catch (cause) {
				error = cause;
			} finally {
				if (previous === undefined) Reflect.deleteProperty(Object.prototype, "toJSON");
				else Object.defineProperty(Object.prototype, "toJSON", previous);
			}
			// Assertions and test-runner error formatting execute only after global restoration.
			expect(reads).toBe(0);
			expect(error).toBeInstanceOf(XlsxError);
			expect(error).toMatchObject({ code: "invalid-input" });
			expect(error instanceof XlsxError && error.message).toContain(`(got ${description})`);
		},
	);
});
