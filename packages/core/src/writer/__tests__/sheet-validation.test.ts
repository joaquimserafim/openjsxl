import { describe, expect, it } from "vitest";
import { checkKeys, sheetInvalid, shortened } from "../sheet-validation";

describe("worksheet writer validation helpers", () => {
	it("scopes invalid-input errors to the worksheet", () => {
		expect(() => sheetInvalid("Data", "bad value")).toThrow('sheet "Data": bad value');
	});

	it("accepts allowed keys and names the first unknown key", () => {
		expect(() => checkKeys("Data", "comment", { ref: "A1" }, ["ref"])).not.toThrow();
		expect(() => checkKeys("Data", "comment", { ref: "A1", typo: true }, ["ref"])).toThrow(
			'comment has an unknown property "typo"',
		);
	});

	it("keeps short values and truncates long values to 24 characters", () => {
		expect(shortened("A1")).toBe("A1");
		expect(shortened("123456789012345678901234567890")).toBe("123456789012345678901234…");
	});
});
