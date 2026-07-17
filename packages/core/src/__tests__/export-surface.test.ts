import { dirname } from "node:path";
import { fileURLToPath } from "node:url";
import ts from "typescript";
import { describe, expect, it } from "vitest";

// A type that appears in an EXPORTED signature but is not itself exported from the package index is
// unnameable: a consumer can hold the value the API hands them but cannot write its type, cannot
// declare a variable for it, cannot wrap the call. Eight M10 model types (CellProtection, PageSetup,
// …) shipped that way before F10.6. This walks the real public surface with the TypeScript compiler
// (no build step — it reads the source `.d.ts`-equivalent) and fails if any in-package type reachable
// from an exported declaration is missing from the index's exports. Self-maintaining: a future
// exported accessor that returns a new internal type trips it with no list to update here.
const HERE = dirname(fileURLToPath(import.meta.url));
const SRC = dirname(HERE); // packages/core/src
const INDEX = `${SRC}/index.ts`;

function program(): ts.Program {
	const tsconfig = `${dirname(SRC)}/tsconfig.json`;
	const parsed = ts.parseJsonConfigFileContent(
		ts.readConfigFile(tsconfig, ts.sys.readFile).config,
		ts.sys,
		dirname(tsconfig),
	);
	return ts.createProgram([INDEX], { ...parsed.options, noEmit: true });
}

// A symbol declared inside packages/core/src (our own code), not a lib/node_modules type.
function isInPackage(sym: ts.Symbol): boolean {
	return (sym.declarations ?? []).some((d) => {
		const f = d.getSourceFile().fileName;
		return f.includes("/packages/core/src/") && !f.includes("/node_modules/");
	});
}

// Nameable TYPE declarations — the kind a consumer would need to import. Type parameters, values,
// and property signatures are excluded, so a generic `<T>` or a referenced function never trips it.
function isNameableType(sym: ts.Symbol): boolean {
	return (sym.declarations ?? []).some(
		(d) =>
			ts.isInterfaceDeclaration(d) || ts.isTypeAliasDeclaration(d) || ts.isEnumDeclaration(d),
	);
}

// A type its own module marks `export` — i.e. one the author declared as a named, importable entity.
// This scopes the pin to the real gap: a type meant to be nameable that the package index forgot to
// re-surface (the eight M10 types). A file-private structural base (e.g. `CellBase`, folded into the
// exported `Cell` union and emitted alongside it) is deliberately internal and never flagged.
function isExportedFromOwnModule(sym: ts.Symbol): boolean {
	return (sym.declarations ?? []).some(
		(d) => ts.getCombinedModifierFlags(d) & ts.ModifierFlags.Export,
	);
}

// A private (`#name` or `private`) class member — never part of the public type surface.
function isPrivateMember(node: ts.Node): boolean {
	return (
		(ts.isPropertyDeclaration(node) ||
			ts.isMethodDeclaration(node) ||
			ts.isGetAccessorDeclaration(node) ||
			ts.isSetAccessorDeclaration(node)) &&
		(ts.isPrivateIdentifier(node.name) ||
			Boolean(ts.getCombinedModifierFlags(node) & ts.ModifierFlags.Private))
	);
}

describe("public export surface", () => {
	it("every type referenced by an exported signature is itself exported from index.ts", () => {
		const prog = program();
		const checker = prog.getTypeChecker();
		const indexSource = prog.getSourceFile(INDEX);
		expect(indexSource, "index.ts must be in the program").toBeDefined();
		const indexSymbol = checker.getSymbolAtLocation(indexSource as ts.SourceFile);
		expect(indexSymbol, "index.ts must be a module").toBeDefined();

		const exports = checker.getExportsOfModule(indexSymbol as ts.Symbol);
		const exportedNames = new Set(exports.map((s) => s.name));

		// Resolve each export alias to the real declaration whose body we walk for type references.
		const real = exports.map((s) =>
			s.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(s) : s,
		);

		const offenders = new Set<string>();
		const visit = (node: ts.Node): void => {
			// Only the public type SIGNATURE counts: a function/method body's locals and any private
			// (`#`) member are invisible to a consumer, so a type used only there is not a gap.
			if (ts.isBlock(node) || isPrivateMember(node)) return;
			if (ts.isTypeReferenceNode(node)) {
				const name = ts.isQualifiedName(node.typeName)
					? node.typeName.right
					: node.typeName;
				const sym = checker.getSymbolAtLocation(name);
				const resolved =
					sym && sym.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(sym) : sym;
				if (
					resolved &&
					isInPackage(resolved) &&
					isNameableType(resolved) &&
					isExportedFromOwnModule(resolved) &&
					!exportedNames.has(resolved.name)
				) {
					offenders.add(resolved.name);
				}
			}
			ts.forEachChild(node, visit);
		};

		for (const sym of real) {
			for (const decl of sym.declarations ?? []) {
				if (decl.getSourceFile().fileName.includes("/packages/core/src/")) visit(decl);
			}
		}

		expect([...offenders].sort()).toEqual([]);
	});
});
