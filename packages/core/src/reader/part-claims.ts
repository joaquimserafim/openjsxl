// One open reads each package part at most once (F11.95). Each open keeps a set of the part paths it
// has read. The optional shared-strings, styles and theme loads, and every sheet, sheet-rels, comments
// and table load, claim their path first and skip the part when it is already claimed, so a package
// that names one part many times decodes it once.

/**
 * Claim the part at `path` for this open: `true` the first time (read it), `false` when this open has
 * already claimed it (skip it).
 *
 * @example claimPart(claimed, "xl/worksheets/sheet1.xml") // true: first claim, read the part
 * @example claimPart(claimed, "xl/worksheets/sheet1.xml") // false: already read in this open
 */
export function claimPart(claimed: Set<string>, path: string): boolean {
	if (claimed.has(path)) return false;
	claimed.add(path);
	return true;
}
