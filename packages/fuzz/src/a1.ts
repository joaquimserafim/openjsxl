// A minimal 1-based column-number → A1 letters helper, local to the harness (the core equivalent is
// internal). 1 → "A", 26 → "Z", 27 → "AA".
export function colToA1(col: number): string {
	let n = col;
	let s = "";
	while (n > 0) {
		const rem = (n - 1) % 26;
		s = String.fromCharCode(65 + rem) + s;
		n = Math.floor((n - 1) / 26);
	}
	return s;
}
