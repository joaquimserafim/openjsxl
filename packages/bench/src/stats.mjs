// Tiny statistics + formatting helpers. No dependency; the harness stays as dep-light as the
// library it measures (competitors aside).

/** Median of a numeric array (returns NaN for an empty array). */
export function median(values) {
	if (values.length === 0) return Number.NaN;
	const sorted = [...values].sort((a, b) => a - b);
	const mid = sorted.length >> 1;
	return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

/** Min of a numeric array (NaN for empty). */
export function min(values) {
	return values.length === 0 ? Number.NaN : Math.min(...values);
}

/** Max of a numeric array (NaN for empty). */
export function max(values) {
	return values.length === 0 ? Number.NaN : Math.max(...values);
}

/** Milliseconds → a compact human string ("42 ms", "1.3 s"). */
export function fmtMs(ms) {
	if (!Number.isFinite(ms)) return "—";
	if (ms < 1000) return `${ms < 10 ? ms.toFixed(1) : Math.round(ms)} ms`;
	return `${(ms / 1000).toFixed(ms < 10_000 ? 2 : 1)} s`;
}

/** Bytes → a compact human string ("512 KB", "4.2 MB"). */
export function fmtBytes(bytes) {
	if (!Number.isFinite(bytes)) return "—";
	if (bytes < 1024) return `${bytes} B`;
	if (bytes < 1024 ** 2) return `${Math.round(bytes / 1024)} KB`;
	if (bytes < 1024 ** 3) return `${(bytes / 1024 ** 2).toFixed(1)} MB`;
	return `${(bytes / 1024 ** 3).toFixed(2)} GB`;
}
