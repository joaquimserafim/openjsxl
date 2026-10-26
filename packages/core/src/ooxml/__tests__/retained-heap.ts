import { setFlagsFromString } from "node:v8";
import { runInNewContext } from "node:vm";

// Heap measurement for memory pins (F11.58). Vitest's workers start without --expose-gc, so V8's
// gc is exposed at run time; heap use is read after a full collection, so only live objects count.

// Keeps a measured result reachable across the second collection.
const holder: { value: unknown } = { value: undefined };

/**
 * Run a full garbage collection from a test.
 *
 * @example collectGarbage() // a full GC has run
 */
export function collectGarbage(): void {
	setFlagsFromString("--expose-gc");
	const gc: unknown = runInNewContext("gc");
	if (typeof gc === "function") gc();
}

/**
 * The heap bytes that `build`'s result keeps alive: heap use after a full GC while holding the
 * result, minus heap use after a full GC before building it. Objects `build` drops are not counted.
 *
 * @example retainedHeapBytes(() => Array.from({ length: 1_000 }, () => "x").join("")) // about 1,000
 */
export function retainedHeapBytes(build: () => unknown): number {
	collectGarbage();
	const before = process.memoryUsage().heapUsed;
	holder.value = build();
	collectGarbage();
	const after = process.memoryUsage().heapUsed;
	holder.value = undefined;
	return after - before;
}
