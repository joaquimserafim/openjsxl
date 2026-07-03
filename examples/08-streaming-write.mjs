// 08 — Streaming write (0.5): export a large sheet with roughly constant memory.
//
//   node 08-streaming-write.mjs          (from ./examples)
//   pnpm --filter openjsxl-examples stream-write
//
// `streamXlsx` returns a ReadableStream and pulls each sheet's `rows` only as the output is
// consumed — so the row source can be an async generator over a database cursor, a paged API, or
// any lazy sequence, and the whole sheet never lives in memory at once. Here we pipe the stream
// straight to a file; the peak working set stays flat no matter how many rows we export.

import { createWriteStream } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Writable } from "node:stream";
import { openXlsx, streamXlsx } from "openjsxl";

// A stand-in for a DB cursor: yields rows one at a time, never building an array of them all.
async function* salesCursor(count) {
	for (let i = 1; i <= count; i++) {
		// e.g. `await cursor.next()` in real code — the writer pulls this only as fast as it drains.
		yield [i, `Customer ${i}`, Math.round(Math.random() * 10_000) / 100, i % 2 === 0];
	}
}

const path = join(tmpdir(), "openjsxl-streamed.xlsx");

const stream = streamXlsx({
	sheets: [
		{
			name: "Sales",
			rows: salesCursor(200_000), // 200k rows, pulled lazily
			columns: [{ min: 2, max: 2, width: 24 }],
			freeze: { rows: 1 },
		},
	],
});

// Pipe the Web ReadableStream to the file — end-to-end streaming, so memory stays bounded.
await stream.pipeTo(Writable.toWeb(createWriteStream(path)));
console.log(`Streamed 200,000 rows to ${path}`);

// Read it back to confirm it is a normal, valid .xlsx.
const { readFile } = await import("node:fs/promises");
const wb = await openXlsx(await readFile(path));
const sheet = wb.sheet("Sales");
console.log("First row:", sheet.cell("A1").value, sheet.cell("B1").value, sheet.cell("C1").value);
console.log("Row 200000, column A:", sheet.cell("A200000").value);
