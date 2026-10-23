import { describe, expect, it } from "vitest";
import { RecordData, readRecords } from "../../biff";
import { CELLXFS, CELLXFS_END, FMT, XF } from "../records";
import { record, recordId, recordLength, stylesBin, u16, u32, wideString } from "./styles-bin";

describe("styles.bin builders", () => {
	it("encodes ids, varint lengths and little-endian integers", () => {
		expect(recordId(0x2c)).toEqual([0x2c]);
		expect(recordId(0x04e9)).toEqual([0xe9, 0x04]);
		expect(recordLength(5)).toEqual([5]);
		expect(recordLength(200)).toEqual([0xc8, 0x01]);
		expect(u16(0x0102)).toEqual([0x02, 0x01]);
		expect(u32(0x01020304)).toEqual([0x04, 0x03, 0x02, 0x01]);
		expect(wideString("Ab")).toEqual([2, 0, 0, 0, 0x41, 0, 0x62, 0]);
		expect(record(0x2c, [7])).toEqual([0x2c, 1, 7]);
	});

	it("builds BrtFmt and BrtXF records the BIFF12 reader walks back", () => {
		const records = [...readRecords(stylesBin([[164, "0.0"]], [0, 164]))];
		expect(records.map((r) => r.id)).toEqual([FMT, CELLXFS, XF, XF, CELLXFS_END]);
		const [fmtRecord, , , xfRecord] = records;
		if (fmtRecord === undefined || xfRecord === undefined) throw new Error("records missing");
		const fmt = new RecordData(fmtRecord.data);
		expect([fmt.u16(), fmt.wideString()]).toEqual([164, "0.0"]);
		const xf = new RecordData(xfRecord.data);
		expect([xf.u16(), xf.u16()]).toEqual([0, 164]);
	});
});
