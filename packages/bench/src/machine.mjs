import { cpus, totalmem } from "node:os";

// A machine-info stamp so a published table can be reproduced (or contextualized) later. Numbers
// are only meaningful next to the hardware and runtime that produced them.
export function machineStamp() {
	const list = cpus();
	const model = list[0]?.model?.trim() ?? "unknown CPU";
	return {
		cpu: model,
		cores: list.length,
		memoryGiB: Math.round((totalmem() / 1024 ** 3) * 10) / 10,
		platform: `${process.platform} ${process.arch}`,
		node: process.version,
	};
}

/** One-line human summary of the stamp, for the report header. */
export function machineLine(stamp = machineStamp()) {
	return `${stamp.cpu} · ${stamp.cores} cores · ${stamp.memoryGiB} GiB · ${stamp.platform} · Node ${stamp.node}`;
}
