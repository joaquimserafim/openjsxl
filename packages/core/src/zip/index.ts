// zip layer — OPC/ZIP container reading and the platform-backed inflate primitive
// (DecompressionStream).

export type { ZipArchive, ZipEntry, ZipLimits, ZipReadOptions } from "./central-directory";
export { openZip, zipLimits } from "./central-directory";
export { inflateRaw, inflateRawStream } from "./inflate";
