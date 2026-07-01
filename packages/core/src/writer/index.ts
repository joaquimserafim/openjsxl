// writer layer — build the OPC/ZIP container and (from F3.2) the OOXML parts inside it.
// Internal for now: the public `openjsxl/write` surface arrives with the workbook writer (F3.2).

export { crc32 } from './crc32'
export { deflateRaw } from './deflate'
export { writeZip, type ZipInput } from './zip'
