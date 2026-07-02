// xml layer — streaming tokenizer + entity decoding for the OOXML subset.
// Populated by F1.1 / F1.2; chunk-fed streaming wrapper added in F2.2.
export { decodeXmlEntities } from "./entities"
export { createXmlStream, type XmlStream } from "./stream"
export type { XmlToken } from "./tokenizer"
export { tokenize } from "./tokenizer"
