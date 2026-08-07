// Largest entry count representable without using the ZIP64 u16 sentinel (0xffff).
// Shared by both writers and the reader so they agree on the classic-ZIP boundary.
export const MAX_CLASSIC_ZIP_ENTRIES = 0xfffe;
