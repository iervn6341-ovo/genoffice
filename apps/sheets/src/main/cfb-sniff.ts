/**
 * CFB (OLE2 compound document) sniffing: legacy .xls and encrypted OOXML
 * (password-protected xlsx) share the same magic number.
 * The latter has an EncryptedPackage stream in its directory (stream name in UTF-16LE).
 */
const CFB_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
const ENCRYPTED_STREAM_UTF16 = Buffer.from('EncryptedPackage', 'utf16le')

export function isCfbHeader(head: Buffer): boolean {
  return head.length >= 8 && head.subarray(0, 8).equals(CFB_MAGIC)
}

export function cfbKind(bytes: Buffer): 'legacy' | 'encrypted' | null {
  if (!isCfbHeader(bytes)) return null
  return bytes.includes(ENCRYPTED_STREAM_UTF16) ? 'encrypted' : 'legacy'
}
