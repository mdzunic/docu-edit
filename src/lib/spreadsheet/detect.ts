// The spreadsheet product's refusal signatures (SPEC-116), apart from the
// readers for the same reason presentation/keynote/detect.ts is: the open
// funnel asks these questions of files that may be anything at all, and
// answering must not load a workbook codec. (Workbook *detection* — which
// packages are .xlsx/.ods — is core/packageDetect.ts's spreadsheetFormatOf.)

/** The Compound File Binary signature: a legacy .xls (BIFF8) — or an
 * encrypted .xlsx/.xlsm, which Office wraps in the same container. */
const COMPOUND_FILE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** Whether `bytes` is an OLE compound file named as an Excel workbook
 * (`.xls`, `.xlsx`, `.xlsm`) — an old binary workbook or a password-protected
 * one. Any other compound file keeps the .docx path's own `.doc` message. */
export function isLegacyOrEncryptedWorkbook(bytes: Uint8Array, fileName: string): boolean {
  return (
    bytes.length >= COMPOUND_FILE_MAGIC.length &&
    COMPOUND_FILE_MAGIC.every((byte, index) => bytes[index] === byte) &&
    /\.(xls|xlsx|xlsm)$/i.test(fileName)
  );
}

/** The calculation-engine entry a Numbers document writes — the EXACT name:
 * Pages writes `Index/CalculationEngine-<id>.iwa`, so a prefix match would
 * call a Pages document Numbers (e2e/fixtures/spec116-numbers.md). */
export const NUMBERS_ENGINE_ENTRY = "Index/CalculationEngine.iwa";

/** Whether an already-unzipped package is an Apple Numbers document: it
 * holds `Index/CalculationEngine.iwa` and no `Index/Slide…` entry (Keynote
 * writes the same engine plus its slides). A `.numbers`-named ZIP with
 * neither marker counts too — the one place the extension steers, and only
 * toward a better message, as `.key` does for Keynote. */
export function isNumbersPackage(parts: Map<string, Uint8Array>, fileName: string): boolean {
  for (const name of parts.keys()) {
    if (name.startsWith("Index/Slide")) {
      return false;
    }
  }
  return parts.has(NUMBERS_ENGINE_ENTRY) || /\.numbers$/i.test(fileName);
}
