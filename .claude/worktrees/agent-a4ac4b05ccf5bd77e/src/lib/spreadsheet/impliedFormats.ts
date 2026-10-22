// Implied number formats (SPEC-118 "Implied formats"): a value its source
// marks as a date, a time or a percentage but leaves unformatted gets a
// format of its own, so a date never arrives as a bare serial because its
// producer skipped the style. Both readers apply it — an ODS
// `office:value-type="date"`/`"time"`/`"percentage"` cell without a data
// style, an xlsx `t="d"` cell with a General style — by pointing the cell at
// a derived style: one per (base style, implied code) pair, appended to the
// style table with `numberFormat` set and no `sourceXf`/`sourceStyleName`
// (nothing in the file names it). Reading writes nothing, so a clean sheet
// still downloads byte-identical.
import type { CellStyle } from "./model";

export const IMPLIED_DATE = "yyyy-mm-dd";
export const IMPLIED_DATE_TIME = "yyyy-mm-dd hh:mm:ss";
export const IMPLIED_TIME = "[hh]:mm:ss";
export const IMPLIED_PERCENTAGE = "0.00%";

const MS_PER_DAY = 86_400_000;

/** A date value's implied code: the date, plus the time when it has one. */
export function impliedDateCode(serial: number): string {
  const msOfDay = ((Math.round(serial * MS_PER_DAY) % MS_PER_DAY) + MS_PER_DAY) % MS_PER_DAY;
  return msOfDay === 0 ? IMPLIED_DATE : IMPLIED_DATE_TIME;
}

/** Whether a style shows General (no number format of any kind). */
export function isGeneralStyle(style: CellStyle | undefined): boolean {
  return style?.numberFormat === undefined && style?.builtinFormatId === undefined;
}

/** The derived styles of one workbook read, appended to its style table. */
export class ImpliedFormatStyles {
  private readonly derived = new Map<string, number>();

  constructor(private readonly styles: CellStyle[]) {}

  /** The style a value with an implied `code` shows: `baseId` itself when
   * that style formats it already, else the derived copy carrying `code`. */
  styleFor(baseId: number, code: string): number {
    const base = this.styles[baseId] ?? this.styles[0];
    if (!isGeneralStyle(base)) {
      return baseId;
    }
    const key = `${baseId}:${code}`;
    let id = this.derived.get(key);
    if (id === undefined) {
      const copy: CellStyle = base ? structuredClone(base) : {};
      delete copy.sourceXf;
      delete copy.sourceStyleName;
      copy.numberFormat = code;
      id = this.styles.length;
      this.styles.push(copy);
      this.derived.set(key, id);
    }
    return id;
  }
}
