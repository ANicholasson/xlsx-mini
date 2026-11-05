import type { Workbook } from "../src/index.js";

/** 100 000 rows by 10 mixed columns; `scale` multiplies the row count. */
export declare function wideWorkbook(scale?: number): Workbook;
/** 40 sheets of 2 500 rows by 6 columns; `scale` multiplies the rows per sheet. */
export declare function manySheetsWorkbook(scale?: number): Workbook;
/** 200 000 rows by 5 text columns drawn from 2 000 phrases; `scale` multiplies the row count. */
export declare function stringHeavyWorkbook(scale?: number): Workbook;

export type SuiteName = "wide" | "many-sheets" | "strings";
export declare const suites: Readonly<Record<SuiteName, () => Workbook>>;
