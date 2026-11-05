import type { Workbook } from "../src/index.js";

/** 100 000 rows by 10 mixed columns. */
export declare function wideWorkbook(): Workbook;
/** 40 sheets of 2 500 rows by 6 columns. */
export declare function manySheetsWorkbook(): Workbook;
/** 200 000 rows by 5 text columns drawn from 2 000 phrases. */
export declare function stringHeavyWorkbook(): Workbook;

export type SuiteName = "wide" | "many-sheets" | "strings";
export declare const suites: Readonly<Record<SuiteName, () => Workbook>>;
