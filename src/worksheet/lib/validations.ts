import { XlsxMiniError, describeValue } from "../../errors/index.js";
import { escapeXml } from "../../xml/index.js";
import { absolute, boxProblem, parseReference, quoteSheetName } from "./cell-ref.js";
import type { DataValidation, ValidationMessage } from "./model.js";
import { serialFromEpochMillis } from "./serial.js";

const MAX_TITLE_LENGTH = 32;
const MAX_MESSAGE_LENGTH = 255;
/** Excel's limit on the source of a literal list, quotes included. */
const MAX_LIST_SOURCE_LENGTH = 255;
const ERROR_STYLES: ReadonlySet<string> = new Set(["stop", "warning", "information"]);
const SOURCE_RANGE =
  /^(?:'((?:[^']|'')+)'!|([^'!]+)!)?([A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?|[A-Z]{1,3}:[A-Z]{1,3})$/;

function invalid(path: string, message: string): XlsxMiniError {
  return new XlsxMiniError("INVALID_VALIDATION", path, message);
}

/**
 * The `<dataValidations>` element, or an empty string. Every rule is validated here, up front,
 * against what Excel's dialog accepts; the rules are written after the rows, so nothing about the
 * rows is needed.
 */
export function renderValidations(
  rules: readonly DataValidation[] | undefined,
  sheetIndex: number,
  sheetNames: readonly string[],
): string {
  if (rules === undefined || rules.length === 0) return "";
  let xml = `<dataValidations count="${rules.length}">`;
  for (let i = 0; i < rules.length; i++) {
    xml += renderRule(rules[i]!, `sheets[${sheetIndex}].validations[${i}]`, sheetNames);
  }
  return `${xml}</dataValidations>`;
}

function renderRule(rule: DataValidation, path: string, sheetNames: readonly string[]): string {
  if (typeof rule !== "object" || rule === null) {
    throw invalid(path, `A validation is an object, not ${describeValue(rule)}`);
  }
  const box = parseReference(rule.range);
  if (box === undefined) {
    throw invalid(
      `${path}.range`,
      `A validation applies to a cell, a range or whole columns in A1 form such as "B2:B100", not ${describeValue(rule.range)}`,
    );
  }
  const problem = boxProblem(box);
  if (problem !== undefined) throw invalid(`${path}.range`, problem);

  const kinds = ("list" in rule ? 1 : 0) + ("number" in rule ? 1 : 0) + ("date" in rule ? 1 : 0);
  if (kinds !== 1) {
    throw invalid(path, "A validation has exactly one of list, number and date");
  }
  let attributes: string;
  let formulas: string;
  if ("list" in rule) {
    attributes = ' type="list"';
    formulas = `<formula1>${listSource(rule.list, `${path}.list`, sheetNames)}</formula1>`;
  } else {
    const bounds = "number" in rule ? rule.number : rule.date;
    const boundsPath = `${path}.${"number" in rule ? "number" : "date"}`;
    if (typeof bounds !== "object" || bounds === null) {
      throw invalid(
        boundsPath,
        `Bounds are an object with min and max, not ${describeValue(bounds)}`,
      );
    }
    const min =
      "number" in rule
        ? bound(bounds.min, `${boundsPath}.min`)
        : date(bounds.min, `${boundsPath}.min`);
    const max =
      "number" in rule
        ? bound(bounds.max, `${boundsPath}.max`)
        : date(bounds.max, `${boundsPath}.max`);
    if (min === undefined && max === undefined) {
      throw invalid(boundsPath, "Bounds need a min, a max or both");
    }
    if (min !== undefined && max !== undefined && max < min) {
      throw invalid(boundsPath, `max ${max} is below min ${min}`);
    }
    const type = "date" in rule ? "date" : rule.number.integer === true ? "whole" : "decimal";
    if (
      "number" in rule &&
      rule.number.integer !== undefined &&
      typeof rule.number.integer !== "boolean"
    ) {
      throw invalid(
        `${boundsPath}.integer`,
        `integer is true or false, not ${describeValue(rule.number.integer)}`,
      );
    }
    if (min !== undefined && max !== undefined) {
      attributes = ` type="${type}" operator="between"`;
      formulas = `<formula1>${min}</formula1><formula2>${max}</formula2>`;
    } else if (min !== undefined) {
      attributes = ` type="${type}" operator="greaterThanOrEqual"`;
      formulas = `<formula1>${min}</formula1>`;
    } else {
      attributes = ` type="${type}" operator="lessThanOrEqual"`;
      formulas = `<formula1>${max}</formula1>`;
    }
  }

  if (rule.allowBlank !== undefined && typeof rule.allowBlank !== "boolean") {
    throw invalid(
      `${path}.allowBlank`,
      `allowBlank is true or false, not ${describeValue(rule.allowBlank)}`,
    );
  }
  if (rule.allowBlank !== false) attributes += ' allowBlank="1"';
  if (rule.input !== undefined) {
    attributes += ` showInputMessage="1"${message(rule.input, `${path}.input`, "promptTitle", "prompt")}`;
  }
  if (rule.error !== undefined) {
    const style = rule.error.style;
    if (style !== undefined && !ERROR_STYLES.has(style)) {
      throw invalid(
        `${path}.error.style`,
        `Error styles are stop, warning or information, not ${describeValue(style)}`,
      );
    }
    attributes +=
      ` showErrorMessage="1"${style === undefined || style === "stop" ? "" : ` errorStyle="${style}"`}` +
      message(rule.error, `${path}.error`, "errorTitle", "error");
  }
  return `<dataValidation${attributes} sqref="${box.ref}">${formulas}</dataValidation>`;
}

/** A literal list as Excel's quoted comma-separated source, or a range of cells as a reference. */
function listSource(list: unknown, path: string, sheetNames: readonly string[]): string {
  if (typeof list === "string") {
    const match = SOURCE_RANGE.exec(list);
    if (match === null) {
      throw invalid(
        path,
        `A list is an array of values or a range in A1 form such as "Lists!A2:A9", not ${describeValue(list)}`,
      );
    }
    const quoted = match[1];
    const sheet = quoted === undefined ? match[2] : quoted.replaceAll("''", "'");
    const box = parseReference(match[3])!;
    const problem = boxProblem(box);
    if (problem !== undefined) throw invalid(path, problem);
    if (
      sheet !== undefined &&
      !sheetNames.some((name) => name.toLowerCase() === sheet.toLowerCase())
    ) {
      throw invalid(path, `The workbook has no sheet named ${describeValue(sheet)}`);
    }
    const reference = absolute(box.ref);
    return escapeXml(sheet === undefined ? reference : `${quoteSheetName(sheet)}!${reference}`);
  }
  if (!Array.isArray(list) || list.length === 0) {
    throw invalid(
      path,
      `A list is a non-empty array of values or a range, not ${describeValue(list)}`,
    );
  }
  for (let i = 0; i < list.length; i++) {
    const value: unknown = list[i];
    if (typeof value !== "string") {
      throw invalid(`${path}[${i}]`, `List values are strings, not ${describeValue(value)}`);
    }
    if (value.includes(",")) {
      throw invalid(
        `${path}[${i}]`,
        `A list value cannot contain a comma; read the list from a range instead`,
      );
    }
  }
  // Excel's source syntax: the values comma-separated inside quotes, with quotes doubled.
  const source = `"${(list as string[]).join(",").replaceAll('"', '""')}"`;
  if (source.length > MAX_LIST_SOURCE_LENGTH) {
    throw invalid(
      path,
      `A literal list is at most ${MAX_LIST_SOURCE_LENGTH} characters including commas; read a longer list from a range`,
    );
  }
  return escapeXml(source);
}

function bound(value: unknown, path: string): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw invalid(path, `A bound is a finite number, not ${describeValue(value)}`);
  }
  return value;
}

function date(value: unknown, path: string): number | undefined {
  if (value === undefined) return undefined;
  if (!(value instanceof Date) || Number.isNaN(value.getTime())) {
    throw invalid(path, `A date bound is a valid Date, not ${describeValue(value)}`);
  }
  const serial = serialFromEpochMillis(value.getTime());
  if (Number.isNaN(serial)) throw invalid(path, "Excel dates run from 1900-01-01 to 9999-12-31");
  return serial;
}

/** The title and body attributes of a message, checked against Excel's lengths. */
function message(
  value: ValidationMessage,
  path: string,
  titleAttribute: string,
  bodyAttribute: string,
): string {
  if (typeof value !== "object" || value === null || typeof value.message !== "string") {
    throw invalid(path, "A message is an object with a message string and an optional title");
  }
  if (value.message.length === 0 || value.message.length > MAX_MESSAGE_LENGTH) {
    throw invalid(`${path}.message`, `A message is 1 to ${MAX_MESSAGE_LENGTH} characters`);
  }
  const title = value.title;
  if (title !== undefined && (typeof title !== "string" || title.length > MAX_TITLE_LENGTH)) {
    throw invalid(`${path}.title`, `A title is at most ${MAX_TITLE_LENGTH} characters`);
  }
  return (
    (title === undefined || title.length === 0 ? "" : ` ${titleAttribute}="${escapeXml(title)}"`) +
    ` ${bodyAttribute}="${escapeXml(value.message)}"`
  );
}
