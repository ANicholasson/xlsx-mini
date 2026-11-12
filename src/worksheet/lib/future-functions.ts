/**
 * Functions Excel added after 2007. A file stores each under a prefixed name, and one that stores
 * the bare name shows `#NAME?` in Excel and LibreOffice until the cell is edited (ADR 0023). The
 * names are Excel's own list of such functions, as its documentation and the two writers that
 * follow it publish them.
 */
const FUTURE_FUNCTIONS =
  "ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE " +
  "BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH CEILING.PRECISE " +
  "CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHOOSECOLS CHOOSEROWS COMBINA " +
  "CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL " +
  "DROP ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST " +
  "FIELDVALUE FILTERXML FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT " +
  "FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST " +
  "GAMMA.INV GAMMALN.PRECISE GAUSS HSTACK HYPGEOM.DIST IFNA IFS IMAGE IMCOSH IMCOT IMCSC IMCSCH " +
  "IMSEC IMSECH IMSINH IMTAN ISFORMULA ISOMITTED ISOWEEKNUM LOGNORM.DIST LOGNORM.INV MAKEARRAY " +
  "MAP MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NORM.DIST NORM.INV NORM.S.DIST " +
  "NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC " +
  "PERCENTRANK.INC PERMUTATIONA PHI POISSON.DIST QUARTILE.EXC QUARTILE.INC QUERYSTRING " +
  "RANDARRAY RANK.AVG RANK.EQ REDUCE RRI SCAN SEC SECH SEQUENCE SHEET SHEETS SKEW.P SORTBY " +
  "STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TEXTAFTER " +
  "TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW UNICHAR UNICODE UNIQUE VALUETOTEXT VAR.P VAR.S " +
  "VSTACK WEBSERVICE WEIBULL.DIST WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST";

/** The two that are stored under a longer prefix. */
const WORKSHEET_FUNCTIONS = "FILTER SORT";

/** `LET` and `LAMBDA` name their own parameters, each stored under a prefix as well. */
const PARAMETERISED = new Set(["LET", "LAMBDA", "_XLFN.LET", "_XLFN.LAMBDA"]);

let stored: Map<string, string> | undefined;

/** The stored name of each function, built on the first formula that calls anything. */
function storedNames(): Map<string, string> {
  if (stored === undefined) {
    stored = new Map();
    for (const name of FUTURE_FUNCTIONS.split(" ")) stored.set(name, `_xlfn.${name}`);
    for (const name of WORKSHEET_FUNCTIONS.split(" ")) stored.set(name, `_xlfn._xlws.${name}`);
  }
  return stored;
}

const QUOTE = 34;
const APOSTROPHE = 39;
const PARENTHESIS = 40;

/** A letter or an underscore: where a function name starts. */
function startsName(code: number): boolean {
  return (code >= 65 && code <= 90) || (code >= 97 && code <= 122) || code === 95;
}

/** What a name goes on with: letters, digits, underscores and the dots of `STDEV.S` or a prefix. */
function continuesName(code: number): boolean {
  return startsName(code) || (code >= 48 && code <= 57) || code === 46;
}

/**
 * A formula as the file stores it: every call of a function newer than Excel 2007 under its
 * prefixed name, so a formula works as it is typed. Text in double quotes and sheet names in
 * apostrophes are passed over, a name that already carries its prefix is left as it is, and so
 * is a whole formula that calls `LET` or `LAMBDA`: their parameters need prefixes too, and a
 * formula prefixed by halves is one Excel repairs.
 */
export function storedFormula(formula: string): string {
  if (!formula.includes("(")) return formula;
  const names = storedNames();
  const length = formula.length;
  let out = "";
  let copied = 0;
  let i = 0;
  while (i < length) {
    const code = formula.charCodeAt(i);
    if (code === QUOTE || code === APOSTROPHE) {
      // To the closing quote; a doubled quote is one inside the text and closes nothing.
      i++;
      while (i < length) {
        if (formula.charCodeAt(i) !== code) i++;
        else if (formula.charCodeAt(i + 1) === code) i += 2;
        else break;
      }
      i++;
    } else if (startsName(code)) {
      const start = i;
      i++;
      while (i < length && continuesName(formula.charCodeAt(i))) i++;
      if (formula.charCodeAt(i) !== PARENTHESIS) continue;
      const name = formula.slice(start, i).toUpperCase();
      if (PARAMETERISED.has(name)) return formula;
      const prefixed = names.get(name);
      if (prefixed === undefined) continue;
      out += formula.slice(copied, start) + prefixed;
      copied = i;
    } else {
      i++;
    }
  }
  return copied === 0 ? formula : out + formula.slice(copied);
}
