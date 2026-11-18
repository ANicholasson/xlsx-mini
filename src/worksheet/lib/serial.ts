const MILLISECONDS_PER_DAY = 86_400_000;
/** Days from Excel's day zero (1899-12-30) to the Unix epoch. */
const EPOCH_OFFSET_DAYS = 25_569;
/** Serials below this are shifted down by one: Excel believes 1900-02-29 existed. */
const LEAP_BUG_BOUNDARY = 61;
const MIN_SERIAL = 1; // 1900-01-01
const MAX_SERIAL = 2_958_465; // 9999-12-31
const DATE_1904_OFFSET_DAYS = 1462;
const MARCH_FIRST_1900_MILLIS = (LEAP_BUG_BOUNDARY - EPOCH_OFFSET_DAYS) * MILLISECONDS_PER_DAY;
const YEAR_10000_MILLIS = (MAX_SERIAL + 1 - EPOCH_OFFSET_DAYS) * MILLISECONDS_PER_DAY;

/**
 * Excel serial for an instant given as milliseconds since the Unix epoch, read as UTC.
 * Returns NaN for a non-finite input or a day outside 1900-01-01 to 9999-12-31.
 */
export function serialFromEpochMillis(millis: number): number {
  if (!Number.isFinite(millis)) return Number.NaN;
  const days = Math.floor(millis / MILLISECONDS_PER_DAY);
  const fraction = (millis - days * MILLISECONDS_PER_DAY) / MILLISECONDS_PER_DAY;
  return finish(days + EPOCH_OFFSET_DAYS, fraction);
}

/**
 * Excel serial for wall-clock fields (month 1 to 12), with no time-zone conversion.
 * Returns NaN for a day outside 1900-01-01 to 9999-12-31.
 */
export function serialFromCivil(
  year: number,
  month: number,
  day: number,
  hour: number,
  minute: number,
  second: number,
  millisecond: number,
): number {
  const days = daysFromCivil(year, month, day) + EPOCH_OFFSET_DAYS;
  const fraction =
    (((hour * 60 + minute) * 60 + second) * 1000 + millisecond) / MILLISECONDS_PER_DAY;
  return finish(days, fraction);
}

/**
 * The instant a serial names, as milliseconds since the Unix epoch read as UTC and rounded to the
 * millisecond. The inverse of `serialFromEpochMillis`. Returns NaN for a serial that names no
 * instant the writer can write. In the 1900 date system that is a serial below 1, from 60 up to 61
 * (29 February 1900, which never existed) or past 9999-12-31. With `date1904` it is a serial below 0
 * or past 9999-12-31. The range holds for the rounded instant too, so a serial a fraction of a
 * millisecond below 60, below 61 or below the end of 9999 gives NaN.
 */
export function epochMillisFromSerial(serial: number, date1904: boolean): number {
  let days = serial;
  let end = YEAR_10000_MILLIS;
  if (date1904) {
    if (serial < 0) return Number.NaN;
    days += DATE_1904_OFFSET_DAYS;
  } else if (serial < LEAP_BUG_BOUNDARY - 1) {
    if (serial < MIN_SERIAL) return Number.NaN;
    days += 1;
    end = MARCH_FIRST_1900_MILLIS;
  } else if (serial < LEAP_BUG_BOUNDARY) {
    return Number.NaN;
  }
  const millis = Math.round((days - EPOCH_OFFSET_DAYS) * MILLISECONDS_PER_DAY);
  return millis < end ? millis : Number.NaN;
}

function finish(days: number, fraction: number): number {
  const serialDay = days < LEAP_BUG_BOUNDARY ? days - 1 : days;
  if (serialDay < MIN_SERIAL || serialDay > MAX_SERIAL) return Number.NaN;
  return serialDay + fraction;
}

/** Days since 1970-01-01 for a proleptic Gregorian date (Howard Hinnant's algorithm). */
function daysFromCivil(year: number, month: number, day: number): number {
  const shiftedYear = month <= 2 ? year - 1 : year;
  const era = Math.floor(shiftedYear / 400);
  const yearOfEra = shiftedYear - era * 400;
  const shiftedMonth = month > 2 ? month - 3 : month + 9;
  const dayOfYear = Math.floor((153 * shiftedMonth + 2) / 5) + day - 1;
  const dayOfEra =
    yearOfEra * 365 + Math.floor(yearOfEra / 4) - Math.floor(yearOfEra / 100) + dayOfYear;
  return era * 146_097 + dayOfEra - 719_468;
}
