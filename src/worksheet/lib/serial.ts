const MILLISECONDS_PER_DAY = 86_400_000;
/** Days from Excel's day zero (1899-12-30) to the Unix epoch. */
const EPOCH_OFFSET_DAYS = 25_569;
/** Serials below this are shifted down by one: Excel believes 1900-02-29 existed. */
const LEAP_BUG_BOUNDARY = 61;
const MIN_SERIAL = 1; // 1900-01-01
const MAX_SERIAL = 2_958_465; // 9999-12-31

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
