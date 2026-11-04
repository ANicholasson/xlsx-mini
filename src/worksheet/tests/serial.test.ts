import fc from "fast-check";
import { describe, expect, it } from "vite-plus/test";
import { serialFromCivil, serialFromEpochMillis } from "../index.js";

const DAY = 86_400_000;
const utc = (iso: string) => Date.parse(`${iso}Z`);

describe("Excel serial dates", () => {
  it("matches published anchors, including the 1900 leap-year quirk", () => {
    expect(serialFromEpochMillis(utc("1900-01-01T00:00:00"))).toBe(1);
    expect(serialFromEpochMillis(utc("1900-02-28T00:00:00"))).toBe(59);
    expect(serialFromEpochMillis(utc("1900-03-01T00:00:00"))).toBe(61);
    expect(serialFromEpochMillis(utc("1970-01-01T00:00:00"))).toBe(25569);
    expect(serialFromEpochMillis(utc("2000-01-01T00:00:00"))).toBe(36526);
    expect(serialFromEpochMillis(utc("9999-12-31T00:00:00"))).toBe(2958465);
  });

  it("puts the time of day in the fraction", () => {
    expect(serialFromEpochMillis(utc("2000-01-01T12:00:00"))).toBe(36526.5);
    expect(serialFromEpochMillis(utc("2000-01-01T06:00:00"))).toBe(36526.25);
  });

  it("converts civil fields the same way, down to milliseconds", () => {
    expect(serialFromCivil(1900, 1, 1, 0, 0, 0, 0)).toBe(1);
    expect(serialFromCivil(1900, 3, 1, 0, 0, 0, 0)).toBe(61);
    expect(serialFromCivil(2000, 1, 1, 12, 0, 0, 0)).toBe(36526.5);
    expect(serialFromCivil(9999, 12, 31, 0, 0, 0, 0)).toBe(2958465);
  });

  it("returns NaN outside Excel's range or for an invalid instant", () => {
    expect(serialFromEpochMillis(utc("1899-12-31T23:59:59"))).toBeNaN();
    expect(serialFromEpochMillis(utc("10000-01-01T00:00:00"))).toBeNaN();
    expect(serialFromEpochMillis(Number.NaN)).toBeNaN();
    expect(serialFromCivil(1899, 12, 31, 0, 0, 0, 0)).toBeNaN();
  });

  it("agrees between the two algorithms for any date in range", () => {
    fc.assert(
      fc.property(
        fc.date({
          min: new Date(utc("1900-01-01T00:00:00")),
          max: new Date(utc("9999-12-31T23:59:59.999")),
          noInvalidDate: true,
        }),
        (date) => {
          const fromMillis = serialFromEpochMillis(date.getTime());
          const fromCivil = serialFromCivil(
            date.getUTCFullYear(),
            date.getUTCMonth() + 1,
            date.getUTCDate(),
            date.getUTCHours(),
            date.getUTCMinutes(),
            date.getUTCSeconds(),
            date.getUTCMilliseconds(),
          );
          expect(fromCivil).toBeCloseTo(fromMillis, 9);
        },
      ),
    );
  });

  it("advances by one per day, within the precision a double keeps near serial 3 million", () => {
    fc.assert(
      fc.property(
        fc.integer({ min: utc("1900-03-01T00:00:00"), max: utc("9999-12-30T00:00:00") }),
        (millis) => {
          expect(serialFromEpochMillis(millis + DAY) - serialFromEpochMillis(millis)).toBeCloseTo(
            1,
            9,
          );
        },
      ),
    );
  });
});
