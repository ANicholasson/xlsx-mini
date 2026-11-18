import fc from "fast-check";
import { describe, expect, it } from "vite-plus/test";
import { epochMillisFromSerial, serialFromCivil, serialFromEpochMillis } from "../index.js";

const DAY = 86_400_000;
const utc = (iso: string) => Date.parse(`${iso}Z`);
const FIRST = utc("1900-01-01T00:00:00");
const LAST = utc("9999-12-31T23:59:59.999");

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

  it("reads published anchors back as instants, in both date systems", () => {
    expect(epochMillisFromSerial(1, false)).toBe(FIRST);
    expect(epochMillisFromSerial(59.5, false)).toBe(utc("1900-02-28T12:00:00"));
    expect(epochMillisFromSerial(61, false)).toBe(utc("1900-03-01T00:00:00"));
    expect(epochMillisFromSerial(25569, false)).toBe(0);
    expect(epochMillisFromSerial(36526.25, false)).toBe(utc("2000-01-01T06:00:00"));
    expect(epochMillisFromSerial(2958465, false)).toBe(utc("9999-12-31T00:00:00"));
    expect(epochMillisFromSerial(0, true)).toBe(utc("1904-01-01T00:00:00"));
    expect(epochMillisFromSerial(35064.5, true)).toBe(utc("2000-01-01T12:00:00"));
    expect(epochMillisFromSerial(2957003, true)).toBe(utc("9999-12-31T00:00:00"));
  });

  it("gives no instant for a serial the writer cannot write, even after rounding", () => {
    for (const serial of [
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.NEGATIVE_INFINITY,
      -1,
      0,
      0.999999995,
      59.999999995,
      60,
      60.5,
      60.999999995,
      2958465.999999995,
      2958466,
    ]) {
      expect(epochMillisFromSerial(serial, false), String(serial)).toBeNaN();
    }
    for (const serial of [Number.NaN, -0.000000001, 2957003.999999995, 2957004]) {
      expect(epochMillisFromSerial(serial, true), String(serial)).toBeNaN();
    }
  });

  it("gives back exactly the instant it wrote, over the whole range", () => {
    const nearEdge = (iso: string) =>
      fc.integer({ min: utc(iso) - 2 * DAY, max: utc(iso) + 2 * DAY });
    fc.assert(
      fc.property(
        fc.oneof(
          fc.integer({ min: FIRST, max: LAST }),
          nearEdge("1900-01-01T00:00:00"),
          nearEdge("1900-03-01T00:00:00"),
          nearEdge("9999-12-31T00:00:00"),
        ),
        (millis) => {
          const serial = serialFromEpochMillis(millis);
          if (millis < FIRST || millis > LAST) expect(serial).toBeNaN();
          else expect(epochMillisFromSerial(serial, false)).toBe(millis);
        },
      ),
      { numRuns: 10_000 },
    );
  });

  it("gives only instants the writer writes, each within half a millisecond and a double's step of the serial", () => {
    const nearEdge = (edge: number) =>
      fc.double({ min: -2, max: 2, noNaN: true }).map((millis) => edge + millis / DAY);
    const serials = (edges: readonly number[], max: number) =>
      fc.oneof(fc.double({ min: -2, max, noNaN: true }), ...edges.map(nearEdge));
    const system = fc.oneof(
      serials([1, 60, 61, 2958466], 2958467).map((serial) => ({ serial, date1904: false })),
      serials([0, 2957004], 2957005).map((serial) => ({ serial, date1904: true })),
    );
    fc.assert(
      fc.property(system, ({ serial, date1904 }) => {
        const millis = epochMillisFromSerial(serial, date1904);
        fc.pre(!Number.isNaN(millis));
        const written = serialFromEpochMillis(millis);
        expect(written).not.toBeNaN();
        expect(Math.abs(written - (date1904 ? 1462 : 0) - serial) * DAY).toBeLessThan(0.6);
      }),
      { numRuns: 10_000 },
    );
  });
});
