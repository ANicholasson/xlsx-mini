import { writeXlsx } from "xlsx-mini";

export default function dates(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "Dates",
        rows: [
          // A Date is written as its UTC instant and shown with the built-in date-time format.
          ["Shipped", new Date("2026-11-05T14:30:00Z")],
          // A Temporal.PlainDate shows as a date; PlainDateTime, ZonedDateTime and Instant as
          // date-times. They are recognised by their Symbol.toStringTag, so no polyfill is needed.
          ["Due", { [Symbol.toStringTag]: "Temporal.PlainDate", year: 2026, month: 12, day: 1 }],
        ],
      },
    ],
  });
}
