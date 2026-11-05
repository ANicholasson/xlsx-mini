import { type Coerce, writeXlsx } from "xlsx-mini";

export class Money {
  constructor(readonly cents: number) {}
}

// Let your own types into cells; the types then accept them everywhere a cell value goes.
declare module "xlsx-mini" {
  interface Register {
    cellValue: Money | bigint;
  }
}

// Then say how they become something a cell can hold.
const coerce: Coerce = (value) => {
  if (value instanceof Money) return value.cents / 100;
  if (typeof value === "bigint") return Number(value);
  return value;
};

export default function extended(): Promise<Uint8Array> {
  return writeXlsx(
    {
      sheets: [
        {
          name: "Ledger",
          rows: [
            ["Rent", new Money(1_250_000)],
            ["Views", 1234567890123n],
          ],
        },
      ],
    },
    { coerce },
  );
}
