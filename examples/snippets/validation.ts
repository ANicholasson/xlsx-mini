import { type DataValidation, writeXlsx } from "xlsx-mini";

const status: DataValidation = {
  range: "B2:B1000",
  list: ["Open", "Closed", "On hold"],
  input: { title: "Status", message: "Pick one from the list" },
};
const quantity: DataValidation = {
  range: "C2:C1000",
  number: { min: 1, max: 1000, integer: true },
  error: { title: "Quantity", message: "A whole number from 1 to 1000" },
};
// A dropdown read from cells, here a column on another sheet.
const owner: DataValidation = { range: "D2:D1000", list: "People!A2:A50" };
// Whole columns are written as E1:E1048576; a warning lets the entry through after a prompt.
const due: DataValidation = {
  range: "E:E",
  date: { min: new Date(Date.UTC(2026, 0, 1)) },
  error: { message: "Due dates fall in 2026 or later", style: "warning" },
};

export default function validated(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "Orders",
        rows: [
          ["Order", "Status", "Qty", "Owner", "Due"],
          ["A-1001", "Open", 2, "Kari", new Date(Date.UTC(2026, 11, 1))],
        ],
        validations: [status, quantity, owner, due],
      },
      { name: "People", rows: [["Name"], ["Kari"], ["Ola"]] },
    ],
  });
}
