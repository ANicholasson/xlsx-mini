import { type PrintSetup, writeXlsx } from "xlsx-mini";

// &L, &C and &R start the left, centre and right sections; &P, &N, &D and &A are the page
// number, page count, date and sheet name; && is an ampersand.
const print: PrintSetup = {
  orientation: "landscape",
  paperSize: "A4",
  fitTo: { width: 1 },
  margins: { left: 0.5, right: 0.5 },
  header: "&LQuarterly report&R&D",
  footer: "&L&A&RPage &P of &N",
  repeatRows: 1,
};

export default function printable(): Promise<Uint8Array> {
  const rows = Array.from({ length: 200 }, (_, i) => [
    i + 1,
    `Region ${(i % 4) + 1}`,
    (i + 1) * 125,
  ]);
  return writeXlsx({
    sheets: [{ name: "Report", rows: [["Week", "Region", "Revenue"], ...rows], print }],
  });
}
