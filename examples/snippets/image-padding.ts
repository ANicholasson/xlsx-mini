import { sheetFromObjects, writeXlsx } from "xlsx-mini";

interface Item {
  name: string;
  /** The bytes of a PNG the application produced, such as a code to scan. */
  label: Uint8Array;
}

export default function labels(items: Item[]): Promise<Uint8Array> {
  const sheet = sheetFromObjects(items, {
    name: "Labels",
    autoWidth: true,
    columns: [
      { key: "name", header: "Item" },
      {
        key: "label",
        header: "Label",
        // 10 pixels on every side: each row is 60 pixels tall, and one label is 20 from the next.
        value: (item) => ({ image: item.label, height: 40, padding: 10 }),
      },
    ],
  });
  return writeXlsx({ sheets: [sheet] });
}
