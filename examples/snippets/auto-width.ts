import { sheetFromObjects, writeXlsx } from "xlsx-mini";

interface Product {
  sku: string;
  name: string;
  price: number;
  restocked: Date;
  note: string;
}

export default function inventory(products: Product[]): Promise<Uint8Array> {
  // Every column without a numeric width fits its widest cell, bold header included; the note
  // column keeps the width it is given. On a single column, width: "auto" does the same.
  const sheet = sheetFromObjects(products, {
    name: "Inventory",
    autoWidth: true,
    columns: [
      { key: "sku", header: "SKU" },
      { key: "name", header: "Product" },
      { key: "price", header: "Unit price", style: { numberFormat: '#,##0.00 "kr"' } },
      { key: "restocked", header: "Restocked", style: { numberFormat: "yyyy-mm-dd" } },
      { key: "note", header: "Note", width: 40, style: { wrap: true } },
    ],
  });
  return writeXlsx({ sheets: [sheet] });
}
