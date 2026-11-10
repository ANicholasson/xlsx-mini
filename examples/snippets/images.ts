import { sheetFromObjects, writeXlsx } from "xlsx-mini";

interface Product {
  name: string;
  code: string;
  /** The bytes of a PNG the application produced; the writer reads only its size. */
  picture: Uint8Array;
}

export default function catalogue(products: Product[]): Promise<Uint8Array> {
  // A picture is a cell like any other: the row grows to its height and the column to its width.
  const sheet = sheetFromObjects(products, {
    name: "Catalogue",
    autoWidth: true,
    columns: [
      { key: "name", header: "Product" },
      { key: "code", header: "Code" },
      {
        key: "picture",
        header: "Picture",
        value: (product) => ({ image: product.picture, height: 40 }),
      },
    ],
  });
  return writeXlsx({ sheets: [sheet] });
}
