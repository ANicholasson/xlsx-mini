import { XML_DECLARATION, escapeXml, needsSpacePreserve } from "../xml/index.js";

/** Interns cell text into the workbook's shared-strings table and renders that part. */
export interface StringTable {
  /** Returns the zero-based index to reference from a cell, adding the text on first sight. */
  intern(text: string): number;
  /** Renders `xl/sharedStrings.xml`. Call once, after every sheet has been rendered. */
  render(): string;
}

export function createStringTable(): StringTable {
  const indexes = new Map<string, number>();
  const items: string[] = [];
  let references = 0;

  return {
    intern(text) {
      references++;
      const existing = indexes.get(text);
      if (existing !== undefined) return existing;
      const index = items.length;
      indexes.set(text, index);
      items.push(renderItem(text));
      return index;
    },
    render() {
      return `${XML_DECLARATION}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${references}" uniqueCount="${items.length}">${items.join("")}</sst>`;
    },
  };
}

function renderItem(text: string): string {
  return needsSpacePreserve(text)
    ? `<si><t xml:space="preserve">${escapeXml(text)}</t></si>`
    : `<si><t>${escapeXml(text)}</t></si>`;
}
