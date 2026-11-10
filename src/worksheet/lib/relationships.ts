import { XML_DECLARATION, escapeXml } from "../../xml/index.js";

const RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";

interface Relationship {
  readonly type: string;
  readonly target: string;
  readonly external: boolean;
}

/** The worksheet part's relationships: external link targets and the sheet's drawing, numbered as added. */
export interface SheetRelationships {
  /** Adds one and returns its id, `rIdN`. */
  add(type: string, target: string, external: boolean): string;
  /** `xl/worksheets/_rels/sheetN.xml.rels`, or undefined when the sheet relates to nothing. */
  render(): string | undefined;
}

/** A collector for one worksheet part; ids are handed out in the order relationships are added. */
export function createSheetRelationships(): SheetRelationships {
  const relationships: Relationship[] = [];
  return {
    add(type, target, external) {
      relationships.push({ type, target, external });
      return `rId${relationships.length}`;
    },
    render() {
      if (relationships.length === 0) return undefined;
      let xml = `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">`;
      for (let i = 0; i < relationships.length; i++) {
        const r = relationships[i]!;
        xml += `<Relationship Id="rId${i + 1}" Type="${r.type}" Target="${escapeXml(r.target)}"${r.external ? ' TargetMode="External"' : ""}/>`;
      }
      return `${xml}</Relationships>`;
    },
  };
}
