import { XlsxReadError } from "../../errors/read.js";
import type { XmlVisitor } from "../../xml/read.js";

interface Relationship {
  readonly type: string | undefined;
  readonly target: string | undefined;
}

/** A `Map`, so an id from the file never reaches a prototype. */
type Relationships = ReadonlyMap<string | undefined, Relationship>;

export function refuse(message: string): never {
  throw new XlsxReadError("INVALID_WORKBOOK", message);
}

export function relationshipsOf(part: string): string {
  const folder = part.lastIndexOf("/") + 1;
  return `${part.slice(0, folder)}_rels/${part.slice(folder)}.rels`;
}

/**
 * The part holds `Relationship` elements directly inside its root and nothing else, as the schema
 * says: anything else is refused, never skipped. An id stated twice is refused, since a sheet names
 * its relationship by id.
 */
export function relationshipsVisitor(): XmlVisitor<never, Relationships> {
  const relationships = new Map<string | undefined, Relationship>();
  return {
    open(xml) {
      const root = xml.depth === 1;
      if (xml.depth > 2 || xml.name !== (root ? "Relationships" : "Relationship")) {
        refuse("a relationships part is damaged");
      }
      if (root) return;
      const id = xml.attribute("Id");
      if (id === undefined || relationships.has(id)) refuse("a relationship has no id of its own");
      const mode = xml.attribute("TargetMode");
      relationships.set(id, {
        type: xml.attribute("Type"),
        target: mode === undefined || mode === "Internal" ? xml.attribute("Target") : undefined,
      });
    },
    close: () => undefined,
    end: () => relationships,
  };
}

const UNFOLLOWED = "a relationship does not name a part of the package";

/**
 * A target holding a character another reader could take another way is refused: `%` decoded or
 * not, `?` and `#` cut off or not, `\` as a folder or not, `:` as a scheme.
 */
export function resolveTarget(source: string, target: string | undefined): string {
  if (target === undefined) refuse(UNFOLLOWED);
  for (const character of "%?#\\:") if (target.includes(character)) refuse(UNFOLLOWED);
  const absolute = target.startsWith("/");
  const steps = absolute ? [] : source.split("/").slice(0, -1);
  const path = target.slice(absolute ? 1 : 0).split("/");
  for (const step of path) {
    if (step === "" || (step === ".." && steps.pop() === undefined)) refuse(UNFOLLOWED);
    if (step !== "." && step !== "..") steps.push(step);
  }
  const last = path[path.length - 1];
  if (last === "." || last === "..") refuse(UNFOLLOWED);
  return steps.join("/");
}
