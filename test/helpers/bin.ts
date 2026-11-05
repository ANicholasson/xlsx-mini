import { join } from "node:path";

const windows = process.platform === "win32";

/** On Windows the shims in node_modules/.bin are .cmd files, which only run through a shell. */
export const shell = windows;

/** The path of a command from node_modules/.bin, to run with `shell` from the helpers. */
export function binary(root: string, name: string): string {
  return join(root, "node_modules", ".bin", windows ? `${name}.cmd` : name);
}
