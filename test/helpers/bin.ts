import { join } from "node:path";

const windows = process.platform === "win32";

/** A program to spawn and whether it needs a shell: `.cmd` shims do, executables must not. */
export interface Command {
  readonly file: string;
  readonly shell: boolean;
}

/** An executable, such as `process.execPath`, run directly so a space in its path is harmless. */
export function direct(file: string): Command {
  return { file, shell: false };
}

/**
 * A command from node_modules/.bin. On Windows the shim is a `.cmd` file, which only runs through
 * a shell; the path is quoted there because a shell splits an unquoted one at its first space.
 */
export function binary(root: string, name: string): Command {
  const file = join(root, "node_modules", ".bin", windows ? `${name}.cmd` : name);
  return { file: windows ? `"${file}"` : file, shell: windows };
}
