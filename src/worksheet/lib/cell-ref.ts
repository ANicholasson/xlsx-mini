const letters: string[] = [];

/**
 * Excel column letters for a zero-based column index: 0 is `A`, 25 is `Z`, 26 is `AA`.
 * Results are cached, so a wide sheet pays the conversion once per column.
 */
export function columnLetter(index: number): string {
  const cached = letters[index];
  if (cached !== undefined) return cached;
  let result = "";
  let remaining = index;
  do {
    result = String.fromCharCode(65 + (remaining % 26)) + result;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  letters[index] = result;
  return result;
}
