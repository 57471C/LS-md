export type EditorPatch = {
  value: string;
  start: number;
  end: number;
};

export function wrapSelection(
  value: string,
  start: number,
  end: number,
  before: string,
  after: string = before,
): EditorPatch {
  const selected = value.slice(start, end) || "text";
  const next = value.slice(0, start) + before + selected + after + value.slice(end);
  return {
    value: next,
    start: start + before.length,
    end: start + before.length + selected.length,
  };
}

export function toggleLinePrefix(
  value: string,
  start: number,
  end: number,
  prefix: string,
): EditorPatch {
  const lineStart = value.lastIndexOf("\n", Math.max(0, start - 1)) + 1;
  const foundEnd = value.indexOf("\n", end);
  const lineEnd = foundEnd === -1 ? value.length : foundEnd;
  const block = value.slice(lineStart, lineEnd);
  const lines = block.split("\n");
  const allPrefixed = lines.every((line) => line.startsWith(prefix) || line.length === 0);
  const nextLines = lines.map((line) => {
    if (line.length === 0) return line;
    if (allPrefixed) {
      return line.startsWith(prefix) ? line.slice(prefix.length) : line;
    }
    return line.startsWith(prefix) ? line : prefix + line;
  });
  const joined = nextLines.join("\n");
  const next = value.slice(0, lineStart) + joined + value.slice(lineEnd);
  return {
    value: next,
    start: lineStart,
    end: lineStart + joined.length,
  };
}

export function insertSnippet(
  value: string,
  start: number,
  end: number,
  snippet: string,
): EditorPatch {
  const next = value.slice(0, start) + snippet + value.slice(end);
  const caret = start + snippet.length;
  return { value: next, start: caret, end: caret };
}

export function headingPrefix(level: 1 | 2 | 3): string {
  return `${"#".repeat(level)} `;
}
