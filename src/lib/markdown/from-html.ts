function escapeMd(text: string): string {
  return text.replace(/([\\`*_[\]#])/g, "\\$1");
}

function styleLooksBold(style: string): boolean {
  return /font-weight\s*:\s*(bold|[6-9]00)/i.test(style);
}

function styleLooksItalic(style: string): boolean {
  return /font-style\s*:\s*italic/i.test(style);
}

function preprocessWordHtml(html: string): string {
  return html
    .replace(/<xml[\s\S]*?<\/xml>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<!--\[if[\s\S]*?<!\[endif\]-->/gi, "")
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/<\/?o:p[^>]*>/gi, "")
    .replace(/<\/?w:[^>]*>/gi, "");
}

function serializeTable(el: HTMLElement): string {
  const rows = [...el.querySelectorAll("tr")];
  if (rows.length === 0) return "";
  const lines: string[] = [""];
  rows.forEach((row, index) => {
    const cells = [...row.querySelectorAll("th,td")].map((cell) =>
      serialize(cell).replace(/\n+/g, " ").trim(),
    );
    lines.push(`| ${cells.join(" | ")} |`);
    if (index === 0) {
      lines.push(`| ${cells.map(() => "---").join(" | ")} |`);
    }
  });
  lines.push("");
  return lines.join("\n");
}

function serialize(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) {
    return escapeMd(node.textContent ?? "");
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return "";
  const el = node as HTMLElement;
  const tag = el.tagName.toLowerCase();
  const style = el.getAttribute("style") ?? "";
  const inner = [...el.childNodes].map(serialize).join("");

  if (tag === "span") {
    let out = inner;
    if (styleLooksBold(style)) out = `**${out}**`;
    if (styleLooksItalic(style)) out = `*${out}*`;
    return out;
  }

  switch (tag) {
    case "h1":
      return `\n\n# ${inner.trim()}\n\n`;
    case "h2":
      return `\n\n## ${inner.trim()}\n\n`;
    case "h3":
      return `\n\n### ${inner.trim()}\n\n`;
    case "h4":
      return `\n\n#### ${inner.trim()}\n\n`;
    case "h5":
      return `\n\n##### ${inner.trim()}\n\n`;
    case "h6":
      return `\n\n###### ${inner.trim()}\n\n`;
    case "p":
    case "div":
      return `\n\n${inner}\n\n`;
    case "br":
      return "  \n";
    case "strong":
    case "b":
      return `**${inner}**`;
    case "em":
    case "i":
      return `*${inner}*`;
    case "s":
    case "del":
    case "strike":
      return `~~${inner}~~`;
    case "code":
      return el.closest("pre") ? inner : `\`${inner}\``;
    case "pre": {
      const text = (el.textContent ?? "").replace(/\n$/, "");
      return `\n\n\`\`\`\n${text}\n\`\`\`\n\n`;
    }
    case "blockquote": {
      const quoted = inner
        .trim()
        .split("\n")
        .map((line) => `> ${line}`)
        .join("\n");
      return `\n\n${quoted}\n\n`;
    }
    case "a": {
      const href = el.getAttribute("href") ?? "";
      return `[${inner}](${href})`;
    }
    case "img": {
      const alt = el.getAttribute("alt") ?? "";
      const src = el.getAttribute("src") ?? "";
      return `![${alt}](${src})`;
    }
    case "ul":
      return (
        "\n" +
        [...el.children]
          .filter((child) => child.tagName.toLowerCase() === "li")
          .map((li) => `- ${serialize(li).trim()}`)
          .join("\n") +
        "\n"
      );
    case "ol":
      return (
        "\n" +
        [...el.children]
          .filter((child) => child.tagName.toLowerCase() === "li")
          .map((li, index) => `${index + 1}. ${serialize(li).trim()}`)
          .join("\n") +
        "\n"
      );
    case "li":
      return inner;
    case "hr":
      return "\n\n---\n\n";
    case "table":
      return serializeTable(el);
    case "thead":
    case "tbody":
    case "tr":
    case "td":
    case "th":
    case "body":
    case "html":
    case "section":
    case "article":
      return inner;
    default:
      return inner;
  }
}

export function htmlToMarkdown(html: string): string {
  if (typeof DOMParser === "undefined") return "";
  const cleaned = preprocessWordHtml(html);
  const doc = new DOMParser().parseFromString(cleaned, "text/html");
  const markdown = serialize(doc.body)
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
  return markdown.length > 0 ? `${markdown}\n` : "";
}
