import { marked } from "marked";

marked.use({
  gfm: true,
  breaks: false,
});

export function markdownToHtml(markdown: string): string {
  try {
    return String(marked.parse(markdown ?? "", { async: false }));
  } catch {
    return "<p></p>";
  }
}

export function toWordHtml(html: string): string {
  const body = html
    .replaceAll("<strong>", "<b>")
    .replaceAll("</strong>", "</b>")
    .replaceAll("<em>", "<i>")
    .replaceAll("</em>", "</i>")
    .replaceAll("<del>", "<s>")
    .replaceAll("</del>", "</s>")
    .replaceAll(
      "<table>",
      '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%;">',
    );
  return `<div>${body}</div>`;
}

export function sanitizeHtml(html: string): string {
  if (typeof DOMParser === "undefined") return html;
  const parser = new DOMParser();
  const doc = parser.parseFromString(`<div id="inkline-root">${html}</div>`, "text/html");
  const root = doc.getElementById("inkline-root");
  if (!root) return "";

  root
    .querySelectorAll("script, iframe, object, embed, link, meta, form, style")
    .forEach((node) => node.remove());

  root.querySelectorAll("*").forEach((el) => {
    for (const attr of [...el.attributes]) {
      const name = attr.name.toLowerCase();
      const value = attr.value;
      if (name.startsWith("on") || name === "srcdoc") {
        el.removeAttribute(attr.name);
        continue;
      }
      if (
        (name === "href" || name === "src") &&
        /^\s*javascript:/i.test(value)
      ) {
        el.removeAttribute(attr.name);
      }
    }
  });

  return root.innerHTML;
}

export function renderDocumentHtml(markdown: string): string {
  return sanitizeHtml(markdownToHtml(markdown));
}

export function renderWordHtml(markdown: string): string {
  return toWordHtml(sanitizeHtml(markdownToHtml(markdown)));
}
