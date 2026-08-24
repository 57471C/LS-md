export type HostKind = "web" | "word";
export type SyncMode = "replace" | "insert";

type WordInsertLoc = "Start" | "End" | "Replace";

type WordBody = {
  clear: () => void;
  insertHtml: (html: string, loc: WordInsertLoc) => void;
  getHtml: () => { value?: string };
};

type WordRange = {
  insertHtml: (html: string, loc: WordInsertLoc) => void;
};

type WordContext = {
  document: {
    body: WordBody;
    getSelection: () => WordRange;
  };
  sync: () => Promise<void>;
};

type WordApi = {
  run: (callback: (context: WordContext) => Promise<void>) => Promise<void>;
};

type OfficeApi = {
  onReady: (callback: (info: { host?: string }) => void) => void;
};

function getOffice(): OfficeApi | undefined {
  return (window as unknown as { Office?: OfficeApi }).Office;
}

function getWord(): WordApi | undefined {
  return (window as unknown as { Word?: WordApi }).Word;
}

const OFFICE_READY_MS = 1800;
const OFFICE_JS = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";

function loadScript(src: string): Promise<void> {
  const existing = document.querySelector(`script[src="${src}"]`);
  if (existing) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("Could not load Office.js"));
    document.head.appendChild(script);
  });
}

export function isOfficeHostQuery(): boolean {
  return /(?:\?|&)_host_Info=/i.test(window.location.search);
}

export async function bootOfficeHost(): Promise<HostKind> {
  if (getOffice()?.onReady) return detectOfficeHost();
  if (!isOfficeHostQuery()) return "web";
  try {
    await loadScript(OFFICE_JS);
  } catch {
    return "web";
  }
  return detectOfficeHost();
}

export async function detectOfficeHost(): Promise<HostKind> {
  const office = getOffice();
  if (!office?.onReady) return "web";

  return await Promise.race([
    new Promise<HostKind>((resolve) => {
      office.onReady((info) => {
        const host = String(info?.host ?? "").toLowerCase();
        resolve(host.includes("word") ? "word" : "web");
      });
    }),
    new Promise<HostKind>((resolve) => {
      window.setTimeout(() => resolve("web"), OFFICE_READY_MS);
    }),
  ]);
}

export async function applyHtmlToWord(
  html: string,
  mode: SyncMode,
): Promise<void> {
  const Word = getWord();
  if (!Word?.run) {
    throw new Error("Word is not available in this window.");
  }
  await Word.run(async (context) => {
    if (mode === "insert") {
      context.document.getSelection().insertHtml(html, "Replace");
    } else {
      context.document.body.clear();
      context.document.body.insertHtml(html, "Start");
    }
    await context.sync();
  });
}

export async function readHtmlFromWord(): Promise<string> {
  const Word = getWord();
  if (!Word?.run) {
    throw new Error("Word is not available in this window.");
  }
  let html = "";
  await Word.run(async (context) => {
    const result = context.document.body.getHtml();
    await context.sync();
    html = result.value ?? "";
  });
  return html;
}
