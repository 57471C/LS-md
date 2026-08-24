import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  type MouseEvent,
} from "react";
import {
  Bold,
  CheckSquare,
  Code,
  Download,
  FileText,
  Heading1,
  Heading2,
  Heading3,
  ImageIcon,
  Italic,
  Link2,
  List,
  ListOrdered,
  Minus,
  Quote,
  Strikethrough,
  Table,
  Upload,
} from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip";
import { LsMark, LsWordmark } from "@/components/ls-mark";
import {
  MarkdownEditor,
  type MarkdownEditorHandle,
} from "@/components/markdown-editor";
import {
  headingPrefix,
  insertSnippet,
  toggleLinePrefix,
  wrapSelection,
  type EditorPatch,
} from "@/lib/markdown/edit";
import { htmlToMarkdown } from "@/lib/markdown/from-html";
import { renderWordHtml } from "@/lib/markdown/parse";
import { TEMPLATES } from "@/lib/markdown/samples";
import {
  bootOfficeHost,
  applyHtmlToWord,
  readHtmlFromWord,
} from "@/lib/office/bridge";
import { useInkline } from "@/lib/store";
import { cn, countWords, formatClock } from "@/lib/utils";

type MarkdownPaneProps = {
  variant?: "workspace" | "taskpane";
  onClose?: () => void;
};

type Tool = {
  label: string;
  shortcut?: string;
  icon: typeof Bold;
  run: (value: string, start: number, end: number) => EditorPatch;
};

const TOOLS: Tool[] = [
  {
    label: "Heading 1",
    icon: Heading1,
    run: (v, s, e) => toggleLinePrefix(v, s, e, headingPrefix(1)),
  },
  {
    label: "Heading 2",
    icon: Heading2,
    run: (v, s, e) => toggleLinePrefix(v, s, e, headingPrefix(2)),
  },
  {
    label: "Heading 3",
    icon: Heading3,
    run: (v, s, e) => toggleLinePrefix(v, s, e, headingPrefix(3)),
  },
  {
    label: "Bold",
    shortcut: "Ctrl+B",
    icon: Bold,
    run: (v, s, e) => wrapSelection(v, s, e, "**"),
  },
  {
    label: "Italic",
    shortcut: "Ctrl+I",
    icon: Italic,
    run: (v, s, e) => wrapSelection(v, s, e, "*"),
  },
  {
    label: "Strikethrough",
    icon: Strikethrough,
    run: (v, s, e) => wrapSelection(v, s, e, "~~"),
  },
  {
    label: "Inline code",
    icon: Code,
    run: (v, s, e) => wrapSelection(v, s, e, "`"),
  },
  {
    label: "Link",
    shortcut: "Ctrl+K",
    icon: Link2,
    run: (v, s, e) => wrapSelection(v, s, e, "[", "](https://)"),
  },
  {
    label: "Image",
    icon: ImageIcon,
    run: (v, s, e) => insertSnippet(v, s, e, "![alt](https://)"),
  },
  {
    label: "Bulleted list",
    icon: List,
    run: (v, s, e) => toggleLinePrefix(v, s, e, "- "),
  },
  {
    label: "Numbered list",
    icon: ListOrdered,
    run: (v, s, e) => toggleLinePrefix(v, s, e, "1. "),
  },
  {
    label: "Checklist",
    icon: CheckSquare,
    run: (v, s, e) => toggleLinePrefix(v, s, e, "- [ ] "),
  },
  {
    label: "Quote",
    icon: Quote,
    run: (v, s, e) => toggleLinePrefix(v, s, e, "> "),
  },
  {
    label: "Code block",
    icon: Code,
    run: (v, s, e) =>
      insertSnippet(v, s, e, "\n```ts\n\n```\n"),
  },
  {
    label: "Table",
    icon: Table,
    run: (v, s, e) =>
      insertSnippet(v, s, e, "\n| Column | Column |\n| --- | --- |\n|  |  |\n"),
  },
  {
    label: "Horizontal rule",
    icon: Minus,
    run: (v, s, e) => insertSnippet(v, s, e, "\n\n---\n\n"),
  },
];

export function MarkdownPane({ variant = "workspace", onClose }: MarkdownPaneProps) {
  const editorRef = useRef<MarkdownEditorHandle>(null);
  const draft = useInkline((s) => s.draft);
  const applied = useInkline((s) => s.applied);
  const liveSync = useInkline((s) => s.liveSync);
  const syncMode = useInkline((s) => s.syncMode);
  const host = useInkline((s) => s.host);
  const syncState = useInkline((s) => s.syncState);
  const lastError = useInkline((s) => s.lastError);
  const lastSyncedAt = useInkline((s) => s.lastSyncedAt);
  const setDraft = useInkline((s) => s.setDraft);
  const apply = useInkline((s) => s.apply);
  const setLiveSync = useInkline((s) => s.setLiveSync);
  const setSyncMode = useInkline((s) => s.setSyncMode);
  const setHost = useInkline((s) => s.setHost);
  const loadTemplate = useInkline((s) => s.loadTemplate);
  const markSyncing = useInkline((s) => s.markSyncing);
  const markSynced = useInkline((s) => s.markSynced);
  const markError = useInkline((s) => s.markError);

  const words = useMemo(() => countWords(draft), [draft]);
  const inWord = host === "word";
  const dirty = draft !== applied;

  useEffect(() => {
    let cancelled = false;
    void bootOfficeHost().then((detected) => {
      if (!cancelled) setHost(detected);
    });
    return () => {
      cancelled = true;
    };
  }, [setHost]);

  useEffect(() => {
    if (host !== "word") return;
    const handle = window.setTimeout(() => {
      void (async () => {
        try {
          markSyncing();
          await applyHtmlToWord(renderWordHtml(applied), syncMode);
          markSynced();
        } catch (error) {
          const message =
            error instanceof Error ? error.message : "Could not update Word.";
          markError(message);
        }
      })();
    }, 420);
    return () => window.clearTimeout(handle);
  }, [applied, host, syncMode, markError, markSynced, markSyncing]);

  const applyPatch = useCallback((patch: EditorPatch) => {
    editorRef.current?.apply(patch);
  }, []);

  const runTool = useCallback(
    (tool: Tool) => {
      const handle = editorRef.current;
      const value = handle?.getValue() ?? draft;
      const sel = handle?.getSelection() ?? {
        start: value.length,
        end: value.length,
      };
      applyPatch(tool.run(value, sel.start, sel.end));
    },
    [applyPatch, draft],
  );

  const onApply = () => {
    apply();
    toast.success(inWord ? "Pushed to Word" : "Document updated");
  };

  const onPull = async () => {
    if (!inWord) {
      toast.message("Pull is available inside Word.");
      return;
    }
    try {
      markSyncing();
      const html = await readHtmlFromWord();
      const markdown = htmlToMarkdown(html);
      setDraft(markdown);
      apply();
      markSynced();
      toast.success("Pulled document into Markdown");
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Could not read Word.";
      markError(message);
      toast.error(message);
    }
  };

  const onDownload = () => {
    const blob = new Blob([draft], { type: "text/markdown;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "document.md";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  };

  const keepFocus = (event: MouseEvent) => {
    event.preventDefault();
  };

  const hostLabel = host === "word" ? "Word" : "Simulator";

  return (
    <section className="flex h-full min-h-0 flex-col overflow-hidden bg-pane text-pane-fg">
      <header className="flex shrink-0 items-center gap-2 border-b border-pane-border bg-pane-elevated px-3 py-2.5">
        <LsMark className="size-6" />
        <div className="min-w-0 flex-1">
          <p className="text-sm">
            <LsWordmark />
          </p>
          <p className="text-[0.7rem] text-pane-muted">Markdown for Word</p>
        </div>
        <span className="rounded-full bg-pane-elevated px-2 py-0.5 text-[0.65rem] tracking-wide text-pane-muted uppercase">
          {hostLabel}
        </span>
        {onClose ? (
          <Button variant="pane-ghost" size="icon" onClick={onClose} aria-label="Close pane">
            <Minus className="size-4 rotate-45" />
          </Button>
        ) : null}
      </header>

      <div className="flex shrink-0 flex-nowrap items-center gap-0.5 overflow-x-auto border-b border-pane-border bg-pane-elevated px-2 py-1.5">
        {TOOLS.map((tool) => (
          <Tooltip key={tool.label}>
            <TooltipTrigger asChild>
              <Button
                variant="pane-ghost"
                size="icon"
                aria-label={tool.label}
                onMouseDown={keepFocus}
                onClick={() => runTool(tool)}
              >
                <tool.icon className="size-3.5" />
              </Button>
            </TooltipTrigger>
            <TooltipContent>
              {tool.label}
              {tool.shortcut ? ` · ${tool.shortcut}` : ""}
            </TooltipContent>
          </Tooltip>
        ))}
        <DropdownMenu>
          <Tooltip>
            <TooltipTrigger asChild>
              <DropdownMenuTrigger asChild>
                <Button variant="pane-ghost" size="icon" aria-label="Templates">
                  <FileText className="size-3.5" />
                </Button>
              </DropdownMenuTrigger>
            </TooltipTrigger>
            <TooltipContent>Templates</TooltipContent>
          </Tooltip>
          <DropdownMenuContent align="end">
            {TEMPLATES.map((template) => (
              <DropdownMenuItem
                key={template.id}
                onSelect={() => loadTemplate(template.markdown)}
              >
                {template.label}
              </DropdownMenuItem>
            ))}
            <DropdownMenuSeparator />
            <DropdownMenuItem onSelect={onDownload}>
              <Download className="size-3.5" />
              Download .md
            </DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
      </div>

      <div className="relative min-h-0 flex-1">
        <MarkdownEditor
          ref={editorRef}
          value={draft}
          onChange={setDraft}
          onBold={() => runTool(TOOLS[3])}
          onItalic={() => runTool(TOOLS[4])}
          onLink={() => runTool(TOOLS[7])}
          onApply={onApply}
        />
      </div>

      <footer className="shrink-0 border-t border-pane-border px-3 py-3">
        <div className="mb-3 flex items-center justify-between gap-3">
          <label className="flex items-center gap-2 text-xs text-pane-muted">
            <button
              type="button"
              role="switch"
              aria-checked={liveSync}
              aria-label="Live update"
              onClick={() => setLiveSync(!liveSync)}
              className={cn(
                "inline-flex h-6 w-10 shrink-0 items-center rounded-full border px-0.5 transition-colors duration-150",
                liveSync
                  ? "border-accent bg-accent"
                  : "border-pane-border bg-pane-elevated",
              )}
            >
              <span
                className={cn(
                  "block size-4 rounded-full bg-pane-fg transition-transform duration-150",
                  liveSync && "translate-x-4 bg-accent-fg",
                )}
              />
            </button>
            Live update
          </label>
          <div className="flex rounded-full bg-pane-elevated p-0.5">
            <button
              type="button"
              onClick={() => setSyncMode("replace")}
              className={cn(
                "rounded-full px-2.5 py-1 text-[0.7rem]",
                syncMode === "replace"
                  ? "bg-accent text-accent-fg"
                  : "text-pane-muted",
              )}
            >
              Replace doc
            </button>
            <button
              type="button"
              onClick={() => setSyncMode("insert")}
              className={cn(
                "rounded-full px-2.5 py-1 text-[0.7rem]",
                syncMode === "insert"
                  ? "bg-accent text-accent-fg"
                  : "text-pane-muted",
              )}
            >
              Insert at cursor
            </button>
          </div>
        </div>
        <div className="flex gap-2">
          <Button
            variant="pane-primary"
            className="flex-1"
            onClick={onApply}
          >
            Apply to document
          </Button>
          <Button
            variant="pane"
            onClick={() => void onPull()}
            aria-label="Pull from Word"
          >
            <Upload className="size-3.5" />
          </Button>
        </div>
        <div className="mt-2 flex items-center justify-between text-[0.7rem] text-pane-muted tabular-nums">
          <span>
            {words} words
            {variant === "taskpane" ? " · Ctrl+S apply" : ""}
          </span>
          <span
            className={cn(
              syncState === "error" && "text-danger",
              syncState === "syncing" && "text-pane-fg",
            )}
          >
            {syncState === "syncing"
              ? "Updating…"
              : syncState === "error"
                ? lastError
                : liveSync
                  ? inWord
                    ? `Live · ${formatClock(lastSyncedAt)}`
                    : "Live"
                  : dirty
                    ? "Unapplied"
                    : "In sync"}
          </span>
        </div>
      </footer>
    </section>
  );
}
