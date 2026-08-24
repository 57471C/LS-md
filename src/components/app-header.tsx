import { Link } from "@tanstack/react-router";
import { PanelRight } from "lucide-react";
import { LsMark, LsWordmark } from "@/components/ls-mark";
import { Button } from "@/components/ui/button";
import { useInkline } from "@/lib/store";
import { firstHeading } from "@/lib/utils";

export function AppHeader({
  title,
}: {
  title?: string;
}) {
  const draft = useInkline((s) => s.draft);
  const paneOpen = useInkline((s) => s.paneOpen);
  const setPaneOpen = useInkline((s) => s.setPaneOpen);
  const host = useInkline((s) => s.host);
  const docTitle = title ?? firstHeading(draft);

  return (
    <header className="flex h-14 shrink-0 items-center gap-3 border-b border-pane-border bg-pane px-3 text-pane-fg sm:px-5">
      <Link to="/" className="flex items-center gap-2.5">
        <LsMark />
        <LsWordmark className="text-base" />
      </Link>
      <a
        href="https://lean.studio"
        className="hidden font-lean text-sm tracking-tight sm:inline"
        aria-label="Lean Studio"
      >
        <span>LEAN</span>
        <span className="text-[0.78em] text-pane-muted">.STUDIO</span>
      </a>
      <span className="hidden text-pane-border sm:inline" aria-hidden>
        /
      </span>
      <p className="hidden min-w-0 truncate text-sm text-pane-muted sm:block">
        {docTitle}
      </p>
      <div className="ml-auto flex items-center gap-2">
        <span className="hidden rounded-full bg-pane-elevated px-2.5 py-1 text-[0.65rem] tracking-wide text-pane-muted uppercase sm:inline">
          {host === "word" ? "Word desktop" : "Win11 preview"}
        </span>
        <Button
          variant="pane-ghost"
          size="icon-lg"
          className="hidden min-[900px]:inline-flex"
          aria-label={paneOpen ? "Hide Markdown pane" : "Show Markdown pane"}
          onClick={() => setPaneOpen(!paneOpen)}
        >
          <PanelRight className="size-4" />
        </Button>
        <Button asChild size="sm" className="rounded-full">
          <Link to="/install">Install for Word</Link>
        </Button>
      </div>
    </header>
  );
}
