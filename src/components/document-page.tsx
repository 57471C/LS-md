import { useMemo } from "react";
import { useInkline } from "@/lib/store";
import { renderDocumentHtml } from "@/lib/markdown/parse";
import { cn, firstHeading } from "@/lib/utils";

export function DocumentPage({ compact = false }: { compact?: boolean }) {
  const applied = useInkline((s) => s.applied);
  const html = useMemo(() => renderDocumentHtml(applied), [applied]);
  const title = firstHeading(applied);
  const empty = applied.trim().length === 0;

  return (
    <article
      className={cn(
        "paper-page mx-auto w-full",
        compact ? "max-w-none rounded-sm px-6 py-8" : "paper-letter rounded-sm px-8 py-10 sm:px-14 sm:py-16",
      )}
    >
      <p className="mb-8 font-sans text-xs tracking-[0.18em] text-subtle uppercase">
        {title}
      </p>
      {empty ? (
        <p className="font-sans text-sm text-muted">
          Start writing in the Markdown pane. This page updates with the
          document.
        </p>
      ) : (
        <div
          className="word-doc"
          dangerouslySetInnerHTML={{ __html: html }}
        />
      )}
      <p className="mt-16 text-center font-sans text-xs text-subtle tabular-nums">
        1
      </p>
    </article>
  );
}
