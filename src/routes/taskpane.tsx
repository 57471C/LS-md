import { createFileRoute, Link } from "@tanstack/react-router";
import { DocumentPage } from "@/components/document-page";
import { MarkdownPane } from "@/components/markdown-pane";
import { useInklineHydration } from "@/lib/hooks";
import { useInkline } from "@/lib/store";

export const Route = createFileRoute("/taskpane")({
  head: () => ({
    meta: [
      { title: "LS.md" },
      { name: "viewport", content: "width=device-width, initial-scale=1" },
    ],
  }),
  component: TaskpanePage,
});

function TaskpanePage() {
  useInklineHydration();
  const host = useInkline((s) => s.host);
  const showPreview = host !== "word";

  return (
    <div className="flex h-dvh min-h-0 flex-col overflow-hidden bg-pane">
      <div
        className={
          showPreview
            ? "flex h-full min-h-0 flex-[1.2] flex-col overflow-hidden"
            : "flex h-full min-h-0 flex-1 flex-col overflow-hidden"
        }
      >
        <MarkdownPane variant="taskpane" />
      </div>
      {showPreview ? (
        <div className="flex min-h-0 flex-1 flex-col overflow-hidden border-t border-pane-border bg-canvas">
          <div className="flex items-center justify-between px-3 py-2">
            <p className="text-xs text-muted">Document preview</p>
            <Link to="/" className="text-xs text-accent hover:underline">
              Full workspace
            </Link>
          </div>
          <div className="min-h-0 flex-1 overflow-auto px-3 pb-3">
            <DocumentPage compact />
          </div>
        </div>
      ) : null}
    </div>
  );
}
