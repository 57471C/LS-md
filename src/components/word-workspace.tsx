import { Group, Panel, Separator } from "react-resizable-panels";
import { AppHeader } from "@/components/app-header";
import { DocumentPage } from "@/components/document-page";
import { MarkdownPane } from "@/components/markdown-pane";
import { Button } from "@/components/ui/button";
import { useInklineHydration } from "@/lib/hooks";
import { useInkline } from "@/lib/store";
import { cn } from "@/lib/utils";

export function WordWorkspace() {
  useInklineHydration();
  const paneOpen = useInkline((s) => s.paneOpen);
  const setPaneOpen = useInkline((s) => s.setPaneOpen);
  const mobileTab = useInkline((s) => s.mobileTab);
  const setMobileTab = useInkline((s) => s.setMobileTab);

  return (
    <div className="relative flex h-dvh min-h-0 flex-col overflow-hidden bg-bg text-fg">
      <AppHeader />
      <div className="hidden min-h-0 min-[900px]:flex min-[900px]:flex-1">
        <Group className="h-full w-full" orientation="horizontal">
          <Panel
            id="document"
            minSize={360}
            defaultSize="68"
            className="min-w-0 overflow-hidden"
          >
            <div className="h-full overflow-auto bg-canvas px-4 py-8 sm:px-10">
              <DocumentPage />
            </div>
          </Panel>
          {paneOpen ? (
            <>
              <Separator className="w-1 bg-border hover:bg-accent/40" />
              <Panel
                id="pane"
                minSize={280}
                defaultSize="32"
                className="min-w-0 overflow-hidden bg-pane"
              >
                <MarkdownPane
                  variant="workspace"
                  onClose={() => setPaneOpen(false)}
                />
              </Panel>
            </>
          ) : null}
        </Group>
      </div>
      <div className="min-[900px]:hidden max-[899px]:flex max-[899px]:min-h-0 max-[899px]:flex-1 max-[899px]:flex-col">
        <div className="flex border-b border-border bg-bg-elevated p-1">
          {(["write", "document"] as const).map((tab) => (
            <button
              key={tab}
              type="button"
              onClick={() => setMobileTab(tab)}
              className={cn(
                "h-11 flex-1 rounded-sm text-sm font-medium",
                mobileTab === tab
                  ? "bg-paper text-fg shadow-[var(--shadow-border)]"
                  : "text-muted",
              )}
            >
              {tab === "write" ? "Markdown" : "Document"}
            </button>
          ))}
        </div>
        <div className="min-h-0 flex-1 overflow-hidden">
          {mobileTab === "write" ? (
            <MarkdownPane variant="workspace" />
          ) : (
            <div className="h-full overflow-auto bg-canvas px-3 py-6">
              <DocumentPage />
            </div>
          )}
        </div>
      </div>
      {!paneOpen ? (
        <div className="pointer-events-none absolute right-4 bottom-4 hidden min-[900px]:block">
          <Button
            className="pointer-events-auto shadow-[var(--shadow-border)]"
            onClick={() => setPaneOpen(true)}
          >
            Open Markdown pane
          </Button>
        </div>
      ) : null}
    </div>
  );
}
