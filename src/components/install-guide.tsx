import { useEffect, useState } from "react";
import { Download, Monitor, ShieldCheck } from "lucide-react";
import { AppHeader } from "@/components/app-header";
import { Button } from "@/components/ui/button";
import {
  buildManifestXml,
  downloadManifest,
  isLocalOrigin,
  normalizeAddinOrigin,
} from "@/lib/office/manifest";
import { useInklineHydration } from "@/lib/hooks";

const STEPS = [
  {
    title: "Download the manifest",
    body: "LS.md is a Lean Studio Word task pane add-in. The XML file tells Word where the pane lives and to request document access.",
  },
  {
    title: "Open a .docx in Word",
    body: "Use Microsoft 365 or Word 2021+ on Windows 11. Older .doc files do not host modern add-ins.",
  },
  {
    title: "Upload My Add-in",
    body: "Home or Insert → Add-ins → My Add-ins → Upload My Add-in. Choose ls-md-manifest.xml and allow it.",
  },
  {
    title: "Write in the pane",
    body: "LS.md docks on the right. With Live update on, the open document restyles as you type Markdown.",
  },
];

export function InstallGuide() {
  useInklineHydration();
  const [origin, setOrigin] = useState("");

  useEffect(() => {
    setOrigin(window.location.origin);
  }, []);

  const normalized = origin ? normalizeAddinOrigin(origin) : "";
  const httpWarning = origin ? isLocalOrigin(origin) : false;

  return (
    <div className="min-h-dvh bg-bg text-fg">
      <AppHeader title="Install for Word" />
      <main className="mx-auto w-full max-w-3xl px-4 py-10 sm:px-6 sm:py-14">
        <p className="text-xs tracking-[0.18em] text-muted uppercase">
          Windows 11 · Word desktop
        </p>
        <h1 className="mt-3 max-w-xl text-4xl leading-tight font-medium tracking-tight text-balance text-heading">
          Sideload LS.md into Word
        </h1>
        <p className="mt-4 max-w-xl text-base leading-relaxed text-pretty text-muted">
          This preview is a full Word workspace so you can try the pane now.
          The same pane installs into Microsoft Word as a{" "}
          <a
            href="https://lean.studio"
            className="text-accent hover:underline"
          >
            Lean Studio
          </a>{" "}
          Office add-in — no COM/.NET installer, just an XML manifest pointing
          at this host.
        </p>

        <div className="mt-8 rounded-xl bg-paper p-4 shadow-[var(--shadow-border)] sm:p-6">
          <label className="text-xs font-medium tracking-wide text-muted uppercase">
            Add-in host URL
          </label>
          <input
            value={origin}
            onChange={(event) => setOrigin(event.target.value)}
            className="mt-2 h-11 w-full rounded-md border border-border bg-bg-elevated px-3 font-mono text-sm outline-none focus:ring-2 focus:ring-accent/30"
            aria-label="Add-in host URL"
          />
          <p className="mt-2 text-sm text-muted">
            Word loads the pane from{" "}
            <span className="font-mono text-fg">{normalized || "…"}/taskpane</span>.
            Use this site’s published HTTPS address.
          </p>
          {httpWarning ? (
            <p className="mt-2 text-sm text-danger">
              Word rejects HTTP except on localhost. Publish the app, then
              paste the https URL before downloading.
            </p>
          ) : null}
          <div className="mt-4 flex flex-col gap-2 sm:flex-row">
            <Button
              onClick={() => downloadManifest(origin || window.location.origin)}
              className="h-11"
            >
              <Download className="size-4" />
              Download manifest
            </Button>
            <Button variant="secondary" className="h-11" asChild>
              <a href="/office/manifest.xml">Open XML in browser</a>
            </Button>
          </div>
        </div>

        <ol className="mt-10 space-y-4">
          {STEPS.map((step, index) => (
            <li
              key={step.title}
              className="flex gap-4 rounded-xl bg-bg-elevated p-4 shadow-[var(--shadow-border)] sm:p-5"
            >
              <span className="flex size-10 shrink-0 items-center justify-center rounded-md bg-paper font-medium tabular-nums text-heading">
                {index + 1}
              </span>
              <div>
                <h2 className="font-medium tracking-tight">{step.title}</h2>
                <p className="mt-1 text-sm leading-relaxed text-pretty text-muted">
                  {step.body}
                </p>
              </div>
            </li>
          ))}
        </ol>

        <section className="mt-10 grid gap-4 sm:grid-cols-2">
          <div className="rounded-xl bg-bg-elevated p-5 shadow-[var(--shadow-border)]">
            <Monitor className="size-5 text-accent" />
            <h2 className="mt-3 font-medium">Keep it installed</h2>
            <p className="mt-2 text-sm leading-relaxed text-muted">
              For a permanent catalog: put the XML in a shared folder such as
              {" "}
              <span className="font-mono text-fg">C:\Add-ins\LS.md</span>,
              share it, then Word → File → Options → Trust Center → Trusted
              Add-in Catalogs. Restart Word and add LS.md from Shared Folder.
            </p>
          </div>
          <div className="rounded-xl bg-bg-elevated p-5 shadow-[var(--shadow-border)]">
            <ShieldCheck className="size-5 text-accent" />
            <h2 className="mt-3 font-medium">Requirements</h2>
            <ul className="mt-2 list-disc space-y-1 pl-4 text-sm text-muted">
              <li>Windows 11 with Microsoft 365 or Word 2021+</li>
              <li>WordApi 1.3 (included in those builds)</li>
              <li>ReadWriteDocument permission when prompted</li>
              <li>HTTPS host, unless you are on localhost</li>
            </ul>
          </div>
        </section>

        <p className="mt-10 text-sm text-muted">
          Ribbon placement: Home → LS.md. The pane can be torn off and
          resized like any Word task pane. Pull from Word converts the
          current document back into Markdown.
        </p>

        <pre className="mt-6 overflow-x-auto rounded-lg bg-pane p-4 font-mono text-[0.7rem] leading-relaxed text-pane-muted">
          {origin ? buildManifestXml(origin).slice(0, 420) + "\n…" : "Loading manifest preview…"}
        </pre>
      </main>
    </div>
  );
}
