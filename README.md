# LS.md

**Markdown for Word.** A [Lean Studio](https://lean.studio) task pane add-in.

Type GitHub-flavored Markdown in a pane; the open Word document updates live — headings, lists, tables, quotes, emphasis, and fenced code.

```
LS.md          Word on Windows 11
Markdown pane  →  live document
```

Provider: **Lean Studio** · Ribbon: **Home → LS.md**

## Sideload

Requires Microsoft 365 or Word 2021+ on Windows 11 (WordApi 1.3). The pane is a web page; Word loads it from an HTTPS host (HTTP is allowed only on localhost).

1. Host this app over HTTPS.
2. Open `/office/manifest.xml` on that host, or fill `office/ls-md-manifest.template.xml` and replace `YOUR-HTTPS-HOST`.
3. In Word: **Home** or **Insert → Add-ins → My Add-ins → Upload My Add-in**. Choose `ls-md-manifest.xml`.
4. The **LS.md** button appears on the Home tab. Live update writes every change; turn it off to draft, then **Apply to document**.

For a shared catalog: put the XML in a folder such as `C:\Add-ins\LS.md`, then Word → File → Options → Trust Center → Trusted Add-in Catalogs.

## Source

| Path | Role |
| --- | --- |
| [`src/lib/office/bridge.ts`](src/lib/office/bridge.ts) | Office.js host detect, `insertHtml` / `getHtml` |
| [`src/lib/office/manifest.ts`](src/lib/office/manifest.ts) | Task pane manifest (GUID, icons, Home ribbon) |
| [`src/lib/markdown/parse.ts`](src/lib/markdown/parse.ts) | GFM → Word-friendly HTML |
| [`src/lib/markdown/from-html.ts`](src/lib/markdown/from-html.ts) | Pull the document back to Markdown |
| [`src/lib/markdown/edit.ts`](src/lib/markdown/edit.ts) | Toolbar wrap / prefix helpers |
| [`src/components/markdown-editor.tsx`](src/components/markdown-editor.tsx) | CodeMirror Markdown highlighting |
| [`src/components/markdown-pane.tsx`](src/components/markdown-pane.tsx) | Task pane chrome and live sync |
| [`src/components/word-workspace.tsx`](src/components/word-workspace.tsx) | Desktop / mobile Word workspace |
| [`public/office/`](public/office/) | Ribbon icons 16 / 32 / 80 |
| [`office/ls-md-manifest.template.xml`](office/ls-md-manifest.template.xml) | Sideload XML (replace the host) |

Permissions: `ReadWriteDocument`. Add-in id: `8f3c1e2a-9b47-4d6e-a1f0-5c2d8e7b4a91`.

## Brand

LS.md follows the Lean Studio `LS.*` naming (LS.TimeStudy, LS.Video, …). Product colour is `#1f7a54`. Wordmark: **LS.md**.

See [lean.studio](https://lean.studio).
