# LS.md

**Markdown for Outlook.** A [Lean Studio](https://lean.studio) compose-pane add-in.

Type GitHub-flavoured Markdown in the task pane — headings, lists, tables, quotes, emphasis, checklists, and fenced code. **Apply to email** writes HTML into the message.

```
LS.md pane     Outlook (classic, on-prem Exchange)
Markdown       →  HTML block in the compose body
```

Apply modes:

- **Update block** — replace the previous LS.md block, leave signature and quoted thread alone
- **Insert at cursor** — insert at the caret
- **Replace body** — replace the whole body

**Live update** applies as you type. **Pull** reads the current block back to Markdown. The pane opens blank.

## Requirements

- Outlook classic on Windows (tested on Version 2609 build 20430 with on-prem Exchange)
- Mailbox requirement set 1.3
- HTTPS host for the pane (Cloudflare Worker in this repo)

Not Word. Permission is `ReadWriteItem`, add-in id `a7c4e2b1-6d38-4f91-9c2a-8b5e1d0f3a47`.

On-prem Exchange rejects `RequestedHeight` on `ItemEdit` and `SupportsPinning` on VersionOverrides 1.0. The generated manifest omits both.

## Sideload

1. From `outlook/` run `npx wrangler deploy`.
2. Open the live manifest and save it:
   `https://ls-md-outlook.terry-b10.workers.dev/manifest.xml`
3. Outlook: **Get Add-ins → My Add-ins → Add a custom add-in → Add from File**. Pick that XML.
4. Open a **new mail**. **LS.md** is on the compose ribbon.

On-prem often cannot sideload from a URL. Use the saved file. Re-add the file after a manifest version bump.

Trusted catalog (shared PC): put the XML in a folder such as `C:\Add-ins\LS.md`, then File → Options → Trust Center → Trusted Add-in Catalogs.

## Deploy

Work in this repo. Deploy only from `outlook/`.

```powershell
cd outlook
npm install
npx wrangler login   # first time
npx wrangler deploy
```

The Worker serves `public/` and builds `/manifest.xml` so icon and pane URLs match the Worker origin.

## Source

| Path | Role |
| --- | --- |
| [`outlook/public/taskpane.html`](outlook/public/taskpane.html) | Task pane shell |
| [`outlook/public/taskpane.js`](outlook/public/taskpane.js) | Markdown → HTML, Office.js apply / pull |
| [`outlook/public/taskpane.css`](outlook/public/taskpane.css) | Pane + preview styles |
| [`outlook/public/vendor/`](outlook/public/vendor/) | marked + highlight.js |
| [`outlook/public/icons/`](outlook/public/icons/) | Ribbon icons 16 / 32 / 80 |
| [`outlook/public/commands.html`](outlook/public/commands.html) | Function file |
| [`outlook/worker.js`](outlook/worker.js) | Manifest + static assets |
| [`outlook/wrangler.toml`](outlook/wrangler.toml) | Cloudflare Worker |

`src/` and `office/` are the older Word pane. They are not what Outlook loads.
