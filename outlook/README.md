# LS.md for Outlook

Word add-in stays at the repo root. This folder is the Outlook compose pane and Cloudflare Worker.

`RequestedHeight` is illegal on `ItemEdit` (compose). On-prem Exchange rejects the manifest if it is present.

## Workflow

Work in Antigravity on this repo. Deploy from `outlook/`.

```powershell
cd outlook
npm install
npx wrangler login
npx wrangler deploy
```

Sideload only the live Worker URL:

`https://ls-md-outlook.<account>.workers.dev/manifest.xml`
