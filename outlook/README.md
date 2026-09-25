# Outlook Worker

Compose-pane add-in + Cloudflare Worker. Product notes are in the root [README](../README.md).

`RequestedHeight` is illegal on `ItemEdit`. `SupportsPinning` is illegal on VersionOverrides 1.0. On-prem Exchange rejects the manifest if either is present. `worker.js` generates XML without them.

## Deploy

```powershell
cd outlook
npm install
npx wrangler login   # first time
npx wrangler deploy
```

Sideload the live manifest file (not a localhost URL):

`https://ls-md-outlook.terry-b10.workers.dev/manifest.xml`

Then a **new mail**, not a window that was already open.
