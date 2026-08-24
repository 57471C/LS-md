import { createFileRoute } from "@tanstack/react-router";
import { buildManifestXml } from "@/lib/office/manifest";

export const Route = createFileRoute("/office/manifest.xml")({
  server: {
    handlers: {
      GET: async ({ request }) => {
        const url = new URL(request.url);
        const forwarded = request.headers.get("x-forwarded-host") ?? url.host;
        const proto =
          request.headers.get("x-forwarded-proto") ??
          url.protocol.replace(":", "");
        const origin = `${proto}://${forwarded}`;
        return new Response(buildManifestXml(origin), {
          headers: {
            "content-type": "text/xml; charset=utf-8",
            "cache-control": "no-store",
            "access-control-allow-origin": "*",
          },
        });
      },
    },
  },
});
