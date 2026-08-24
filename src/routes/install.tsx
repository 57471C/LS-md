import { createFileRoute } from "@tanstack/react-router";
import { InstallGuide } from "@/components/install-guide";

export const Route = createFileRoute("/install")({
  head: () => ({
    meta: [{ title: "Install LS.md for Word" }],
  }),
  component: InstallPage,
});

function InstallPage() {
  return <InstallGuide />;
}
