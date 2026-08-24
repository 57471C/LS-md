import { createFileRoute } from "@tanstack/react-router";
import { WordWorkspace } from "@/components/word-workspace";

export const Route = createFileRoute("/")({ component: Home });

function Home() {
  return <WordWorkspace />;
}
