import { cn } from "@/lib/utils";

export function LsMark({ className }: { className?: string }) {
  return (
    <svg
      viewBox="0 0 32 32"
      className={cn("size-7", className)}
      aria-hidden="true"
    >
      <rect width="32" height="32" rx="7" fill="#0a0a0a" />
      <rect
        x="1.1"
        y="1.1"
        width="29.8"
        height="29.8"
        rx="6.2"
        fill="none"
        stroke="#1f7a54"
        strokeWidth="1.4"
      />
      <path
        d="M9 8.2h2.6v12.2H19.2V23H9z"
        fill="#f4f2ec"
      />
      <rect x="21.2" y="20.6" width="3.6" height="3.6" rx="0.7" fill="#1f7a54" />
    </svg>
  );
}

export function LsWordmark({
  className,
  mutedClassName,
}: {
  className?: string;
  mutedClassName?: string;
}) {
  return (
    <span className={cn("font-lean tracking-tight", className)}>
      <span>LS</span>
      <span className={cn("text-[0.82em] text-accent", mutedClassName)}>
        .md
      </span>
    </span>
  );
}
