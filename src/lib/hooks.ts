import { useEffect, useState } from "react";
import { useInkline } from "@/lib/store";

export function useInklineHydration() {
  useEffect(() => {
    void Promise.resolve(useInkline.persist.rehydrate());
  }, []);
}

export function useIsDesktop(query = "(min-width: 900px)") {
  const [desktop, setDesktop] = useState(true);

  useEffect(() => {
    const media = window.matchMedia(query);
    const update = () => setDesktop(media.matches);
    update();
    media.addEventListener("change", update);
    return () => media.removeEventListener("change", update);
  }, [query]);

  return desktop;
}
