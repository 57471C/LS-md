import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import { SAMPLE_BRIEF } from "@/lib/markdown/samples";
import type { HostKind, SyncMode } from "@/lib/office/bridge";

type SyncState = "idle" | "syncing" | "error";

type InklineState = {
  draft: string;
  applied: string;
  liveSync: boolean;
  syncMode: SyncMode;
  host: HostKind;
  paneOpen: boolean;
  mobileTab: "write" | "document";
  syncState: SyncState;
  lastError: string | null;
  lastSyncedAt: number | null;
  setDraft: (value: string) => void;
  apply: () => void;
  setLiveSync: (value: boolean) => void;
  setSyncMode: (value: SyncMode) => void;
  setHost: (value: HostKind) => void;
  setPaneOpen: (value: boolean) => void;
  setMobileTab: (value: "write" | "document") => void;
  loadTemplate: (value: string) => void;
  markSyncing: () => void;
  markSynced: () => void;
  markError: (message: string) => void;
};

export const useInkline = create<InklineState>()(
  persist(
    (set, get) => ({
      draft: SAMPLE_BRIEF,
      applied: SAMPLE_BRIEF,
      liveSync: true,
      syncMode: "replace" as SyncMode,
      host: "web",
      paneOpen: true,
      mobileTab: "write",
      syncState: "idle",
      lastError: null,
      lastSyncedAt: null,
      setDraft: (value) => {
        const live = get().liveSync;
        set({
          draft: value,
          applied: live ? value : get().applied,
          lastError: null,
        });
      },
      apply: () => set({ applied: get().draft, lastError: null }),
      setLiveSync: (value) =>
        set({
          liveSync: value,
          applied: value ? get().draft : get().applied,
        }),
      setSyncMode: (value) => set({ syncMode: value }),
      setHost: (value) => set({ host: value }),
      setPaneOpen: (value) => set({ paneOpen: value }),
      setMobileTab: (value) => set({ mobileTab: value }),
      loadTemplate: (value) => {
        const live = get().liveSync;
        set({
          draft: value,
          applied: live ? value : get().applied,
          lastError: null,
        });
      },
      markSyncing: () => set({ syncState: "syncing" }),
      markSynced: () =>
        set({
          syncState: "idle",
          lastSyncedAt: Date.now(),
          lastError: null,
        }),
      markError: (message) => set({ syncState: "error", lastError: message }),
    }),
    {
      name: "ls-md-v1",
      storage: createJSONStorage(() => localStorage),
      skipHydration: true,
      partialize: (state) => ({
        draft: state.draft,
        applied: state.applied,
        liveSync: state.liveSync,
        syncMode: state.syncMode,
      }),
    },
  ),
);
