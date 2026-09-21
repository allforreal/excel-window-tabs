import { normalizeDisplayMode, type Bounds, type DisplayMode } from "./logic";

const STORAGE_KEY = "excel-window-tabs:v1";

export interface PersistedState {
  alignEnabled: boolean;
  frame: Bounds | null;
  originalBounds: Record<string, Bounds>;
  displayMode: DisplayMode;
}

export function defaultState(): PersistedState {
  return {
    alignEnabled: false,
    frame: null,
    originalBounds: {},
    displayMode: "full"
  };
}

export function loadState(): PersistedState {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return defaultState();
    }

    const parsed = JSON.parse(raw) as Partial<PersistedState>;
    return {
      alignEnabled: parsed.alignEnabled === true,
      frame: parsed.frame ?? null,
      originalBounds: parsed.originalBounds ?? {},
      displayMode: normalizeDisplayMode(parsed.displayMode)
    };
  } catch {
    return defaultState();
  }
}

export function saveState(state: PersistedState): void {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    // 存储不可用时忽略：标签功能本身不受影响。
  }
}
