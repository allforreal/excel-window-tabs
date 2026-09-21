export type WindowState = "maximized" | "minimized" | "normal";
export type DisplayMode = "full" | "compact";

export interface Bounds {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface WindowInfo {
  index: number;
  name: string;
  windowState: WindowState;
  isActive: boolean;
  geometry?: Bounds;
}

export interface WindowDiff {
  added: number[];
  removed: number[];
  needsRender: boolean;
}

export interface PollState {
  visible: boolean;
  focused: boolean;
  failStreak: number;
}

export const POLL_FOCUSED_MS = 800;
export const POLL_UNFOCUSED_MS = 3000;
export const POLL_BACKOFF_MAX_MS = 5000;
export const DEFAULT_EPSILON = 1;

const WINDOW_SUFFIX = /:\d+$/;

export function sortWindows(windows: WindowInfo[]): WindowInfo[] {
  return [...windows].sort((left, right) => left.index - right.index);
}

export function indexesOf(windows: WindowInfo[]): number[] {
  return sortWindows(windows).map((window) => window.index);
}

export function sameOrder(left: number[], right: number[]): boolean {
  return left.length === right.length && left.every((value, position) => value === right[position]);
}

export function displayName(window: WindowInfo): string {
  const raw = (window.name ?? "").trim();
  const cleaned = raw.replace(WINDOW_SUFFIX, "").trim();
  return cleaned.length > 0 ? cleaned : "工作簿";
}

export function uniqueLabels(windows: WindowInfo[]): Map<number, string> {
  const seen = new Map<string, number>();
  const labels = new Map<number, string>();

  for (const window of sortWindows(windows)) {
    const base = displayName(window);
    const count = (seen.get(base) ?? 0) + 1;
    seen.set(base, count);
    labels.set(window.index, count === 1 ? base : `${base} (${count})`);
  }

  return labels;
}

export function boundsFromWindow(window: WindowInfo): Bounds | null {
  if (!window.geometry) {
    return null;
  }
  return { ...window.geometry };
}

export function boundsEqual(left: Bounds, right: Bounds, epsilon: number = DEFAULT_EPSILON): boolean {
  return (
    Math.abs(left.left - right.left) <= epsilon &&
    Math.abs(left.top - right.top) <= epsilon &&
    Math.abs(left.width - right.width) <= epsilon &&
    Math.abs(left.height - right.height) <= epsilon
  );
}

export function diffWindows(previous: WindowInfo[], next: WindowInfo[]): WindowDiff {
  const previousByIndex = new Map(previous.map((window) => [window.index, window]));
  const nextByIndex = new Map(next.map((window) => [window.index, window]));

  const added: number[] = [];
  const removed: number[] = [];

  for (const index of nextByIndex.keys()) {
    if (!previousByIndex.has(index)) {
      added.push(index);
    }
  }
  for (const index of previousByIndex.keys()) {
    if (!nextByIndex.has(index)) {
      removed.push(index);
    }
  }

  let needsRender = added.length > 0 || removed.length > 0;

  if (!needsRender) {
    for (const [index, window] of nextByIndex) {
      const before = previousByIndex.get(index);
      if (!before) {
        continue;
      }
      if (
        before.name !== window.name ||
        before.windowState !== window.windowState ||
        before.isActive !== window.isActive
      ) {
        needsRender = true;
        break;
      }
    }
  }

  if (!needsRender && !sameOrder(indexesOf(previous), indexesOf(next))) {
    needsRender = true;
  }

  return {
    added: added.sort((left, right) => left - right),
    removed: removed.sort((left, right) => left - right),
    needsRender
  };
}

export function nextPollDelay(state: PollState): number | null {
  if (!state.visible) {
    return null;
  }

  const base = state.focused ? POLL_FOCUSED_MS : POLL_UNFOCUSED_MS;
  if (state.failStreak <= 0) {
    return base;
  }

  const doubled = base * 2 ** Math.min(state.failStreak, 4);
  return Math.min(doubled, POLL_BACKOFF_MAX_MS);
}

export function normalizeDisplayMode(value: unknown): DisplayMode {
  return value === "compact" ? "compact" : "full";
}

export function alignSummary(windowCount: number, alignEnabled: boolean): string {
  const countText = `${windowCount} 个窗口`;
  return alignEnabled ? `${countText} · 对齐模式已开启` : `${countText} · 对齐模式已关闭`;
}
