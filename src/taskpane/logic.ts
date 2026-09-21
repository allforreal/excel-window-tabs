export type WindowState = "maximized" | "minimized" | "normal";

export interface WindowInfo {
  index: number;
  name: string;
  left: number;
  top: number;
  width: number;
  height: number;
  windowState: WindowState;
  isActive: boolean;
}

export interface Bounds {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface WindowDiff {
  added: number[];
  removed: number[];
  needsRender: boolean;
}

export const DEFAULT_EPSILON = 1;

const WINDOW_SUFFIX = /:\d+$/;

export function sortWindows(windows: WindowInfo[]): WindowInfo[] {
  return [...windows].sort((left, right) => left.index - right.index);
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

export function boundsFromWindow(window: WindowInfo): Bounds {
  return {
    left: window.left,
    top: window.top,
    width: window.width,
    height: window.height
  };
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

  if (!needsRender) {
    const previousOrder = sortWindows(previous)
      .map((window) => window.index)
      .join(",");
    const nextOrder = sortWindows(next)
      .map((window) => window.index)
      .join(",");
    if (previousOrder !== nextOrder) {
      needsRender = true;
    }
  }

  return {
    added: added.sort((left, right) => left - right),
    removed: removed.sort((left, right) => left - right),
    needsRender
  };
}

export function alignSummary(windowCount: number, alignEnabled: boolean): string {
  const countText = `${windowCount} 个窗口`;
  return alignEnabled ? `${countText} · 对齐模式已开启` : `${countText} · 对齐模式已关闭`;
}
