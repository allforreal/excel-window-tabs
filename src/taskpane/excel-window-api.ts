import type { Bounds, WindowInfo, WindowState } from "./logic";

const BASE_PROPS = "items/index,items/name,items/windowState";
const GEOMETRY_PROPS = ",items/left,items/top,items/width,items/height";
const ACTIVATE_SETTLE_MS = 140;

function delay(milliseconds: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, milliseconds);
  });
}

export interface WindowListResult {
  windows: WindowInfo[];
  activeIndex: number | null;
}

export interface ListWindowOptions {
  withGeometry?: boolean;
}

/**
 * 目标窗口引用。
 *
 * `Window.index` 是位置索引，活动窗口永远是 1 号，激活/开关窗口都会让它变化，
 * 因此匹配时优先使用更稳定的文档名 `name`，索引只作为兜底。
 */
export interface WindowTarget {
  index: number;
  name: string;
}

export function isSupported(): boolean {
  const requirements = Office?.context?.requirements;
  return Boolean(requirements && requirements.isSetSupported("ExcelApiDesktop", "1.1"));
}

function normalizeWindowState(value: string): WindowState {
  if (value === "maximized" || value === "minimized") {
    return value;
  }
  return "normal";
}

function matchWindow(items: Excel.Window[], target: WindowTarget): Excel.Window | undefined {
  const name = (target.name ?? "").trim();
  if (name) {
    const byName = items.find((item) => (item.name ?? "").trim() === name);
    if (byName) {
      return byName;
    }
  }
  return items.find((item) => item.index === target.index);
}

export async function listWindows(options?: ListWindowOptions): Promise<WindowListResult> {
  const withGeometry = options?.withGeometry === true;

  return Excel.run(async (context) => {
    const application = context.workbook.application;
    const windows = application.windows;
    const activeWindow = application.activeWindow;

    windows.load(withGeometry ? BASE_PROPS + GEOMETRY_PROPS : BASE_PROPS);
    activeWindow.load("index");
    await context.sync();

    const activeIndex = activeWindow.index ?? null;
    const items: WindowInfo[] = windows.items.map((item) => {
      const info: WindowInfo = {
        index: item.index,
        name: item.name,
        windowState: normalizeWindowState(String(item.windowState)),
        isActive: item.index === activeIndex
      };

      if (withGeometry) {
        info.geometry = {
          left: item.left,
          top: item.top,
          width: item.width,
          height: item.height
        };
      }

      return info;
    });

    return { windows: items, activeIndex };
  });
}

function applyBounds(item: Excel.Window, frame: Bounds): void {
  item.left = frame.left;
  item.top = frame.top;
  item.width = frame.width;
  item.height = frame.height;
}

export async function getActiveWindowBounds(): Promise<Bounds | null> {
  return Excel.run(async (context) => {
    const activeWindow = context.workbook.application.activeWindow;
    activeWindow.load("left,top,width,height");
    await context.sync();

    return {
      left: activeWindow.left,
      top: activeWindow.top,
      width: activeWindow.width,
      height: activeWindow.height
    };
  });
}

/**
 * 切换到目标窗口。
 *
 * macOS 版 Excel 的窗口几何写入只对前台窗口生效，而且窗口索引会在激活后变化，
 * 因此这里先按文档名激活目标窗口，待窗口稳定后再对齐"当前活动窗口"。
 */
export async function activateWindow(target: WindowTarget, frame?: Bounds | null): Promise<void> {
  await Excel.run(async (context) => {
    const windows = context.workbook.application.windows;
    windows.load("items/index,items/name,items/windowState");
    await context.sync();

    const window = matchWindow(windows.items, target);
    if (!window) {
      throw new Error("窗口已不存在，请刷新后重试。");
    }

    if (String(window.windowState) !== "normal") {
      for (const item of windows.items) {
        if (String(item.windowState) !== "normal") {
          item.windowState = "normal";
        }
      }
      await context.sync();
    }

    window.activate();
    await context.sync();
  });

  if (frame) {
    await delay(ACTIVATE_SETTLE_MS);
    await alignActiveWindow(frame);
  }
}

/**
 * 把对齐基准套用到当前活动窗口。
 *
 * 故意不按索引查找：`Window.index` 会随激活变化，激活后再按索引找会命中另一个窗口。
 */
export async function alignActiveWindow(frame: Bounds): Promise<void> {
  await Excel.run(async (context) => {
    const activeWindow = context.workbook.application.activeWindow;
    activeWindow.load("windowState");
    await context.sync();

    if (String(activeWindow.windowState) !== "normal") {
      activeWindow.windowState = "normal";
      await context.sync();
    }

    applyBounds(activeWindow, frame);
    await context.sync();
  });
}

export async function applyFrameToAll(frame: Bounds, onlyIndexes?: number[]): Promise<void> {
  await Excel.run(async (context) => {
    const windows = context.workbook.application.windows;
    windows.load("items/index,items/windowState");
    await context.sync();

    const targets =
      onlyIndexes && onlyIndexes.length > 0
        ? windows.items.filter((item) => onlyIndexes.includes(item.index))
        : windows.items;

    const needsRestore = targets.some((item) => String(item.windowState) !== "normal");
    if (needsRestore) {
      for (const item of targets) {
        if (String(item.windowState) !== "normal") {
          item.windowState = "normal";
        }
      }
      await context.sync();
    }

    for (const item of targets) {
      applyBounds(item, frame);
    }
    await context.sync();
  });
}

export async function restoreBounds(bounds: Record<string, Bounds>): Promise<void> {
  await Excel.run(async (context) => {
    const windows = context.workbook.application.windows;
    windows.load("items/index,items/windowState");
    await context.sync();

    const restorable = windows.items.filter((item) => Boolean(bounds[String(item.index)]));
    if (restorable.length === 0) {
      return;
    }

    const needsRestore = restorable.some((item) => String(item.windowState) !== "normal");
    if (needsRestore) {
      for (const item of restorable) {
        if (String(item.windowState) !== "normal") {
          item.windowState = "normal";
        }
      }
      await context.sync();
    }

    for (const item of restorable) {
      applyBounds(item, bounds[String(item.index)]);
    }
    await context.sync();
  });
}

export async function closeWindow(target: WindowTarget): Promise<void> {
  await Excel.run(async (context) => {
    const windows = context.workbook.application.windows;
    windows.load("items/index,items/name");
    await context.sync();

    const window = matchWindow(windows.items, target);
    if (!window) {
      throw new Error("窗口已不存在。");
    }

    window.close();
    await context.sync();
  });
}

export async function createWorkbook(): Promise<void> {
  await Excel.createWorkbook();
}
