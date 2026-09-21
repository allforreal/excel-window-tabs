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
 * macOS 版 Excel 的窗口几何写入只对前台窗口生效，因此这里改成
 * “先激活、待窗口稳定后再套用对齐基准”，避免写入被激活动作覆盖。
 */
export async function activateWindow(index: number, frame?: Bounds | null): Promise<void> {
  await Excel.run(async (context) => {
    const windows = context.workbook.application.windows;
    windows.load("items/index,items/windowState");
    await context.sync();

    const target = windows.items.find((item) => item.index === index);
    if (!target) {
      throw new Error("窗口已不存在，请刷新后重试。");
    }

    if (String(target.windowState) !== "normal") {
      const needsRestore = windows.items.some((item) => String(item.windowState) !== "normal");
      if (needsRestore) {
        for (const item of windows.items) {
          if (String(item.windowState) !== "normal") {
            item.windowState = "normal";
          }
        }
        await context.sync();
      }
    }

    target.activate();
    await context.sync();
  });

  if (frame) {
    await delay(ACTIVATE_SETTLE_MS);
    await alignWindow(index, frame);
  }
}

/** 把对齐基准套用到单个窗口（调用方需保证该窗口处于前台，几何写入才会生效）。 */
export async function alignWindow(index: number, frame: Bounds): Promise<void> {
  await Excel.run(async (context) => {
    const windows = context.workbook.application.windows;
    windows.load("items/index,items/windowState");
    await context.sync();

    const target = windows.items.find((item) => item.index === index);
    if (!target) {
      return;
    }

    if (String(target.windowState) !== "normal") {
      target.windowState = "normal";
      await context.sync();
    }

    applyBounds(target, frame);
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

export async function closeWindow(index: number): Promise<void> {
  await Excel.run(async (context) => {
    const windows = context.workbook.application.windows;
    windows.load("items/index");
    await context.sync();

    const target = windows.items.find((item) => item.index === index);
    if (!target) {
      throw new Error("窗口已不存在。");
    }

    target.close();
    await context.sync();
  });
}

export async function createWorkbook(): Promise<void> {
  await Excel.createWorkbook();
}
