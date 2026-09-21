import type { Bounds, WindowInfo, WindowState } from "./logic";

const WINDOW_PROPS =
  "items/index,items/name,items/left,items/top,items/width,items/height,items/windowState";

export interface WindowListResult {
  windows: WindowInfo[];
  activeIndex: number | null;
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

export async function listWindows(): Promise<WindowListResult> {
  return Excel.run(async (context) => {
    const application = context.workbook.application;
    const windows = application.windows;
    const activeWindow = application.activeWindow;

    windows.load(WINDOW_PROPS);
    activeWindow.load("index");
    await context.sync();

    const activeIndex = activeWindow.index ?? null;
    const items: WindowInfo[] = windows.items.map((item) => ({
      index: item.index,
      name: item.name,
      left: item.left,
      top: item.top,
      width: item.width,
      height: item.height,
      windowState: normalizeWindowState(String(item.windowState)),
      isActive: item.index === activeIndex
    }));

    return { windows: items, activeIndex };
  });
}

function applyBounds(item: Excel.Window, frame: Bounds): void {
  item.left = frame.left;
  item.top = frame.top;
  item.width = frame.width;
  item.height = frame.height;
}

async function findWindow(
  context: Excel.RequestContext,
  index: number
): Promise<Excel.Window | undefined> {
  const windows = context.workbook.application.windows;
  windows.load("items/index,items/windowState");
  await context.sync();
  return windows.items.find((item) => item.index === index);
}

export async function activateWindow(index: number, frame?: Bounds | null): Promise<void> {
  await Excel.run(async (context) => {
    const windows = context.workbook.application.windows;
    windows.load("items/index,items/windowState");
    await context.sync();

    const target = windows.items.find((item) => item.index === index);
    if (!target) {
      throw new Error("窗口已不存在，请刷新后重试。");
    }

    if (frame) {
      for (const item of windows.items) {
        if (String(item.windowState) !== "normal") {
          item.windowState = "normal";
        }
      }
      await context.sync();

      for (const item of windows.items) {
        applyBounds(item, frame);
      }
    } else if (String(target.windowState) === "minimized") {
      target.windowState = "normal";
      await context.sync();
    }

    target.activate();
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

    for (const item of targets) {
      if (String(item.windowState) !== "normal") {
        item.windowState = "normal";
      }
    }
    await context.sync();

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

    for (const item of restorable) {
      if (String(item.windowState) !== "normal") {
        item.windowState = "normal";
      }
    }
    await context.sync();

    for (const item of restorable) {
      applyBounds(item, bounds[String(item.index)]);
    }
    await context.sync();
  });
}

export async function closeWindow(index: number): Promise<void> {
  await Excel.run(async (context) => {
    const target = await findWindow(context, index);
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
