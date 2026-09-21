import {
  alignSummary,
  boundsFromWindow,
  diffWindows,
  sortWindows,
  uniqueLabels,
  type Bounds,
  type WindowInfo
} from "./logic";
import * as api from "./excel-window-api";
import { loadState, saveState, type PersistedState } from "./state";

const POLL_INTERVAL_MS = 1500;
const AUTO_OPEN_KEY = "Office.AutoShowTaskpaneWithDocument";

const elements = {
  app: document.getElementById("app") as HTMLElement,
  unsupported: document.getElementById("unsupported") as HTMLElement,
  unsupportedText: document.getElementById("unsupported-text") as HTMLElement,
  list: document.getElementById("tab-list") as HTMLElement,
  status: document.getElementById("status") as HTMLElement,
  newButton: document.getElementById("btn-new") as HTMLButtonElement,
  alignButton: document.getElementById("btn-align") as HTMLButtonElement,
  refreshButton: document.getElementById("btn-refresh") as HTMLButtonElement,
  autoOpen: document.getElementById("chk-auto-open") as HTMLInputElement
};

let persisted: PersistedState = loadState();
let currentWindows: WindowInfo[] = [];
let lastWindows: WindowInfo[] = [];
let errorMessage = "";
let hasRendered = false;
let inFlight = false;
let timerId: number | undefined;

function errorMessageOf(error: unknown): string {
  if (error instanceof Error && error.message) {
    return error.message;
  }
  if (typeof error === "string" && error.length > 0) {
    return error;
  }
  if (error && typeof error === "object" && "message" in error) {
    return String((error as { message: unknown }).message);
  }
  return "发生未知错误。";
}

function showUnsupported(message: string): void {
  elements.unsupportedText.textContent = message;
  elements.unsupported.classList.remove("hidden");
  elements.app.classList.add("hidden");
}

function updateStatus(): void {
  elements.status.textContent = errorMessage || alignSummary(currentWindows.length, persisted.alignEnabled);
  elements.status.classList.toggle("error", errorMessage.length > 0);
  elements.alignButton.textContent = persisted.alignEnabled ? "对齐：开" : "对齐：关";
  elements.alignButton.classList.toggle("on", persisted.alignEnabled);
  elements.alignButton.setAttribute("aria-pressed", String(persisted.alignEnabled));
}

function render(): void {
  const windows = sortWindows(currentWindows);
  const labels = uniqueLabels(windows);

  elements.list.textContent = "";

  if (windows.length === 0) {
    const empty = document.createElement("p");
    empty.className = "empty";
    empty.textContent = "未检测到其他窗口";
    elements.list.appendChild(empty);
  }

  windows.forEach((window, position) => {
    const row = document.createElement("div");
    row.className = window.isActive ? "tab active" : "tab";
    row.tabIndex = 0;
    row.setAttribute("role", "tab");
    row.setAttribute("aria-selected", String(window.isActive));

    const badge = document.createElement("span");
    badge.className = "tab-index";
    badge.textContent = String(position + 1);

    const name = document.createElement("span");
    name.className = "tab-name";
    name.textContent = labels.get(window.index) ?? "";
    name.title = `${name.textContent}（窗口 ${window.index}）`;

    const state = document.createElement("span");
    state.className = "tab-state";
    state.textContent = window.windowState === "minimized" ? "已最小化" : "";

    const close = document.createElement("button");
    close.type = "button";
    close.className = "tab-close";
    close.textContent = "×";
    close.title = "关闭该工作簿窗口";
    close.addEventListener("click", (event) => {
      event.stopPropagation();
      void closeAt(window.index);
    });

    row.append(badge, name, state, close);
    row.addEventListener("click", () => {
      void switchTo(window.index);
    });
    row.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        void switchTo(window.index);
      }
    });

    elements.list.appendChild(row);
  });

  updateStatus();
}

async function tick(): Promise<void> {
  if (inFlight || document.hidden) {
    return;
  }

  inFlight = true;
  try {
    const result = await api.listWindows();
    const diff = diffWindows(lastWindows, result.windows);

    currentWindows = result.windows;
    lastWindows = result.windows;

    if (persisted.alignEnabled && persisted.frame && diff.added.length > 0) {
      for (const index of diff.added) {
        const added = result.windows.find((window) => window.index === index);
        if (added && !persisted.originalBounds[String(index)]) {
          persisted.originalBounds[String(index)] = boundsFromWindow(added);
        }
      }
      saveState(persisted);
      await api.applyFrameToAll(persisted.frame, diff.added);
    }

    errorMessage = "";
    if (diff.needsRender || !hasRendered) {
      hasRendered = true;
      render();
    } else {
      updateStatus();
    }
  } catch (error) {
    errorMessage = errorMessageOf(error);
    updateStatus();
  } finally {
    inFlight = false;
  }
}

async function switchTo(index: number): Promise<void> {
  try {
    const frame = persisted.alignEnabled ? persisted.frame : null;
    await api.activateWindow(index, frame);
    errorMessage = "";
    updateStatus();
    await tick();
  } catch (error) {
    errorMessage = errorMessageOf(error);
    updateStatus();
  }
}

async function closeAt(index: number): Promise<void> {
  try {
    await api.closeWindow(index);
    errorMessage = "";
    window.setTimeout(() => {
      void tick();
    }, 300);
  } catch (error) {
    errorMessage = errorMessageOf(error);
    updateStatus();
  }
}

async function createNewWorkbook(): Promise<void> {
  try {
    await api.createWorkbook();
    errorMessage = "";
    window.setTimeout(() => {
      void tick();
    }, 300);
  } catch (error) {
    errorMessage = errorMessageOf(error);
    updateStatus();
  }
}

async function toggleAlign(): Promise<void> {
  try {
    if (persisted.alignEnabled) {
      const savedBounds = persisted.originalBounds;
      await api.restoreBounds(savedBounds);
      persisted.alignEnabled = false;
      persisted.frame = null;
      persisted.originalBounds = {};
      saveState(persisted);
      updateStatus();
    } else {
      const result = await api.listWindows();
      currentWindows = result.windows;
      lastWindows = result.windows;

      const active = result.windows.find((window) => window.isActive) ?? result.windows[0];
      if (!active) {
        throw new Error("没有可对齐的窗口。");
      }

      const originals: Record<string, Bounds> = {};
      for (const window of result.windows) {
        originals[String(window.index)] = boundsFromWindow(window);
      }

      persisted.frame = boundsFromWindow(active);
      persisted.originalBounds = originals;
      persisted.alignEnabled = true;
      saveState(persisted);
      updateStatus();
      await api.applyFrameToAll(persisted.frame);
    }

    errorMessage = "";
    await tick();
  } catch (error) {
    errorMessage = errorMessageOf(error);
    updateStatus();
  }
}

function readAutoOpen(): boolean {
  try {
    const value = Office.context.document?.settings?.get(AUTO_OPEN_KEY);
    return value === true || value === "true";
  } catch {
    return false;
  }
}

function writeAutoOpen(enabled: boolean): void {
  try {
    Office.context.document.settings.set(AUTO_OPEN_KEY, enabled);
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        errorMessage = `自动打开设置保存失败：${result.error?.message ?? "未知原因"}`;
        updateStatus();
      }
    });
  } catch (error) {
    errorMessage = errorMessageOf(error);
    updateStatus();
  }
}

function bindUi(): void {
  elements.newButton.addEventListener("click", () => {
    void createNewWorkbook();
  });
  elements.alignButton.addEventListener("click", () => {
    void toggleAlign();
  });
  elements.refreshButton.addEventListener("click", () => {
    void tick();
  });
  elements.autoOpen.checked = readAutoOpen();
  elements.autoOpen.addEventListener("change", () => {
    writeAutoOpen(elements.autoOpen.checked);
  });
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) {
      void tick();
    }
  });
}

async function init(): Promise<void> {
  if (!api.isSupported()) {
    showUnsupported(
      "当前 Excel 不支持 ExcelApiDesktop 1.1（需要桌面版 Microsoft 365 for Mac / Windows，且为较新版本）。网页版和旧版 Excel 无法使用窗口标签功能。"
    );
    return;
  }

  bindUi();
  await tick();
  timerId = window.setInterval(() => {
    void tick();
  }, POLL_INTERVAL_MS);
}

Office.onReady(() => {
  void init();
});

window.addEventListener("unload", () => {
  if (timerId !== undefined) {
    window.clearInterval(timerId);
  }
});
