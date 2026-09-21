import {
  alignSummary,
  boundsFromWindow,
  diffWindows,
  needsRealign,
  nextPollDelay,
  sameOrder,
  sortWindows,
  uniqueLabels,
  type Bounds,
  type DisplayMode,
  type WindowInfo
} from "./logic";
import * as api from "./excel-window-api";
import { loadState, saveState, type PersistedState } from "./state";

const AUTO_OPEN_KEY = "Office.AutoShowTaskpaneWithDocument";
const ACTION_REFRESH_MS = 120;
const LEAVING_MS = 150;
const TOAST_OK_MS = 2500;
const TOAST_ERROR_MS = 8000;
const QUICK_SWITCH_MAX = 9;
const SKELETON_ROWS = 3;
const ALIGN_DEBOUNCE_MS = 150;
const ALIGN_MIN_INTERVAL_MS = 900;

const elements = {
  app: document.getElementById("app") as HTMLElement,
  unsupported: document.getElementById("unsupported") as HTMLElement,
  unsupportedText: document.getElementById("unsupported-text") as HTMLElement,
  list: document.getElementById("tab-list") as HTMLElement,
  toastArea: document.getElementById("toast-area") as HTMLElement,
  status: document.getElementById("status") as HTMLElement,
  newButton: document.getElementById("btn-new") as HTMLButtonElement,
  alignButton: document.getElementById("btn-align") as HTMLButtonElement,
  realignButton: document.getElementById("btn-realign") as HTMLButtonElement,
  setBaseButton: document.getElementById("btn-setbase") as HTMLButtonElement,
  refreshButton: document.getElementById("btn-refresh") as HTMLButtonElement,
  autoOpen: document.getElementById("chk-auto-open") as HTMLInputElement,
  displayModeButton: document.getElementById("btn-display-mode") as HTMLButtonElement,
  shortcutHint: document.getElementById("shortcut-hint") as HTMLElement
};

interface RowRefs {
  row: HTMLElement;
  badge: HTMLElement;
  name: HTMLElement;
  state: HTMLElement;
  close: HTMLButtonElement;
}

let persisted: PersistedState = loadState();
let windows: WindowInfo[] = [];
let lastSnapshot: WindowInfo[] = [];
let rowNodes = new Map<number, RowRefs>();
let emptyNode: HTMLElement | null = null;
let skeletonNodes: HTMLElement[] = [];
let inFlight = false;
let queuedTick = false;
let failStreak = 0;
let pendingIndex: number | null = null;
let focusedIndex: number | null = null;
let pollTimer: number | undefined;
let toastTimer: number | undefined;
let activeToast: { kind: "success" | "error"; message: string } | null = null;
let hasRendered = false;
let lastRetry: (() => void) | null = null;
let lastActiveIndex: number | null = null;
let lastAlignAt = 0;
let alignTimer: number | undefined;
let alignTargetIndex: number | null = null;

const busy = {
  creating: false,
  aligning: false,
  closing: new Set<number>()
};

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

function isMacPlatform(): boolean {
  try {
    return Office.context.platform === Office.PlatformType.Mac;
  } catch {
    return true;
  }
}

async function copyToClipboard(text: string): Promise<void> {
  try {
    await navigator.clipboard.writeText(text);
    showToast("success", "已复制错误信息");
  } catch {
    showToast("error", "复制失败，请手动选择文本");
  }
}

function clearToasts(): void {
  if (toastTimer !== undefined) {
    window.clearTimeout(toastTimer);
    toastTimer = undefined;
  }
  elements.toastArea.textContent = "";
  activeToast = null;
}

function resetToastTimer(kind: "success" | "error"): void {
  if (toastTimer !== undefined) {
    window.clearTimeout(toastTimer);
  }
  toastTimer = window.setTimeout(() => {
    clearToasts();
  }, kind === "error" ? TOAST_ERROR_MS : TOAST_OK_MS);
}

function showToast(kind: "success" | "error", message: string, retry?: () => void): void {
  if (activeToast && activeToast.kind === kind && activeToast.message === message) {
    resetToastTimer(kind);
    return;
  }

  clearToasts();

  const toast = document.createElement("div");
  toast.className = `toast ${kind}`;

  const text = document.createElement("span");
  text.className = "toast-text";
  text.textContent = message;
  toast.appendChild(text);

  if (kind === "error") {
    if (retry) {
      const retryButton = document.createElement("button");
      retryButton.type = "button";
      retryButton.textContent = "重试";
      retryButton.addEventListener("click", () => {
        clearToasts();
        retry();
      });
      toast.appendChild(retryButton);
    }

    const copyButton = document.createElement("button");
    copyButton.type = "button";
    copyButton.textContent = "复制";
    copyButton.addEventListener("click", () => {
      void copyToClipboard(`[${new Date().toLocaleTimeString()}] ${message}`);
    });
    toast.appendChild(copyButton);
  }

  elements.toastArea.appendChild(toast);
  activeToast = { kind, message };
  resetToastTimer(kind);
}

function reportError(error: unknown, retry?: () => void): void {
  const message = errorMessageOf(error);
  if (retry) {
    lastRetry = retry;
  }
  showToast("error", message, retry);
}

function showUnsupported(message: string): void {
  elements.unsupportedText.textContent = message;
  elements.unsupported.classList.remove("hidden");
  elements.app.classList.add("hidden");
}

function updateStatus(): void {
  const summary = alignSummary(windows.length, persisted.alignEnabled);

  if (busy.creating) {
    elements.status.textContent = "正在新建工作簿…";
  } else if (busy.aligning) {
    elements.status.textContent = "正在处理对齐…";
  } else {
    elements.status.textContent = summary;
  }

  elements.status.classList.toggle("busy", busy.creating || busy.aligning);
  elements.alignButton.textContent = persisted.alignEnabled ? "对齐：开" : "对齐：关";
  elements.alignButton.classList.toggle("on", persisted.alignEnabled);
  elements.alignButton.setAttribute("aria-pressed", String(persisted.alignEnabled));
  elements.alignButton.classList.toggle("busy", busy.aligning);
  elements.newButton.classList.toggle("busy", busy.creating);
  elements.realignButton.classList.toggle("hidden", !persisted.alignEnabled);
  elements.setBaseButton.classList.toggle("hidden", !persisted.alignEnabled);
  elements.displayModeButton.textContent = persisted.displayMode === "compact" ? "显示：紧凑" : "显示：完整";
  elements.shortcutHint.textContent = isMacPlatform()
    ? "⌘1–9 切换 · ↑↓ 选择 · Enter 激活"
    : "Ctrl+1–9 切换 · ↑↓ 选择 · Enter 激活";
}

function applyDisplayMode(): void {
  elements.list.classList.toggle("compact", persisted.displayMode === "compact");
  updateStatus();
}

function createRow(info: WindowInfo): RowRefs {
  const row = document.createElement("div");
  row.className = "tab";
  row.dataset.index = String(info.index);
  row.setAttribute("role", "tab");

  const badge = document.createElement("span");
  badge.className = "tab-index";

  const name = document.createElement("span");
  name.className = "tab-name";

  const state = document.createElement("span");
  state.className = "tab-state";

  const close = document.createElement("button");
  close.type = "button";
  close.className = "tab-close";
  close.textContent = "×";
  close.title = "关闭该工作簿窗口";
  close.addEventListener("click", (event) => {
    event.stopPropagation();
    void closeAt(info.index);
  });

  row.append(badge, name, state, close);
  row.addEventListener("click", () => {
    void switchTo(info.index);
  });

  const refs: RowRefs = { row, badge, name, state, close };
  rowNodes.set(info.index, refs);
  elements.list.appendChild(row);
  return refs;
}

function updateRow(
  refs: RowRefs,
  info: WindowInfo,
  label: string,
  position: number,
  activeIndex: number | null
): void {
  const rawName = (info.name ?? "").trim() || label;
  const tooltip = `${rawName}（窗口 ${info.index}）`;

  refs.row.dataset.index = String(info.index);
  refs.badge.textContent = String(position);
  refs.name.textContent = label;
  refs.name.title = tooltip;
  refs.row.title = tooltip;
  refs.row.setAttribute(
    "aria-label",
    `${label}，窗口 ${info.index}${info.isActive ? "，当前窗口" : ""}`
  );
  refs.state.textContent = info.windowState === "minimized" ? "已最小化" : "";
  refs.row.classList.toggle("active", info.index === activeIndex);
  refs.row.classList.toggle("pending", info.index === pendingIndex);
  refs.row.classList.toggle("busy", busy.closing.has(info.index));
  refs.row.tabIndex =
    focusedIndex === info.index || (focusedIndex === null && info.index === activeIndex) ? 0 : -1;
}

function removeRow(index: number, refs: RowRefs): void {
  rowNodes.delete(index);
  if (focusedIndex === index) {
    focusedIndex = null;
  }
  refs.row.classList.add("leaving");
  window.setTimeout(() => {
    refs.row.remove();
  }, LEAVING_MS);
}

function ensureOrder(desired: number[]): void {
  const current = Array.from(elements.list.children)
    .filter((node) => !node.classList.contains("leaving") && !node.classList.contains("skeleton"))
    .map((node) => Number((node as HTMLElement).dataset.index))
    .filter((value) => !Number.isNaN(value));

  if (sameOrder(current, desired)) {
    return;
  }

  for (const index of desired) {
    const refs = rowNodes.get(index);
    if (refs) {
      elements.list.appendChild(refs.row);
    }
  }
}

function showEmpty(): void {
  if (emptyNode) {
    return;
  }
  emptyNode = document.createElement("p");
  emptyNode.className = "empty";
  emptyNode.textContent = "未检测到其他窗口";
  elements.list.appendChild(emptyNode);
}

function hideEmpty(): void {
  if (emptyNode) {
    emptyNode.remove();
    emptyNode = null;
  }
}

function renderSkeleton(): void {
  clearSkeleton();
  for (let index = 0; index < SKELETON_ROWS; index += 1) {
    const node = document.createElement("div");
    node.className = "tab skeleton";
    elements.list.appendChild(node);
    skeletonNodes.push(node);
  }
}

function clearSkeleton(): void {
  for (const node of skeletonNodes) {
    node.remove();
  }
  skeletonNodes = [];
}

function render(): void {
  const sorted = sortWindows(windows);
  const labels = uniqueLabels(sorted);
  const activeIndex = pendingIndex ?? sorted.find((window) => window.isActive)?.index ?? null;
  const nextIndexes = new Set(sorted.map((window) => window.index));

  clearSkeleton();

  for (const [index, refs] of Array.from(rowNodes.entries())) {
    if (!nextIndexes.has(index)) {
      removeRow(index, refs);
    }
  }

  if (sorted.length === 0) {
    showEmpty();
    updateStatus();
    return;
  }

  hideEmpty();

  sorted.forEach((info, position) => {
    const refs = rowNodes.get(info.index) ?? createRow(info);
    updateRow(refs, info, labels.get(info.index) ?? "", position + 1, activeIndex);
  });

  ensureOrder(sorted.map((window) => window.index));
  updateStatus();
}

function scheduleTick(delay?: number): void {
  if (pollTimer !== undefined) {
    window.clearTimeout(pollTimer);
    pollTimer = undefined;
  }

  if (document.hidden) {
    return;
  }

  const wait =
    delay ??
    nextPollDelay({
      visible: true,
      focused: document.hasFocus(),
      failStreak,
      alignEnabled: persisted.alignEnabled
    }) ??
    0;

  pollTimer = window.setTimeout(() => {
    void tick();
  }, wait);
}

async function applySnapshot(next: WindowInfo[]): Promise<void> {
  const diff = diffWindows(lastSnapshot, next);
  const activeWindow = next.find((item) => item.isActive) ?? null;
  const activeChanged =
    activeWindow !== null && lastActiveIndex !== null && activeWindow.index !== lastActiveIndex;
  const firstSnapshot = lastActiveIndex === null;

  windows = next;
  lastSnapshot = next;

  if (pendingIndex !== null && next.some((item) => item.index === pendingIndex && item.isActive)) {
    pendingIndex = null;
  }

  if (persisted.alignEnabled && persisted.frame && diff.added.length > 0) {
    for (const index of diff.added) {
      const added = next.find((item) => item.index === index);
      if (!added) {
        continue;
      }
      const geometry = boundsFromWindow(added);
      if (geometry && !persisted.originalBounds[String(index)]) {
        persisted.originalBounds[String(index)] = geometry;
      }
    }
    saveState(persisted);
  }

  // macOS 版 Excel 的窗口几何写入只对前台窗口生效，
  // 因此改为“谁被激活，就把基准套用到谁”，避免给后台窗口写无效几何。
  if (persisted.alignEnabled && persisted.frame && activeWindow) {
    if (activeChanged || firstSnapshot || needsRealign(activeWindow.geometry, persisted.frame)) {
      scheduleAlign(activeWindow.index);
    }
  }

  lastActiveIndex = activeWindow?.index ?? null;

  if (diff.needsRender || !hasRendered) {
    hasRendered = true;
    render();
  } else {
    updateStatus();
  }
}

function scheduleAlign(index: number): void {
  if (!persisted.frame || !persisted.alignEnabled) {
    return;
  }

  alignTargetIndex = index;
  if (alignTimer !== undefined) {
    return;
  }

  const elapsed = Date.now() - lastAlignAt;
  const wait = Math.max(ALIGN_DEBOUNCE_MS, ALIGN_MIN_INTERVAL_MS - elapsed);

  alignTimer = window.setTimeout(() => {
    alignTimer = undefined;
    const target = alignTargetIndex;
    alignTargetIndex = null;

    if (!persisted.frame || !persisted.alignEnabled || target === null) {
      return;
    }

    lastAlignAt = Date.now();
    // 自动纠正失败不打扰用户：下一次轮询会重试。
    void api.alignWindow(target, persisted.frame).catch(() => undefined);
  }, wait);
}

async function tick(force = false): Promise<void> {
  if (inFlight) {
    if (force) {
      queuedTick = true;
    }
    return;
  }

  inFlight = true;

  try {
    const result = await api.listWindows({ withGeometry: persisted.alignEnabled });
    failStreak = 0;
    await applySnapshot(result.windows);
  } catch (error) {
    failStreak += 1;
    reportError(error, () => {
      void tick(true);
    });
  } finally {
    inFlight = false;
    scheduleTick();
    if (queuedTick) {
      queuedTick = false;
      scheduleTick(ACTION_REFRESH_MS);
    }
  }
}

async function switchTo(index: number): Promise<void> {
  const previousActive = lastSnapshot.find((window) => window.isActive)?.index ?? null;

  pendingIndex = index;
  render();

  try {
    const frame = persisted.alignEnabled ? persisted.frame : null;
    await api.activateWindow(index, frame);
    pendingIndex = null;
    render();
    scheduleTick(ACTION_REFRESH_MS);
  } catch (error) {
    pendingIndex = previousActive;
    render();
    reportError(error, () => {
      void switchTo(index);
    });
    scheduleTick(ACTION_REFRESH_MS);
  }
}

async function closeAt(index: number): Promise<void> {
  if (busy.closing.has(index)) {
    return;
  }

  busy.closing.add(index);
  const refs = rowNodes.get(index);
  refs?.row.classList.add("leaving");

  try {
    await api.closeWindow(index);
    windows = windows.filter((window) => window.index !== index);
    lastSnapshot = lastSnapshot.filter((window) => window.index !== index);
    if (refs) {
      removeRow(index, refs);
    }
    updateStatus();
    showToast("success", "已关闭该工作簿窗口");
    scheduleTick(ACTION_REFRESH_MS);
  } catch (error) {
    refs?.row.classList.remove("leaving");
    reportError(error, () => {
      void closeAt(index);
    });
  } finally {
    busy.closing.delete(index);
  }
}

async function createNewWorkbook(): Promise<void> {
  busy.creating = true;
  updateStatus();

  try {
    await api.createWorkbook();
    showToast("success", "已新建工作簿");
    scheduleTick(ACTION_REFRESH_MS);
  } catch (error) {
    reportError(error, () => {
      void createNewWorkbook();
    });
  } finally {
    busy.creating = false;
    updateStatus();
  }
}

async function toggleAlign(): Promise<void> {
  busy.aligning = true;
  updateStatus();

  try {
    if (persisted.alignEnabled) {
      const savedBounds = persisted.originalBounds;
      await api.restoreBounds(savedBounds);
      persisted.alignEnabled = false;
      persisted.frame = null;
      persisted.originalBounds = {};
      saveState(persisted);
      showToast("success", "已关闭对齐模式，恢复各窗口原位置");
    } else {
      const result = await api.listWindows({ withGeometry: true });
      windows = result.windows;
      lastSnapshot = result.windows;

      const active = result.windows.find((window) => window.isActive) ?? result.windows[0];
      const frame = active ? boundsFromWindow(active) : null;
      if (!frame) {
        throw new Error("无法读取当前窗口的位置和大小。");
      }

      const originals: Record<string, Bounds> = {};
      for (const window of result.windows) {
        const geometry = boundsFromWindow(window);
        if (geometry) {
          originals[String(window.index)] = geometry;
        }
      }

      persisted.frame = frame;
      persisted.originalBounds = originals;
      persisted.alignEnabled = true;
      saveState(persisted);

      await api.applyFrameToAll(frame);
      showToast("success", "已开启对齐模式");
    }

    render();
    scheduleTick(ACTION_REFRESH_MS);
  } catch (error) {
    reportError(error, () => {
      void toggleAlign();
    });
  } finally {
    busy.aligning = false;
    updateStatus();
  }
}

async function realignNow(): Promise<void> {
  if (!persisted.frame) {
    return;
  }

  try {
    await api.applyFrameToAll(persisted.frame);
    showToast("success", "已重新对齐全部窗口");
    scheduleTick(ACTION_REFRESH_MS);
  } catch (error) {
    reportError(error, () => {
      void realignNow();
    });
  }
}

async function setBaseFromActive(): Promise<void> {
  try {
    const bounds = await api.getActiveWindowBounds();
    if (!bounds) {
      throw new Error("无法读取当前窗口的位置和大小。");
    }

    persisted.frame = bounds;
    saveState(persisted);

    await api.applyFrameToAll(bounds);
    showToast("success", "已把当前窗口设为对齐基准");
    scheduleTick(ACTION_REFRESH_MS);
  } catch (error) {
    reportError(error, () => {
      void setBaseFromActive();
    });
  }
}

function setDisplayMode(mode: DisplayMode): void {
  persisted.displayMode = mode;
  saveState(persisted);
  applyDisplayMode();
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
        reportError(result.error?.message ?? "自动打开设置保存失败。");
        return;
      }

      if (enabled) {
        showToast("success", "已设置：打开该工作簿时自动显示标签栏（记得保存工作簿）");
      } else {
        showToast("success", "已取消自动打开");
      }
    });
  } catch (error) {
    reportError(error);
  }
}

function orderedIndexes(): number[] {
  return sortWindows(windows).map((window) => window.index);
}

function focusRow(index: number): void {
  const refs = rowNodes.get(index);
  if (!refs) {
    return;
  }
  focusedIndex = index;
  for (const [rowIndex, rowRefs] of rowNodes) {
    rowRefs.row.tabIndex = rowIndex === index ? 0 : -1;
  }
  refs.row.focus();
}

function handleListKeydown(event: KeyboardEvent): void {
  const target = event.target as HTMLElement | null;
  if (target?.classList.contains("tab-close")) {
    return;
  }
  const row = target?.closest(".tab") as HTMLElement | null;
  if (!row) {
    return;
  }

  const index = Number(row.dataset.index);
  if (Number.isNaN(index)) {
    return;
  }

  const indexes = orderedIndexes();
  const position = indexes.indexOf(index);

  switch (event.key) {
    case "ArrowDown":
    case "ArrowUp": {
      event.preventDefault();
      if (indexes.length === 0) {
        return;
      }
      const step = event.key === "ArrowDown" ? 1 : -1;
      const nextPosition = (position + step + indexes.length) % indexes.length;
      focusRow(indexes[nextPosition]);
      break;
    }
    case "Home":
      event.preventDefault();
      if (indexes.length > 0) {
        focusRow(indexes[0]);
      }
      break;
    case "End":
      event.preventDefault();
      if (indexes.length > 0) {
        focusRow(indexes[indexes.length - 1]);
      }
      break;
    case "Enter":
    case " ":
      event.preventDefault();
      void switchTo(index);
      break;
    default:
      break;
  }
}

function handleDocumentKeydown(event: KeyboardEvent): void {
  if (event.key === "Escape") {
    clearToasts();
    return;
  }

  const modifier = isMacPlatform() ? event.metaKey : event.ctrlKey;
  if (!modifier || event.altKey || event.shiftKey) {
    return;
  }

  if (!/^[1-9]$/.test(event.key)) {
    return;
  }

  const position = Number(event.key) - 1;
  if (position >= QUICK_SWITCH_MAX) {
    return;
  }

  const indexes = orderedIndexes();
  const targetIndex = indexes[position];
  if (targetIndex === undefined) {
    return;
  }

  event.preventDefault();
  void switchTo(targetIndex);
}

function bindUi(): void {
  elements.newButton.addEventListener("click", () => {
    void createNewWorkbook();
  });
  elements.alignButton.addEventListener("click", () => {
    void toggleAlign();
  });
  elements.realignButton.addEventListener("click", () => {
    void realignNow();
  });
  elements.setBaseButton.addEventListener("click", () => {
    void setBaseFromActive();
  });
  elements.refreshButton.addEventListener("click", () => {
    scheduleTick(0);
  });
  elements.displayModeButton.addEventListener("click", () => {
    setDisplayMode(persisted.displayMode === "compact" ? "full" : "compact");
  });

  elements.autoOpen.checked = readAutoOpen();
  elements.autoOpen.addEventListener("change", () => {
    writeAutoOpen(elements.autoOpen.checked);
  });

  elements.list.addEventListener("keydown", handleListKeydown);
  elements.list.addEventListener("focusin", (event) => {
    const row = (event.target as HTMLElement | null)?.closest(".tab") as HTMLElement | null;
    if (row) {
      const index = Number(row.dataset.index);
      if (!Number.isNaN(index)) {
        focusedIndex = index;
      }
    }
  });

  document.addEventListener("keydown", handleDocumentKeydown);
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) {
      scheduleTick();
    } else {
      scheduleTick(ACTION_REFRESH_MS);
    }
  });

  window.addEventListener("focus", () => {
    scheduleTick(ACTION_REFRESH_MS);
  });
  window.addEventListener("blur", () => {
    scheduleTick();
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
  applyDisplayMode();
  renderSkeleton();
  await tick(true);
  scheduleTick();
}

Office.onReady(() => {
  void init();
});

window.addEventListener("unload", () => {
  if (pollTimer !== undefined) {
    window.clearTimeout(pollTimer);
  }
  if (toastTimer !== undefined) {
    window.clearTimeout(toastTimer);
  }
});
