import {
  alignSummary,
  boundsFromWindow,
  diffWindows,
  findWindowByKey,
  needsRealign,
  nextPollDelay,
  sameOrder,
  sortWindows,
  uniqueLabels,
  windowKey,
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
const ALIGN_RETRY_DELAYS = [0, 800, 1800];
const HIDDEN_POLL_MS = 5000;

// 构建时由 HtmlWebpackPlugin 注入清单版本号：面板里显示它，便于确认 webview
// 加载的是哪一版资源（Office 内置浏览器缓存曾多次导致"改了却没生效"）。
const APP_VERSION =
  document.querySelector('meta[name="x-app-version"]')?.getAttribute("content")?.trim() ?? "";

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
let rowNodes = new Map<string, RowRefs>();
let emptyNode: HTMLElement | null = null;
let skeletonNodes: HTMLElement[] = [];
let inFlight = false;
let queuedTick = false;
let failStreak = 0;
let pendingKey: string | null = null;
let focusedKey: string | null = null;
let lastActiveKey: string | null = null;
let pollTimer: number | undefined;
let toastTimer: number | undefined;
let activeToast: { kind: "success" | "error"; message: string } | null = null;
let hasRendered = false;
let lastAlignAt = 0;
let alignTimer: number | undefined;

const busy = {
  creating: false,
  aligning: false,
  closing: new Set<string>()
};

function keyOf(window: WindowInfo, list: WindowInfo[] = windows): string {
  return windowKey(window, list);
}

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
  showToast("error", errorMessageOf(error), retry);
}

function showUnsupported(message: string): void {
  elements.unsupportedText.textContent = message;
  elements.unsupported.classList.remove("hidden");
  elements.app.classList.add("hidden");
}

function updateStatus(): void {
  const summary = alignSummary(windows.length, persisted.alignEnabled);
  const active = windows.find((window) => window.isActive) ?? null;
  const frameText = persisted.frame
    ? `基准 ${Math.round(persisted.frame.width)}×${Math.round(persisted.frame.height)} @(${Math.round(
        persisted.frame.left
      )},${Math.round(persisted.frame.top)})`
    : "未设置基准";
  const activeText = active?.geometry
    ? `活动窗口 ${Math.round(active.geometry.width)}×${Math.round(active.geometry.height)} @(${Math.round(
        active.geometry.left
      )},${Math.round(active.geometry.top)})`
    : "活动窗口几何未知";

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
  elements.alignButton.title = `对齐模式：把所有工作簿窗口叠放到同一位置和大小；${frameText}；${activeText}`;
  elements.alignButton.classList.toggle("busy", busy.aligning);
  elements.newButton.classList.toggle("busy", busy.creating);
  elements.realignButton.classList.toggle("hidden", !persisted.alignEnabled);
  elements.setBaseButton.classList.toggle("hidden", !persisted.alignEnabled);
  elements.displayModeButton.textContent = persisted.displayMode === "compact" ? "显示：紧凑" : "显示：完整";
  const hint = isMacPlatform()
    ? "⌘1–9 切换 · ↑↓ 选择 · Enter 激活"
    : "Ctrl+1–9 切换 · ↑↓ 选择 · Enter 激活";
  elements.shortcutHint.textContent = APP_VERSION ? `${hint} · v${APP_VERSION}` : hint;
}

function applyDisplayMode(): void {
  elements.list.classList.toggle("compact", persisted.displayMode === "compact");
  updateStatus();
}

function createRow(key: string): RowRefs {
  const row = document.createElement("div");
  row.className = "tab";
  row.dataset.key = key;
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

  // 行内元素必须显式挂载：之前漏了这三行 appendChild，标签就成了空白框。
  row.appendChild(badge);
  row.appendChild(name);
  row.appendChild(state);
  row.appendChild(close);

  const refs: RowRefs = { row, badge, name, state, close };

  close.addEventListener("click", (event) => {
    event.stopPropagation();
    const target = findWindowByKey(windows, row.dataset.key ?? "");
    if (target) {
      void closeAt(target);
    }
  });

  row.addEventListener("click", () => {
    const target = findWindowByKey(windows, row.dataset.key ?? "");
    if (target) {
      void switchTo(target);
    }
  });

  rowNodes.set(key, refs);
  elements.list.appendChild(row);
  return refs;
}

function updateRow(
  refs: RowRefs,
  info: WindowInfo,
  key: string,
  label: string,
  position: number,
  activeKey: string | null
): void {
  const rawName = (info.name ?? "").trim() || label;
  const tooltip = `${rawName}（窗口 ${info.index}）`;

  refs.row.dataset.key = key;
  refs.badge.textContent = String(position);
  refs.name.textContent = label;
  refs.name.title = tooltip;
  refs.row.title = tooltip;
  refs.row.setAttribute(
    "aria-label",
    `${label}，窗口 ${info.index}${info.isActive ? "，当前窗口" : ""}`
  );
  refs.state.textContent = info.windowState === "minimized" ? "已最小化" : "";
  refs.row.classList.toggle("active", key === activeKey);
  refs.row.classList.toggle("pending", key === pendingKey);
  refs.row.classList.toggle("busy", busy.closing.has(key));
  refs.row.tabIndex = focusedKey === key || (focusedKey === null && key === activeKey) ? 0 : -1;
}

function removeRow(key: string, refs: RowRefs): void {
  rowNodes.delete(key);
  if (focusedKey === key) {
    focusedKey = null;
  }
  if (pendingKey === key) {
    pendingKey = null;
  }
  refs.row.classList.add("leaving");
  window.setTimeout(() => {
    refs.row.remove();
  }, LEAVING_MS);
}

function ensureOrder(desired: string[]): void {
  const current = Array.from(elements.list.children)
    .filter((node) => !node.classList.contains("leaving") && !node.classList.contains("skeleton"))
    .map((node) => (node as HTMLElement).dataset.key ?? "");

  if (sameOrder(current, desired)) {
    return;
  }

  for (const key of desired) {
    const refs = rowNodes.get(key);
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
  const keys = sorted.map((window) => windowKey(window, sorted));
  const activeWindow = sorted.find((window) => window.isActive) ?? null;
  const activeKey = pendingKey ?? (activeWindow ? windowKey(activeWindow, sorted) : null);
  const nextKeys = new Set(keys);

  clearSkeleton();

  for (const [key, refs] of Array.from(rowNodes.entries())) {
    if (!nextKeys.has(key)) {
      removeRow(key, refs);
    }
  }

  if (sorted.length === 0) {
    showEmpty();
    updateStatus();
    return;
  }

  hideEmpty();

  sorted.forEach((info, position) => {
    const key = keys[position];
    const refs = rowNodes.get(key) ?? createRow(key);
    updateRow(refs, info, key, labels.get(info.index) ?? "", position + 1, activeKey);
  });

  ensureOrder(keys);
  updateStatus();
}

function scheduleTick(delay?: number): void {
  if (pollTimer !== undefined) {
    window.clearTimeout(pollTimer);
    pollTimer = undefined;
  }

  // 窗口完全被遮挡（document.hidden）时放慢到 5s，仍保持同步与自动对齐，
  // 避免"窗格在后台就再也不纠正新窗口"。
  const wait = document.hidden
    ? HIDDEN_POLL_MS
    : delay ??
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

/** 把基准套用到当前活动窗口（macOS 只对前台窗口生效，且窗口索引会随激活变化）。 */
function scheduleAlign(): void {
  if (!persisted.frame || !persisted.alignEnabled) {
    return;
  }

  if (alignTimer !== undefined) {
    return;
  }

  const frame = persisted.frame;
  const elapsed = Date.now() - lastAlignAt;
  const wait = Math.max(ALIGN_DEBOUNCE_MS, ALIGN_MIN_INTERVAL_MS - elapsed);

  alignTimer = window.setTimeout(() => {
    alignTimer = undefined;
    void runAlignAttempts(frame, 0);
  }, wait);
}

/**
 * Excel 打开/激活窗口后可能延迟恢复窗口位置，单次写入会被覆盖，
 * 因此连续重试几次；失败只提示一次，避免刷屏。
 */
async function runAlignAttempts(frame: Bounds, attempt: number): Promise<void> {
  if (!persisted.alignEnabled || !persisted.frame) {
    return;
  }

  lastAlignAt = Date.now();

  try {
    await api.alignActiveWindow(frame);
  } catch (error) {
    if (attempt === 0) {
      showToast("error", `自动对齐失败：${errorMessageOf(error)}`);
    }
  }

  const next = attempt + 1;
  if (next < ALIGN_RETRY_DELAYS.length) {
    window.setTimeout(() => {
      void runAlignAttempts(frame, next);
    }, ALIGN_RETRY_DELAYS[next]);
  }
}

async function applySnapshot(next: WindowInfo[]): Promise<void> {
  const diff = diffWindows(lastSnapshot, next);
  const sorted = sortWindows(next);
  const activeWindow = sorted.find((window) => window.isActive) ?? null;
  const activeKey = activeWindow ? windowKey(activeWindow, sorted) : null;
  const activeChanged = activeKey !== lastActiveKey;

  windows = next;
  lastSnapshot = next;

  if (pendingKey !== null) {
    const pendingWindow = findWindowByKey(sorted, pendingKey);
    if (pendingWindow && pendingWindow.isActive) {
      pendingKey = null;
    }
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

  // macOS 上几何写入只对前台窗口生效：谁被激活，就把基准套用到谁身上。
  if (persisted.alignEnabled && persisted.frame && activeWindow) {
    if (activeChanged || needsRealign(activeWindow.geometry, persisted.frame)) {
      scheduleAlign();
    }
  }

  lastActiveKey = activeKey;

  if (diff.needsRender || !hasRendered) {
    hasRendered = true;
    render();
  } else {
    updateStatus();
  }
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

async function switchTo(target: WindowInfo): Promise<void> {
  if (busy.closing.has(keyOf(target))) {
    return;
  }

  const targetKey = keyOf(target);
  pendingKey = targetKey;
  render();

  try {
    const frame = persisted.alignEnabled ? persisted.frame : null;
    await api.activateWindow({ index: target.index, name: target.name }, frame);
    pendingKey = null;
    render();
    scheduleTick(ACTION_REFRESH_MS);
  } catch (error) {
    pendingKey = null;
    reportError(error, () => {
      void switchTo(target);
    });
    scheduleTick(ACTION_REFRESH_MS);
  }
}

async function closeAt(target: WindowInfo): Promise<void> {
  const targetKey = keyOf(target);
  if (busy.closing.has(targetKey)) {
    return;
  }

  busy.closing.add(targetKey);
  const refs = rowNodes.get(targetKey);
  refs?.row.classList.add("leaving");

  try {
    await api.closeWindow({ index: target.index, name: target.name });
    windows = windows.filter((window) => windowKey(window, windows) !== targetKey);
    lastSnapshot = lastSnapshot.filter((window) => windowKey(window, lastSnapshot) !== targetKey);
    if (refs) {
      removeRow(targetKey, refs);
    }
    updateStatus();
    showToast("success", "已关闭该工作簿窗口");
    scheduleTick(ACTION_REFRESH_MS);
  } catch (error) {
    refs?.row.classList.remove("leaving");
    reportError(error, () => {
      void closeAt(target);
    });
  } finally {
    busy.closing.delete(targetKey);
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

function orderedKeys(): string[] {
  const sorted = sortWindows(windows);
  return sorted.map((window) => windowKey(window, sorted));
}

function focusRow(key: string): void {
  const refs = rowNodes.get(key);
  if (!refs) {
    return;
  }
  focusedKey = key;
  for (const [rowKey, rowRefs] of rowNodes) {
    rowRefs.row.tabIndex = rowKey === key ? 0 : -1;
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

  const key = row.dataset.key ?? "";
  const targetWindow = findWindowByKey(windows, key);
  if (!targetWindow) {
    return;
  }

  const keys = orderedKeys();
  const position = keys.indexOf(key);

  switch (event.key) {
    case "ArrowDown":
    case "ArrowUp": {
      event.preventDefault();
      if (keys.length === 0) {
        return;
      }
      const step = event.key === "ArrowDown" ? 1 : -1;
      const nextPosition = (position + step + keys.length) % keys.length;
      focusRow(keys[nextPosition]);
      break;
    }
    case "Home":
      event.preventDefault();
      if (keys.length > 0) {
        focusRow(keys[0]);
      }
      break;
    case "End":
      event.preventDefault();
      if (keys.length > 0) {
        focusRow(keys[keys.length - 1]);
      }
      break;
    case "Enter":
    case " ":
      event.preventDefault();
      void switchTo(targetWindow);
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

  const sorted = sortWindows(windows);
  const target = sorted[position];
  if (!target) {
    return;
  }

  event.preventDefault();
  void switchTo(target);
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
    if (row?.dataset.key) {
      focusedKey = row.dataset.key;
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
  if (alignTimer !== undefined) {
    window.clearTimeout(alignTimer);
  }
});
