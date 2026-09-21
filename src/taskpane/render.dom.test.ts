// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

/**
 * 任务窗格渲染冒烟测试。
 *
 * 背景：曾经出现过"标签行只有空白框"的回归——createRow() 创建了序号/名称/关闭
 * 按钮却没有 append 到行里，DOM 里只剩下空 div。纯函数测试覆盖不到这种问题，
 * 所以这里用 jsdom 把真实的 taskpane 模块跑起来，直接检查渲染结果。
 */

interface FakeWindow {
  index: number;
  name: string;
  windowState: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

const PANE_HTML = `
  <main id="app">
    <header class="toolbar">
      <button id="btn-new" type="button"></button>
      <button id="btn-align" type="button"></button>
      <button id="btn-realign" type="button" class="icon-btn hidden"></button>
      <button id="btn-setbase" type="button" class="icon-btn hidden"></button>
      <button id="btn-refresh" type="button" class="icon-btn"></button>
    </header>
    <div id="toast-area" class="toast-area"></div>
    <section id="tab-list" class="tab-list"></section>
    <footer class="footer">
      <label class="auto-open"><input id="chk-auto-open" type="checkbox" /></label>
      <p id="status" class="status"></p>
      <div class="footer-row">
        <button id="btn-display-mode" type="button" class="link-btn"></button>
        <span id="shortcut-hint" class="hint"></span>
      </div>
    </footer>
  </main>
  <section id="unsupported" class="unsupported hidden">
    <p id="unsupported-text"></p>
  </section>
`;

let sheetWindows: FakeWindow[] = [];
let activeIndex = 1;
const activateCalls: string[] = [];

function makeContext() {
  const items = sheetWindows.map((definition) => ({
    ...definition,
    load() {
      /* 桩：真实 API 里由 Office 填充属性 */
    },
    activate() {
      activateCalls.push(definition.name);
      activeIndex = definition.index;
    }
  }));

  return {
    workbook: {
      application: {
        windows: {
          items,
          load() {
            /* 桩 */
          }
        },
        activeWindow: items.find((item) => item.index === activeIndex) ?? items[0]
      }
    },
    sync: async () => {
      /* 桩：立即完成 */
    }
  };
}

async function flush(times = 4): Promise<void> {
  for (let index = 0; index < times; index += 1) {
    await Promise.resolve();
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

async function startPane(): Promise<void> {
  document.body.innerHTML = PANE_HTML;

  (globalThis as unknown as { Excel: unknown }).Excel = {
    run: async (callback: (context: unknown) => Promise<unknown>) => callback(makeContext())
  };

  (globalThis as unknown as { Office: unknown }).Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    PlatformType: { Mac: "Mac", PC: "PC" },
    context: {
      platform: "Mac",
      requirements: { isSetSupported: () => true },
      document: {
        settings: {
          get: () => undefined,
          set: () => undefined,
          saveAsync: (callback: (result: { status: string }) => void) => callback({ status: "succeeded" })
        }
      }
    },
    onReady: (callback: () => void) => {
      void Promise.resolve().then(callback);
    }
  };

  vi.resetModules();
  await import("./taskpane");
  await flush();
}

function rows(): HTMLElement[] {
  return Array.from(document.querySelectorAll<HTMLElement>("#tab-list .tab"));
}

describe("任务窗格渲染", () => {
  beforeEach(() => {
    window.localStorage.clear();
    activateCalls.length = 0;
    activeIndex = 1;
    sheetWindows = [
      {
        index: 1,
        name: "2026年第三季度销售预测与渠道返点结算表(最终修订版)",
        windowState: "normal",
        left: 100,
        top: 80,
        width: 1200,
        height: 800
      },
      { index: 2, name: "库存台账", windowState: "normal", left: 140, top: 120, width: 1200, height: 800 },
      { index: 3, name: "Book3", windowState: "minimized", left: 180, top: 160, width: 900, height: 700 }
    ];
  });

  afterEach(() => {
    window.dispatchEvent(new Event("unload"));
    document.body.innerHTML = "";
  });

  it("每个标签行都渲染出序号、完整文档名和关闭按钮", async () => {
    await startPane();

    const list = rows();
    expect(list).toHaveLength(3);

    for (const row of list) {
      expect(row.querySelector(".tab-index")?.textContent).toBeTruthy();
      expect(row.querySelector(".tab-name")?.textContent?.trim()).toBeTruthy();
      expect(row.querySelector(".tab-close")?.textContent).toBe("×");
    }

    const names = list.map((row) => row.querySelector(".tab-name")?.textContent ?? "");
    expect(names[0]).toBe("2026年第三季度销售预测与渠道返点结算表(最终修订版)");
    expect(names[1]).toBe("库存台账");
    expect(names[2]).toBe("Book3");

    expect(list.map((row) => row.querySelector(".tab-index")?.textContent)).toEqual(["1", "2", "3"]);
    expect(list[0].classList.contains("active")).toBe(true);
    expect(list[2].querySelector(".tab-state")?.textContent).toBe("已最小化");
  });

  it("重名窗口追加序号后缀", async () => {
    sheetWindows = [
      { index: 1, name: "报表", windowState: "normal", left: 0, top: 0, width: 800, height: 600 },
      { index: 2, name: "报表", windowState: "normal", left: 0, top: 0, width: 800, height: 600 }
    ];

    await startPane();

    const names = rows().map((row) => row.querySelector(".tab-name")?.textContent ?? "");
    expect(names).toEqual(["报表", "报表 (2)"]);
  });

  it("点击标签行会激活对应窗口", async () => {
    await startPane();

    const second = rows()[1];
    second.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    // 点击后立刻乐观高亮，不必等 Office 往返。
    expect(rows()[1].classList.contains("active")).toBe(true);

    await flush();

    expect(activateCalls).toEqual(["库存台账"]);

    // 轮询刷新拿回真实活动窗口后，高亮仍然停在这一行。
    await new Promise((resolve) => setTimeout(resolve, 300));
    expect(rows()[1].classList.contains("active")).toBe(true);
  });
});
