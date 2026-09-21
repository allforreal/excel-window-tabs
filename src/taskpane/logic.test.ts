import { describe, expect, it } from "vitest";
import {
  boundsEqual,
  boundsFromWindow,
  diffWindows,
  displayName,
  findWindowByKey,
  indexesOf,
  needsRealign,
  nextPollDelay,
  normalizeDisplayMode,
  sameOrder,
  sortWindows,
  uniqueLabels,
  windowKey,
  POLL_BACKOFF_MAX_MS,
  POLL_ALIGN_UNFOCUSED_MS,
  POLL_FOCUSED_MS,
  POLL_UNFOCUSED_MS,
  type WindowInfo
} from "./logic";

function makeWindow(partial: Partial<WindowInfo> & { index: number; name: string }): WindowInfo {
  return {
    windowState: "normal",
    isActive: false,
    geometry: { left: 0, top: 0, width: 1200, height: 800 },
    ...partial
  };
}

describe("sortWindows", () => {
  it("按 index 升序排序且不修改原数组", () => {
    const input = [makeWindow({ index: 3, name: "C" }), makeWindow({ index: 1, name: "A" })];
    const sorted = sortWindows(input);

    expect(sorted.map((window) => window.index)).toEqual([1, 3]);
    expect(input.map((window) => window.index)).toEqual([3, 1]);
  });
});

describe("displayName", () => {
  it("去掉窗口编号后缀", () => {
    expect(displayName(makeWindow({ index: 2, name: "Book1:2" }))).toBe("Book1");
  });

  it("空名称回退为“工作簿”", () => {
    expect(displayName(makeWindow({ index: 1, name: "   " }))).toBe("工作簿");
  });
});

describe("uniqueLabels", () => {
  it("重名窗口依次追加序号", () => {
    const labels = uniqueLabels([
      makeWindow({ index: 2, name: "预算.xlsx" }),
      makeWindow({ index: 1, name: "预算.xlsx" }),
      makeWindow({ index: 3, name: "预算.xlsx" }),
      makeWindow({ index: 4, name: "报表.xlsx" })
    ]);

    expect(labels.get(1)).toBe("预算.xlsx");
    expect(labels.get(2)).toBe("预算.xlsx (2)");
    expect(labels.get(3)).toBe("预算.xlsx (3)");
    expect(labels.get(4)).toBe("报表.xlsx");
  });
});

describe("boundsFromWindow", () => {
  it("有几何信息时返回副本", () => {
    const window = makeWindow({ index: 1, name: "A" });
    expect(boundsFromWindow(window)).toEqual({ left: 0, top: 0, width: 1200, height: 800 });
  });

  it("未加载几何信息时返回 null", () => {
    const window: WindowInfo = { index: 1, name: "A", windowState: "normal", isActive: true };
    expect(boundsFromWindow(window)).toBeNull();
  });
});

describe("boundsEqual", () => {
  it("允许默认 1pt 误差", () => {
    const base = { left: 10, top: 20, width: 1200, height: 800 };
    expect(boundsEqual(base, { left: 10.6, top: 20, width: 1200.4, height: 799.5 })).toBe(true);
    expect(boundsEqual(base, { left: 12, top: 20, width: 1200, height: 800 })).toBe(false);
  });
});

describe("indexesOf / sameOrder", () => {
  it("按顺序输出 index 并正确比较", () => {
    const windows = [makeWindow({ index: 5, name: "E" }), makeWindow({ index: 2, name: "B" })];
    expect(indexesOf(windows)).toEqual([2, 5]);
    expect(sameOrder([2, 5], [2, 5])).toBe(true);
    expect(sameOrder([2, 5], [5, 2])).toBe(false);
    expect(sameOrder([2], [2, 5])).toBe(false);
  });
});

describe("diffWindows", () => {
  it("识别新增窗口并触发渲染", () => {
    const before = [makeWindow({ index: 1, name: "A", isActive: true })];
    const after = [
      makeWindow({ index: 1, name: "A" }),
      makeWindow({ index: 2, name: "B", isActive: true })
    ];

    const diff = diffWindows(before, after);
    expect(diff.added).toEqual([2]);
    expect(diff.removed).toEqual([]);
    expect(diff.needsRender).toBe(true);
  });

  it("活动窗口变化触发渲染", () => {
    const before = [
      makeWindow({ index: 1, name: "A", isActive: true }),
      makeWindow({ index: 2, name: "B" })
    ];
    const after = [
      makeWindow({ index: 1, name: "A" }),
      makeWindow({ index: 2, name: "B", isActive: true })
    ];

    expect(diffWindows(before, after).needsRender).toBe(true);
  });

  it("仅几何变化不触发渲染", () => {
    const before = [makeWindow({ index: 1, name: "A" })];
    const after = [
      makeWindow({ index: 1, name: "A", geometry: { left: 120, top: 60, width: 1200, height: 800 } })
    ];

    const diff = diffWindows(before, after);
    expect(diff.needsRender).toBe(false);
    expect(diff.added).toEqual([]);
    expect(diff.removed).toEqual([]);
  });
});

describe("nextPollDelay", () => {
  it("隐藏时暂停轮询", () => {
    expect(nextPollDelay({ visible: false, focused: true, failStreak: 0 })).toBeNull();
  });

  it("聚焦 800ms、失焦 3000ms", () => {
    expect(nextPollDelay({ visible: true, focused: true, failStreak: 0 })).toBe(POLL_FOCUSED_MS);
    expect(nextPollDelay({ visible: true, focused: false, failStreak: 0 })).toBe(POLL_UNFOCUSED_MS);
  });

  it("失败退避按倍数增长并封顶", () => {
    expect(nextPollDelay({ visible: true, focused: true, failStreak: 1 })).toBe(POLL_FOCUSED_MS * 2);
    expect(nextPollDelay({ visible: true, focused: true, failStreak: 9 })).toBe(POLL_BACKOFF_MAX_MS);
    expect(nextPollDelay({ visible: true, focused: false, failStreak: 3 })).toBe(POLL_BACKOFF_MAX_MS);
  });

  it("对齐开启时失焦间隔更短，便于及时纠正窗口位置", () => {
    expect(
      nextPollDelay({ visible: true, focused: false, failStreak: 0, alignEnabled: true })
    ).toBe(POLL_ALIGN_UNFOCUSED_MS);
    expect(
      nextPollDelay({ visible: true, focused: false, failStreak: 0, alignEnabled: false })
    ).toBe(POLL_UNFOCUSED_MS);
  });
});

describe("needsRealign", () => {
  const frame = { left: 10, top: 20, width: 1200, height: 800 };

  it("缺少几何信息时不触发纠正（例如未按需加载）", () => {
    expect(needsRealign(undefined, frame)).toBe(false);
  });

  it("偏差在 2pt 以内视为已对齐", () => {
    expect(needsRealign({ left: 11, top: 21, width: 1200.5, height: 799.5 }, frame)).toBe(false);
  });

  it("偏差超过 2pt 时需要纠正", () => {
    expect(needsRealign({ left: 14, top: 20, width: 1200, height: 800 }, frame)).toBe(true);
    expect(needsRealign({ left: 10, top: 20, width: 1100, height: 800 }, frame)).toBe(true);
  });
});

describe("normalizeDisplayMode", () => {
  it("只接受 compact，其它值回退为 full", () => {
    expect(normalizeDisplayMode("compact")).toBe("compact");
    expect(normalizeDisplayMode("full")).toBe("full");
    expect(normalizeDisplayMode(undefined)).toBe("full");
    expect(normalizeDisplayMode(42)).toBe("full");
  });
});

describe("windowKey", () => {
  it("名字唯一时直接使用文档名（不受索引变化影响）", () => {
    const windows = [makeWindow({ index: 3, name: "A.xlsx" }), makeWindow({ index: 1, name: "B.xlsx" })];
    expect(windowKey(windows[0], windows)).toBe("A.xlsx");
    expect(windowKey(windows[1], windows)).toBe("B.xlsx");
  });

  it("重名时拼接索引以便区分", () => {
    const windows = [makeWindow({ index: 2, name: "同名.xlsx" }), makeWindow({ index: 5, name: "同名.xlsx" })];
    expect(windowKey(windows[0], windows)).toBe("同名.xlsx#2");
    expect(windowKey(windows[1], windows)).toBe("同名.xlsx#5");
  });

  it("空名字退化为索引标识", () => {
    const windows = [makeWindow({ index: 7, name: "  " })];
    expect(windowKey(windows[0], windows)).toBe("#7");
  });

  it("激活导致索引变化后仍能按名字找回同一个窗口", () => {
    const before = [makeWindow({ index: 2, name: "报告.xlsx" }), makeWindow({ index: 1, name: "汇总.xlsx" })];
    const after = [makeWindow({ index: 1, name: "报告.xlsx" }), makeWindow({ index: 2, name: "汇总.xlsx" })];
    const key = windowKey(before[0], before);
    const found = findWindowByKey(after, key);
    expect(found?.name).toBe("报告.xlsx");
    expect(found?.index).toBe(1);
  });
});
