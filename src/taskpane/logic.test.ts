import { describe, expect, it } from "vitest";
import {
  boundsEqual,
  diffWindows,
  displayName,
  sortWindows,
  uniqueLabels,
  type WindowInfo
} from "./logic";

function makeWindow(partial: Partial<WindowInfo> & { index: number; name: string }): WindowInfo {
  return {
    left: 0,
    top: 0,
    width: 1200,
    height: 800,
    windowState: "normal",
    isActive: false,
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

describe("boundsEqual", () => {
  it("允许默认 1pt 误差", () => {
    const base = { left: 10, top: 20, width: 1200, height: 800 };
    expect(boundsEqual(base, { left: 10.6, top: 20, width: 1200.4, height: 799.5 })).toBe(true);
    expect(boundsEqual(base, { left: 12, top: 20, width: 1200, height: 800 })).toBe(false);
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
    const after = [makeWindow({ index: 1, name: "A", left: 120, top: 60 })];

    const diff = diffWindows(before, after);
    expect(diff.needsRender).toBe(false);
    expect(diff.added).toEqual([]);
    expect(diff.removed).toEqual([]);
  });
});
