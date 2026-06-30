import { describe, expect, it } from "vitest";
import { rotateCropPoints, stackedPageLayout } from "./imageUtils";

describe("PDF page stacking", () => {
  it("centers PDF pages and stacks them into one tall image layout", () => {
    const layout = stackedPageLayout(
      [
        { width: 100, height: 200 },
        { width: 50, height: 80 },
        { width: 80, height: 120 }
      ],
      10
    );

    expect(layout.width).toBe(100);
    expect(layout.height).toBe(420);
    expect(layout.placements).toEqual([
      { x: 0, y: 0 },
      { x: 25, y: 210 },
      { x: 10, y: 300 }
    ]);
  });
});

describe("crop rotation", () => {
  it("keeps four crop corners independent when rotating", () => {
    const rotated = rotateCropPoints(
      [
        { x: 10, y: 20 },
        { x: 90, y: 30 },
        { x: 80, y: 180 },
        { x: 5, y: 170 }
      ],
      100,
      200,
      90
    );

    expect(rotated).toEqual([
      { x: 30, y: 5 },
      { x: 180, y: 10 },
      { x: 170, y: 90 },
      { x: 20, y: 80 }
    ]);
  });
});
