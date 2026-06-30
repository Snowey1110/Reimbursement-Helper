import { describe, expect, it } from "vitest";
import { cropOutputSize, normalizedCropPoints, perspectiveCoefficients, rotateCropPoints, stackedPageLayout } from "./imageUtils";

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

  it("keeps a four-corner crop as a rectangular output size", () => {
    const points = [
      { x: 10, y: 20 },
      { x: 90, y: 30 },
      { x: 80, y: 180 },
      { x: 5, y: 170 }
    ];

    expect(cropOutputSize(points)).toEqual({ width: 81, height: 150 });
  });

  it("solves perspective coefficients for each crop corner", () => {
    const points = [
      { x: 10, y: 20 },
      { x: 90, y: 30 },
      { x: 80, y: 180 },
      { x: 5, y: 170 }
    ];
    const coefficients = perspectiveCoefficients(points, 81, 150);
    const map = (x: number, y: number) => {
      const [a, b, c, d, e, f, g, h] = coefficients;
      const denominator = g * x + h * y + 1;
      return {
        x: (a * x + b * y + c) / denominator,
        y: (d * x + e * y + f) / denominator
      };
    };

    expect(map(0, 0).x).toBeCloseTo(10);
    expect(map(0, 0).y).toBeCloseTo(20);
    expect(map(81, 150).x).toBeCloseTo(80);
    expect(map(81, 150).y).toBeCloseTo(180);
  });

  it("treats unchanged full-image crop points as no crop", () => {
    expect(
      normalizedCropPoints(
        [
          { x: 0, y: 0 },
          { x: 100, y: 0 },
          { x: 100, y: 200 },
          { x: 0, y: 200 }
        ],
        100,
        200
      )
    ).toBeUndefined();
  });
});
