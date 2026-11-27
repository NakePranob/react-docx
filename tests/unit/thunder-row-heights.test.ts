import { describe, expect, it } from "vitest";

import { measuredRowHeightsByModelRow } from "../../packages/react-viewer/src/editor";

describe("measuredRowHeightsByModelRow", () => {
  it("files each drawn row under its model row, not its DOM position", () => {
    // Rows 1 and 3 hidden by a condition: only 0, 2, 4 are drawn.
    const drawn = [
      { rowIndex: 0, heightPx: 20 },
      { rowIndex: 2, heightPx: 39 },
      { rowIndex: 4, heightPx: 21 },
    ];
    expect(measuredRowHeightsByModelRow(drawn, [24, 24, 24, 24, 24])).toEqual([20, 24, 39, 24, 21]);
  });
  it("adds up a row drawn on two pages", () => {
    const drawn = [
      { rowIndex: 0, heightPx: 20 },
      { rowIndex: 1, heightPx: 12 },
      { rowIndex: 1, heightPx: 9 },
    ];
    expect(measuredRowHeightsByModelRow(drawn, [24, 24])).toEqual([20, 21]);
  });
});
