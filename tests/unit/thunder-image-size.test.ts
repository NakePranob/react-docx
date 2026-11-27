import { describe, expect, it } from "vitest";

import { fitImageSize } from "../../packages/react-viewer/src/editor";

describe("fitImageSize", () => {
  it("keeps a square picture square", () => {
    expect(fitImageSize({ width: 128, height: 128 }, 240)).toEqual({ widthPx: 128, heightPx: 128 });
  });
  it("shrinks a picture wider than the room, keeping its shape", () => {
    expect(fitImageSize({ width: 128, height: 128 }, 105)).toEqual({ widthPx: 105, heightPx: 105 });
    expect(fitImageSize({ width: 1200, height: 600 }, 240)).toEqual({ widthPx: 240, heightPx: 120 });
  });
  it("falls back to a width alone when the file could not be read", () => {
    expect(fitImageSize(undefined, 240)).toEqual({ widthPx: 240 });
  });
});
