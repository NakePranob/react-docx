import { describe, expect, it } from "vitest";

import { graphemeBoundaryAtOrAfter } from "../../packages/react-viewer/src/pretext-layout";

describe("graphemeBoundaryAtOrAfter", () => {
  const text = "ถ้าผู้ให้กู้ผ่อนเวลาให้แก่";
  it("moves an offset between a Thai consonant and its marks to after them", () => {
    const word = text.indexOf("ผ่อน");
    expect(graphemeBoundaryAtOrAfter(text, word - 2)).toBe(word);
    expect(graphemeBoundaryAtOrAfter(text, word - 1)).toBe(word);
    expect(graphemeBoundaryAtOrAfter(text, text.indexOf("แก่") - 1)).toBe(text.indexOf("แก่"));
  });
  it("leaves a boundary, the ends and plain text alone", () => {
    expect(graphemeBoundaryAtOrAfter(text, text.indexOf("ผ่อน"))).toBe(text.indexOf("ผ่อน"));
    expect(graphemeBoundaryAtOrAfter(text, 0)).toBe(0);
    expect(graphemeBoundaryAtOrAfter(text, text.length)).toBe(text.length);
    expect(graphemeBoundaryAtOrAfter("abc", 1)).toBe(1);
  });
});
