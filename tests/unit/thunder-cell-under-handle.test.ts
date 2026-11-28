import { describe, expect, it } from "vitest";

import { firstTableCellLocationAt } from "../../packages/react-viewer/src/editor";

// Just enough of an element for the lookup: closest() and attributes.
function el(kind: "handle" | "cell" | "other", at?: [number, number, number]): Element {
  const cell = at
    ? {
        getAttribute: (name: string) =>
          ({ "data-docx-table-index": String(at[0]), "data-docx-row-index": String(at[1]), "data-docx-cell-index": String(at[2]) })[name] ?? null,
      }
    : null;
  return {
    closest: (selector: string) =>
      selector === "[data-docx-table-resize-handle]" ? (kind === "handle" ? {} : null) : kind === "cell" ? cell : null,
  } as unknown as Element;
}

describe("firstTableCellLocationAt", () => {
  it("looks past a column-resize handle laid over a narrow column", () => {
    expect(firstTableCellLocationAt([el("handle"), el("cell", [1, 4, 6])])).toEqual({
      tableIndex: 1,
      rowIndex: 4,
      cellIndex: 6,
    });
  });
  it("finds nothing where there is no cell", () => {
    expect(firstTableCellLocationAt([el("handle"), el("other")])).toBeUndefined();
  });
});

import { columnHandleHitSizePx } from "../../packages/react-viewer/src/editor";

describe("columnHandleHitSizePx", () => {
  it("keeps the full zone between wide columns", () => {
    expect(columnHandleHitSizePx(80, 120)).toBe(16);
  });
  it("leaves part of a narrow column to its cells", () => {
    expect(columnHandleHitSizePx(80, 10)).toBe(4);
    expect(columnHandleHitSizePx(20, 80)).toBe(12);
  });
  it("keeps the full zone at the table's right edge", () => {
    expect(columnHandleHitSizePx(80, undefined)).toBe(16);
  });
});
