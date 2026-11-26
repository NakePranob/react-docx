import { describe, expect, it } from "vitest";

import { columnEdgesOfRow } from "../../packages/react-viewer/src/editor";

const cell = (gridSpan?: number) => ({ type: "table-cell", style: gridSpan ? { gridSpan } : undefined, nodes: [] });

describe("columnEdgesOfRow", () => {
  it("names only the boundaries a row's cells end on", () => {
    // the cover sheet's fee row on a 12-column grid: 4 | 2 | 1 | 5
    const row = { type: "table-row", cells: [cell(4), cell(2), cell(), cell(5)] };
    const table = { type: "table", rows: [row] };
    expect([...columnEdgesOfRow(table as never, row as never)].sort((a, b) => a - b)).toEqual([3, 5, 6, 11]);
  });
  it("starts after the grid columns a row skips", () => {
    const row = { type: "table-row", style: { gridBefore: 2 }, cells: [cell(2), cell()] };
    const table = { type: "table", rows: [row] };
    expect([...columnEdgesOfRow(table as never, row as never)].sort((a, b) => a - b)).toEqual([1, 3, 4]);
  });
});
