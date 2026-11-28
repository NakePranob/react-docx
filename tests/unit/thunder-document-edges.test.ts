import { describe, expect, it } from "vitest";

import {
  documentEdgeForKey,
  wholeDocumentRange,
} from "../../packages/react-viewer/src/editor";

const key = (key: string, mods: Partial<Record<"ctrlKey" | "metaKey" | "shiftKey" | "altKey", boolean>> = {}) => ({
  key,
  ctrlKey: false,
  metaKey: false,
  shiftKey: false,
  altKey: false,
  ...mods,
});

const paragraph = (text: string) => ({
  type: "paragraph" as const,
  children: [{ type: "text" as const, text }],
});

describe("documentEdgeForKey", () => {
  it("reads Ctrl+Home/End and Cmd+Up/Down as the document's ends", () => {
    expect(documentEdgeForKey(key("Home", { ctrlKey: true }))).toBe("start");
    expect(documentEdgeForKey(key("End", { ctrlKey: true }))).toBe("end");
    expect(documentEdgeForKey(key("ArrowUp", { metaKey: true }))).toBe("start");
    expect(documentEdgeForKey(key("ArrowDown", { metaKey: true }))).toBe("end");
  });

  it("leaves plain Home/End, arrows and Shift extensions to the line", () => {
    expect(documentEdgeForKey(key("Home"))).toBeUndefined();
    expect(documentEdgeForKey(key("End"))).toBeUndefined();
    expect(documentEdgeForKey(key("ArrowDown", { ctrlKey: true }))).toBeUndefined();
    expect(documentEdgeForKey(key("End", { ctrlKey: true, shiftKey: true }))).toBeUndefined();
    expect(documentEdgeForKey(key("ArrowUp", { metaKey: true, altKey: true }))).toBeUndefined();
  });
});

describe("wholeDocumentRange", () => {
  it("runs from the first paragraph to the end of the last, past a table", () => {
    const model = {
      nodes: [
        paragraph("สัญญาค้ำประกันเงินกู้"),
        {
          type: "table" as const,
          rows: [{ cells: [{ nodes: [paragraph("ลงชื่อ")] }] }],
        },
        paragraph("ท้าย"),
      ],
      metadata: {},
    };
    const range = wholeDocumentRange(model as never);
    expect(range?.start).toEqual({ location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 });
    expect(range?.end).toEqual({ location: { kind: "paragraph", nodeIndex: 2 }, offset: 4 });
  });

  it("ends inside the last cell when the document ends on a table", () => {
    const model = {
      nodes: [
        paragraph("หัว"),
        {
          type: "table" as const,
          rows: [{ cells: [{ nodes: [paragraph("ก")] }, { nodes: [paragraph("ขค")] }] }],
        },
      ],
      metadata: {},
    };
    const range = wholeDocumentRange(model as never);
    expect(range?.end.offset).toBe(2);
    expect(range?.end.location).toMatchObject({ tableIndex: 1, rowIndex: 0, cellIndex: 1 });
  });

  it("has nothing to select in an empty document", () => {
    expect(wholeDocumentRange({ nodes: [], metadata: {} } as never)).toBeUndefined();
  });
});
