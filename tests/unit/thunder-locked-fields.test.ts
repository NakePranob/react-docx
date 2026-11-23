import { describe, expect, it } from "vitest";

import type { ParagraphNode } from "../../packages/doc-model/src";
import { editAroundLockedFields } from "../../packages/react-viewer/src/locked-fields";

// THUNDER: locked text fields edit as one unit (the template editor's chips).
const chip = (value: string, tag = `{{${value}}}`) => ({
  type: "form-field" as const,
  fieldType: "text" as const,
  tag,
  value,
  widget: { enabled: false },
});

const paragraph = (children: ParagraphNode["children"]): ParagraphNode => ({
  type: "paragraph",
  children,
});

// "ยอด [จำนวน] บาท"
const sample = paragraph([
  { type: "text", text: "ยอด " },
  chip("จำนวน"),
  { type: "text", text: " บาท", style: { bold: true } },
]);
const text = "ยอด จำนวน บาท";
const chipStart = 4;
const chipEnd = 9;

const shape = (children: ParagraphNode["children"]) =>
  children.map((child) => (child.type === "text" ? child.text : `[${child.type === "form-field" ? child.value : ""}]`));

describe("editAroundLockedFields", () => {
  it("Backspace after a chip removes the whole chip", () => {
    const edit = editAroundLockedFields(sample, text, chipEnd - 1, chipEnd, "");
    expect(shape(edit!.children)).toEqual(["ยอด ", " บาท"]);
    expect(edit!.caret).toBe(chipStart);
  });

  it("Delete before a chip removes the whole chip", () => {
    const edit = editAroundLockedFields(sample, text, chipStart, chipStart + 1, "");
    expect(shape(edit!.children)).toEqual(["ยอด ", " บาท"]);
  });

  it("typing inside a chip lands after it, the chip intact", () => {
    const edit = editAroundLockedFields(sample, text, chipStart + 2, chipStart + 2, "บาท");
    expect(shape(edit!.children)).toEqual(["ยอด ", "[จำนวน]", "บาท", " บาท"]);
    expect(edit!.caret).toBe(chipEnd + 3);
  });

  it("typing over a selection that covers part of a chip replaces the whole chip", () => {
    const edit = editAroundLockedFields(sample, text, 2, chipStart + 1, "X");
    expect(shape(edit!.children)).toEqual(["ยอ", "X", " บาท"]);
  });

  it("places plain text by offset, even where it reads like a chip's label", () => {
    // "เลขที่ [เลขที่]": typing after the chip must not move the chip.
    const same = paragraph([{ type: "text", text: "เลขที่ " }, chip("เลขที่")]);
    const edit = editAroundLockedFields(same, "เลขที่ เลขที่", 13, 13, "Y");
    expect(shape(edit!.children)).toEqual(["เลขที่ ", "[เลขที่]", "Y"]);
  });

  it("leaves a paragraph with no chip to the ordinary path", () => {
    const plain = paragraph([{ type: "text", text: "ยอด บาท" }]);
    expect(editAroundLockedFields(plain, "ยอด บาท", 1, 1, "X")).toBeUndefined();
  });

  it("inserts given nodes at the caret, a chip next to a chip", () => {
    const edit = editAroundLockedFields(sample, text, chipEnd, chipEnd, [chip("โทร")]);
    expect(shape(edit!.children)).toEqual(["ยอด ", "[จำนวน]", "[โทร]", " บาท"]);
    expect(edit!.caret).toBe(chipEnd + 3);
  });

  it("gives up when the children do not add up to the text", () => {
    expect(editAroundLockedFields(sample, "something else", 0, 1, "")).toBeUndefined();
  });
});
