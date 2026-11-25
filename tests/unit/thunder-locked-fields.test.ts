import { describe, expect, it } from "vitest";

import type { ParagraphNode } from "../../packages/doc-model/src";
import {
  editAroundLockedFields,
  editableTextLength,
  joinParagraphChildren,
  mergeTextRuns,
} from "../../packages/react-viewer/src/locked-fields";

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
    expect(shape(edit!.children)).toEqual(["ยอด ", "[จำนวน]", "บาท บาท"]);
    expect(edit!.caret).toBe(chipEnd + 3);
  });

  it("typing over a selection that covers part of a chip replaces the whole chip", () => {
    const edit = editAroundLockedFields(sample, text, 2, chipStart + 1, "X");
    expect(shape(edit!.children)).toEqual(["ยอX", " บาท"]);
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

describe("joinParagraphChildren", () => {
  it("keeps the chips of both paragraphs, in order", () => {
    const joined = joinParagraphChildren(
      [{ type: "text", text: "วันที่ " }, chip("date")],
      [{ type: "text", text: "เลขที่ " }, chip("number")]
    );
    expect(shape(joined)).toEqual(["วันที่ ", "[date]", "เลขที่ ", "[number]"]);
  });

  it("drops empty runs, and joins runs that meet with the same style", () => {
    const joined = joinParagraphChildren(
      [{ type: "text", text: "" }, chip("a"), { type: "text", text: "x" }],
      [{ type: "text", text: "y" }]
    );
    expect(shape(joined)).toEqual(["[a]", "xy"]);
  });

  it("keeps one run when both paragraphs are empty", () => {
    expect(shape(joinParagraphChildren([{ type: "text", text: "" }], [{ type: "text", text: "" }]))).toEqual([""]);
  });
});

describe("editableTextLength", () => {
  it("counts text runs and chips as drawn", () => {
    expect(editableTextLength(sample)).toBe(text.length);
  });
});

describe("typing next to a chip", () => {
  it("joins typed text to the run it follows", () => {
    let current = sample;
    let currentText = text;
    for (const [index, key] of ["ก", "ข"].entries()) {
      const edit = editAroundLockedFields(current, currentText, chipEnd + 4 + index, chipEnd + 4 + index, key);
      current = paragraph(edit!.children);
      currentText = currentText + key;
    }
    expect(shape(current.children)).toEqual(["ยอด ", "[จำนวน]", " บาทกข"]);
  });

  it("joins text typed after a chip that ends the paragraph into one run", () => {
    let current = paragraph([{ type: "text", text: "ef" }, chip("ชื่อ")]);
    for (const key of ["g", "h", "i"]) {
      const at = editableTextLength(current);
      current = paragraph(editAroundLockedFields(current, "ef" + "ชื่อ" + current.children.slice(2).map((c) => (c.type === "text" ? c.text : "")).join(""), at, at, key)!.children);
    }
    expect(shape(current.children)).toEqual(["ef", "[ชื่อ]", "ghi"]);
  });

  it("does not join runs of different styles", () => {
    expect(
      shape(mergeTextRuns([{ type: "text", text: "a" }, { type: "text", text: "b", style: { bold: true } }]))
    ).toEqual(["a", "b"]);
  });
});

describe("style of typed text", () => {
  it("after a chip that ends the paragraph takes the chip's style", () => {
    const styled = paragraph([{ type: "text", text: "ef" }, { ...chip("ชื่อ"), style: { fontFamily: "TH SarabunPSK" } }]);
    const edit = editAroundLockedFields(styled, "efชื่อ", 6, 6, "g");
    expect(edit!.children[2]).toEqual({ type: "text", text: "g", style: { fontFamily: "TH SarabunPSK" } });
  });
});
