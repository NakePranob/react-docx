// THUNDER: locked text fields — inline units the caret steps over.
//
// A form field of type "text" with `widget.enabled === false` is drawn as its
// value and edited as one unit: an edit that reaches into it removes or
// replaces the whole field, never part of its text. The editable DOM text of a
// paragraph includes the field's value, so edits arrive as text offsets; this
// maps them back onto the paragraph's children.

import type {
  FormFieldRunNode,
  TextRunNode,
  ParagraphNode,
} from "@extend-ai/react-docx-doc-model";

type ParagraphChild = ParagraphNode["children"][number];

export function isLockedField(child: ParagraphChild): child is FormFieldRunNode {
  return (
    child.type === "form-field" &&
    child.fieldType === "text" &&
    child.widget?.enabled === false
  );
}

/** What a locked field shows, as the viewer draws it. */
export function lockedFieldText(field: FormFieldRunNode): string {
  return field.value || " ";
}

/** A child's share of the paragraph's editable text, or undefined when this
 * mapping cannot place it (the edit then takes the ordinary path). */
function childEditableText(child: ParagraphChild): string | undefined {
  if (child.type === "text") {
    return child.fieldInstruction === undefined && !child.noteReference
      ? child.text
      : undefined;
  }
  if (child.type === "image") return child.floating ? "" : undefined;
  return isLockedField(child) ? lockedFieldText(child) : undefined;
}

function insertedLength(nodes: ParagraphChild[]): number {
  return nodes.reduce(
    (sum, node) =>
      sum +
      (node.type === "text"
        ? node.text.length
        : isLockedField(node)
        ? lockedFieldText(node).length
        : 0),
    0
  );
}

/**
 * Replaces `[start, end)` of the paragraph's editable text with `inserted`,
 * widened to whole locked fields wherever it reaches into one. An insertion
 * that lands inside a field goes after it.
 *
 * Undefined when the paragraph's children do not add up to `text`, or when
 * it holds no locked field and `inserted` is plain text — both cases belong
 * to the ordinary text path.
 */
export function editAroundLockedFields(
  paragraph: ParagraphNode,
  text: string,
  start: number,
  end: number,
  inserted: string | ParagraphChild[]
): { children: ParagraphChild[]; caret: number } | undefined {
  const pieces = paragraph.children.map(childEditableText);
  if (pieces.some((piece) => piece === undefined)) return undefined;
  if (pieces.join("") !== text) return undefined;

  const spans: { start: number; end: number }[] = [];
  let offset = 0;
  paragraph.children.forEach((child, index) => {
    const length = (pieces[index] as string).length;
    if (isLockedField(child)) spans.push({ start: offset, end: offset + length });
    offset += length;
  });

  let editStart = Math.max(0, Math.min(start, end));
  let editEnd = Math.min(text.length, Math.max(start, end));
  for (const span of spans) {
    if (editStart === editEnd) {
      if (editStart > span.start && editStart < span.end) {
        editStart = editEnd = span.end;
      }
    } else if (span.start < editEnd && span.end > editStart) {
      editStart = Math.min(editStart, span.start);
      editEnd = Math.max(editEnd, span.end);
    }
  }
  // Every edit of a paragraph that holds a field comes through here, not
  // only those that touch one: the text path finds a field again by
  // searching for its label, which lands on any plain text that reads the same.
  if (spans.length === 0 && typeof inserted === "string") return undefined;

  // Inserted text takes the style of the text it follows, as typing does.
  let insertedStyle: TextRunNode["style"];
  offset = 0;
  for (let index = 0; index < paragraph.children.length; index++) {
    const child = paragraph.children[index];
    const length = (pieces[index] as string).length;
    if (child.type === "text" && offset < editStart && editStart <= offset + length) {
      insertedStyle = child.style;
    }
    if (insertedStyle === undefined && offset >= editStart && child.type !== "image") {
      insertedStyle = child.style;
    }
    offset += length;
  }
  const insertion: ParagraphChild[] =
    typeof inserted === "string"
      ? inserted
        ? [{ type: "text", text: inserted, ...(insertedStyle ? { style: { ...insertedStyle } } : {}) }]
        : []
      : inserted.map((node) =>
          node.type !== "image" && !node.style && insertedStyle
            ? { ...node, style: { ...insertedStyle } }
            : node
        );

  const children: ParagraphChild[] = [];
  let placed = false;
  const place = () => {
    if (!placed) children.push(...insertion);
    placed = true;
  };
  offset = 0;
  paragraph.children.forEach((child, index) => {
    const piece = pieces[index] as string;
    const childStart = offset;
    const childEnd = offset + piece.length;
    offset = childEnd;
    if (child.type === "text") {
      if (childEnd <= editStart || childStart >= editEnd) {
        if (childStart >= editEnd) place();
        children.push(child);
        if (childEnd === editStart && editStart === editEnd) place();
        return;
      }
      const left = piece.slice(0, Math.max(0, editStart - childStart));
      const right = piece.slice(Math.max(0, editEnd - childStart));
      if (left) children.push({ ...child, text: left });
      place();
      if (right) children.push({ ...child, text: right });
      return;
    }
    if (isLockedField(child) && childStart >= editStart && childEnd <= editEnd && childEnd > childStart) {
      place();
      return;
    }
    if (childStart >= editEnd && piece.length > 0) place();
    children.push(child);
  });
  place();

  return { children, caret: editStart + insertedLength(insertion) };
}
