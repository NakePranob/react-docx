import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import { useDocxEditor, type DocxEditorController } from "../../packages/react-viewer/src";

// THUNDER: the extension points the loan-management template editor builds on.
describe("controller extension points", () => {
  it("exposes model edits and the viewer's cell selection", () => {
    let editor: DocxEditorController | undefined;
    function Probe() {
      editor = useDocxEditor();
      return React.createElement("div");
    }
    renderToStaticMarkup(React.createElement(Probe));

    expect(typeof editor?.applyModelChange).toBe("function");
    expect(typeof editor?.reportTableCellSelection).toBe("function");
    expect(editor?.tableCellSelection).toBeUndefined();
  });
});
