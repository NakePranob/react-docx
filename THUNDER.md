# THUNDER fork of react-docx

A fork of [extend-hq/react-docx](https://github.com/extend-hq/react-docx) (MIT)
for the loan-management back office's document-template editor. Upstream has no
cell merge, shading, repeating header rows or cell padding, and keeps the
editor's model edits internal; this fork opens extension points so the
application builds those tools itself. Why, and what it decided:
`loan-management/docs/adr/0009-forked-docx-editor.md`.

## Branches and releases

- `main` follows upstream untouched.
- `thunder` is our work: upstream's release tag plus one commit per change
  below, each small and on its own so a rebase stays readable.
- Releases are tags `v<upstream>-thunder.<n>`, made by
  `scripts/thunder-release.sh <n>` from a clean `thunder`. A release commit's
  root is the built `@extend-ai/react-docx` package (package.json, dist/,
  LICENSE) — a git dependency installs the repository root — and is on no
  branch. The application pins one:

  ```json
  "@extend-ai/react-docx": "github:NakePranob/react-docx#v0.15.2-thunder.0"
  ```

## Building

Node from `.nvmrc` and pnpm 9.15.0. The WASM binary and its bindings are
committed by upstream, so no Rust toolchain is needed:

```bash
npx -y pnpm@9.15.0 install --frozen-lockfile
npx -y pnpm@9.15.0 --filter @extend-ai/react-docx-wasm run build:js
npx -y pnpm@9.15.0 -r --filter "./packages/*" --filter "!@extend-ai/react-docx-wasm" run build
```

## Upgrading to a new upstream release

```bash
git fetch upstream --tags
git rebase --onto <new-upstream-tag> <old-upstream-tag> thunder
scripts/thunder-release.sh 0      # v<new>-thunder.0
```

Then compare every document the application prints, page by page, before and
after (the application's tasks name them).

## Changes on `thunder`

| | change | files |
| --- | --- | --- |
| — | release script and this file | `scripts/thunder-release.sh`, `THUNDER.md` |
| 1 | controller exposes `applyModelChange` (one undoable step) and the viewer's selected cell block (`tableCellSelection`, reported by `DocxEditorViewer` through `reportTableCellSelection`); `DocxTableCellSelectionRange` exported | `react-viewer/src/editor.tsx`, `index.tsx`, `tests/unit/thunder-controller-api.test.ts` |
| 2 | a drag that starts on a cell's text and reaches another cell selects the block of cells (upstream only did so from a cell's padding, which a table with no cell padding does not have), and the selected cells are drawn — upstream tracked the block but never showed it | `react-viewer/src/editor.tsx` |
| 3 | a table with no cell spacing always collapses its borders, as Word does; cells with borders of their own no longer push the table into the separated model (1px between cells), which drew frames dotted and such tables wider | `react-viewer/src/editor.tsx` |
| 4 | Tab / Shift+Tab move between cells from a cell in edit mode too (upstream did so only from a cell's paragraph, so the second Tab left the table); Tab in the last cell adds a row | `react-viewer/src/editor.tsx` |
| 5 | Up / Down at a cell's first or last line go to the cell above or below in the same grid column; reading order sent them to the next cell, which could not be focused as a paragraph, and the caret was lost | `react-viewer/src/editor.tsx` |
| 6 | locked text fields (`form-field`, `fieldType: "text"`, `widget.enabled === false`) are one unit: drawn as their value with a highlight, a click selects the whole field, the caret goes beside it never into it, and an edit that reaches into one removes or replaces the whole field (`locked-fields.ts`); a paragraph that holds one is edited on the contenteditable path and every edit in it is placed by offset (the text path finds a field by searching for its label). Copy and cut write the field's `tag`; a style command over a range styles the fields in it, and a click on a field gives the editor's text range the field. `commitParagraphChildrenAtRange` on the controller; `inlineNodesFromText` on the viewer turns typed or pasted text into nodes | `react-viewer/src/editor.tsx`, `locked-fields.ts`, `tests/unit/thunder-locked-fields.test.ts` |
| 7 | `useDocxEditor({ transformImportedModel })` rewrites each imported model before the editor holds it, outside the history | `react-viewer/src/editor.tsx` |
| 8 | a column or row border dragged in any table commits to the model (upstream's own `commitTableGridResize`), so it is one undoable step and what is saved — upstream did so only for tables it had already canonicalised and kept the rest in view state; the two columns' widths show while dragging, worded by the viewer's `formatColumnWidth` | `react-viewer/src/editor.tsx` |
| 9 | `displayModel` on the viewer: a model it lays out and draws in place of the document's, of the same shape (so the places an edit names are the document's), for the template editor's "view as a case"; `isLockedParagraph` refuses typing in paragraphs it reduced to strips; a locked field committed from what is drawn keeps the document's value; with a display model the page is measured rather than estimated | `react-viewer/src/editor.tsx` |
| 10 | an exact table row is as tall as it says, however short: the 14px floor no longer applies to a row whose height rule is exact, in the estimate or the measurement (a strip of the case view is a one-twip row) | `react-viewer/src/editor.tsx` |
| 11 | a paragraph with a first-line indent (and no hanging indent, numbering or line break) has its lines counted by pretext with the indent as a first-row exclusion, the way it renders — upstream skipped pretext for it and counted by breaking at spaces, which over-counts Thai and broke pages a paragraph early | `react-viewer/src/editor.tsx` |
| 12 | editing a paragraph through its textarea: focus follows Enter onto the new paragraph (the focus attempt read a model from before the split); text `inlineNodesFromText` recognises becomes nodes there too, as on the contenteditable path, so the first variable put into a paragraph is a field; the input textarea takes no clicks (it sat 1px wide on the caret, and a click at a line's end put the caret at the paragraph's start); focus retries for 30 frames, so a paragraph Enter puts on a new page still gets it | `react-viewer/src/editor.tsx` |
| 13 | Enter under a display model splits the document's paragraph (`offsetsInSource`), by its own children when it holds a field (`splitAroundLockedFields`) — the drawn text was split as the document's, which put a field's sample value into the paragraph as text and moved fields; an empty half of a split takes the style at the caret, with a pending style laid over it, so the line Enter made keeps the font and size | `react-viewer/src/editor.tsx`, `locked-fields.ts` |
| 14 | Backspace or Delete in an empty body paragraph beside a table removes it (not the document's last), so two tables can stand one under the other; a drag from an empty cell selects cells, as one from a cell's text does | `react-viewer/src/editor.tsx` |
| 15 | editing around locked fields: Backspace / Delete that joins two paragraphs keeps every child of both (`joinParagraphChildren`) — the text-only rebuild dropped the fields — and puts the caret where the first one ended as drawn; a caret beside a field (element-anchored, `caretBesideLockedField`) is a real caret, not one a re-render destroyed, so it is no longer replaced by a stale one, and every text-offset walk steps around a field (`positionBesideLockedField`); a committed range in a paragraph with a field keeps its drawn offsets (it was clamped to the text runs and fell back before an earlier field); typed text joins the neighbouring run of the same style (`mergeTextRuns`), and after a field that ends a paragraph takes the field's style | `react-viewer/src/editor.tsx`, `locked-fields.ts`, `tests/unit/thunder-locked-fields.test.ts` |
| 16 | Enter in a table cell keeps the text: an edit through the paragraph's textarea clears the contenteditable path's continuation, which was left from a first keystroke there and, stale, made the next keystroke after Enter rewrite line 1 from its old text (`abc` Enter `Z` read `aZ`); a cell paragraph holding a variable (no textarea) splits at Enter in the document's paragraph as the textarea does, instead of committing its drawn text with a line break, which turned the variable's sample into plain text | `react-viewer/src/editor.tsx` |
| 17 | typing after Enter in a paragraph holding a field: the split-off half with no field is drawn again as a wrapped paragraph a few frames later, and the host focused before that went with it (focus fell to the body, keystrokes lost); `focusParagraphAtOffset` waits for the wrapped surface of a field-less paragraph and refocuses when the host it focused is replaced | `react-viewer/src/editor.tsx` |
| 18 | Backspace / Delete in the empty paragraph between two tables shrinks it to a 1pt line (`HAIRLINE_TWIPS`: exact 20-twip line, no spacing, 1pt mark) instead of removing it — two tables with nothing between them are one table to Word; a second press does nothing | `react-viewer/src/editor.tsx` |
| 19 | a block of selected cells is dropped when the selected table's shape changes (a merge or split renumbers each row's cells, and the kept block covered cells nobody selected); right-click on the text of a wrapped cell paragraph opens the table menu, as on the cell's empty part, unless text is selected | `react-viewer/src/editor.tsx` |
| 20 | click in a table cell: off the text (an empty paragraph, the blank beside or under a line) the caret goes to the nearest line's start or end, the cell selected first as Tab does; on a wrapped cell paragraph the caret is placed the same way; the click's own reconcile is skipped after it | `react-viewer/src/editor.tsx` |
| 21 | focus follows the wrapped textarea: `focusWrappedParagraphTextarea` waits a few frames for the textarea to reach the paragraph instead of giving up, a textarea mounted while focus is on the body takes it, and `focusTableCellParagraphAtOffset` retries a cell element that was replaced | `react-viewer/src/editor.tsx` |
| 22 | typing in a table cell after a click: the wrapped textarea's inline ref puts the session's selection back only when it has just mounted (it ran on every render, moving the caret behind the first character); a model sync takes the caret from the focused textarea, not the session from before the typing; a queued textarea focus gives up once typing starts; the pending click selection is dropped on input; `focusTableCellParagraphAtOffset` uses the latest `focusWrappedParagraphRange` (the one closed over missed the paragraph Enter just made); a click off the text finds the cell's paragraphs from the cell itself, wrapped cells having no contenteditable root | `react-viewer/src/editor.tsx` |
| 23 | formatting a selection: a point's offset is never inside a character cluster (`graphemeBoundaryAtOrAfter` in `resolveOffsetAtPoint`) — Thai marks have no width, so a drag took the marks before a word and left its own last mark out; `setDisplayRangeMapper` lets the viewer map a range in what it draws onto the document, so a style lands on the selected text where a field is drawn with a sample of another length; a style command reads a range just set before it renders; a double click selects the Thai word under it (`Intl.Segmenter`) | `react-viewer/src/pretext-layout.ts`, `react-viewer/src/editor.tsx`, `tests/unit/thunder-grapheme.test.ts` |
| 24 | a click on a page off the text (below the last line, as on a blank page) puts the caret at the end of the nearest body line above it; the paragraphs one numbered-list toggle makes count up (1., 2., …), carrying on from a numbered line just above, where each was given "1." | `react-viewer/src/editor.tsx` |
| 25 | a style pressed with nothing typed yet (ตัวหนา) goes onto what is inserted next: a variable from `inlineNodesFromText` (`withPendingStyle`) and text typed into a paragraph holding a field — both came out in the neighbouring run's style; the row resize zone across a row border is 8px, not 16 — a click into the top of a compact row's cell (an 8pt summary sheet) hit the zone, the caret stayed in the row above, and the text went there | `react-viewer/src/editor.tsx` |
| 26 | a second Backspace / Delete on the 1pt line between two tables (row 18) takes the line out, leaving the tables straight one under the other — the cover sheet is drawn so; the first press still only shrinks it | `react-viewer/src/editor.tsx` |
| 27 | a row inserted from the cell menu can be typed in: its empty paragraphs get a mark style from their run (they were drawn bare, nothing to click into), and an empty line is laid out the left-aligned way whatever its alignment (an empty right-aligned cell was a bare span, its click went to the cell beside it); a column resize handle catches the pointer only in a row its boundary is an edge of (`columnEdgesOfRow`) — on a 12-column grid of merged cells the 16px handles covered the middle of most cells, and a click there started a resize instead of putting the caret in | `react-viewer/src/editor.tsx`, `tests/unit/thunder-column-edges.test.ts` |
| 28 | a style or size on a selection that ends on a locked field — a selected chip, a line holding only a chip, a chip at the end of its line — reaches the field: range offsets were clamped to the paragraph's text without its fields, so the range stopped before the chip and nothing changed | `react-viewer/src/editor.tsx` |

`v0.15.2-thunder.0` is upstream 0.15.2 as built here.
