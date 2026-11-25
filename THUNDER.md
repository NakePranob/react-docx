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

`v0.15.2-thunder.0` is upstream 0.15.2 as built here.
