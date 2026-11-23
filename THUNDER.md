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

Nothing in the editor itself is changed yet: `v0.15.2-thunder.0` is upstream
0.15.2 as built here.
