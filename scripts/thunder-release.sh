#!/usr/bin/env bash
# Build packages/react-viewer and tag a commit whose root IS that package, so
# a project installs it straight from git:
#
#   "@extend-ai/react-docx": "github:NakePranob/react-docx#v0.15.2-thunder.0"
#
# A git dependency installs the repository root, and this repository's root is
# the monorepo; the tagged commit holds only package.json, README, LICENSE and
# dist/. It is not on any branch: its parent is the previous release tag, so
# the release tags form their own history, and the source it was built from is
# named in the message.
#
# usage: scripts/thunder-release.sh <n>      → tag v<upstream>-thunder.<n>
set -euo pipefail

n="${1:?usage: scripts/thunder-release.sh <n>}"
root="$(git rev-parse --show-toplevel)"
cd "$root"

[ -z "$(git status --porcelain)" ] || { echo "working tree is not clean" >&2; exit 1; }
[ "$(git rev-parse --abbrev-ref HEAD)" = "thunder" ] || { echo "release from the thunder branch" >&2; exit 1; }

pnpm="npx -y pnpm@9.15.0"
$pnpm install --frozen-lockfile
# The WASM binary and its bindings are committed; building them needs Rust and
# is upstream's job. Only the JS around them is built here.
$pnpm --filter @extend-ai/react-docx-wasm run build:js
$pnpm -r --filter "./packages/*" --filter "!@extend-ai/react-docx-wasm" run build

upstream="$(node -p "require('./packages/react-viewer/package.json').version")"
tag="v${upstream}-thunder.${n}"
git rev-parse -q --verify "refs/tags/$tag" >/dev/null && { echo "$tag exists" >&2; exit 1; }

stage="$(mktemp -d)"
trap 'rm -rf "$stage"' EXIT
cp -R packages/react-viewer/dist "$stage/dist"
cp LICENSE "$stage/LICENSE"
cp packages/react-viewer/README.md "$stage/README.md"
cp THUNDER.md "$stage/THUNDER.md"
node -e '
  const fs = require("fs");
  const pkg = JSON.parse(fs.readFileSync("packages/react-viewer/package.json", "utf8"));
  pkg.version = process.argv[1];
  delete pkg.scripts;
  delete pkg.devDependencies;
  fs.writeFileSync(process.argv[2], JSON.stringify(pkg, null, 2) + "\n");
' "${upstream}-thunder.${n}" "$stage/package.json"

source="$(git rev-parse HEAD)"
index="$(mktemp -u)"
GIT_INDEX_FILE="$index" git --work-tree="$stage" add -A -f .
tree="$(GIT_INDEX_FILE="$index" git write-tree)"
rm -f "$index"

parent=()
previous="$(git tag --list "v*-thunder.*" --sort=-creatordate | head -1)"
[ -n "$previous" ] && parent=(-p "$(git rev-list -n 1 "$previous")")

commit="$(git commit-tree "$tree" ${parent[@]+"${parent[@]}"} -m "release ${tag}

Built from thunder ${source}.")"
git tag -a "$tag" "$commit" -m "$tag"
echo "tagged $tag → $commit (source $source)"
echo "push with: git push origin thunder $tag"
