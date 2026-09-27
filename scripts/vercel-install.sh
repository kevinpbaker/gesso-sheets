#!/usr/bin/env bash
#
# The install step for a build that has no sibling Gesso checkout.
#
# vendor/ is not committed — it is 2MB of tarballs that change on
# every engine iteration — so a fresh clone has no engine to build
# against. Locally `vendor-gesso.sh` packs one out of ../gesso; here
# there is no ../gesso, so this fetches the exact commit the last
# vendoring packed from, recorded in gesso.lock, and hands it to the
# same script. The deployed build therefore runs the engine the
# working copy ran, and not whatever is on Gesso's main branch.
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SHA="$(tr -d '[:space:]' < "$HERE/gesso.lock")"
GESSO="${TMPDIR:-/tmp}/gesso-$SHA"

if [ ! -d "$GESSO" ]; then
  echo "fetching gesso $SHA…"
  mkdir -p "$GESSO"
  git -C "$GESSO" init --quiet
  git -C "$GESSO" remote add origin https://github.com/kevinpbaker/gesso.git
  git -C "$GESSO" fetch --quiet --depth 1 origin "$SHA"
  git -C "$GESSO" checkout --quiet FETCH_HEAD
fi

echo "installing gesso's own dependencies…"
(cd "$GESSO" && pnpm install --frozen-lockfile)

# vendor-gesso.sh builds the packages, packs them under content
# hashes, rewrites package.json and pnpm-workspace.yaml to match, and
# installs. The rewrite is what the lockfile cannot survive: the
# hashes are a property of the bytes just built, so the committed
# lockfile names different tarballs and pnpm — which refuses to move
# a lockfile under CI — has to be told this once.
export NPM_CONFIG_FROZEN_LOCKFILE=false
GESSO_REPO="$GESSO" bash "$HERE/scripts/vendor-gesso.sh"
