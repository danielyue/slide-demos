#!/bin/bash
# Start a new stepped demo from the template.
#   tools/new_demo.sh NAME "Title for the demo list"    public: NAME/index.html, listed in demos.json
#   tools/new_demo.sh --private DIR/NAME.html           private (unpublished papers): copied outside this repo
set -euo pipefail
here="$(cd "$(dirname "$0")/.." && pwd)"
tpl="$here/templates/stepped-explainer/index.html"
if [ "${1:-}" = "--private" ]; then
  out="$2"; mkdir -p "$(dirname "$out")"; [ -e "$out" ] && { echo "exists: $out"; exit 1; }
  cp "$tpl" "$out"
  echo "created $out"
  echo "serve it privately at: https://mdr.dny/raw${out/#\~/$HOME}"
  exit 0
fi
name="$1"; title="${2:-$1}"
[ -e "$here/$name" ] && { echo "exists: $here/$name"; exit 1; }
mkdir -p "$here/$name" && cp "$tpl" "$here/$name/index.html"
uv run --no-project python -I - "$here/demos.json" "$title" "$name/" <<'PY'
import json, sys
p, title, path = sys.argv[1:4]
d = json.load(open(p)); d.append({"title": title, "path": path})
json.dump(d, open(p, "w"), indent=2); open(p, "a").write("\n")
PY
echo "created $here/$name/index.html; after pushing: https://danielyue.github.io/slide-demos/$name/"
