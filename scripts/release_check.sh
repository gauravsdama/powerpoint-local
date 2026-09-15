#!/bin/sh
set -eu

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
project_root=$(CDPATH= cd -- "$script_dir/.." && pwd)
mode=${1:-technical}
case "$mode" in
  technical|--publish) ;;
  *) echo "Usage: ./scripts/release_check.sh [--publish]" >&2; exit 64 ;;
esac
install_root=$(mktemp -d "${TMPDIR:-/tmp}/powerpoint-local-release.XXXXXX")
trap 'rm -rf "$install_root"' EXIT HUP INT TERM

cd "$project_root"
git diff --check
python3 -m compileall -q src tests
python3 -m unittest discover -s tests -v
./scripts/setup.sh --prefix "$install_root"
printf '%s\n' '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-03-26"}}' | "$install_root/bin/powerpoint-local" | grep -q '"name":"powerpoint-local"'

if git grep -n '/Users/' -- . ':!scripts/release_check.sh'; then
  echo "Release check failed: a tracked personal home path was found." >&2
  exit 1
fi

if git ls-files | grep -Eq '(^|/)(upstream|exports|build|dist|[^/]+\.egg-info)(/|$)|\.deck-manifest\.json$'; then
  echo "Release check failed: generated, private, or study-only output is tracked." >&2
  exit 1
fi

if [ "$mode" = "--publish" ] && [ ! -f LICENSE ]; then
  echo "Publish check failed: the repository owner has not selected a LICENSE." >&2
  exit 1
fi
if [ "$mode" = "--publish" ] && [ -n "$(git status --porcelain --untracked-files=all)" ]; then
  echo "Publish check failed: the working tree is not clean." >&2
  exit 1
fi

echo "PowerPoint Local technical release checks passed."
