#!/bin/sh
set -eu

usage() {
  cat <<'EOF'
Usage: ./scripts/setup.sh [--check] [--prefix DIR]

Copies the dependency-free PowerPoint Local package and launcher under DIR.
The default prefix is $HOME/.local. No package download is performed.
EOF
}

prefix="${HOME}/.local"
check_only=false
while [ "$#" -gt 0 ]; do
  case "$1" in
    --check) check_only=true ;;
    --prefix)
      [ "$#" -ge 2 ] || { echo "--prefix requires a directory" >&2; exit 64; }
      prefix=$2
      shift
      ;;
    -h|--help) usage; exit 0 ;;
    *) echo "Unknown option: $1" >&2; usage >&2; exit 64 ;;
  esac
  shift
done

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
project_root=$(CDPATH= cd -- "$script_dir/.." && pwd)
python_path=$(command -v python3 || true)
[ -n "$python_path" ] || { echo "Python 3 is required." >&2; exit 1; }
"$python_path" -c 'import sys; raise SystemExit(sys.version_info < (3, 11))' || {
  echo "Python 3.11 or newer is required." >&2
  exit 1
}

if [ "$check_only" = true ]; then
  PYTHONPATH="$project_root/src" "$python_path" -c 'from powerpoint_local.mermaid_bridge import bridge_status; print(bridge_status())'
  if [ -d "/Applications/Microsoft PowerPoint.app" ]; then
    echo "Microsoft PowerPoint: installed"
  else
    echo "Microsoft PowerPoint: not found (required for open, export, and capture tools)"
  fi
  exit 0
fi

package_parent="$prefix/lib/powerpoint-local"
package_root="$package_parent/powerpoint_local"
mkdir -p "$prefix/bin" "$package_parent"
rm -rf "$package_root"
mkdir -p "$package_root"
cp "$project_root"/src/powerpoint_local/*.py "$package_root/"

cat > "$prefix/bin/powerpoint-local" <<EOF
#!/bin/sh
PYTHONPATH="$package_parent" exec "$python_path" -m powerpoint_local.server "\$@"
EOF
chmod 755 "$prefix/bin/powerpoint-local"

echo "Installed executable: $prefix/bin/powerpoint-local"
