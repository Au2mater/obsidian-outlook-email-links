#!/usr/bin/env bash
# Copy the plugin into a vault's config folder: ./install.sh <vault>/.obsidian
set -euo pipefail
config_dir="${1:?usage: ./install.sh <vault config folder, e.g. vault/.obsidian>}"
[ -d "$config_dir" ] || { echo "No such folder: $config_dir" >&2; exit 1; }
target="$config_dir/plugins/outlook-email-links"
mkdir -p "$target"
here="$(cd "$(dirname "$0")" && pwd)"
node "$here/scripts/build.js" >/dev/null
cp "$here/manifest.json" "$here/main.js" "$target/"
echo "Installed into $target — reload Obsidian or toggle the plugin to pick up changes."
