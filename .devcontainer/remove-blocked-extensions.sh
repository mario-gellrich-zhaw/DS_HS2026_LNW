#!/usr/bin/env bash
# Uninstall every extension listed with a "-" prefix in devcontainer.json.
# The "-" prefix only stops the devcontainer from installing an extension;
# extensions from Settings Sync or user defaults still get installed.
cd "$(dirname "$0")"
python3 -c '
import json
exts = json.load(open("devcontainer.json"))["customizations"]["vscode"]["extensions"]
print("\n".join(e[1:] for e in exts if e.startswith("-")))
' | while read -r ext; do
  code --uninstall-extension "$ext" >/dev/null 2>&1 && echo "Removed $ext"
done
true
