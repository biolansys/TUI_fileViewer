#!/usr/bin/env bash
set -euo pipefail

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOG_FILE="$APP_DIR/example.log"

echo "[$(date +"%Y-%m-%d %H:%M:%S")] INFO starting demo" | tee -a "$LOG_FILE"

for file in "$APP_DIR"/*; do
  if [[ -f "$file" ]]; then
    echo "- $(basename "$file")"
  fi
done

echo "Done."
