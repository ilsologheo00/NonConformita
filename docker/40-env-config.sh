#!/bin/sh
# Run by the nginx image at start-up: writes the container's VITE_* variables
# into env.js so one image can be configured for any tenant/site/list.
set -e

out=/usr/share/nginx/html/env.js

escape() {
  printf '%s' "$1" | sed 's/\\/\\\\/g; s/"/\\"/g'
}

{
  echo "window.__ENV__ = {"
  for key in VITE_AZURE_CLIENT_ID VITE_AZURE_TENANT_ID VITE_SP_SITE VITE_SP_LIST VITE_SP_FIELDS VITE_SP_OPERATORS_LIST VITE_SP_ACTIONS_LIST VITE_SP_ACTIVITIES_LIST; do
    value=$(printenv "$key" || true)
    echo "  \"$key\": \"$(escape "$value")\","
  done
  echo "};"
} > "$out"

echo "env-config: wrote $out"
