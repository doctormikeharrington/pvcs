#!/usr/bin/env bash
#
# Build script for the PVCS site.
# Encrypts every page in source/ with a single password and writes the
# password-protected pages to the site root (where GitHub Pages serves from).
#
# Usage:
#   STATICRYPT_PASSWORD='your-team-password' ./build.sh
# or run without the variable and you'll be prompted for the password.
#
set -euo pipefail
cd "$(dirname "$0")"

# --- Make Node/npx findable -------------------------------------------------
# When this file is double-clicked, macOS starts a shell that usually does NOT
# load your normal profile or nvm, so `npx` can appear "not found" even when
# Node is installed. Add the common Node locations to PATH before we use npx.
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
export NVM_DIR="${NVM_DIR:-$HOME/.nvm}"
if [ -s "$NVM_DIR/nvm.sh" ]; then . "$NVM_DIR/nvm.sh" >/dev/null 2>&1 || true; fi
for d in "$NVM_DIR"/versions/node/*/bin; do
  if [ -d "$d" ]; then export PATH="$d:$PATH"; fi
done
if command -v brew >/dev/null 2>&1; then export PATH="$(brew --prefix)/bin:$PATH"; fi

if ! command -v npx >/dev/null 2>&1; then
  echo
  echo "ERROR: Node.js (the 'npx' command) was not found on this Mac."
  echo "Install the LTS version from https://nodejs.org , then run this again."
  echo
  exit 1
fi
# ---------------------------------------------------------------------------

# --- Team password: pull from the macOS Keychain if available ---------------
# Store it once and you'll never be prompted again. Run this in Terminal:
#   security add-generic-password -a "$USER" -s pvcs-staticrypt -w
# (it will ask you to type the team password, then store it securely)
# To change it later, delete it first then re-add:
#   security delete-generic-password -s pvcs-staticrypt
if [ -z "${STATICRYPT_PASSWORD:-}" ]; then
  kc_pw="$(security find-generic-password -s pvcs-staticrypt -w 2>/dev/null || true)"
  if [ -n "$kc_pw" ]; then
    export STATICRYPT_PASSWORD="$kc_pw"
    echo "Using the team password stored in your macOS Keychain."
  fi
fi

echo "Encrypting pages from source/ ..."

# Note: css/ and js/ live at the repo root and are served directly
# (they contain no confidential information). Only the HTML pages are encrypted.

# Encrypt each HTML page in source/ into the repo root.
# --remember 30  : "Remember me" keeps the user logged in for 30 days (per browser)
# --config       : stores a stable salt so the remembered login keeps working
npx staticrypt source/*.html \
  --directory . \
  --config .staticrypt.json \
  --remember 30 \
  --short \
  --template login-template.html \
  --template-title "PVCS — Provincial Virtual Crisis Service" \
  --template-instructions "Please enter the team password." \
  --template-button "Enter" \
  --template-placeholder "Password" \
  --template-color-primary "#1f3a5f" \
  --template-color-secondary "#edf2f7"

echo "Encrypting the admin page with its own (separate) password ..."

# The admin page lives at source/admin/index.html (a subfolder, so it is NOT
# matched by the source/*.html glob above) and is encrypted with a DIFFERENT
# password so only administrators can open pvcsmanitoba.ca/admin .
# Its salt/config is kept separate so its "remember me" works independently.
mkdir -p admin
STATICRYPT_PASSWORD='admin' npx staticrypt source/admin/index.html \
  --directory admin \
  --config .staticrypt-admin.json \
  --remember 30 \
  --short \
  --template login-template.html \
  --template-title "PVCS — Administration" \
  --template-instructions "Administrators only. Please enter the admin password." \
  --template-button "Enter" \
  --template-placeholder "Password" \
  --template-color-primary "#1f3a5f" \
  --template-color-secondary "#edf2f7"

echo "Patching login persistence (session login + separate admin login) ..."

# Post-process the generated pages so that:
#   * the admin page keeps its own saved-login keys (so you can stay logged into
#     both the main site and /admin at once), and
#   * visitors stay logged in for the rest of the browser session even if they
#     don't tick "remember me".
# Safe to run every build; it skips pages that are already patched.
node patch-staticrypt.js *.html admin/index.html

echo "Done. Encrypted pages written to the site root (admin page at /admin)."
