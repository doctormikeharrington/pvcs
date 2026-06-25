#!/usr/bin/env node
/*
 * patch-staticrypt.js
 *
 * Post-build patch for the staticrypt-encrypted pages. Run automatically by
 * build.sh after staticrypt generates the pages. Safe to run repeatedly
 * (each patch is skipped if already applied).
 *
 * It fixes two login annoyances:
 *
 *   1. "Stay logged into BOTH the main site and /admin at the same time."
 *      The main pages and the admin page are on the same domain and, by
 *      default, staticrypt stores the saved login under the SAME browser key
 *      for every page. Logging into one therefore overwrites the other and
 *      forces a re-login when you switch. We give the admin page its own
 *      storage keys so the two logins no longer clobber each other.
 *
 *   2. "Stay logged in for the session even without ticking 'remember me'."
 *      Without the checkbox staticrypt saves nothing, so every page load asks
 *      for the password again. We also save the login in sessionStorage, which
 *      survives navigation within the visit and is automatically cleared when
 *      the browser is closed. Ticking "remember me" still gives the longer
 *      (localStorage) persistence as before.
 *
 * Usage:  node patch-staticrypt.js <file1.html> <file2.html> ...
 */

const fs = require("fs");

function patchFile(path) {
  let html;
  try {
    html = fs.readFileSync(path, "utf8");
  } catch (e) {
    console.error("  ! could not read " + path);
    return;
  }

  // Only touch generated staticrypt pages.
  if (html.indexOf("decryptOnLoadFromRememberMe") === -1) {
    return;
  }

  const isAdmin = /(^|\/)admin\//.test(path) || /admin[\\/]index\.html$/.test(path);
  let changed = false;

  // --- Fix 1: give the admin page its own storage keys ----------------------
  if (isAdmin && html.indexOf("staticrypt_admin_passphrase") === -1) {
    html = html
      .replace(/"staticrypt_passphrase"/g, '"staticrypt_admin_passphrase"')
      .replace(/"staticrypt_expiration"/g, '"staticrypt_admin_expiration"');
    changed = true;
  }

  // --- Fix 2a: read a saved login from sessionStorage as a fallback ---------
  const readNeedle =
    "const hashedPassword = localStorage.getItem(rememberPassphraseKey);";
  if (html.indexOf(readNeedle) !== -1) {
    html = html.replace(
      readNeedle,
      "const hashedPassword = localStorage.getItem(rememberPassphraseKey) || sessionStorage.getItem(rememberPassphraseKey);"
    );
    changed = true;
  }

  // --- Fix 2b: always save the login for the browser session ----------------
  const writeNeedle = "if (isRememberEnabled && isRememberChecked) {";
  if (
    html.indexOf("stay logged in for the browser session") === -1 &&
    html.indexOf(writeNeedle) !== -1
  ) {
    html = html.replace(
      writeNeedle,
      '// stay logged in for the browser session even if "remember me" is unchecked\n' +
        "        window.sessionStorage.setItem(rememberPassphraseKey, hashedPassword);\n\n" +
        "        " +
        writeNeedle
    );
    changed = true;
  }

  // --- Fix 2c: also clear sessionStorage on logout --------------------------
  const clearNeedle =
    "localStorage.removeItem(rememberPassphraseKey);\n            localStorage.removeItem(rememberExpirationKey);";
  if (
    html.indexOf("sessionStorage.removeItem(rememberPassphraseKey)") === -1 &&
    html.indexOf(clearNeedle) !== -1
  ) {
    html = html.replace(
      clearNeedle,
      clearNeedle +
        "\n            sessionStorage.removeItem(rememberPassphraseKey);"
    );
    changed = true;
  }

  if (changed) {
    fs.writeFileSync(path, html);
    console.log("  patched " + path);
  } else {
    console.log("  (already patched) " + path);
  }
}

const files = process.argv.slice(2);
if (!files.length) {
  console.error("Usage: node patch-staticrypt.js <file.html> ...");
  process.exit(1);
}
files.forEach(patchFile);
