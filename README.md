# PVCS — Provincial Virtual Crisis Service (internal team site)

A password-protected static website for psychiatrists and physician assistants
working on Manitoba's Provincial Virtual Crisis Service. Same approach as the
mikeandkari.ca site: plain HTML/CSS/JS hosted free on **GitHub Pages**, custom
domain (**pvcsmanitoba.ca**) via **Cloudflare DNS** — with one addition: **the pages
are encrypted** so a single team password is required to read anything.

## How the password protection works

This uses **StatiCrypt**. Each page is encrypted with AES‑256 using your team
password. A visitor sees only a login box; nothing on the page can be read
(even by "View Source") until the correct password is entered. The password is
**never stored** in the files — only a random "salt" is. Once entered, the
"Remember me" option keeps a clinician logged in on their own browser for 30
days. Free, and works on plain GitHub Pages.

> One shared password, no usernames or accounts — as requested. A shared
> password is only as private as the people who hold it; if someone leaves the
> team, change it (see below) and everyone re-enters the new one.

## Important: public repo, encrypted-only

GitHub Pages is free only on a **public** repository. So this repo contains
**only the ENCRYPTED pages** — the plain, readable versions (`source/`) are
kept **on your Mac and never committed** (they're listed in `.gitignore`).
That way the public repo gives nothing away. Don't add the `source/` folder to
the repo.

## File structure

```
PVCS Website/
├── source/                 # PLAINTEXT pages — local only, NOT in the repo
│   ├── index.html          #   Home
│   ├── announcements.html  #   Team announcements
│   ├── schedule.html       #   On-call / coverage schedule
│   ├── resources.html      #   Protocols, guidelines, links
│   └── contacts.html       #   Team directory / key contacts
├── index.html ...          # ENCRYPTED pages (these get published)
├── css/style.css           # Styling (not secret)
├── js/main.js              # Navigation (not secret)
├── CNAME                   # Custom domain (pvcsmanitoba.ca)
├── .staticrypt.json        # Random salt (safe; not the password)
├── build.sh                # Encrypt source/ → encrypted pages
└── README.md
```

## Publishing / updating content

Because the build (encryption) happens before publishing, updates go through a
quick rebuild-and-push. The easiest way: ask me (Claude) to make the change —
I edit the `source/` pages, re-encrypt, and hand you the push command. Manually:

1. Edit the relevant page in `source/`.
2. Re-encrypt with the team password:
   ```bash
   STATICRYPT_PASSWORD='your-team-password' ./build.sh
   ```
   (Requires Node.js. This regenerates the encrypted pages in the repo root.)
3. Publish:
   ```bash
   cd "/Users/Mike/Desktop/Claude Projects/PVCS Website/PVCS Website"
   git add -A && git commit -m "Update content" && git push
   ```
GitHub Pages redeploys automatically within ~1–2 minutes.

## Changing the password

Re-run the build with the new password and push:
```bash
STATICRYPT_PASSWORD='new-team-password' ./build.sh
git add -A && git commit -m "Rotate password" && git push
```

## First-time deployment (one time)

1. **Register pvcsmanitoba.ca** at Cloudflare (Canadian presence required for .ca).
2. **Create a public GitHub repo** `pvcs` under `doctormikeharrington`.
3. **Push this folder** (plaintext `source/` is auto-excluded by `.gitignore`):
   ```bash
   cd "/Users/Mike/Desktop/Claude Projects/PVCS Website/PVCS Website"
   git remote add origin https://github.com/doctormikeharrington/pvcs.git
   git push -u origin main
   ```
4. **Enable Pages:** repo **Settings → Pages → Source = Deploy from a branch →
   main / (root) → Save.**
5. **Custom domain:** Settings → Pages → Custom domain → `pvcsmanitoba.ca` → Save;
   tick **Enforce HTTPS** when available.
6. **Cloudflare DNS** for pvcsmanitoba.ca:

   | Type  | Name | Content                          |
   |-------|------|----------------------------------|
   | CNAME | www  | doctormikeharrington.github.io   |
   | A     | @    | 185.199.108.153                  |
   | A     | @    | 185.199.109.153                  |
   | A     | @    | 185.199.110.153                  |
   | A     | @    | 185.199.111.153                  |

   Set Cloudflare SSL/TLS mode to **Full**.

## Security notes
- Appropriate for non-PHI team information (announcements, schedules,
  protocols, contacts). **Do not post patient-identifiable health information.**
- Pages include `noindex` so search engines won't list them.
- The encrypted pages currently use a **temporary** password until you set the
  real one (see "Changing the password").
