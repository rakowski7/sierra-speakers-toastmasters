# Sierra Speakers Toastmasters — Project & Deployment Context

This file exists so that anyone (including an AI assistant in a fresh session, or
future-Mateusz after losing chat history) can understand how this project is wired
and — most importantly — **which Apps Script project to push to, and when.**

> ⚠️ The single most confusing thing about this repo: the code is deployed to **TWO
> separate Apps Script projects**, and `clasp` can only point at one at a time. Read
> the "Which project do I push to?" section before running `clasp push`.

---

## TL;DR — Which project do I push to?

| What you changed | Set `.clasp.json` scriptId to | After `clasp push`, also… |
|---|---|---|
| **Web app** look/behavior (`Styles.html`, `Index.html`, `JavaScript.html`, `WebApp.gs`, `Manifest.html`, `ServiceWorker.html`) | **OLD COPY** → `16wL_ABg-wQNZnNInRaWw8BZQVfTa0jcTy8aLqCdMLi7apoPoeuO1B4Gp` | **Redeploy the web app**: Apps Script editor → Deploy → Manage deployments → edit → **New version** → Deploy |
| **Menu / macros** (`Code.gs`, `GeminiResilience.gs`) | **CLUB SHEET** → `1NqhaX7d8f3OdLttJh2jr0s9VwyoZrwqsuALNQgq1TEZJtJ1_fz2zF95g` | Nothing — just reload the sheet; the menu rebuilds via `onOpen` |

`clasp push` only sends files to whatever `scriptId` is in `.clasp.json`. So before pushing,
open `.clasp.json` and make sure the `scriptId` matches the table above. GitHub is the
source of truth for the code; either project can be brought up to date by pushing from this repo.

---

## Architecture overview

There are **two Google Sheets** and **two Apps Script projects**:

1. **The master club sheet** — the real schedule the club uses.
   - This is where the **menu/macros** are bound (the "Toastmasters" menu: Start Role
     Confirmations, Generate Meeting Agenda, Draft Club Meeting Email, etc.).
   - **Owned by Steve** (the club treasurer). Mateusz is only an *editor*. Members have edit access.

2. **The "old copy" sheet** ("Copy of Sierra Speakers TM Schedule") — owned by Mateusz.
   - Its bound script **hosts the deployed web app** (the mobile PWA).
   - ⚠️ **DO NOT DELETE this sheet.** It is doing nothing visible, but it is the host for
     the live web app. Recommend renaming it to something like
     "⚠️ Hosts Sierra Speakers web app — do not delete".

**Important:** the web app does NOT read the old copy's data. It reads/writes the **master
club sheet** by ID. So no matter which project hosts the code, the data always comes from
the master sheet:
- `WebApp.gs` → `WEBAPP_CONFIG_.SPREADSHEET_ID` = master club sheet ID.
- `Code.gs` → `SCHEDULING_SHEET_URL_` = master club sheet URL.

---

## Key IDs & links

| Thing | Value |
|---|---|
| Master club sheet ID | `1BAQPmjadcvjnPOsj16H6y22Ls-t34EeY63jlGPHPGU4` |
| Old copy sheet ID (web-app host) | `1gLWiPXAzW_LXw-7_GDG7ykaGOB2faa3WcY6h18qH-rw` |
| Club-sheet bound script ID (MENU) | `1NqhaX7d8f3OdLttJh2jr0s9VwyoZrwqsuALNQgq1TEZJtJ1_fz2zF95g` |
| Old-copy bound script ID (WEB APP) | `16wL_ABg-wQNZnNInRaWw8BZQVfTa0jcTy8aLqCdMLi7apoPoeuO1B4Gp` |
| Web app URL (`/exec`) | `https://script.google.com/macros/s/AKfycbzrzHWHYb_rTtJx6o5kySKuF9qIHKz12JbiGmOTgmxefi_cwRykDC0s09IwkwDxp2zP/exec` |
| GitHub repo | `github.com/rakowski7/sierra-speakers-toastmasters` (branch `main`) |
| Local path (Windows) | `C:\Users\rakow\OneDrive\Desktop\Projects\Toastmasters\github-repo` |

---

## Why two projects? (the ownership constraint)

Ideally everything (menu + web app) would live in **one** project bound to the master club
sheet. The menu got there fine — an *editor* can add and run a bound script. But **publishing
a web app from a script is owner-only**, and Steve owns the master sheet, so deploying the web
app from the club-sheet project is **blocked** for Mateusz.

So the web app stays deployed from the old copy (which Mateusz owns), and the menu lives on the
club sheet. That split is why `clasp` has to be flipped back and forth.

**The clean fix (someday):** ask Steve to transfer ownership of the master sheet to Mateusz
(there is no "co-owner" on a normal Drive file — only ownership *transfer*). Then the web app
can be deployed from the club-sheet project, `.clasp.json` can point at the club sheet
permanently, and the old copy can be retired. Until then, the two-project dance is expected.

---

## Web app deployment settings ("Plan A")

The web app is deployed with:
- **Execute as:** *User accessing the web app*
- **Who has access:** *Anyone with Google account*

This is "Plan A": each visitor runs as themselves, so `Session.getActiveUser().getEmail()`
returns the visitor's own email, which is matched against the member directory on the sheet.
(An earlier approach using Google Identity Services / GSI sign-in was removed because Google
blocks `googleusercontent.com` as an OAuth origin inside the Apps Script iframe.)
Do NOT reintroduce a name-picker — it allowed impersonation. If a visitor can't be identified,
the app shows a clean "we couldn't identify you" message.

`appsscript.json` `webapp` block should stay: `executeAs: USER_ACCESSING`, `access: ANYONE`
(NOT `ANYONE_ANONYMOUS`).

---

## What members need to use the web app

No per-person permission to grant. Each member just needs to:
1. Open the web app URL **signed into the Google account the club has on file for them**
   (must match the email in the member directory on the master sheet).
2. Click through the **one-time Google authorization** screen (it shows an "unverified app"
   warning — that's expected; Advanced → Go to… → Allow).

They also need **edit access to the master club sheet** (so their "Confirm" can write the
green cell) — members already have this. Edge case: members on Google **Workspace** accounts
(custom domains, not @gmail) may be blocked by their org's admin from authorizing an unverified
app; personal Gmail users are fine.

---

## Script Properties (API keys) — per project!

Script Properties do NOT travel with `clasp push`. Each project keeps its own. The keys that
matter (set in **both** projects' Project Settings → Script Properties if needed):
- `GEMINI_API_KEY` — Google Gemini (agenda word-of-the-day definitions, AI club email)
- `MW_API_KEY` — Merriam-Webster dictionary lookups

Everything else (`GEMINI_MODELS_CACHE`, `LAST_*`, `_confirmState`, ping counters, etc.) is
auto-generated runtime cache — no need to copy it; run **Toastmasters → Update AI Models** to
rebuild the model cache.

> 🔐 TODO: regenerate `GEMINI_API_KEY` in Google AI Studio (it was briefly exposed in a
> screenshot during setup) and update it in the project's Script Properties. Never commit key
> values to this repo.

---

## Common recipes

**I changed how the web app looks or works:**
1. Set `.clasp.json` scriptId → old copy (`16wL…`).
2. `clasp push`
3. Old copy → Extensions → Apps Script → Deploy → Manage deployments → edit → New version → Deploy.
4. Hard-refresh the web app URL to see changes.

**I changed a menu macro (Code.gs):**
1. Set `.clasp.json` scriptId → club sheet (`1Nqh…`).
2. `clasp push`
3. Reload the master sheet; the menu rebuilds automatically.

**Push code to GitHub (always do this after any change):**
```
git add -A
git commit -m "describe the change"
git push
```

**clasp auth issues:** run `clasp login` once (opens a browser; sign in as Mateusz).
**Stale `index.lock` error:** delete `.git/index.lock` (`del .git\index.lock` on Windows), then retry.

---

## Known rough edges / future work

- **Admin generators still reference "the active sheet" in spots.** A few functions in
  `Code.gs` use `SpreadsheetApp.getActiveSpreadsheet()` instead of the master-sheet ID. When
  triggered from the **web app** (hosted on the old copy), those could read the old copy
  instead of the master sheet. The core schedule view + role confirmation are fine (they go
  through `WEBAPP_CONFIG_.SPREADSHEET_ID`). Clean these up if the agenda/email admin features
  misbehave from the app.
- **Two-project drift.** Keep GitHub as the source of truth; re-push from the repo to whichever
  project is behind.
- **Native app** is the eventual goal (replacing the PWA web app).
- **Consolidation** onto a single project depends on Steve transferring sheet ownership.
