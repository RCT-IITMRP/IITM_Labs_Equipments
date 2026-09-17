# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A static, single-page directory of IIT Madras / IITM Research Park labs and equipment. An Excel workbook is the only data source; `regenerate_browser.py` turns it into a fully self-contained `index.html` (data inlined as JSON, no build step, no runtime fetches except the AI assistant), which is served from GitHub Pages.

```
OneDrive/SharePoint .xlsx ──(GitHub Actions)──▶ regenerate_browser.py ──▶ index.html ──▶ GitHub Pages
                                                                                │
                                         AI assistant: browser ──▶ iitm-worker (Cloudflare) ──▶ Gemini
                                         Contact form: browser ──▶ iitm-worker /contact ──▶ Gmail SMTP ──▶ team inbox
```

## Commands

```bash
# Regenerate index.html from the workbook in this folder (local venv already has openpyxl)
.venv/bin/python regenerate_browser.py

# Syntax-check the JavaScript embedded in the generated page
python3 -c "import re;open('/tmp/page.js','w').write(re.search(r'<script>\n(.*?)\n</script>',open('index.html',encoding='utf-8').read(),re.S).group(1))" && node --check /tmp/page.js

# Cloudflare Worker (AI proxy + contact form)
cd iitm-worker && npm install && wrangler deploy
wrangler secret put GEMINI_API_KEY      # the key lives only as a Worker secret
# contact form secrets: TURNSTILE_SECRET_KEY, GMAIL_USER, GMAIL_APP_PASSWORD, CONTACT_TO_EMAIL
cp .dev.vars.example .dev.vars && wrangler dev   # local test: Turnstile test keys, CONTACT_DRY_RUN=1
```

There is no test suite, linter, or package manifest. Preview by opening `index.html` directly in a browser.

## Things that are easy to get wrong

- **Never edit `index.html` by hand.** It is generated output and CI overwrites it. All page HTML/CSS/JS lives in the `HTML_TEMPLATE` string inside `regenerate_browser.py`.
- `HTML_TEMPLATE` is a raw string (`r"""…"""`) filled with plain `str.replace` on `__DATA__`, `__FILTER_OPTIONS__`, `__DATE__`, `__ASSISTANT_PROXY_URL__`, `__CONTACT_PROXY_URL__`, `__TURNSTILE_SITE_KEY__` — not `.format()`, so JS braces need no escaping, and JS regex backslashes are written exactly as they should appear in the browser.
- **Two workflow files have diverged.** The repo now runs the Python `requests` variant (daily at 00:07 UTC), which matches the root `update-equipment.yml` copy; the local `.github/workflows/update-equipment.yml` is the older curl/every-30-min version and is *not* what GitHub executes. Only the copy inside `.github/workflows/` in the GitHub repo runs.
- The workbook must be loaded **without** `read_only=True`: openpyxl only exposes `cell.hyperlink` with the full parser, and hyperlink targets are needed for clickable links.
- The AI assistant returns 403 when the page is opened locally — the Worker's `ALLOWED_ORIGINS` (`iitm-worker/worker.js`) only permits `https://rct-iitmrp.github.io` (site: `https://rct-iitmrp.github.io/IITM_Labs_Equipments/`, repo `RCT-IITMRP/IITM_Labs_Equipments`). The page calls the URL in `ASSISTANT_PROXY_URL` near the top of `regenerate_browser.py`.
- **The destination email address must never appear in `HTML_TEMPLATE`.** The "Contact Us" modal posts to the Worker's `/contact` route (`iitm-worker/contact.js`), which verifies Turnstile, rate-limits (binding + KV daily caps), validates, and sends via Gmail SMTP; the destination is the `CONTACT_TO_EMAIL` Worker secret. Field limits are duplicated in `CONTACT_LIMITS` (page) and `LIMITS` (`contact.js`) — change both.
- `DEPLOYMENT.md` is referenced by the script and the Worker but does not exist in this folder.
- This folder is not a git checkout; the GitHub repo is the source of truth, and CI commits `index.html` back to it.

## Data model (Python side)

- Every sheet **except** `Departments_and_Entities` is equipment data. That lookup sheet only populates the Department / Entity Type dropdowns.
- Columns are detected per sheet from row 1. Each row becomes a dict keyed by the cleaned header, so **header names differ between sheets** (`Lab Name` vs `Name of subordinate laboratories`, `PI` vs `Centre`, …). New sheets and new columns are picked up automatically.
- Cell hyperlinks are captured into a reserved per-row key, `LINKS_KEY = '_links'` → `{header: url}`. `cell_link()` drops `about:blank` placeholders and non-http/mailto/ftp schemes, and re-joins the URL fragment Excel stores separately in `hyperlink.location`.

## Page architecture (JS inside `HTML_TEMPLATE`)

- **Field resolution:** `FIELD_KEYS` maps each logical field (name, lab, prof, dept, operator, …) to its header variants in priority order; `eqKey`/`eqField`/`eqName`/… resolve them. When a new sheet uses a new header for a known field, add it to `FIELD_KEYS` — otherwise it renders as a generic info row.
- **Always use `dataKeys(e)`, not `Object.keys(e)`**, when enumerating a row's columns, so `_links` never leaks into table columns, card rows, or search.
- **`cardModel(e)` is the single description of what a card shows.** `renderCards()` draws it and `searchHaystack()` reads it, so the global search matches exactly the words visible on a card — including field labels such as `PI`, `Operator`, `Source`, but never link targets (a `mailto:` href) or headers that are not printed. Any change to card content must go through `cardModel`, or search results will drift from what users see.
- `CARD_TOP_KEYS_SET` lists the columns rendered in the structured card layout; every other column becomes a `"<header>: value"` row.
- **`render()` is the one update path:** `filterData()` → result count → `updateStats()` (the three header stat pills reflect the *filtered* rows) → `renderActiveFilters()` (filter chips, gold-highlighted controls) → card or table view. The unfiltered state is both search boxes empty and both dropdowns on their placeholder; `clearFilter(id)` resets one filter, `clearFilter()` all.
- **Links:** `fieldHtml(e, key, q)` renders a cell — a captured hyperlink makes the whole value clickable, otherwise bare URLs in the text are linkified. All hrefs go through `safeUrl()`, and search highlighting (`hl()`) is preserved inside link text.
- **AI assistant:** `aiCandidateIndex` (one `index|equipment|lab|department|professor` line per row) is prepended to the first user message only. The model must reply with JSON `{"reply": "...", "matches": [{"i": <index>, "why": "..."}]}`, where `i` indexes `equipmentData`.
