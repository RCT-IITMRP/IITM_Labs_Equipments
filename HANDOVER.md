# Project Handover — IITM Labs & Equipment Browser

Everything a new maintainer needs to run, change and deploy this project.
Last updated: 18 September 2026.

Private items (accounts, passwords, keys) are **not** in this file. They are listed in
`HANDOVER-PRIVATE.md`, which is given to the team lead directly and must never be committed.

---

## 1. What this project is

A public, single-page directory of IIT Madras / IITM Research Park labs and equipment,
served from GitHub Pages at **https://rct-iitmrp.github.io/IITM_Labs_Equipments/**.

An Excel workbook is the only data source. A Python script turns it into one self-contained
`index.html`. Two extra features call a Cloudflare Worker: an AI assistant and a contact form.

```
SharePoint/OneDrive .xlsx ──(GitHub Actions, daily)──▶ regenerate_browser.py ──▶ index.html ──▶ GitHub Pages
                                                                                      │
                            AI assistant:  browser ──▶ Worker  /        ──▶ Google Gemini
                            Contact form:  browser ──▶ Worker  /contact ──▶ Gmail SMTP ──▶ team inbox
```

There is no build step, no framework and no database. The page is plain HTML/CSS/JS with the
equipment data inlined as JSON.

---

## 2. Where everything lives

| Component | Location | Notes |
|---|---|---|
| Website source + page code | GitHub repo `RCT-IITMRP/IITM_Labs_Equipments` | `regenerate_browser.py` is the source of truth |
| Published site | GitHub Pages, `https://rct-iitmrp.github.io/IITM_Labs_Equipments/` | built by GitHub automatically |
| Data workbook | SharePoint/OneDrive | URL stored as GitHub secret `ONEDRIVE_EXCEL_URL` |
| Backend (AI + contact form) | Cloudflare Worker `iitm-equipment-assistant` | account `dharman@respark.iitm.ac.in` |
| Worker URL | `https://iitm-equipment-assistant.dharman.workers.dev` | `/` = AI proxy, `/contact` = contact form |
| Bot protection | Cloudflare Turnstile widget | site key is public and sits in the page |
| Outgoing email | Gmail account + app password | held as Worker secrets |

---

## 3. Repository layout

```
regenerate_browser.py     The whole website: Python + the HTML_TEMPLATE string (page HTML/CSS/JS)
index.html                GENERATED OUTPUT — never edit by hand, CI overwrites it
iitm_logo.png             Header logos
iitmrp_logo.png
.github/workflows/update-equipment.yml   Daily job: download workbook, regenerate, commit
iitm-worker/              Cloudflare Worker (backend)
  worker.js                 Entry point: routes /contact, else proxies to Gemini
  contact.js                Contact form: validation, Turnstile, rate limits, sending mail
  wrangler.toml             Worker config: bindings, rate limiter, KV namespace
  package.json              One dependency: worker-mailer (SMTP client)
  .dev.vars.example         Template for local testing (safe, no real values)
  .gitignore                Keeps node_modules/, .wrangler/ and .dev.vars out of git
CLAUDE.md                 Notes for AI coding assistants working on this repo
HANDOVER.md               This file
```

**Never commit** `iitm-worker/.dev.vars` (real credentials), `node_modules/` or `.wrangler/`.

---

## 4. How the website is built and published

The workflow **Update Equipment Browser** runs daily at 00:07 UTC, and can be started by hand
from the repo's **Actions** tab (**Run workflow**). It:

1. downloads the workbook using the `ONEDRIVE_EXCEL_URL` secret,
2. runs `regenerate_browser.py`,
3. commits `index.html` if it changed. GitHub Pages then republishes automatically.

**Rule: never edit `index.html`.** All page HTML, CSS and JavaScript live inside the
`HTML_TEMPLATE` string in `regenerate_browser.py`. Any hand edit is wiped on the next run.

### To change the page

```bash
python3 -m venv .venv && .venv/bin/pip install openpyxl    # first time only
.venv/bin/python regenerate_browser.py                     # rebuild index.html locally
open index.html                                            # preview
```
Then commit `regenerate_browser.py` (uploading through the GitHub web UI is fine) and run the
workflow so CI regenerates `index.html`.

Placeholders substituted at build time: `__DATA__`, `__FILTER_OPTIONS__`, `__DATE__`,
`__ASSISTANT_PROXY_URL__`, `__CONTACT_PROXY_URL__`, `__TURNSTILE_SITE_KEY__`.

---

## 5. The Cloudflare Worker

One Worker serves both features:

| Route | Purpose |
|---|---|
| `POST /` | AI assistant: forwards the request to Google Gemini, keeping the API key server-side |
| `POST /contact` | Contact form: validates, checks Turnstile, rate-limits, sends the email |

Configuration in `iitm-worker/wrangler.toml`:
- `compatibility_flags = ["nodejs_compat"]` — required by the mail library.
- Rate limiter `CONTACT_RATE_LIMITER`: 3 requests per 60 seconds per IP.
- KV namespace `CONTACT_KV` (id `197f66b890ca480e800064ca17ddbc7f`): daily submission counters.
- Variable `GEMINI_MODEL = "gemini-2.5-flash"`.

Secrets (values are set on Cloudflare, never in the repo):
`GEMINI_API_KEY`, `TURNSTILE_SECRET_KEY`, `GMAIL_USER`, `GMAIL_APP_PASSWORD`, `CONTACT_TO_EMAIL`.

### Deploying the Worker

```bash
cd iitm-worker
npm install
wrangler login          # sign in as the Cloudflare account owner
wrangler deploy
wrangler tail           # live logs, useful while testing
```

The allow-list in `worker.js` (`ALLOWED_ORIGINS`) must contain the exact site origin
`https://rct-iitmrp.github.io`. If the site ever moves, update it here **and** in the
Turnstile widget's hostname list, or both features stop working with a 403.

---

## 6. How the contact form works

The team's email address is never in the page. It exists only as the Worker secret
`CONTACT_TO_EMAIL`, so scrapers cannot harvest it.

Visitor fills the form → the page posts to `/contact` → the Worker checks everything → it sends
one plain-text email through Gmail to the team inbox.

- **From:** `"<Visitor name> via IITM Labs Website" <the website Gmail address>`
- **Reply-To:** the visitor's address, so **Reply** in Outlook goes straight to them.
- The visitor's address cannot be the sender: other providers would treat it as forged and
  Outlook would junk it.

Fields and limits (kept in sync in two places — `CONTACT_LIMITS` in `regenerate_browser.py`
and `LIMITS` in `contact.js`; change both):

| Field | Required | Max length |
|---|---|---|
| Name | yes | 100 |
| Title/Designation | yes | 200 |
| Email | yes | 254 |
| Company Name | yes | 150 |
| Message | yes | 5000 |

Protections, all enforced on the server:
- Requests accepted only from the site's own address.
- Cloudflare Turnstile token verified server-side, single use, checked for this form and site.
- 3 submissions per minute per IP; 5 per IP per day; 100 per day overall (protects the Gmail quota).
- All fields type-checked, length-checked and stripped of control characters, so nothing can
  inject extra email headers. Bodies over 16 KB rejected.
- A hidden honeypot field: bots that fill it get a success response and nothing is sent.
- Errors shown to visitors are generic; details go only to the Worker log.

### To change the email wording
Edit `iitm-worker/contact.js`: the `text` array (body), the `subject` line, and the sender
display name, all inside `sendEmail()`. Then `wrangler deploy`. No GitHub upload needed.

---

## 7. Local development and testing

```bash
# Terminal 1 — run the Worker locally with test credentials
cd iitm-worker
cp .dev.vars.example .dev.vars     # Turnstile test keys, CONTACT_DRY_RUN=1
wrangler dev --port 8788

# Terminal 2 — build a test page pointed at the local Worker and serve it
cd ..
mkdir -p ~/labs-test && cp iitm_logo.png iitmrp_logo.png ~/labs-test/
.venv/bin/python -c "
import regenerate_browser as r
r.CONTACT_PROXY_URL = 'http://localhost:8788/contact'
r.TURNSTILE_SITE_KEY = '1x00000000000000000000AA'
eq, f = r.extract_data(r.SOURCE_FILE)
r.generate_html(eq, f, '$HOME/labs-test/index.html')"
python3 -m http.server 8000 --directory ~/labs-test
```
Open http://localhost:8000. Opening `index.html` as a file does **not** work: the Worker rejects
requests that don't come from a real web address, and Turnstile won't run.

**Safety rule:** keep `CONTACT_DRY_RUN=1` in `.dev.vars`, which prints the email instead of
sending it. With `0`, local test submissions land in the real team inbox.

Check the page's JavaScript still compiles after editing the template:
```bash
python3 -c "import re;open('/tmp/page.js','w').write(re.search(r'<script>\n(.*?)\n</script>',open('index.html',encoding='utf-8').read(),re.S).group(1))" && node --check /tmp/page.js
```

---

## 8. Troubleshooting

| Symptom | Likely cause and fix |
|---|---|
| Contact form or AI assistant returns 403 | Site origin missing from `ALLOWED_ORIGINS` in `worker.js`, or the Turnstile widget's hostname list. Update and redeploy. |
| "Verification failed" on submit | Turnstile site key in `regenerate_browser.py` and secret key on the Worker are from different widgets. |
| "Too many attempts" | Rate limit hit: 3/minute per IP. Expected during testing; wait a minute. |
| "The contact form is not available right now" | A Worker secret is missing. Check with `wrangler secret list`. |
| Form worked, no email arrives | Gmail app password revoked or the account hit its 500/day limit. Check `wrangler tail` during a submission. |
| Email lands in Junk | Mark "Not junk" and add the sender to Outlook Safe Senders. Nothing to change in code. |
| Site data is stale | The daily workflow failed. Open Actions and read the run; the download step is the usual failure. |
| Workflow fails at "Download latest Excel" | The SharePoint share link in `ONEDRIVE_EXCEL_URL` expired or its owner lost access. Create a fresh link and update the secret. |

---

## 9. Known issues and possible next steps

- **Mobile layout:** the long header title makes the page about 713 px wide on a 390 px phone,
  so the page scrolls sideways. Pre-existing, unrelated to the contact form.
- **Gmail app passwords:** Google has said it will retire them eventually. If that happens,
  the alternative is a domain on Cloudflare plus Cloudflare Email Service.
- **`DEPLOYMENT.md`** is referenced in a few code comments but does not exist; this file replaces it.
- **Duplicate workflow file:** an extra copy of `update-equipment.yml` may exist in the project
  root. Only `.github/workflows/update-equipment.yml` is ever executed by GitHub.
- The AI assistant has no rate limiting of its own; only the contact form does.

---

## 10. Account ownership

As of the handover on 18 September 2026, every account this project depends on is owned by the
team. Nothing is tied to the departing intern:

| Thing | Owned by |
|---|---|
| GitHub account `RCT-IITMRP` (repo + published site) | Team |
| Cloudflare account (Worker, KV, Turnstile) | `dharman@respark.iitm.ac.in` |
| Sending Gmail account + app password | Team lead |
| Gemini API key | Team lead's Google account |
| Data workbook + its share link | Team lead's OneDrive |
| Destination inbox (`CONTACT_TO_EMAIL`) | Team |

Secret **names** are listed in section 5; their values live only in Cloudflare and in the GitHub
Actions secret `ONEDRIVE_EXCEL_URL`. To change any of them:

```bash
cd iitm-worker && wrangler secret put <NAME>     # prompts for the value, never stored in git
```
