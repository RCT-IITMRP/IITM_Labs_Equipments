/**
 * Contact form → email to the team inbox   (POST /contact)
 * ─────────────────────────────────────────────────────────────────────────
 * The GitHub Pages site posts the "Contact Us" form here. Checks run
 * cheapest first: origin → method/type → burst rate limit → body size →
 * field validation → daily caps → Turnstile → send via Gmail SMTP.
 *
 * Nothing sensitive lives in the page or in this repository. Worker secrets
 * (set with `wrangler secret put <NAME>`):
 *   TURNSTILE_SECRET_KEY  Turnstile widget secret key
 *   GMAIL_USER            the website's dedicated Gmail address (the sender)
 *   GMAIL_APP_PASSWORD    a Google App Password for that account
 *   CONTACT_TO_EMAIL      destination address (the team inbox)
 * Bindings (wrangler.toml): CONTACT_RATE_LIMITER (optional), CONTACT_KV.
 * Local-only vars (.dev.vars): DEV_ALLOWED_ORIGIN, CONTACT_DRY_RUN.
 */

import { WorkerMailer } from "worker-mailer";

// Field limits — keep in sync with CONTACT_LIMITS in regenerate_browser.py.
const LIMITS = { name: 100, title: 200, email: 254, company: 150, message: 5000 };
const MAX_BODY_BYTES = 16 * 1024;
const MAX_TOKEN_LENGTH = 2048;

// Daily submission caps (UTC day). The per-IP cap stops one sender flooding
// the inbox; the global cap protects the Gmail account's sending reputation.
const PER_IP_DAILY_LIMIT = 5;
const GLOBAL_DAILY_LIMIT = 100;

const TURNSTILE_VERIFY_URL = "https://challenges.cloudflare.com/turnstile/v0/siteverify";
const TURNSTILE_ACTION = "contact";
// Cloudflare's documented always-pass test secret. Its responses carry a
// placeholder hostname/action, so only for it are those two checks skipped.
const TURNSTILE_TEST_SECRET = "1x0000000000000000000000000000000AA";

const EMAIL_RE = /^[A-Za-z0-9.!#$%&'*+\/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$/;
const TLD_RE = /\.(?:[A-Za-z]{2,63}|xn--[A-Za-z0-9-]{1,59})$/;

// Control, zero-width and bidi-override characters, as code point ranges.
const UNSAFE_RANGES = [[0x00, 0x1f], [0x7f, 0x9f], [0x200b, 0x200f], [0x2028, 0x202e], [0x2060, 0x2069], [0xfeff, 0xfeff]];
const isUnsafe = cp => UNSAFE_RANGES.some(([lo, hi]) => cp >= lo && cp <= hi);

function corsHeaders(origin) {
  return {
    "Access-Control-Allow-Origin": origin,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Max-Age": "86400",
    "Vary": "Origin",
  };
}

function json(status, body, headers) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store", ...headers },
  });
}

/** Single-line field: unsafe characters become spaces (so CR/LF can never
 *  reach an email header), runs of whitespace collapse. */
function cleanLine(value) {
  return Array.from(value, ch => (isUnsafe(ch.codePointAt(0)) ? " " : ch)).join("").replace(/\s+/g, " ").trim();
}

/** Message: keep line breaks and tabs, drop other unsafe characters. */
function cleanText(value) {
  return Array.from(value.replace(/\r\n?/g, "\n"), ch => {
    const cp = ch.codePointAt(0);
    return cp === 0x0a || cp === 0x09 || !isUnsafe(cp) ? ch : "";
  }).join("").replace(/\n{4,}/g, "\n\n\n").trim();
}

/** Display names sit inside From / Reply-To headers: also drop characters
 *  that are special in an address header. */
function headerName(value) {
  return value.replace(/["<>()[\]\\,;:@]/g, "").replace(/\s+/g, " ").trim().slice(0, 60);
}

function isValidEmail(email) {
  if (email.length > LIMITS.email || !EMAIL_RE.test(email) || !TLD_RE.test(email)) return false;
  const local = email.slice(0, email.lastIndexOf("@"));
  return local.length <= 64 && !local.startsWith(".") && !local.endsWith(".") && !local.includes("..");
}

/** Returns { data } or { error, field } with a visitor-facing message. */
function validate(input) {
  for (const key of ["name", "title", "email", "company", "message", "turnstileToken"]) {
    if (input[key] !== undefined && typeof input[key] !== "string") return { error: "Invalid request." };
  }
  const name = cleanLine(input.name || "");
  const title = cleanLine(input.title || "");
  const email = cleanLine(input.email || "");
  const company = cleanLine(input.company || "");
  const message = cleanText(input.message || "");
  const token = (input.turnstileToken || "").trim();

  if (!name) return { error: "Please enter your name.", field: "name" };
  if (name.length > LIMITS.name) return { error: `Name must be ${LIMITS.name} characters or fewer.`, field: "name" };
  if (!title) return { error: "Please enter your title or designation.", field: "title" };
  if (title.length > LIMITS.title) return { error: `Title/Designation must be ${LIMITS.title} characters or fewer.`, field: "title" };
  if (!email) return { error: "Please enter your email address.", field: "email" };
  if (!isValidEmail(email)) return { error: "Please enter a valid email address.", field: "email" };
  if (!company) return { error: "Please enter your company name.", field: "company" };
  if (company.length > LIMITS.company) return { error: `Company name must be ${LIMITS.company} characters or fewer.`, field: "company" };
  if (!message) return { error: "Please enter your message.", field: "message" };
  if (message.length > LIMITS.message) return { error: `Message must be ${LIMITS.message} characters or fewer.`, field: "message" };
  if (!token || token.length > MAX_TOKEN_LENGTH) return { error: "Please complete the verification check.", field: "turnstile" };
  return { data: { name, title, email, company, message, token } };
}

/** Read the body, giving up (null) as soon as it exceeds MAX_BODY_BYTES. */
async function readBody(request) {
  if (Number(request.headers.get("Content-Length") || 0) > MAX_BODY_BYTES) return null;
  if (!request.body) return "";
  const reader = request.body.getReader();
  const chunks = [];
  let size = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    size += value.byteLength;
    if (size > MAX_BODY_BYTES) {
      await reader.cancel();
      return null;
    }
    chunks.push(value);
  }
  const bytes = new Uint8Array(size);
  let offset = 0;
  for (const c of chunks) { bytes.set(c, offset); offset += c.byteLength; }
  return new TextDecoder().decode(bytes);
}

async function sha256Hex(text) {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return [...new Uint8Array(digest)].map(b => b.toString(16).padStart(2, "0")).join("");
}

/** Per-IP and global counters for today. The IP is stored only as a salted
 *  hash. KV is eventually consistent, so the caps are approximate — they are
 *  abuse protection, not accounting. */
async function dailyCaps(env, ip) {
  const day = new Date().toISOString().slice(0, 10);
  const ipKey = `contact:ip:${day}:${(await sha256Hex(`${day}|${ip}`)).slice(0, 32)}`;
  const allKey = `contact:all:${day}`;
  const [ipCount, allCount] = (await Promise.all([env.CONTACT_KV.get(ipKey), env.CONTACT_KV.get(allKey)]))
    .map(v => parseInt(v || "0", 10) || 0);
  const ttl = { expirationTtl: 2 * 24 * 60 * 60 };
  return {
    ok: ipCount < PER_IP_DAILY_LIMIT && allCount < GLOBAL_DAILY_LIMIT,
    record: () => Promise.all([
      env.CONTACT_KV.put(ipKey, String(ipCount + 1), ttl),
      env.CONTACT_KV.put(allKey, String(allCount + 1), ttl),
    ]),
  };
}

/** Returns "ok", "failed" (bad/expired/replayed token) or "unavailable". */
async function verifyTurnstile(token, ip, env, allowedOrigins) {
  const form = new FormData();
  form.append("secret", env.TURNSTILE_SECRET_KEY);
  form.append("response", token);
  if (ip) form.append("remoteip", ip);

  let outcome;
  try {
    const res = await fetch(TURNSTILE_VERIFY_URL, { method: "POST", body: form });
    outcome = await res.json();
  } catch (e) {
    console.error("Turnstile siteverify unreachable:", e && e.message);
    return "unavailable";
  }
  if (!outcome.success) {
    const codes = outcome["error-codes"] || [];
    console.warn("Turnstile rejected token:", codes.join(","));
    const configProblem = codes.some(c => ["missing-input-secret", "invalid-input-secret", "internal-error"].includes(c));
    return configProblem ? "unavailable" : "failed";
  }
  if (env.TURNSTILE_SECRET_KEY === TURNSTILE_TEST_SECRET) return "ok";

  const hosts = allowedOrigins.map(o => new URL(o).hostname);
  if (outcome.action !== TURNSTILE_ACTION || !hosts.includes(outcome.hostname)) {
    console.warn("Turnstile token for wrong action/hostname:", outcome.action, outcome.hostname);
    return "failed";
  }
  return "ok";
}

async function sendEmail(env, d, request) {
  const submitted = new Date().toISOString().replace("T", " ").slice(0, 16) + " UTC";
  const country = (request.cf && request.cf.country) || "unknown";
  const text = [
    "New enquiry from the IITM/IITMRP Labs & Equipment website.",
    "",
    `Name:    ${d.name}`,
    `Title:   ${d.title}`,
    `Email:   ${d.email}`,
    `Company: ${d.company}`,
    "",
    "Message:",
    d.message || "(no message)",
    "",
    "--",
    "Reply to this email to respond directly to the sender.",
    `Submitted ${submitted}, visitor country: ${country}`,
  ].join("\n");

  const email = {
    from: { name: `${headerName(d.name) || "Visitor"} via IITM/IITMRP Labs & Equipment Website`, email: env.GMAIL_USER },
    to: env.CONTACT_TO_EMAIL,
    reply: { name: headerName(d.name), email: d.email },
    subject: `[IITM/IITMRP Labs & Equipment] Enquiry from ${d.name} (${d.company})`.slice(0, 200),
    text,
  };

  if (env.CONTACT_DRY_RUN === "1") {
    console.log("CONTACT_DRY_RUN — email not sent:\n" + JSON.stringify({ ...email, to: "(CONTACT_TO_EMAIL)" }, null, 2));
    return;
  }

  await WorkerMailer.send({
    host: "smtp.gmail.com",
    port: 465,
    secure: true,
    credentials: {
      username: env.GMAIL_USER,
      password: env.GMAIL_APP_PASSWORD.replace(/\s+/g, ""), // Google displays it in 4-letter groups
    },
    authType: "plain",
    socketTimeoutMs: 15000,
    responseTimeoutMs: 15000,
  }, email);
}

export async function handleContact(request, env, baseOrigins) {
  const allowedOrigins = env.DEV_ALLOWED_ORIGIN ? [...baseOrigins, env.DEV_ALLOWED_ORIGIN] : baseOrigins;
  const origin = request.headers.get("Origin") || "";

  // Exact-origin gate. Non-browser clients can forge Origin, which is why
  // Turnstile, rate limits and validation below are the real protection.
  if (!allowedOrigins.includes(origin)) {
    return json(403, { ok: false, error: "Forbidden origin." }, { "Vary": "Origin" });
  }
  const cors = corsHeaders(origin);

  if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: cors });
  if (request.method !== "POST") {
    return json(405, { ok: false, error: "Method not allowed." }, { ...cors, "Allow": "POST, OPTIONS" });
  }
  if (!(request.headers.get("Content-Type") || "").toLowerCase().startsWith("application/json")) {
    return json(415, { ok: false, error: "Unsupported content type." }, cors);
  }

  const dryRun = env.CONTACT_DRY_RUN === "1";
  if (!env.TURNSTILE_SECRET_KEY || !env.CONTACT_KV ||
      (!dryRun && (!env.GMAIL_USER || !env.GMAIL_APP_PASSWORD || !env.CONTACT_TO_EMAIL))) {
    console.error("Contact form is missing a secret or the CONTACT_KV binding.");
    return json(500, { ok: false, error: "The contact form is not available right now. Please try again later." }, cors);
  }

  try {
    const ip = request.headers.get("CF-Connecting-IP") || "unknown";

    if (env.CONTACT_RATE_LIMITER) {
      const { success } = await env.CONTACT_RATE_LIMITER.limit({ key: `contact:${ip}` });
      if (!success) {
        return json(429, { ok: false, error: "Too many attempts. Please wait a minute and try again." }, cors);
      }
    }

    const raw = await readBody(request);
    if (raw === null) return json(413, { ok: false, error: "Your message is too long." }, cors);

    let input;
    try { input = JSON.parse(raw); } catch { input = null; }
    if (!input || typeof input !== "object" || Array.isArray(input)) {
      return json(400, { ok: false, error: "Invalid request." }, cors);
    }

    // Honeypot: a hidden field people never see. Bots that fill it in get a
    // normal-looking success so they have nothing to adapt to.
    if (typeof input.website === "string" && input.website.trim()) return json(200, { ok: true }, cors);

    const checked = validate(input);
    if (checked.error) return json(400, { ok: false, error: checked.error, field: checked.field }, cors);

    const caps = await dailyCaps(env, ip);
    if (!caps.ok) {
      return json(429, { ok: false, error: "The daily submission limit has been reached. Please try again tomorrow." }, cors);
    }

    const turnstile = await verifyTurnstile(checked.data.token, ip, env, allowedOrigins);
    if (turnstile === "unavailable") {
      return json(503, { ok: false, error: "Verification is temporarily unavailable. Please try again shortly." }, cors);
    }
    if (turnstile !== "ok") {
      return json(403, { ok: false, error: "Verification failed. Please complete the check again and resubmit.", field: "turnstile" }, cors);
    }

    await caps.record();

    try {
      await sendEmail(env, checked.data, request);
    } catch (e) {
      console.error("Contact email send failed:", e && e.message);
      return json(502, { ok: false, error: "Sorry, your message could not be sent right now. Please try again later." }, cors);
    }
    return json(200, { ok: true }, cors);
  } catch (e) {
    console.error("Contact handler error:", e && e.message);
    return json(500, { ok: false, error: "Something went wrong. Please try again later." }, cors);
  }
}
