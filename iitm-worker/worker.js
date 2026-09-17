/**
 * IIT-M Equipment Assistant — Gemini API proxy
 * ─────────────────────────────────────────────────────────────────────────
 * Deploys as a Cloudflare Worker. Holds GEMINI_API_KEY as a server-side
 * secret (set via `wrangler secret put GEMINI_API_KEY`) — it is never sent
 * to, or visible from, the browser. The GitHub Pages site calls this
 * Worker's URL instead of calling Gemini directly.
 *
 * Also serves POST /contact — the "Contact Us" form mailer (contact.js).
 *
 * Setup: see DEPLOYMENT.md in this folder.
 */

import { handleContact } from "./contact.js";

// Restrict which sites are allowed to call this proxy. Update this to the
// exact origin(s) your GitHub Pages site is served from before deploying.
// Example: "https://your-username.github.io"
const ALLOWED_ORIGINS = [
  "https://rct-iitmrp.github.io",   // site: https://rct-iitmrp.github.io/IITM_Labs_Equipments/
  // "http://localhost:8000",   // uncomment while testing locally
];

function corsHeaders(origin) {
  const allow = ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0];
  return {
    "Access-Control-Allow-Origin": allow,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
}

export default {
  async fetch(request, env) {
    const origin = request.headers.get("Origin") || "";

    // Contact form has its own route, CORS and checks (contact.js).
    if (new URL(request.url).pathname === "/contact") {
      return handleContact(request, env, ALLOWED_ORIGINS);
    }

    // Preflight
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders(origin) });
    }

    if (request.method !== "POST") {
      return new Response("Method not allowed", {
        status: 405,
        headers: corsHeaders(origin),
      });
    }

    // Origin gate — not a strong security boundary on its own (headers can
    // be forged outside a browser), but it stops the overwhelming majority
    // of casual scraping/abuse of this public endpoint by other sites.
    if (!ALLOWED_ORIGINS.includes(origin)) {
      return new Response("Forbidden origin", {
        status: 403,
        headers: corsHeaders(origin),
      });
    }

    if (!env.GEMINI_API_KEY) {
      return new Response("Proxy is not configured (missing GEMINI_API_KEY secret)", {
        status: 500,
        headers: corsHeaders(origin),
      });
    }

    const model = env.GEMINI_MODEL || "gemini-2.5-flash";
    const upstream = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${env.GEMINI_API_KEY}`;

    let body;
    try {
      body = await request.text();
    } catch (e) {
      return new Response("Invalid request body", { status: 400, headers: corsHeaders(origin) });
    }

    let geminiRes;
    try {
      geminiRes = await fetch(upstream, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body,
      });
    } catch (e) {
      return new Response(JSON.stringify({ error: "Upstream Gemini request failed" }), {
        status: 502,
        headers: { "Content-Type": "application/json", ...corsHeaders(origin) },
      });
    }

    const text = await geminiRes.text();
    return new Response(text, {
      status: geminiRes.status,
      headers: { "Content-Type": "application/json", ...corsHeaders(origin) },
    });
  },
};
