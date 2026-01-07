// deno-lint-ignore-file no-explicit-any
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";

function corsHeaders(origin?: string) {
  return {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": origin || "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  } as Record<string, string>;
}

function b64urlFromString(s: string) {
  const encoded = btoa(unescape(encodeURIComponent(s)));
  return encoded.replace(/=+$/g, "").replace(/\+/g, "-").replace(/\//g, "_");
}

function b64urlFromArrayBuffer(buf: ArrayBuffer) {
  const bytes = new Uint8Array(buf);
  let binary = "";
  for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
  const encoded = btoa(binary);
  return encoded.replace(/=+$/g, "").replace(/\+/g, "-").replace(/\//g, "_");
}

async function signHS256(payloadObj: any, secret: string) {
  const header = { alg: "HS256", typ: "JWT" };
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey(
    "raw",
    enc.encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const headerPart = b64urlFromString(JSON.stringify(header));
  const payloadPart = b64urlFromString(JSON.stringify(payloadObj));
  const data = `${headerPart}.${payloadPart}`;
  const sig = await crypto.subtle.sign("HMAC", key, enc.encode(data));
  const sigPart = b64urlFromArrayBuffer(sig);
  return `${data}.${sigPart}`;
}

serve(async (req: Request) => {
  const origin = req.headers.get("origin") || undefined;
  if (req.method === "OPTIONS") {
    return new Response("", { headers: corsHeaders(origin) });
  }
  const headers = corsHeaders(origin);

  try {
    const secret = Deno.env.get("ONLYOFFICE_JWT") || Deno.env.get("DOCUMENT_SERVER_JWT") || "";
    if (!secret) {
      return new Response(
        JSON.stringify({ error: "server_misconfigured", hint: "Missing ONLYOFFICE_JWT or DOCUMENT_SERVER_JWT" }),
        { status: 500, headers },
      );
    }
    const body = await req.json().catch(() => ({}));
    const cfg = body?.config;
    if (!cfg || typeof cfg !== "object") {
      return new Response(JSON.stringify({ error: "bad_request", hint: "config missing" }), { status: 400, headers });
    }
    // Sign the full config object as payload
    const token = await signHS256(cfg, secret);
    return new Response(JSON.stringify({ token }), { status: 200, headers });
  } catch (e) {
    return new Response(JSON.stringify({ error: "server_error", details: String(e?.message || e) }), { status: 500, headers });
  }
});
