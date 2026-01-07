// deno-lint-ignore-file no-explicit-any
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

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
    const SUPABASE_URL = Deno.env.get("SUPABASE_URL");
    const SERVICE_ROLE = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
    const DS_URL = Deno.env.get("DOCUMENT_SERVER_URL") || Deno.env.get("ONLYOFFICE_SERVER_URL");
    const OO_JWT = Deno.env.get("ONLYOFFICE_JWT") || Deno.env.get("DOCUMENT_SERVER_JWT") || "";
    if (!SUPABASE_URL || !SERVICE_ROLE || !DS_URL) {
      return new Response(JSON.stringify({ error: "server_misconfigured" }), { status: 500, headers });
    }
    const supabase = createClient(SUPABASE_URL, SERVICE_ROLE);

    const body = await req.json().catch(() => ({}));
    const bucket = String(body?.bucket || "docs");
    const key = String(body?.key || "");
    const title = String(body?.title || "document.docx");
    if (!key) return new Response(JSON.stringify({ error: "missing_key" }), { status: 400, headers });

    // Create a short-lived signed URL for DS to fetch (prefer signed to avoid private bucket issues)
    const { data: signed, error: signErr } = await supabase.storage.from(bucket).createSignedUrl(key, 300);
    if (signErr || !signed?.signedUrl) return new Response(JSON.stringify({ error: "signed_url_failed", details: signErr }), { status: 500, headers });
    const srcUrl = signed.signedUrl.startsWith("http") ? signed.signedUrl : `${SUPABASE_URL}${signed.signedUrl}`;

    // Derive a stable DS key from storage key (max ~128 chars)
    const dsKeyBase = `sb-${bucket}-${key}`.slice(-120);

    // Build ConvertService payload (async with polling)
    const convertPayload: Record<string, any> = {
      async: true,
      url: srcUrl,
      filetype: "docx",
      outputtype: "pdf",
      title: title.replace(/\.docx?$/i, ".pdf"),
      key: dsKeyBase,
    };

    // Sign payload if JWT configured
    let token: string | undefined;
    if (OO_JWT) {
      token = await signHS256(convertPayload, OO_JWT);
    }

    const endpoint = `${DS_URL.replace(/\/$/, "")}/ConvertService.ashx`;
    const initialBody = { ...convertPayload } as Record<string, any>;
    let conv: any = null;
    const maxTries = 20; // ~10-15s total with delays
    let lastRaw = "";
    let lastKey: string | undefined;
    for (let attempt = 0; attempt < maxTries; attempt++) {
      // Build body per attempt: first attempt with url; subsequent attempts poll with key
      let payload: Record<string, any>;
      if (attempt > 0 && lastKey) {
        payload = { async: true, key: lastKey };
      } else {
        payload = { ...initialBody };
      }
      let signed: Record<string, any> = payload;
      if (OO_JWT) {
        const t = await signHS256(payload, OO_JWT);
        signed = { ...payload, token: t };
      }
      const resp = await fetch(endpoint, {
        method: "POST",
        headers: { "content-type": "application/json", "accept": "application/json" },
        body: JSON.stringify(signed),
      });
      if (!resp.ok) {
        let text = "";
        try { text = await resp.text(); } catch {}
        return new Response(JSON.stringify({ error: "convert_failed", status: resp.status, body: text, endpoint, srcUrl }), { status: 502, headers });
      }
      let text = "";
      try { text = await resp.text(); } catch { text = ""; }
      lastRaw = text;
      try {
        conv = JSON.parse(text);
      } catch {
        conv = null;
      }
      if (!conv && text && text.trim().startsWith('<')) {
        try {
          const m = text.match(/<Error>(-?\d+)<\/Error>/);
          if (m) conv = { error: Number(m[1]) };
        } catch {}
      }
      // Error mapping from DS (allow -7 to fall back to upload method)
      const convErrVal = (conv && typeof conv.error !== "undefined") ? Number(conv.error) : 0;
      if (convErrVal && convErrVal !== -7) {
        return new Response(JSON.stringify({ error: "convert_failed", conv, raw: lastRaw.slice(0, 400), srcUrl }), { status: 502, headers });
      }
      if (convErrVal === -7) {
        // break early to attempt upload fallback below
        break;
      }
      if (conv && conv.fileUrl) break;
      if (conv && (conv.endConvert === true)) break;
      if (conv && conv.key && typeof conv.key === "string") {
        lastKey = conv.key;
      }
      // Not ready yet, wait and poll again
      await new Promise(r => setTimeout(r, 700));
    }
    // Try alternative properties that some DS builds may return
    let fileUrl: string | undefined = conv?.fileUrl || conv?.fileurl || conv?.Url || conv?.url;
    // If DS couldn't download the source (-7), fall back to uploading the DOCX via multipart/form-data
    const wasMinus7 = (!!conv && Number(conv.error) === -7) || (lastRaw && /<Error>\s*-7\s*<\/Error>/.test(lastRaw));
    if (!fileUrl && wasMinus7) {
      // Download source from storage and upload directly to DS
      const { data: dl, error: dlErr } = await supabase.storage.from(bucket).download(key);
      if (dlErr || !dl) {
        return new Response(JSON.stringify({ error: "download_src_failed", details: dlErr, srcUrl }), { status: 500, headers });
      }
      const fd = new FormData();
      const safeTitle = title.replace(/[^a-zA-Z0-9_.\-\sçğıöşüÇĞİÖŞÜ]/g, '').replace(/\s+/g,' ').trim();
      const inName = safeTitle.replace(/\.pdf$/i, ".docx") || "document.docx";
      let docFile: File;
      try {
        docFile = new File([dl as any], inName, { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
      } catch {
        // Fallback to Blob if File constructor not available
        const b = dl as any;
        docFile = new File([b], inName);
      }
      fd.append("file", docFile, inName);
      fd.append("outputtype", "pdf");
      fd.append("filetype", "docx");
      fd.append("title", safeTitle.replace(/\.docx?$/i, ".pdf"));
      fd.append("key", dsKeyBase);
      fd.append("async", "true");
      const headers2: Record<string,string> = { "accept": "application/json" };
      if (OO_JWT) {
        const payloadJwt = { async: true, outputtype: "pdf", filetype: "docx", title: safeTitle.replace(/\.docx?$/i, ".pdf"), key: dsKeyBase };
        const t2 = await signHS256(payloadJwt, OO_JWT);
        headers2["Authorization"] = `Bearer ${t2}`;
        fd.append("token", t2);
      }
      const resp2 = await fetch(endpoint, { method: "POST", headers: headers2, body: fd });
      if (!resp2.ok) {
        let text2 = "";
        try { text2 = await resp2.text(); } catch {}
        return new Response(JSON.stringify({ error: "convert_failed", status: resp2.status, body: text2, endpoint, method: "formdata_fallback" }), { status: 502, headers });
      }
      let text2 = "";
      try { text2 = await resp2.text(); } catch { text2 = ""; }
      lastRaw = text2;
      let conv2: any = null;
      try { conv2 = JSON.parse(text2); } catch { conv2 = null; }
      if (!conv2 && text2 && text2.trim().startsWith('<')) {
        // Try to extract FileUrl from XML
        try {
          const m2 = text2.match(/<FileUrl>([^<]+)<\/FileUrl>/i) || text2.match(/<Url>([^<]+)<\/Url>/i);
          if (m2) {
            fileUrl = m2[1];
          } else {
            const em = text2.match(/<Error>(-?\d+)<\/Error>/);
            if (em) {
              return new Response(JSON.stringify({ error: "convert_failed", conv: { error: Number(em[1]) }, raw: lastRaw.slice(0, 400), method: "formdata_fallback" }), { status: 502, headers });
            }
          }
        } catch {}
      }
      if (!fileUrl && conv2) {
        if (conv2 && typeof conv2.error !== "undefined" && conv2.error) {
          return new Response(JSON.stringify({ error: "convert_failed", conv: conv2, raw: lastRaw.slice(0, 400), method: "formdata_fallback" }), { status: 502, headers });
        }
        fileUrl = conv2.fileUrl || conv2.fileurl || conv2.Url || conv2.url;
        let k2: string | undefined = conv2.key || undefined;
        // If still not ready, poll by key as before
        if (!fileUrl && !conv2.endConvert && (k2 || lastKey)) {
          k2 = k2 || lastKey;
          const tries2 = 20;
          for (let t = 0; t < tries2; t++) {
            let payload2: Record<string, any> = { async: true, key: k2 };
            let signed2: Record<string, any> = payload2;
            if (OO_JWT) {
              const tt = await signHS256(payload2, OO_JWT);
              signed2 = { ...payload2, token: tt };
            }
            const r = await fetch(endpoint, { method: "POST", headers: { "content-type": "application/json", "accept": "application/json" }, body: JSON.stringify(signed2) });
            if (!r.ok) {
              let tx = ""; try { tx = await r.text(); } catch {}
              return new Response(JSON.stringify({ error: "convert_failed", status: r.status, body: tx, endpoint, method: "formdata_poll" }), { status: 502, headers });
            }
            let tx = ""; try { tx = await r.text(); } catch { tx = ""; }
            lastRaw = tx;
            let cj: any = null; try { cj = JSON.parse(tx); } catch { cj = null; }
            if (!cj && tx && tx.trim().startsWith('<')) {
              try {
                const m3 = tx.match(/<FileUrl>([^<]+)<\/FileUrl>/i) || tx.match(/<Url>([^<]+)<\/Url>/i);
                if (m3) { fileUrl = m3[1]; break; }
              } catch {}
            }
            if (cj) {
              if (cj && typeof cj.error !== "undefined" && cj.error) {
                return new Response(JSON.stringify({ error: "convert_failed", conv: cj, raw: lastRaw.slice(0, 400), method: "formdata_poll" }), { status: 502, headers });
              }
              if (cj.fileUrl) { fileUrl = cj.fileUrl; break; }
              if (cj.endConvert === true) { fileUrl = cj.fileUrl || cj.fileurl || cj.Url || cj.url; break; }
            }
            await new Promise(r2 => setTimeout(r2, 700));
          }
        }
      }
    }
    if (!fileUrl) {
      return new Response(JSON.stringify({ error: "no_file_url", conv: conv || {}, raw: lastRaw.slice(0, 400), srcUrl }), { status: 502, headers });
    }

    // Normalize fileUrl to DS origin if needed
    try {
      const test = new URL(fileUrl);
      // ok absolute
    } catch {
      // relative -> resolve against DS_URL
      fileUrl = new URL(fileUrl, DS_URL).href;
    }
    // If absolute but different origin than DS_URL, prefer DS origin (some proxies return internal hosts)
    try {
      const a = new URL(fileUrl);
      const d = new URL(DS_URL);
      if (a.host !== d.host) {
        a.protocol = d.protocol;
        a.host = d.host;
        fileUrl = a.href;
      }
    } catch {}

    const fileResp = await fetch(fileUrl);
    if (!fileResp.ok) return new Response(JSON.stringify({ error: "download_failed", status: fileResp.status, fileUrl }), { status: 502, headers });
    const arr = new Uint8Array(await fileResp.arrayBuffer());

    const pdfKey = key.replace(/\.docx?$/i, "").replace(/\.$/, "") + `_${Date.now()}.pdf`;
    const { error: upErr } = await supabase.storage.from(bucket).upload(pdfKey, arr, {
      contentType: "application/pdf",
      upsert: true,
      cacheControl: "3600",
    });
    if (upErr) return new Response(JSON.stringify({ error: "upload_failed", details: upErr }), { status: 500, headers });

    // Return a public URL if possible, else signed URL
    const pubPdf = supabase.storage.from(bucket).getPublicUrl(pdfKey);
    let url = pubPdf?.data?.publicUrl as string | undefined;
    if (!url) {
      const { data: s2 } = await supabase.storage.from(bucket).createSignedUrl(pdfKey, 60 * 60 * 24 * 7);
      url = s2?.signedUrl ? (s2.signedUrl.startsWith("http") ? s2.signedUrl : `${SUPABASE_URL}${s2.signedUrl}`) : undefined;
    }

    return new Response(JSON.stringify({ ok: true, url, key: pdfKey }), { status: 200, headers });
  } catch (e) {
    return new Response(
      JSON.stringify({ error: "server_error", details: String(e?.message || e) }),
      { status: 500, headers },
    );
  }
});
