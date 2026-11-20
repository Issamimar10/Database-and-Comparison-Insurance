// Cloudflare Pages Function: /api
// Reverse proxy to Google Apps Script backend with HttpOnly session cookie & proxySecret injection.

function parseCookies(cookieHeader) {
  const out = {};
  if (!cookieHeader) return out;
  cookieHeader.split(";").forEach((part) => {
    const idx = part.indexOf("=");
    if (idx === -1) return;
    const key = part.slice(0, idx).trim();
    const val = part.slice(idx + 1).trim();
    if (key) out[key] = val;
  });
  return out;
}

function getCookieName(isHttps) {
  return isHttps ? "__Host-idcm_session" : "idcm_session";
}

function buildSessionCookie(token, maxAgeSec, isHttps) {
  const name = getCookieName(isHttps);
  const parts = [
    `${name}=${token}`,
    "Path=/",
    "HttpOnly",
    "SameSite=Lax",
    `Max-Age=${Math.max(0, Math.floor(maxAgeSec))}`
  ];
  if (isHttps) parts.push("Secure");
  return parts.join("; ");
}

function jsonResp(obj, status = 200, extraHeaders = {}) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      "X-Content-Type-Options": "nosniff",
      "Vary": "Origin, Cookie, Sec-Fetch-Site",
      ...extraHeaders
    }
  });
}

export async function onRequest(context) {
  const { request, env } = context;

  if (request.method !== "POST") {
    return jsonResp(
      { status: "error", message: "Method Not Allowed. Use POST." },
      405,
      { Allow: "POST" }
    );
  }

  // CSRF / Cross-Site Protection: Fetch Metadata + Same-Origin Check
  const reqUrl = new URL(request.url);
  const isHttps = reqUrl.protocol === "https:";
  const secFetchSite = request.headers.get("Sec-Fetch-Site");
  if (secFetchSite && !["same-origin", "same-site", "none"].includes(secFetchSite)) {
    return jsonResp({ status: "error", code: "FORBIDDEN", message: "Cross-site request rejected." }, 403);
  }
  const origin = request.headers.get("Origin");
  if (origin && origin !== reqUrl.origin) {
    return jsonResp({ status: "error", code: "FORBIDDEN", message: "Invalid request origin." }, 403);
  }

  const gasUrl = String(env.GAS_BACKEND_URL || "").trim();
  if (!gasUrl || !gasUrl.startsWith("https://script.google.com/")) {
    return jsonResp(
      { status: "error", message: "Backend GAS_BACKEND_URL belum dikonfigurasi di Cloudflare Pages." },
      500
    );
  }

  let payload;
  try {
    payload = await request.json();
    if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
      return jsonResp({ status: "error", message: "Payload JSON tidak valid." }, 400);
    }
  } catch (e) {
    return jsonResp({ status: "error", message: "Body bukan JSON yang valid." }, 400);
  }

  const action = String(payload.action || "").trim();
  const cookies = parseCookies(request.headers.get("Cookie") || "");
  const cookieToken = cookies[getCookieName(isHttps)] || cookies["__Host-idcm_session"] || cookies["idcm_session"] || "";

  // Fast path: jika auth_me dipanggil tanpa cookie sama sekali, langsung kembalikan UNAUTHORIZED tanpa hit GAS
  if (action === "auth_me" && !cookieToken) {
    return jsonResp({ status: "error", code: "UNAUTHORIZED", message: "Belum login." }, 200);
  }

  // Jangan pernah terima proxySecret dari klien; sisipkan dari env server
  delete payload.proxySecret;
  if (env.IDCM_PROXY_SECRET) {
    payload.proxySecret = String(env.IDCM_PROXY_SECRET);
  }

  // Sisipkan sessionToken dari HttpOnly cookie (klien tidak pernah menyimpan token di localStorage)
  const PUBLIC_NO_COOKIE_ACTIONS = {
    auth_signin: true,
    auth_signup: true,
    auth_request_otp: true,
    auth_reset_password: true
  };
  if (!PUBLIC_NO_COOKIE_ACTIONS[action]) {
    payload.sessionToken = cookieToken;
  } else {
    delete payload.sessionToken;
  }

  let gasResObj;
  try {
    const upstream = await fetch(gasUrl, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
      redirect: "follow"
    });
    const text = await upstream.text();
    gasResObj = JSON.parse(text);
  } catch (err) {
    return jsonResp(
      { status: "error", message: "Gagal menghubungi server backend Google Apps Script." },
      502
    );
  }

  const extraHeaders = {};

  // Pasang HttpOnly cookie saat Sign In / Sign Up berhasil, lalu HAPUS sessionToken dari body JSON
  if ((action === "auth_signin" || action === "auth_signup") && gasResObj && gasResObj.status === "success" && gasResObj.sessionToken) {
    const rawToken = String(gasResObj.sessionToken);
    const expiresAtMs = Number(gasResObj.expiresAt) || (Date.now() + 24 * 60 * 60 * 1000);
    const maxAgeSec = Math.max(60, Math.min(7 * 24 * 60 * 60, Math.floor((expiresAtMs - Date.now()) / 1000)));
    extraHeaders["Set-Cookie"] = buildSessionCookie(rawToken, maxAgeSec, isHttps);
    delete gasResObj.sessionToken;
  }

  // Hapus HttpOnly cookie saat Logout, UNAUTHORIZED, atau selfDemoted
  if (action === "auth_logout" || (gasResObj && (gasResObj.code === "UNAUTHORIZED" || gasResObj.selfDemoted === true))) {
    extraHeaders["Set-Cookie"] = buildSessionCookie("", 0, isHttps);
  }

  return jsonResp(gasResObj, 200, extraHeaders);
}
