const AIRTABLE_BASE_URL = "https://api.airtable.com/v0";

module.exports = async function (context, req) {
  const pat = process.env.AIRTABLE_PAT;
  if (!pat) {
    context.res = { status: 500, body: "AIRTABLE_PAT is missing" };
    return;
  }

  const path = (context.bindingData.path || "").replace(/^\/+/, "");
  if (!path) {
    context.res = { status: 400, body: "Missing Airtable path" };
    return;
  }

  const allowlist = (process.env.AIRTABLE_BASE_ALLOWLIST || "")
    .split(",")
    .map((v) => v.trim())
    .filter(Boolean);
  if (allowlist.length) {
    const baseId = path.split("/")[0];
    if (!allowlist.includes(baseId)) {
      context.res = { status: 403, body: "Base not allowed" };
      return;
    }
  }

  const qs = req.originalUrl && req.originalUrl.includes("?")
    ? req.originalUrl.substring(req.originalUrl.indexOf("?"))
    : "";
  const targetUrl = `${AIRTABLE_BASE_URL}/${path}${qs}`;

  const headers = {
    Authorization: `Bearer ${pat}`,
    "Content-Type": "application/json",
  };

  const init = {
    method: req.method,
    headers,
  };
  if (req.method !== "GET" && req.method !== "HEAD") {
    init.body = req.rawBody || (req.body ? JSON.stringify(req.body) : undefined);
  }

  try {
    const response = await fetch(targetUrl, init);
    const bodyText = await response.text();
    context.res = {
      status: response.status,
      headers: {
        "Content-Type": response.headers.get("content-type") || "application/json",
      },
      body: bodyText,
    };
  } catch (error) {
    context.res = { status: 500, body: String(error) };
  }
};
