const express = require("express");
const path = require("path");
const helmet = require("helmet");
const session = require("express-session");
const axios = require("axios");
const { v4: uuidv4 } = require("uuid");
const {
  sanitizeDomain,
  normalizeExternalIdDomain,
  deriveTenantDomain,
  buildMetadataUrl,
  parseMetadata,
  generateAuthnRequestXml,
  generateServiceProviderMetadataXml,
  encodeSamlRequestForRedirect,
  decodeBase64Xml,
  parseAssertion,
  newRequestId,
} = require("./saml");

const app = express();
const port = process.env.PORT || 3000;

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

app.use(helmet({ contentSecurityPolicy: false }));
app.use(express.urlencoded({ extended: true, limit: "2mb" }));
app.use(express.static(path.join(__dirname, "public")));
app.use(
  session({
    secret: process.env.SESSION_SECRET || "dev-session-secret-change-me",
    resave: false,
    saveUninitialized: true,
    cookie: { httpOnly: true, sameSite: "lax" },
  }),
);

function getBaseUrl(req) {
  if (process.env.BASE_URL) {
    return process.env.BASE_URL.replace(/\/$/, "");
  }

  return `${req.protocol}://${req.get("host")}`;
}

function renderHome(req, res, overrides = {}) {
  const flash = req.session.flash || {};
  req.session.flash = null;

  res.render("index", {
    error: overrides.error || flash.error || "",
    message: overrides.message || flash.message || "",
    form: overrides.form || flash.form || {},
    result: overrides.result || null,
  });
}

app.get("/", (req, res) => {
  renderHome(req, res);
});

app.get("/README", (_req, res) => {
  res.sendFile(path.join(__dirname, "..", "README.html"));
});

app.get("/README.md", (_req, res) => {
  res.redirect(302, "/README");
});

app.get("/metadata-preview", async (req, res) => {
  const typedDomain = sanitizeDomain(req.query.domain);
  const domain = normalizeExternalIdDomain(typedDomain);
  const clientId = String(req.query.clientId || "").trim();
  const tenantDomain = String(req.query.tenantDomain || "").trim() || deriveTenantDomain(domain);

  if (!domain || !clientId || !tenantDomain) {
    return res.status(400).json({
      error: "domain, clientId, and tenantDomain (or derivable domain) are required",
    });
  }

  const metadataUrl = buildMetadataUrl(domain, tenantDomain, clientId);
  const baseUrl = getBaseUrl(req);
  const acsUrl = `${baseUrl}/acs`;
  const entityId = `urn:saml-sp-test:${clientId}`;
  const spMetadataXml = generateServiceProviderMetadataXml({ entityId, acsUrl });

  try {
    const metadataResponse = await axios.get(metadataUrl, { timeout: 15000 });
    return res.json({
      domain,
      tenantDomain,
      metadataUrl,
      idpMetadataXml: metadataResponse.data,
      spMetadataXml,
    });
  } catch (error) {
    return res.status(502).json({
      domain,
      tenantDomain,
      metadataUrl,
      spMetadataXml,
      error: `Failed to load IdP metadata: ${error.message}`,
    });
  }
});

app.post("/start", async (req, res) => {
  const typedDomain = sanitizeDomain(req.body.domain);
  const domain = normalizeExternalIdDomain(typedDomain);
  const clientId = String(req.body.clientId || "").trim();
  const tenantDomain = String(req.body.tenantDomain || "").trim() || deriveTenantDomain(domain);

  const form = { domain, clientId, tenantDomain };

  if (!domain || !clientId || !tenantDomain) {
    return renderHome(req, res, {
      error: "Domain, client_id, and tenant domain are required (tenant domain can be auto-derived from domain).",
      form,
    });
  }

  const metadataUrl = buildMetadataUrl(domain, tenantDomain, clientId);

  let metadataXml;
  let ssoUrl;

  try {
    const metadataResponse = await axios.get(metadataUrl, { timeout: 15000 });
    metadataXml = metadataResponse.data;
    ({ ssoUrl } = parseMetadata(metadataXml));
  } catch (error) {
    return renderHome(req, res, {
      error: `Failed to load or parse federation metadata from ${metadataUrl}: ${error.message}`,
      form,
    });
  }

  const requestId = newRequestId();
  const issueInstant = new Date().toISOString();
  const acsUrl = `${getBaseUrl(req)}/acs`;
  const issuer = `urn:saml-sp-test:${clientId}`;

  const authnRequestXml = generateAuthnRequestXml({
    id: requestId,
    issueInstant,
    destination: ssoUrl,
    acsUrl,
    issuer,
  });

  const samlRequest = encodeSamlRequestForRedirect(authnRequestXml);
  const relayState = uuidv4();

  req.session.pendingSaml = {
    requestId,
    relayState,
    domain,
    clientId,
    tenantDomain,
    metadataUrl,
    ssoUrl,
    issuer,
    acsUrl,
    authnRequestXml,
  };

  const redirectUrl = new URL(ssoUrl);
  redirectUrl.searchParams.set("SAMLRequest", samlRequest);
  redirectUrl.searchParams.set("RelayState", relayState);

  return res.redirect(redirectUrl.toString());
});

app.post("/acs", (req, res) => {
  const pendingSaml = req.session.pendingSaml || null;
  const samlResponse = req.body.SAMLResponse;
  const relayState = req.body.RelayState;

  if (!samlResponse) {
    return renderHome(req, res, {
      error: "No SAMLResponse was posted to ACS.",
      form: pendingSaml || {},
    });
  }

  let samlXml;
  let parsed;

  try {
    samlXml = decodeBase64Xml(samlResponse);
    parsed = parseAssertion(samlXml);
  } catch (error) {
    return renderHome(req, res, {
      error: `Failed to decode/parse SAMLResponse: ${error.message}`,
      form: pendingSaml || {},
    });
  }

  const requestMatched =
    pendingSaml &&
    parsed.inResponseTo &&
    parsed.inResponseTo === pendingSaml.requestId;

  const relayMatched = pendingSaml && relayState && relayState === pendingSaml.relayState;

  return renderHome(req, res, {
    form: pendingSaml || {},
    result: {
      postedRelayState: relayState || "",
      postedInResponseTo: parsed.inResponseTo || "",
      requestMatched: Boolean(requestMatched),
      relayMatched: Boolean(relayMatched),
      parsed,
      samlXml,
      pendingSaml,
    },
    message: "SAML response received. Signature validation is not performed in this inspection tool.",
  });
});

app.get("/healthz", (_req, res) => {
  res.status(200).json({ ok: true });
});

app.listen(port, () => {
  // eslint-disable-next-line no-console
  console.log(`SAML SP tester listening on http://localhost:${port}`);
});
