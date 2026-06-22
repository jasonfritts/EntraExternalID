const { DOMParser } = require("@xmldom/xmldom");
const xpath = require("xpath");
const zlib = require("zlib");
const crypto = require("crypto");

function sanitizeDomain(input) {
  return String(input || "")
    .trim()
    .replace(/^https?:\/\//i, "")
    .replace(/\/$/, "")
    .toLowerCase();
}

function normalizeExternalIdDomain(input) {
  const domain = sanitizeDomain(input);
  if (!domain) {
    return "";
  }

  if (domain.includes(".")) {
    return domain;
  }

  return `${domain}.ciamlogin.com`;
}

function deriveTenantDomain(domain) {
  const host = sanitizeDomain(domain).split("/")[0].split(":")[0];
  const firstLabel = host.split(".")[0];

  if (!firstLabel || firstLabel.includes(" ")) {
    return "";
  }

  return `${firstLabel}.onmicrosoft.com`;
}

function buildMetadataUrl(domain, tenantDomain, clientId) {
  const cleanDomain = sanitizeDomain(domain);
  const cleanTenant = String(tenantDomain || "").trim().toLowerCase();
  const cleanClientId = String(clientId || "").trim();

  return `https://${cleanDomain}/${cleanTenant}/federationmetadata/2007-06/federationmetadata.xml?appid=${encodeURIComponent(cleanClientId)}`;
}

function parseMetadata(metadataXml) {
  const doc = new DOMParser().parseFromString(metadataXml, "text/xml");

  const redirectNode = xpath.select1(
    "//*[local-name()='SingleSignOnService' and @Binding='urn:oasis:names:tc:SAML:2.0:bindings:HTTP-Redirect']",
    doc,
  );

  const postNode = xpath.select1(
    "//*[local-name()='SingleSignOnService' and @Binding='urn:oasis:names:tc:SAML:2.0:bindings:HTTP-POST']",
    doc,
  );

  const ssoUrl =
    (redirectNode && redirectNode.getAttribute("Location")) ||
    (postNode && postNode.getAttribute("Location"));

  if (!ssoUrl) {
    throw new Error("Could not find SSO endpoint in federation metadata.");
  }

  return { ssoUrl };
}

function generateAuthnRequestXml({ id, issueInstant, destination, acsUrl, issuer }) {
  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<samlp:AuthnRequest',
    ' xmlns:samlp="urn:oasis:names:tc:SAML:2.0:protocol"',
    ' xmlns:saml="urn:oasis:names:tc:SAML:2.0:assertion"',
    ` ID="${id}"`,
    ' Version="2.0"',
    ` IssueInstant="${issueInstant}"`,
    ' ForceAuthn="false"',
    ' IsPassive="false"',
    ` Destination="${destination}"`,
    ` AssertionConsumerServiceURL="${acsUrl}">`,
    `<saml:Issuer>${issuer}</saml:Issuer>`,
    '<samlp:NameIDPolicy AllowCreate="true" Format="urn:oasis:names:tc:SAML:1.1:nameid-format:unspecified" />',
    '</samlp:AuthnRequest>',
  ].join("");
}

function encodeSamlRequestForRedirect(xml) {
  const deflated = zlib.deflateRawSync(Buffer.from(xml, "utf-8"));
  return deflated.toString("base64");
}

function decodeBase64Xml(samlResponse) {
  return Buffer.from(String(samlResponse || ""), "base64").toString("utf-8");
}

function textOrEmpty(node) {
  return node ? node.textContent || "" : "";
}

function parseAssertion(xml) {
  const doc = new DOMParser().parseFromString(xml, "text/xml");

  const inResponseTo = textOrEmpty(
    xpath.select1("//*[local-name()='Response']/@InResponseTo", doc),
  );

  const responseIssuer = textOrEmpty(
    xpath.select1("//*[local-name()='Response']/*[local-name()='Issuer']/text()", doc),
  );

  const assertionIssuer = textOrEmpty(
    xpath.select1("//*[local-name()='Assertion']/*[local-name()='Issuer']/text()", doc),
  );

  const nameId = textOrEmpty(
    xpath.select1("//*[local-name()='Subject']/*[local-name()='NameID']/text()", doc),
  );

  const audience = textOrEmpty(
    xpath.select1("//*[local-name()='AudienceRestriction']/*[local-name()='Audience']/text()", doc),
  );

  const notBefore = textOrEmpty(
    xpath.select1("//*[local-name()='Conditions']/@NotBefore", doc),
  );

  const notOnOrAfter = textOrEmpty(
    xpath.select1("//*[local-name()='Conditions']/@NotOnOrAfter", doc),
  );

  const authnInstant = textOrEmpty(
    xpath.select1("//*[local-name()='AuthnStatement']/@AuthnInstant", doc),
  );

  const attrNodes = xpath.select("//*[local-name()='Attribute']", doc);
  const attributes = attrNodes.map((attributeNode) => {
    const name = attributeNode.getAttribute("Name") || "(unnamed)";
    const valueNodes = xpath.select("./*[local-name()='AttributeValue']", attributeNode);
    const values = valueNodes.map((v) => v.textContent || "");
    return { name, values };
  });

  return {
    inResponseTo,
    responseIssuer,
    assertionIssuer,
    nameId,
    audience,
    notBefore,
    notOnOrAfter,
    authnInstant,
    attributes,
  };
}

function newRequestId() {
  return `_${crypto.randomUUID().replace(/-/g, "")}`;
}

function generateServiceProviderMetadataXml({ entityId, acsUrl }) {
  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<EntityDescriptor xmlns="urn:oasis:names:tc:SAML:2.0:metadata"',
    ` entityID="${entityId}">`,
    '<SPSSODescriptor AuthnRequestsSigned="false" WantAssertionsSigned="false"',
    ' protocolSupportEnumeration="urn:oasis:names:tc:SAML:2.0:protocol">',
    '<NameIDFormat>urn:oasis:names:tc:SAML:1.1:nameid-format:unspecified</NameIDFormat>',
    `<AssertionConsumerService Binding="urn:oasis:names:tc:SAML:2.0:bindings:HTTP-POST" Location="${acsUrl}" index="0" isDefault="true"/>`,
    '</SPSSODescriptor>',
    '</EntityDescriptor>',
  ].join("");
}

module.exports = {
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
};
