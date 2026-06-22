(function () {
  const domainInput = document.getElementById("domainInput");
  const tenantInput = document.getElementById("tenantDomainInput");
  const clientIdInput = document.getElementById("clientIdInput");

  const metadataPanel = document.getElementById("metadataPanel");
  const metadataHint = document.getElementById("metadataHint");
  const idpMetadataPre = document.getElementById("idpMetadataPre");
  const spMetadataPre = document.getElementById("spMetadataPre");
  const samlResponsePre = document.getElementById("samlResponsePre");

  const tabButtons = Array.from(document.querySelectorAll(".tab"));
  const tabPanes = Array.from(document.querySelectorAll(".tab-pane"));

  function baseLabelFromDomain(value) {
    const cleaned = String(value || "")
      .trim()
      .replace(/^https?:\/\//i, "")
      .replace(/\/$/, "")
      .toLowerCase();

    if (!cleaned) {
      return "";
    }

    const host = cleaned.split("/")[0].split(":")[0];
    return host.split(".")[0] || "";
  }

  function normalizeDomain(value) {
    const cleaned = String(value || "")
      .trim()
      .replace(/^https?:\/\//i, "")
      .replace(/\/$/, "")
      .toLowerCase();

    if (!cleaned) {
      return "";
    }

    if (cleaned.includes(".")) {
      return cleaned;
    }

    return `${cleaned}.ciamlogin.com`;
  }

  function updateDerivedFields(options) {
    const shouldNormalizeInputValue = Boolean(options && options.normalizeDomainInput);
    const label = baseLabelFromDomain(domainInput.value);
    const normalizedDomain = normalizeDomain(domainInput.value);

    if (shouldNormalizeInputValue && normalizedDomain) {
      domainInput.value = normalizedDomain;
    }

    if (label) {
      tenantInput.value = `${label}.onmicrosoft.com`;
    } else {
      tenantInput.value = "";
    }
  }

  function setTab(targetId) {
    tabButtons.forEach((btn) => {
      const active = btn.dataset.target === targetId;
      btn.classList.toggle("active", active);
      btn.setAttribute("aria-selected", active ? "true" : "false");
    });

    tabPanes.forEach((pane) => {
      pane.classList.toggle("active", pane.id === targetId);
    });
  }

  function prettyPrintXml(xml) {
    const raw = String(xml || "").trim();
    if (!raw) {
      return "";
    }

    const withBreaks = raw
      .replace(/>\s*</g, "><")
      .replace(/(>)(<)(\/*)/g, "$1\n$2$3");

    const lines = withBreaks.split("\n");
    let indentLevel = 0;
    const formatted = [];

    lines.forEach((line) => {
      const trimmed = line.trim();
      if (!trimmed) {
        return;
      }

      const isClosingTag = /^<\//.test(trimmed);
      const isCommentOrDeclaration = /^<\?|^<!/.test(trimmed);
      const isSelfClosing = /\/>$/.test(trimmed);
      const hasOpenAndCloseSameLine = /^<[^!?/][^>]*>.*<\/[^>]+>$/.test(trimmed);

      if (isClosingTag && indentLevel > 0) {
        indentLevel -= 1;
      }

      const indent = "  ".repeat(indentLevel);
      formatted.push(`${indent}${trimmed}`);

      if (!isClosingTag && !isSelfClosing && !isCommentOrDeclaration && !hasOpenAndCloseSameLine) {
        indentLevel += 1;
      }
    });

    return formatted.join("\n");
  }

  let currentFetchToken = 0;
  let debounceHandle = null;

  async function loadMetadataPreview() {
    const domain = normalizeDomain(domainInput.value);
    const tenantDomain = String(tenantInput.value || "").trim();
    const clientId = String(clientIdInput.value || "").trim();

    if (!domain || !tenantDomain || !clientId) {
      metadataPanel.classList.add("hidden");
      idpMetadataPre.textContent = "";
      spMetadataPre.textContent = "";
      metadataHint.textContent = "Type domain and client ID to load metadata previews.";
      return;
    }

    metadataPanel.classList.remove("hidden");
    metadataHint.textContent = "Loading metadata...";

    const fetchToken = ++currentFetchToken;

    try {
      const url = new URL("/metadata-preview", window.location.origin);
      url.searchParams.set("domain", domain);
      url.searchParams.set("tenantDomain", tenantDomain);
      url.searchParams.set("clientId", clientId);

      const response = await fetch(url.toString());
      const data = await response.json();

      if (fetchToken !== currentFetchToken) {
        return;
      }

      spMetadataPre.textContent = prettyPrintXml(data.spMetadataXml) || "(no SP metadata returned)";

      if (!response.ok) {
        idpMetadataPre.textContent = data.error || "Failed to load IdP metadata.";
        metadataHint.textContent = data.metadataUrl || "Metadata URL unavailable.";
        return;
      }

      idpMetadataPre.textContent = prettyPrintXml(data.idpMetadataXml) || "(empty IdP metadata response)";
      metadataHint.textContent = `Metadata URL: ${data.metadataUrl}`;
    } catch (error) {
      if (fetchToken !== currentFetchToken) {
        return;
      }

      idpMetadataPre.textContent = `Error loading metadata: ${error.message}`;
      metadataHint.textContent = "Could not reach metadata endpoint.";
    }
  }

  function scheduleMetadataPreview() {
    if (debounceHandle) {
      window.clearTimeout(debounceHandle);
    }

    debounceHandle = window.setTimeout(loadMetadataPreview, 300);
  }

  tabButtons.forEach((btn) => {
    btn.addEventListener("click", () => setTab(btn.dataset.target));
  });

  if (samlResponsePre && samlResponsePre.textContent) {
    samlResponsePre.textContent = prettyPrintXml(samlResponsePre.textContent);
  }

  if (!domainInput || !tenantInput || !clientIdInput) {
    return;
  }

  domainInput.addEventListener("input", () => {
    updateDerivedFields();
    scheduleMetadataPreview();
  });

  domainInput.addEventListener("blur", () => {
    updateDerivedFields({ normalizeDomainInput: true });
    scheduleMetadataPreview();
  });

  tenantInput.addEventListener("input", scheduleMetadataPreview);
  clientIdInput.addEventListener("input", scheduleMetadataPreview);

  updateDerivedFields();
  scheduleMetadataPreview();
})();
