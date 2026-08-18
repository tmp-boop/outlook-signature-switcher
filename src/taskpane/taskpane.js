let rawTemplates = { internal: "", external: "" };

// Skal matche hosting-adressen i manifest.xml / HOST_BASE i autorun.js.
const GITHUB_OWNER = "tmp-boop";
const GITHUB_REPO = "outlook-signature-switcher";
const GITHUB_BRANCH = "main";

Office.onReady(() => {
  document.getElementById("ownDomain").textContent = getOwnDomain() || "(ukendt)";
  document.getElementById("userEmail").value = Office.context.mailbox.userProfile.emailAddress || "";

  document.getElementById("userName").value =
    Office.context.roamingSettings.get("userName") || Office.context.mailbox.userProfile.displayName || "";
  document.getElementById("userTitle").value = Office.context.roamingSettings.get("userTitle") || "";
  document.getElementById("userPhone").value = Office.context.roamingSettings.get("userPhone") || "";

  document.getElementById("saveBtn").addEventListener("click", onSave);
  document.getElementById("applyNowBtn").addEventListener("click", onApplyNow);
  ["userName", "userTitle", "userPhone"].forEach((id) =>
    document.getElementById(id).addEventListener("input", renderPreviews)
  );

  document.getElementById("saveInternalTemplateBtn").addEventListener("click", () => onSaveTemplate("internal"));
  document.getElementById("saveExternalTemplateBtn").addEventListener("click", () => onSaveTemplate("external"));

  loadTemplatesAndPreview();
});

async function loadTemplatesAndPreview() {
  const [internal, external] = await Promise.all([getTemplate("internal"), getTemplate("external")]);
  rawTemplates = { internal, external };
  document.getElementById("internalTemplateEditor").value = internal;
  document.getElementById("externalTemplateEditor").value = external;
  renderPreviews();
}

function renderPreviews() {
  const info = currentPersonalInfo();
  document.getElementById("internalPreview").innerHTML =
    fillTemplate(rawTemplates.internal, info) || "<em>Kunne ikke hente skabelon.</em>";
  document.getElementById("externalPreview").innerHTML =
    fillTemplate(rawTemplates.external, info) || "<em>Kunne ikke hente skabelon.</em>";
}

function currentPersonalInfo() {
  return {
    navn: document.getElementById("userName").value,
    titel: document.getElementById("userTitle").value,
    telefon: document.getElementById("userPhone").value,
    email: Office.context.mailbox.userProfile.emailAddress || "",
  };
}

function onSave() {
  Office.context.roamingSettings.set("userName", document.getElementById("userName").value);
  Office.context.roamingSettings.set("userTitle", document.getElementById("userTitle").value);
  Office.context.roamingSettings.set("userPhone", document.getElementById("userPhone").value);
  Office.context.roamingSettings.saveAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      showStatus("Gemt. Signaturen skifter automatisk fra næste nye mail.", "ok");
    } else {
      showStatus("Kunne ikke gemme: " + asyncResult.error.message, "error");
    }
  });
}

function onApplyNow() {
  if (!Office.context.mailbox.item || !Office.context.mailbox.item.body) {
    showStatus("Åbn en ny eller eksisterende mail for at anvende signaturen.", "error");
    return;
  }
  applySignature()
    .then(() => showStatus("Signatur anvendt på den åbne mail.", "ok"))
    .catch((err) => showStatus("Kunne ikke anvende signatur: " + err.message, "error"));
}

function showStatus(text, kind) {
  const el = document.getElementById("status");
  el.textContent = text;
  el.className = kind;
}

async function onSaveTemplate(name) {
  const token = document.getElementById("githubToken").value.trim();
  const editorId = name === "internal" ? "internalTemplateEditor" : "externalTemplateEditor";
  const newContent = document.getElementById(editorId).value;

  if (!token) {
    showTemplateStatus("Indsæt et GitHub-token for at kunne gemme.", "error");
    return;
  }

  const apiUrl = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/templates/${name}.html`;
  const headers = {
    Authorization: `token ${token}`,
    Accept: "application/vnd.github+json",
  };

  showTemplateStatus("Gemmer …", "ok");
  try {
    const current = await fetch(`${apiUrl}?ref=${GITHUB_BRANCH}`, { headers });
    if (!current.ok) throw new Error(`Kunne ikke læse nuværende fil (HTTP ${current.status})`);
    const { sha } = await current.json();

    const put = await fetch(apiUrl, {
      method: "PUT",
      headers: { ...headers, "Content-Type": "application/json" },
      body: JSON.stringify({
        message: `Opdater ${name}-skabelon via taskpane`,
        content: toBase64Utf8(newContent),
        sha,
        branch: GITHUB_BRANCH,
      }),
    });
    if (!put.ok) {
      const body = await put.json().catch(() => ({}));
      throw new Error(body.message || `HTTP ${put.status}`);
    }

    rawTemplates[name] = newContent;
    renderPreviews();

    // Opdater også den lokale reserve-kopi på denne enhed, så vi ikke selv
    // ender med at bruge en forældet cache, hvis det efterfølgende live-hentet
    // (fx i klassisk Outlook, hvor det er kendt at kunne fejle - se README).
    Office.context.roamingSettings.set("templateCache_" + name, newContent);
    Office.context.roamingSettings.saveAsync(() => {});

    showTemplateStatus(
      "Gemt. Kolleger får den nye skabelon, næste gang de åbner Signatur-indstillinger.",
      "ok"
    );
  } catch (err) {
    showTemplateStatus("Kunne ikke gemme skabelon: " + err.message, "error");
  } finally {
    document.getElementById("githubToken").value = "";
  }
}

function toBase64Utf8(str) {
  const bytes = new TextEncoder().encode(str);
  let binary = "";
  bytes.forEach((b) => (binary += String.fromCharCode(b)));
  return btoa(binary);
}

function showTemplateStatus(text, kind) {
  const el = document.getElementById("templateStatus");
  el.textContent = text;
  el.className = kind;
}
