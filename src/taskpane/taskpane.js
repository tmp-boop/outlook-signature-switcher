let rawTemplates = { internal: "", external: "" };

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

  loadTemplatesAndPreview();
});

async function loadTemplatesAndPreview() {
  const [internal, external] = await Promise.all([getTemplate("internal"), getTemplate("external")]);
  rawTemplates = { internal, external };
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
