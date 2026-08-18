Office.onReady(() => {
  document.getElementById("ownDomain").textContent = getOwnDomain() || "(ukendt)";

  getSettings().then((settings) => {
    document.getElementById("internalEditor").innerHTML = settings.internalSignature || "";
    document.getElementById("externalEditor").innerHTML = settings.externalSignature || "";
  });

  document.querySelectorAll(".toolbar button").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.getElementById(btn.closest(".toolbar").dataset.target).focus();
      document.execCommand(btn.dataset.cmd, false, null);
    });
  });

  document.getElementById("saveBtn").addEventListener("click", onSave);
  document.getElementById("applyNowBtn").addEventListener("click", onApplyNow);
});

function onSave() {
  const internalSignature = document.getElementById("internalEditor").innerHTML;
  const externalSignature = document.getElementById("externalEditor").innerHTML;

  Office.context.roamingSettings.set("internalSignature", internalSignature);
  Office.context.roamingSettings.set("externalSignature", externalSignature);
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
