// Køres automatisk af Outlook (event-based activation).
// Sætter intern eller ekstern signatur ud fra modtagernes mail-domæne.

Office.onReady();

Office.actions.associate("onMessageComposeHandler", (event) => runAndComplete(event));
Office.actions.associate("onRecipientsChangedHandler", (event) => runAndComplete(event));

function runAndComplete(event) {
  applySignature()
    .catch((err) => console.error("Signatur-skifter fejlede:", err))
    .finally(() => event.completed());
}

async function applySignature() {
  const settings = await getSettings();
  if (!settings.internalSignature && !settings.externalSignature) {
    // Intet konfigureret endnu i indstillingerne – rør ikke ved mailen.
    return;
  }

  const ownDomain = getOwnDomain();
  const recipients = await getAllRecipients();
  const isInternal =
    recipients.length === 0 || recipients.every((address) => domainOf(address) === ownDomain);

  const html = isInternal ? settings.internalSignature : settings.externalSignature;
  if (!html) return;

  await setSignature(html);
}

function getOwnDomain() {
  const email = Office.context.mailbox.userProfile.emailAddress || "";
  return domainOf(email);
}

function domainOf(email) {
  const at = email.lastIndexOf("@");
  return at === -1 ? "" : email.slice(at + 1).toLowerCase();
}

function getAllRecipients() {
  const item = Office.context.mailbox.item;
  const fields = [item.to, item.cc, item.bcc].filter(Boolean);

  return Promise.all(fields.map(getRecipientsAsync)).then((groups) =>
    groups.flat().map((r) => r.emailAddress || "")
  );
}

function getRecipientsAsync(field) {
  return new Promise((resolve) => {
    field.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(asyncResult.value || []);
      } else {
        resolve([]);
      }
    });
  });
}

function getSettings() {
  return new Promise((resolve) => {
    Office.context.roamingSettings.refreshAsync(() => {
      resolve({
        internalSignature: Office.context.roamingSettings.get("internalSignature") || "",
        externalSignature: Office.context.roamingSettings.get("externalSignature") || "",
      });
    });
  });
}

function setSignature(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSignatureAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(asyncResult.error);
        } else {
          resolve();
        }
      }
    );
  });
}
