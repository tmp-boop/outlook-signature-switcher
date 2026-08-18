// Køres automatisk af Outlook (event-based activation).
// Henter en fælles, centralt hostet skabelon og udfylder den med brugerens
// egne oplysninger, og sætter intern eller ekstern signatur ud fra
// modtagernes mail-domæne.

// Ret denne hvis I flytter hosting et andet sted hen.
const HOST_BASE = "https://tmp-boop.github.io/outlook-signature-switcher";

Office.onReady();

Office.actions.associate("onMessageComposeHandler", (event) => runAndComplete(event));
Office.actions.associate("onRecipientsChangedHandler", (event) => runAndComplete(event));

function runAndComplete(event) {
  applySignature()
    .catch((err) => console.error("Signatur-skifter fejlede:", err))
    .finally(() => event.completed());
}

async function applySignature() {
  const ownDomain = getOwnDomain();
  const recipients = await getAllRecipients();
  const isInternal =
    recipients.length === 0 || recipients.every((address) => domainOf(address) === ownDomain);

  const templateName = isInternal ? "internal" : "external";
  const [template, personalInfo] = await Promise.all([
    getTemplate(templateName),
    getPersonalInfo(),
  ]);
  if (!template) return; // hverken netværk eller cache virkede - rør ikke ved mailen

  await setSignature(fillTemplate(template, personalInfo));
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

// Brugerens egne felter: navn kommer gratis fra Outlook, telefon/titel
// indtastes én gang i taskpane'et og gemmes på brugerens egen postkasse.
function getPersonalInfo() {
  const savedName = Office.context.roamingSettings.get("userName");
  return {
    navn: savedName || Office.context.mailbox.userProfile.displayName || "",
    titel: Office.context.roamingSettings.get("userTitle") || "",
    telefon: Office.context.roamingSettings.get("userPhone") || "",
    email: Office.context.mailbox.userProfile.emailAddress || "",
  };
}

// Henter den fælles skabelon fra hosting. Gemmer en kopi i roamingSettings,
// så signaturen stadig kan sættes, selv hvis hentningen fejler (fx offline).
async function getTemplate(name) {
  const cacheKey = "templateCache_" + name;
  try {
    const response = await fetch(`${HOST_BASE}/templates/${name}.html?t=${Date.now()}`);
    if (!response.ok) throw new Error("HTTP " + response.status);
    const html = await response.text();
    Office.context.roamingSettings.set(cacheKey, html);
    Office.context.roamingSettings.saveAsync(() => {});
    return html;
  } catch (err) {
    console.error("Kunne ikke hente skabelon, bruger cache:", err);
    return Office.context.roamingSettings.get(cacheKey) || "";
  }
}

function fillTemplate(html, values) {
  return html.replace(/\{\{\s*(\w+)\s*\}\}/g, (match, key) =>
    values[key] !== undefined ? escapeHtml(values[key]) : match
  );
}

function escapeHtml(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
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
