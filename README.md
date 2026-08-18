# Signatur-skifter til Outlook

Et Outlook-tilføjelsesprogram (Office Add-in), der automatisk vælger din
**interne** signatur, når *alle* modtagere er på dit eget mail-domæne, og din
**eksterne** signatur i alle andre tilfælde (inkl. Cc/Bcc).

Virker i: nyt Outlook (Windows), Outlook på nettet, klassisk Outlook (Windows)
og Outlook på Mac — kræver en Microsoft 365 / Exchange Online-postkasse.

## Sådan virker det

- `manifest.xml` registrerer et "runtime", der starter automatisk når du:
  - åbner en ny mail (`OnNewMessageCompose`)
  - ændrer modtagere, fx ved svar/videresend eller når du tilføjer/fjerner
    nogen i Til/Cc/Bcc (`OnMessageRecipientsChanged`)
- [src/runtime/autorun.js](src/runtime/autorun.js) sammenligner modtagernes
  domæner med dit eget (`Office.context.mailbox.userProfile.emailAddress`) og
  sætter signaturen med `item.body.setSignatureAsync(...)`.
- [src/taskpane/taskpane.html](src/taskpane/taskpane.html) er indstillingsruden,
  hvor du indtaster de to signaturer. De gemmes på din postkasse via
  `Office.context.roamingSettings`, så de følger med dig mellem enheder.

Reglen for "intern" ligger i `applySignature()` i
[autorun.js](src/runtime/autorun.js): en mail er intern, hvis den har mindst
én modtager, og *alle* modtagere har samme domæne som din egen adresse. En
tom modtagerliste (helt ny, blank mail) tæller som intern som udgangspunkt.

## 1. Host filerne (kræves — Outlook henter dem via HTTPS)

Outlook kan ikke bruge lokale filer direkte; de skal ligge på en HTTPS-adresse.
Nemmeste gratis løsning er GitHub Pages:

```bash
cd outlook-signature-switcher
git init
git add assets src manifest.xml
git commit -m "Signatur-skifter add-in"
git branch -M main
git remote add origin https://github.com/<dit-brugernavn>/outlook-signature-switcher.git
git push -u origin main
```

Slå derefter Pages til: repo → **Settings → Pages → Branch: main / (root)**.
Din side er herefter tilgængelig på:

```
https://<dit-brugernavn>.github.io/outlook-signature-switcher
```

Åbn `manifest.xml` og erstat **alle** forekomster af `REPLACE_WITH_YOUR_HOST`
med den adresse (uden afsluttende skråstreg), fx:

```
https://<dit-brugernavn>.github.io/outlook-signature-switcher
```

Vent et par minutter til GitHub Pages er live, og tjek at
`https://<dit-brugernavn>.github.io/outlook-signature-switcher/src/taskpane/taskpane.html`
kan åbnes i en browser.

> Alternativ: enhver anden HTTPS-hosting duer også (Azure Static Web Apps,
> firmaets egen webserver osv.) — det er kun URL'erne i `manifest.xml`, der
> skal pege det rigtige sted hen.

## 2. Sideload add-in'et i Outlook

**Nyt Outlook / Outlook på nettet:**
Indstillinger (tandhjul) → **Vis alle Outlook-indstillinger** → **Generelt →
Administrer tilføjelsesprogrammer** → **Mine tilføjelsesprogrammer** →
**Tilføj et brugerdefineret tilføjelsesprogram → Tilføj fra fil** → vælg din
opdaterede `manifest.xml`.

**Klassisk Outlook (Windows):**
**Filer → Hent tilføjelsesprogrammer** (eller Bånd → **Hent
tilføjelsesprogrammer**) → **Mine tilføjelsesprogrammer** → **Tilføj et
brugerdefineret tilføjelsesprogram → Tilføj fra fil** → vælg `manifest.xml`.

**Outlook på Mac:** samme sted som klassisk Windows, via
**Hent tilføjelsesprogrammer**.

Hvis din organisation bruger Microsoft 365, kan en administrator i stedet
udrulle den centralt til alle via **Microsoft 365 admin center →
Integrerede apps**, så du (og evt. kolleger) slipper for at sideloade manuelt.

## 3. Konfigurér dine signaturer

1. Opret en ny mail i Outlook.
2. Klik **Signatur-indstillinger** i båndet.
3. Skriv/indsæt din interne signatur i det ene felt og din eksterne i det
   andet (du kan kopiere en eksisterende signatur med logo og formatering
   direkte ind).
4. Klik **Gem indstillinger**.

Fra nu af skifter signaturen automatisk, når du ændrer modtagere. Knappen
**Anvend på denne mail nu** kan bruges til at teste eller til manuelt at
tvinge et skift.

**Tip:** Slå Outlooks egen automatiske signatur fra (Filer → Indstillinger →
Mail → Signaturer, eller i nyt Outlook: Indstillinger → Kompose og svar), så
den ikke konkurrerer med denne add-in.

## Begrænsninger

- Kræver Microsoft 365/Exchange Online — virker ikke med rene POP/IMAP-konti.
- `roamingSettings` har en størrelsesgrænse (få hundrede KB) — hold billeder i
  signaturen små, eller link til dem i stedet for at indsætte dem som base64.
- I klassisk Outlook (Windows) kan `OnMessageRecipientsChanged` udløses igen
  ved svar, selv uden ændringer — helt ufarligt, den sætter blot samme
  signatur igen.
- Vil du regne flere domæner som "interne" (fx søsterselskaber), tilføj en
  liste og tjek mod den i stedet for kun `getOwnDomain()` i `autorun.js`.

## Ikoner

`make_icons.py` genererer `assets/icon-*.png` med Pillow. Kør den igen, hvis
du vil ændre farve/bogstav:

```bash
python make_icons.py
```
