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
- [templates/internal.html](templates/internal.html) og
  [templates/external.html](templates/external.html) er de **fælles
  skabeloner** — samme layout/logo/disclaimer for alle i virksomheden. De
  hostes sammen med resten af filerne og hentes live af alle brugeres
  tilføjelsesprogram, hver gang signaturen skal sættes.
- Skabelonen redigeres **direkte fra Outlook**, ikke ved at redigere filer
  og køre git-kommandoer: åbn indstillingsruden → **Rediger fælles skabelon
  (kun administrator)** → ret den rå HTML → **Gem**. Det gemmer direkte til
  GitHub via deres API (se afsnittet nedenfor) — alle kolleger får ændringen
  automatisk, uden gensideload.
- Skabelonerne bruger pladsholdere: `{{navn}}`, `{{titel}}`, `{{telefon}}`,
  `{{email}}`.
- [src/runtime/autorun.js](src/runtime/autorun.js) sammenligner modtagernes
  domæner med dit eget (`Office.context.mailbox.userProfile.emailAddress`),
  henter den rigtige skabelon, udfylder pladsholderne med brugerens egne
  oplysninger, og sætter signaturen med `item.body.setSignatureAsync(...)`.
- [src/taskpane/taskpane.html](src/taskpane/taskpane.html) er indstillingsruden,
  hvor hver bruger kun udfylder **titel og telefon** (navn hentes automatisk
  fra Outlook). Det gemmes på den enkeltes postkasse via
  `Office.context.roamingSettings` — ingen andre kan se eller ændre det.

Reglen for "intern" ligger i `applySignature()` i
[autorun.js](src/runtime/autorun.js): en mail er intern, hvis den har mindst
én modtager, og *alle* modtagere har samme domæne som din egen adresse. En
tom modtagerliste (helt ny, blank mail) tæller som intern som udgangspunkt.

### Rediger den fælles skabelon fra Outlook (admin)

1. Opret et **fine-grained personal access token** på GitHub: gå til
   [github.com/settings/personal-access-tokens](https://github.com/settings/personal-access-tokens/new) →
   **Only select repositories** → vælg `outlook-signature-switcher` →
   under **Repository permissions** sæt **Contents: Read and write** →
   **Generate token**.
2. Kopiér tokenet (det vises kun én gang).
3. I Outlook: åbn en ny mail → **Signatur-indstillinger** → udvid
   **Rediger fælles skabelon (kun administrator)**.
4. Indsæt tokenet i feltet, og rediger felterne for intern/ekstern skabelon
   direkte, som de kommer til at se ud (fed/kursiv/understreget virker som i
   en normal tekstboks) — ikke som HTML-kode. Pladsholdere som
   `{{navn}}`/`{{titel}}`/`{{telefon}}`/`{{email}}` skrives bare som
   almindelig tekst, der hvor de skal stå.
5. Klik **Gem**. Tokenet bliver **aldrig gemt** nogen steder (ikke i
   `roamingSettings`, ikke i koden) — det bruges kun til det ene gem-kald og
   ryddes fra feltet bagefter. Du skal indsætte det igen næste gang du
   redigerer.
6. Ændringen henter vi fra `raw.githubusercontent.com` (de rå fil-data i
   repoet), ikke fra GitHub Pages — det opdaterer sig praktisk taget
   øjeblikkeligt efter en gemning, i modsætning til Pages, der bygger/udruller
   siden asynkront og kan være flere minutter bagud.

**Hvornår ser kolleger ændringen?** I nyt Outlook, Outlook på nettet og Mac
hentes skabelonen live næsten øjeblikkeligt. I **klassisk Outlook (Windows)**
er den automatiske baggrundsproces, der sætter signaturen, kendt for at have
et upålideligt netværkslag (en dokumenteret Microsoft-begrænsning, ikke noget
i denne kode) — der bruges derfor en lokal reserve-kopi, som kun opdateres,
når man selv åbner **Signatur-indstillinger**. Bed derfor kolleger om at
åbne indstillingsruden én gang efter en skabelon-ændring, hvis de bruger
klassisk Outlook.

**Om tokenet:** det er reelt sådan adgangsstyringen fungerer her — kun den,
der har (eller får) et token med skriveadgang til repoet, kan ændre
skabelonen. Del det ikke, og lav et nyt/roter det via GitHub, hvis du er i
tvivl om det er kompromitteret.

### Om navn/titel/telefon — hvorfor ikke hente det fra AD?

Navn kommer gratis fra Outlook (`userProfile.displayName`). Titel og telefon
findes derimod ikke i Office.js og kræver et Microsoft Graph-opslag mod jeres
Entra ID, hvilket kræver en admin-godkendt app-registrering **og** at felterne
faktisk står udfyldt i AD. Det er ikke sat op her — hver bruger indtaster
derfor selv titel/telefon én gang. Har I senere admin-adgang og pænt udfyldte
AD-profiler, kan det tilføjes som en udvidelse.

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

## 3. Udfyld dine oplysninger

1. Opret en ny mail i Outlook.
2. Klik **Signatur-indstillinger** i båndet.
3. Navn er allerede udfyldt. Skriv din titel (valgfrit) og telefonnummer.
4. Se forhåndsvisningen af begge signaturer nederst i ruden.
5. Klik **Gem mine oplysninger**.

Fra nu af skifter signaturen automatisk, når du ændrer modtagere. Knappen
**Anvend på denne mail nu** kan bruges til at teste eller til manuelt at
tvinge et skift.

**Tip:** Slå Outlooks egen automatiske signatur fra (Filer → Indstillinger →
Mail → Signaturer, eller i nyt Outlook: Indstillinger → Kompose og svar), så
den ikke konkurrerer med denne add-in.

## Begrænsninger

- Kræver Microsoft 365/Exchange Online — virker ikke med rene POP/IMAP-konti.
- Skabelonerne hentes over internettet hver gang signaturen sættes. Er
  brugeren offline — eller i klassisk Outlook (Windows), hvor den slags
  netværkskald i baggrunden er kendt for at fejle (se afsnittet om admin-
  redigering ovenfor) — bruges en cachet kopi fra sidste succesfulde
  hentning (gemt i `roamingSettings`) i stedet.
- Ændringer i `templates/*.html` hentes fra `raw.githubusercontent.com` og
  slår derfor typisk igennem inden for få sekunder, ikke øjeblikkeligt (den
  har en kort CDN-cache på op til 5 minutter).
- `roamingSettings` har en størrelsesgrænse (få hundrede KB) — hold billeder i
  skabelonerne små, eller link til dem i stedet for at indsætte dem som base64.
- I klassisk Outlook (Windows) kan `OnMessageRecipientsChanged` udløses igen
  ved svar, selv uden ændringer — helt ufarligt, den sætter blot samme
  signatur igen.
- Vil du regne flere domæner som "interne" (fx søsterselskaber), tilføj en
  liste og tjek mod den i stedet for kun `getOwnDomain()` i `autorun.js`.
- Titel/telefon kommer ikke fra AD/Entra ID (se afsnittet ovenfor) — hver
  bruger indtaster det selv én gang.
- Skabelon-editoren i taskpane'et kræver et GitHub-token med skriveadgang
  til repoet (se afsnittet ovenfor). Uden et token kan man stadig se/bruge
  signaturerne, men ikke gemme ændringer i skabelonen fra Outlook.

## Ikoner

`make_icons.py` genererer `assets/icon-*.png` med Pillow. Kør den igen, hvis
du vil ændre farve/bogstav:

```bash
python make_icons.py
```
