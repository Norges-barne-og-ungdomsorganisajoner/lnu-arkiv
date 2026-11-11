# LNU Arkiv – Outlook-tillegg for e-postarkiv i SharePoint

Tillegget legger en knapp, **Arkiver i SharePoint**, på e-poster i Outlook. Brukeren velger kommunikasjonstype og prosjekt, skriver eventuelt en kommentar og trykker Arkiver. Da skjer dette:

1. E-posten hentes som en `.eml`-fil (med vedlegg) direkte fra Outlook og lagres i biblioteket *E-postarkiv*.
2. Metadata fylles inn: emne, avsender, mottakere, dato, retning (inn/ut), kommunikasjonstype, prosjekt, kommentar og meldings-ID.
3. E-posten får kategorien «Arkivert i SharePoint» i Outlook, hvis brukeren har opprettet den (se steg 7).
4. Hvis e-posten allerede er arkivert (også av en kollega), vises det med lenke før man arkiverer på nytt.

Tillegget virker i ny og klassisk Outlook for Windows, Outlook for Mac og Outlook på nett, for e-post i egen innboks og i delte postkasser. Oppgaveruten kan festes, slik at man kan arkivere flere e-poster etter hverandre.

### Tilganger

Tillegget er bygget for å ha så lite tilgang som mulig:

| Hvor | Tilgang | Hva det betyr |
|---|---|---|
| Outlook (manifestet) | `ReadWriteItem` | Kan bare lese og sette kategori på **e-posten som er åpen**. Ingen tilgang til resten av postkassen. |
| Microsoft Graph (Entra) | `Sites.Selected` (delegert) | Kan bare nå **arkivsiden**, og bare det brukeren selv har lov til der. Ingen e-posttillatelser. |

Selv om koden skulle bli endret av uvedkommende, kan den ikke lese andre e-poster enn den brukeren har åpen, og ikke nå andre SharePoint-områder enn arkivet.

## Innhold

| Fil / mappe | Hva det er |
|---|---|
| `docs/` | Selve tillegget. Det er denne mappen som publiseres på nett. |
| `docs/taskpane.html`, `.css`, `.js` | Oppgaveruten og all logikk. |
| `docs/lib/msal-browser.min.js` | Microsofts innloggingsbibliotek (MSAL 5.29.0), lagt ved så du slipper byggesteg. |
| `innstillinger.json` | Det eneste stedet du fyller inn verdier. |
| `konfigurer.py` | Lager `manifest.xml` og `docs/config.js` fra innstillingene. |
| `maler/` | Malene skriptet bruker. |

## Oppsett

Regn med 1–2 timer første gang. Gjør stegene i denne rekkefølgen.

### 1. SharePoint: arkivsiden

1. Opprett et teamområde, for eksempel **Arkiv** med adressen `/sites/arkiv`. Gi alle ansatte redigeringstilgang (medlemmer). Tillegget lagrer med brukerens egen tilgang, så den som arkiverer må kunne skrive i biblioteket.

2. Lag to lister, **Kommunikasjonstyper** og **Prosjekter**. Begge får én ekstra kolonne:
   - `Aktiv` – Ja/Nei, standardverdi Ja. Sett den til Nei for prosjekter som er avsluttet: de forsvinner fra nedtrekkslisten, men gamle e-poster beholder koblingen.

   Fyll inn noen verdier. Forslag til typer: Henvendelse, Søknad, Rapport, Avtale/kontrakt, Økonomi/faktura, Møte/innkalling, Høring, Pressehenvendelse.

3. Lag dokumentbiblioteket **E-postarkiv** med disse kolonnene:

   | Kolonnenavn | Type | Merknad |
   |---|---|---|
   | `Avsender` | Én tekstlinje | |
   | `Mottakere` | Flere tekstlinjer | Ren tekst |
   | `Mottatt` | Dato og klokkeslett | Ta med klokkeslett |
   | `Retning` | Valg | Verdier: `Innkommende`, `Utgående` |
   | `Kommunikasjonstype` | Oppslag | Fra *Kommunikasjonstyper*, kolonne Tittel. Ikke kall den «Type»: det navnet er reservert i dokumentbiblioteker. |
   | `Prosjekt` | Oppslag | Fra *Prosjekter*, kolonne Tittel |
   | `Kommentar` | Flere tekstlinjer | Ren tekst |
   | `MeldingsID` | Én tekstlinje | Brukes til å oppdage duplikater |
   | `HarVedlegg` | Ja/Nei | Om e-posten har vedlegg. Visningsnavnet kan endres til «Vedlegg» etterpå. |
   | `Vedleggsnavn` | Flere tekstlinjer | Ren tekst. Navnene på vedleggene, ett per linje. |

   Emnet lagres i bibliotekets innebygde kolonne **Tittel**. Den er skjult i standardvisningen. Legg den til i visningen, og gi den gjerne visningsnavnet «Emne» under bibliotekinnstillingene.

   I listene *Kommunikasjonstyper* og *Prosjekter* skal navnet stå i standardkolonnen **Tittel**. Det er den tillegget og oppslagskolonnene leser.

   > **Viktig:** Skriv kolonnenavnene nøyaktig som i tabellen når du oppretter dem. SharePoint låser det interne navnet ved opprettelse, og tillegget bruker det interne navnet. Ikke bruk æ, ø, å eller mellomrom i første omgang. Du kan endre visningsnavnet etterpå uten at noe slutter å virke.

4. Gå til **Bibliotekinnstillinger → Indekserte kolonner** og lag indekser på `MeldingsID`, `Prosjekt`, `Kommunikasjonstype` og `Mottatt`. `MeldingsID` er påkrevd for duplikatsjekken. De andre gjør at visninger og filtre virker når arkivet passerer 5 000 e-poster.

5. Lag gjerne visninger som *Per prosjekt* (gruppert på Prosjekt) og *Siste 30 dager*.

### 2. Publiser `docs/`-mappen

Tillegget er bare statiske filer, så hosting er gratis. To alternativer:

- **GitHub Pages (enklest):** Lag et repository, for eksempel `lnu-arkiv`, last opp hele prosjektmappen og velg *Settings → Pages → Deploy from a branch → main / docs*. Adressen blir `https://<konto>.github.io/lnu-arkiv`. Gratisversjonen krever at repositoryet er offentlig. Det inneholder ingen hemmeligheter, men viser klient-ID og SharePoint-adressen.
- **Azure Static Web Apps (privat):** Gratisnivået holder. Krever et Azure-abonnement.

Noter adressen. Det er `baseUrl` i steg 4.

### 3. Entra ID: appregistrering

1. Gå til [entra.microsoft.com](https://entra.microsoft.com) → *Applikasjoner → Appregistreringer → Ny registrering*.
   - Navn: `LNU Arkiv (Outlook)`
   - Kontotyper: *Bare kontoer i denne organisasjonskatalogen*
   - Omdirigerings-URI: plattform **Single-page application (SPA)**, verdi `brk-multihub://<konto>.github.io` (bare domenet, uten sti, med små bokstaver). `konfigurer.py` skriver ut riktig verdi.
2. Kopier **Program-ID (klient)** og **Katalog-ID (tenant)** fra oversikten.
3. Under *API-tillatelser*: fjern `User.Read` som ligger der fra før (den trengs ikke), og velg *Legg til en tillatelse → Microsoft Graph → Delegerte tillatelser* → **`Sites.Selected`**.
4. Trykk **Gi administratorsamtykke for LNU**.

`Sites.Selected` gir alene ingen tilgang til noe. Appen får tilgang til arkivsiden først i steg 5.

### 4. Konfigurer

Fyll inn `innstillinger.json`:

```json
{
  "baseUrl": "https://<konto>.github.io/lnu-arkiv",
  "clientId": "<Program-ID fra steg 3>",
  "tenantId": "<Katalog-ID fra steg 3>",
  "sharepoint": {
    "hostname": "lnu.sharepoint.com",
    "sitePath": "/sites/arkiv",
    "library": "E-postarkiv",
    "folder": "",
    "typeList": "Kommunikasjonstyper",
    "projectList": "Prosjekter"
  },
  "category": "Arkivert i SharePoint"
}
```

Ligger navnene i listene i en annen kolonne enn standardkolonnen Title (for eksempel fordi listen er laget fra Excel), legger du til `"projectField"` og/eller `"typeField"` under `sharepoint` med kolonnens **interne** navn, for eksempel `"projectField": "field_0"`. Det interne navnet står etter `Field=` i adressen når du redigerer kolonnen. Oppslagskolonnen i biblioteket bør hente fra den samme kolonnen.

Sjekk at `hostname` og `sitePath` stemmer med adressen til arkivsiden. `folder` kan stå tom (alt havner i roten av biblioteket) eller være et mappenavn. Sett `"category": null` hvis du ikke vil at tillegget skal sette kategori.

Kjør så:

```
python3 konfigurer.py
```

Skriptet lager `manifest.xml` og `docs/config.js`, og lagrer en fast tilleggs-ID i `innstillinger.json`. Det skriver også ut omdirigerings-URI-en og de ferdige kallene til steg 5. Publiser `docs/config.js` sammen med resten (push til GitHub).

### 5. Gi appen tilgang til arkivsiden

Dette gjøres én gang, av en administrator, i [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer).

1. Logg inn i Graph Explorer med administratorkontoen.
2. Åpne fanen **Modify permissions** og gi samtykke til `Sites.FullControl.All`. Hvis du får spørsmål om å samtykke på vegne av organisasjonen, lar du det være **uavkrysset**. Det er bare du som trenger tilgangen, og bare nå.
3. Finn ID-en til arkivsiden. Velg `GET` og kjør:
   ```
   https://graph.microsoft.com/v1.0/sites/lnu.sharepoint.com:/sites/arkiv?$select=id
   ```
   Kopier verdien i `id` (en lang tekst med to komma i).
4. Gi appen skrivetilgang. Velg `POST` og kjør:
   ```
   https://graph.microsoft.com/v1.0/sites/<id fra punkt 3>/permissions
   ```
   med dette i *Request body* (bytt inn klient-ID-en, eller kopier ferdig tekst fra `konfigurer.py`):
   ```json
   {
     "roles": ["write"],
     "grantedToIdentities": [
       { "application": { "id": "<Program-ID fra steg 3>", "displayName": "LNU Arkiv (Outlook)" } }
     ]
   }
   ```
   Svaret skal være `201 Created`.
5. Kontroller med `GET https://graph.microsoft.com/v1.0/sites/<id>/permissions`. Appen skal stå oppført med rollen `write`.
6. **Rydd opp:** Gå til Entra → *Bedriftsapplikasjoner*, fjern filteret for applikasjonstype, søk opp **Graph Explorer** og slett den. Da forsvinner `Sites.FullControl.All`-tilgangen du ga i punkt 2. Graph Explorer opprettes automatisk på nytt neste gang noen bruker den.

Rollen `write` lar appen lese og endre innhold på arkivsiden, men ikke endre tilganger eller innstillinger. Tilgangen kan trekkes tilbake med `DELETE https://graph.microsoft.com/v1.0/sites/<id>/permissions/<permission-id>`.

### 6. Test på egen konto

1. Åpne [aka.ms/olksideload](https://aka.ms/olksideload) i Outlook på nett.
2. Velg *Mine tillegg → Egendefinerte tillegg → Legg til et egendefinert tillegg → Legg til fra fil* og velg `manifest.xml`.
3. Åpne en e-post. I Outlook på nett og ny Outlook ligger tillegget under **Apper**-knappen (eller «…»). I klassisk Outlook dukker knappen opp i båndet.
4. Første gang kan du bli bedt om å logge inn i et lite vindu.

Test disse tilfellene:
- en vanlig e-post i egen innboks
- en e-post du selv har sendt (skal få retning «Utgående»)
- en e-post med store vedlegg (over 4 MB)
- **en e-post i post@lnu.no.** Dette er viktigst å sjekke, fordi det ikke er bekreftet at Outlook lar tillegg hente e-post fra delte postkasser i alle Outlook-varianter.
- med og uten kategorien «Arkivert i SharePoint» opprettet

### 7. Rull ut til alle

1. Fjern testinstallasjonen fra steg 6.
2. Gå til [admin.microsoft.com](https://admin.microsoft.com) → *Innstillinger → Integrerte apper → Last opp egendefinerte apper*.
3. Velg *Office-tillegg*, last opp `manifest.xml`, og tildel tillegget til alle ansatte eller en gruppe.
4. Be alle ansatte opprette kategorien **Arkivert i SharePoint** i Outlook én gang (*Kategoriser → Ny kategori*). Navnet må være helt likt det som står i `innstillinger.json`. Kategorien synkroniseres til alle enhetene deres.
   For delte postkasser må kategorien trolig opprettes i postkassens egen kategoriliste. Åpne postkassen i Outlook på nett og gå til *Innstillinger → Generelt → Kategorier*.

Mangler kategorien, blir e-posten arkivert som vanlig, og brukeren får et tips om å opprette den.

Det kan ta noen timer, opptil et døgn, før knappen dukker opp hos alle.

## Endringer senere

- **Kode og utseende** (`docs/`): Publiser på nytt. Endringen gjelder neste gang tillegget åpnes.
- **Nye prosjekter eller typer:** Legg dem inn i SharePoint-listene. Brukerne trykker *Last inn listene på nytt*.
- **Manifestet** (navn, knapp, ikon, adresse, tilganger): Øk `version` i `innstillinger.json` (f.eks. `1.0.1.0`), kjør `konfigurer.py` og last opp `manifest.xml` på nytt under *Integrerte apper → Oppdater*.

## Personvern

E-postinnholdet går direkte fra Outlook til SharePoint via brukerens egen Outlook. Nettstedet i steg 2 leverer bare programfilene og ser aldri e-post eller innloggingsdata. Ingen tredjepart behandler data, så tillegget trenger ingen egen databehandleravtale utover den dere har med Microsoft. GitHub ser IP-adressen til de som laster tillegget, men ikke noe innhold.

## Kjente begrensninger

- **Eldre Outlook-versjoner:** Tillegget krever en oppdatert Microsoft 365-versjon av Outlook (Mailbox 1.14 og nested app authentication). Er Outlook for gammel, viser tillegget en melding om å oppdatere eller bruke Outlook på nett.
- **Mobil:** Støttes ikke i denne versjonen.
- **Nylig sendt e-post:** Rett etter sending kan meldings-ID-en mangle i Sendte elementer. Da virker ikke duplikatsjekken for den e-posten. Vent et øyeblikk før du arkiverer.
- **Kategori:** Hver postkasse har sin egen kategoriliste, så kategorien må opprettes per bruker (og per delt postkasse).

## Feilsøking

| Melding | Sannsynlig årsak |
|---|---|
| «Tillegget er ikke konfigurert ennå» | `docs/config.js` mangler eller er ikke publisert. Kjør `konfigurer.py` og publiser. |
| «Denne Outlook-versjonen er for gammel» | Oppdater Outlook eller bruk Outlook på nett. |
| «Ingen tilgang til arkivsiden» | Steg 5 er ikke gjort (eller feilet), eller brukeren er ikke medlem av arkivsiden. |
| «Fant ikke arkivsiden» | Feil `hostname` eller `sitePath`. |
| «Fant ikke «E-postarkiv» …» | Navnet på biblioteket eller listene stemmer ikke med innstillingene. |
| «Biblioteket mangler kolonnen MeldingsID …» | Kolonnen mangler, har feil internt navn eller er ikke indeksert. |
| «Du mangler tilgang» | Brukeren kan ikke redigere arkivsiden, eller administratorsamtykke for `Sites.Selected` er ikke gitt. |
| «Outlook klarte ikke å hente e-posten» | Typisk e-post i en delt postkasse eller kryptert e-post som Outlook ikke gir tillegg tilgang til. |
| Lagret, men uten metadata | Et kolonnenavn er feil. Sjekk de interne navnene (se kolonnens URL i bibliotekinnstillingene: `Field=...`). |
| Tips om å opprette kategorien | Kategorien «Arkivert i SharePoint» finnes ikke i brukerens (eller den delte postkassens) kategoriliste. |
| Innloggingsfeil / AADSTS50011 | Omdirigerings-URI-en mangler eller er feil. Den skal være `brk-multihub://` + domenet, type SPA. |

## Videre

- **Oppbevaring:** Legg en Purview-oppbevaringsetikett som standard på biblioteket, slik at arkiverte e-poster ikke kan slettes før oppbevaringstiden er ute.
