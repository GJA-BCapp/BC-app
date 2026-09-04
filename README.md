# BladelsCreatief — Administratie

Leden- en financiële administratie van BladelsCreatief, gebouwd als een gewoon React-project
(Vite), met Cloudflare Pages + Pages Functions + D1 als opslag — volledig los van Claude.

## Wat zit er in deze map

```
├── src/App.jsx                    — de volledige applicatie (alle schermen/logica)
├── src/main.jsx                   — start de app
├── src/index.css                  — Tailwind-basis
├── functions/api/opslag/[key].js  — de API die de D1-database aanspreekt
├── d1-schema.sql                  — het databaseschema (één keer uit te voeren)
├── wrangler.toml                  — Cloudflare-configuratie (D1-koppeling)
├── package.json, vite.config.js, tailwind.config.js, postcss.config.js, index.html
└── .gitignore
```

**Belangrijk om te weten:** alle gegevens (leden, financiën, workshops, etc.) lopen door twee
functies in `src/App.jsx`, met de naam `opslagLezen` en `opslagSchrijven` (zoek er gerust naar).
Die praten met de API in `functions/api/opslag/[key].js`, die op zijn beurt met de D1-database
praat. De rest van de ~5000 regels code in `App.jsx` raakt de opslag nergens rechtstreeks aan —
bij eventuele toekomstige aanpassingen aan hóe/wáár data wordt opgeslagen, zijn dus alleen die
twee plekken relevant.

## Stap voor stap: eerste keer opzetten

### 1. D1-database aanmaken
```
wrangler d1 create bladelscreatief-db
```
Dit geeft een `database_id` terug — vul die in bij `wrangler.toml` (regel met
`database_id = "VUL_HIER_JE_D1_DATABASE_ID_IN"`).

### 2. Schema uitvoeren
```
wrangler d1 execute bladelscreatief-db --file=d1-schema.sql
```
(Lokaal eerst testen kan met `--local` erachter toegevoegd.)

### 3. Naar GitHub pushen
```
git init
git add .
git commit -m "Eerste versie"
git branch -M main
git remote add origin <jouw-github-repo-url>
git push -u origin main
```

### 4. Cloudflare Pages-project aanmaken
In het Cloudflare-dashboard: **Workers & Pages → Create → Pages → Connect to Git**, kies deze
repository. Bouwinstellingen:
- **Build command**: `npm run build`
- **Build output directory**: `dist`

### 5. D1 koppelen aan het Pages-project
Ga naar het Pages-project → **Settings → Functions → D1 database bindings** → voeg een binding
toe met variabelenaam `DB`, gekoppeld aan de `bladelscreatief-db`-database. (Cloudflare's
dashboard-indeling kan wijzigen — zoek desnoods op "D1 binding Pages" in hun documentatie als dit
onderdeel er anders uitziet dan hier beschreven.)

### 6. (Aanbevolen) een gedeelde API-sleutel instellen
Dit is een basale beveiligingsdrempel voor de opslag-API — zonder deze stap is de API in
principe door iedereen te benaderen die de URL-structuur raadt.
- Bij het Pages-project → **Settings → Environment variables**:
  - `API_SECRET` (Functions-kant, geheim) — een lang, willekeurig wachtwoord
  - `VITE_API_SECRET` (build-kant, wordt in de frontend meegebouwd) — **dezelfde waarde**
- Let op: deze sleutel is uiteindelijk wel terug te vinden in de gebouwde frontend-code (dat
  geldt voor elke sleutel die in browsercode gebruikt wordt) — het is een drempel, geen sterke
  beveiliging. Voor iets steviger: overweeg Cloudflare Access (Zero Trust) vóór de site te
  zetten, dat regel je los in het Cloudflare-dashboard, niet in deze code.

### 7. Opnieuw deployen
Na het instellen van de environment variables: trigger een nieuwe deploy (bijvoorbeeld met een
lege commit + push, of via de "Retry deployment"-knop in het dashboard), zodat de
`VITE_API_SECRET` daadwerkelijk meegebouwd wordt.

### 8. Data overzetten vanuit de huidige (Claude-)versie
Download in de huidige app een volledige back-up (Instellingen → Back-up →
"Download volledige back-up"). Dat JSON-bestand bevat alle huidige gegevens, per tabel. Zet
elke tabel handmatig klaar in de nieuwe database, bijvoorbeeld:
```
wrangler d1 execute bladelscreatief-db --command="INSERT INTO opslag (sleutel, waarde) VALUES ('bladels:leden', '<hier de JSON-inhoud van het leden-deel uit de back-up, als string>')"
```
(Vraag gerust om een hulpscriptje dat de volledige back-up-JSON automatisch naar dit soort
commando's omzet, in plaats van dit met de hand te doen.)

## Bij elke volgende wijziging

De opzet blijft: code hier bewerken (of via Claude laten aanpassen en opnieuw aanleveren), dan
`git add . && git commit -m "..." && git push` — Cloudflare Pages bouwt en publiceert dan
automatisch de nieuwe versie. Er is geen aparte "handmatige deploy"-stap nodig zolang de
git-koppeling actief staat.

## Lokaal testen (optioneel, voor een ontwikkelaar)

```
npm install
npm run dev
```
De opslag-API (`functions/...`) draait dan **niet** mee via `npm run dev` — dat is puur de
frontend. Voor een lokale test van frontend + API + D1 samen:
```
npm run build
wrangler pages dev dist --d1 DB=bladelscreatief-db
```
