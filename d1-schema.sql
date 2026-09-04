-- =============================================================================
-- BladelsCreatief — Cloudflare D1-schema
-- =============================================================================
-- Ontwerpkeuze: de app werkt nu al overal met sleutel/waarde-opslag (zie het
-- STORAGE_KEYS-overzicht in de app-code, boven de useStored-functie). Elke
-- "tabel" (leden, workshops, financiën, etc.) is daar al één sleutel met een
-- volledig JSON-blob als waarde, die als geheel wordt gelezen en overschreven.
--
-- Dit schema volgt die opzet bewust één-op-één, in plaats van een aparte
-- SQL-tabel per gegevenstype te maken. Dat is een bewuste keuze:
--   - Het sluit exact aan op hoe de app nu leest/schrijft (hele object per
--     sleutel in- en uitlezen), dus de Worker-API kan heel dun blijven.
--   - Het is precies waarom eerder is voorgesteld dat alléén de opslag-adapter
--     (opslagLezen/opslagSchrijven) hoeft te veranderen — de rest van de
--     ~4800 regels app-code hoeft niet aangepast te worden.
--   - Een "echte" relationele opsplitsing (aparte tabellen met kolommen en
--     vreemde sleutels) zou correcter aanvoelen, maar vergt dat vrijwel elke
--     schrijfactie in de app herschreven wordt naar losse rij-operaties —
--     een veel grotere en risicovollere ingreep dan nodig is.
--
-- Wil je op termijn wél naar een volledig relationeel schema (bijvoorbeeld om
-- rechtstreeks met SQL-rapportages te kunnen werken), dan kan dat later alsnog
-- als tweede stap, zonder dat dit schema opnieuw hoeft.
-- =============================================================================

-- Kolommen: sleutel (primaire sleutel), waarde (volledige JSON-inhoud van deze "tabel"),
-- bijgewerkt_op (UTC-tijdstip van laatste wijziging).
CREATE TABLE IF NOT EXISTS opslag (
  sleutel        TEXT PRIMARY KEY NOT NULL,
  waarde         TEXT NOT NULL,
  bijgewerkt_op  TEXT NOT NULL DEFAULT (datetime('now'))
);

-- =============================================================================
-- De sleutels die de app gebruikt (uit STORAGE_KEYS in de broncode). Dit is
-- puur ter referentie — je hoeft ze niet vooraf aan te maken, de app schrijft
-- ze vanzelf aan bij eerste gebruik, precies zoals nu met Claude's opslag.
-- =============================================================================
-- bladels:leden                    — ledenlijst
-- bladels:workshops                — workshopdefinities
-- bladels:workshopinschrijvingen   — inschrijvingen op workshops
-- bladels:transacties              — financiële boekingen
-- bladels:rekeningen               — grootboekrekeningen
-- bladels:begrotingkoppelingen     — trefwoord-naar-grootboekcode-koppelingen
-- bladels:dagdelen                 — dagdelen/groepen
-- bladels:workshopsoorten          — workshop-soorten
-- bladels:agendapuntenvooraf       — standaard agendapunten (vooraf)
-- bladels:agendapuntenafsluitend   — standaard agendapunten (afsluitend)
-- bladels:begroting                — begrotingsregels per boekjaar
-- bladels:boekjaren                — lijst boekjaren
-- bladels:vergaderingen            — vergaderingen met agenda/notulen
-- bladels:actielijst               — actiepunten
-- bladels:contributies             — contributiebetalingen
-- bladels:pins                     — gehashte pincodes/wachtwoorden (hash + zout, nooit leesbaar)
-- bladels:beveiliging              — 2FA/wachtwoordinstellingen
-- bladels:standaarden              — inactiviteitsduur, logogrootte
-- bladels:workshopsortering        — laatst gebruikte sorteervolgorde workshops
-- bladels:tfa-secrets              — 2FA-geheimen per persoon
-- bladels:tfa-vertrouwd            — 2FA-vertrouwensperiode per persoon
-- bladels:rolpermissies            — rechten per bestuursfunctie
-- bladels:logboek                  — logboek van wijzigingen
-- bladels:prullenbak               — zachte verwijdering (herstelbaar)
-- bladels:sessies                  — "wie is er nu actief"-signaal (presentie)

-- =============================================================================
-- Optioneel: eenmalige index op bijgewerkt_op, alleen relevant als je ooit een
-- overzicht wilt van recent gewijzigde tabellen (niet nodig voor de app zelf,
-- die leest/schrijft altijd op de primaire sleutel).
-- =============================================================================
CREATE INDEX IF NOT EXISTS idx_opslag_bijgewerkt ON opslag (bijgewerkt_op);
