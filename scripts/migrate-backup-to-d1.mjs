#!/usr/bin/env node
// scripts/migrate-backup-to-d1.mjs
//
// Zet een volledige back-up-JSON (gedownload via "Download volledige back-up"
// in de huidige Claude-versie van de app) om naar een .sql-bestand met
// INSERT-statements voor de D1-tabel `opslag` — dit is de uitwerking van
// stap 8 uit de README ("Data overzetten vanuit de huidige (Claude-)versie").
//
// Gebruik:
//   node scripts/migrate-backup-to-d1.mjs <pad-naar-backup.json> [uitvoer.sql]
//
// Daarna, eerst lokaal testen (aanbevolen):
//   wrangler d1 execute bladelscreatief-db --local --file=<uitvoer.sql>
// en als dat goed gaat, naar de live database:
//   wrangler d1 execute bladelscreatief-db --remote --file=<uitvoer.sql>

import { readFileSync, writeFileSync } from 'node:fs';

// Koppeling tussen de veldnamen in de back-up-JSON (zie exportBackup() en de
// backupData-prop in src/App.jsx) en de opslagsleutels uit STORAGE_KEYS.
// Let op: niet alle STORAGE_KEYS zitten in de back-up (bijv. pins, tfa-secrets,
// sessies en de prullenbak bewust niet) — dat is verwacht gedrag.
const VELD_NAAR_SLEUTEL = {
  members: 'bladels:leden',
  workshops: 'bladels:workshops',
  inschrijvingen: 'bladels:workshopinschrijvingen',
  tx: 'bladels:transacties',
  accounts: 'bladels:rekeningen',
  budget: 'bladels:begroting',
  boekjaren: 'bladels:boekjaren',
  vergaderingen: 'bladels:vergaderingen',
  actielijst: 'bladels:actielijst',
  contributies: 'bladels:contributies',
  rolpermissies: 'bladels:rolpermissies',
  begrotingKoppelingen: 'bladels:begrotingkoppelingen',
};

const [, , inputPath, outputPathArg] = process.argv;
if (!inputPath) {
  console.error('Gebruik: node migrate-backup-to-d1.mjs <pad-naar-backup.json> [uitvoer.sql]');
  process.exit(1);
}
const outputPath = outputPathArg || 'migratie-data.sql';

const ruw = readFileSync(inputPath, 'utf8');
const backup = JSON.parse(ruw);

function sqlString(str) {
  // SQLite-string-escaping: enkele quote verdubbelen. Backslash-escaping is niet nodig.
  return `'${str.replace(/'/g, "''")}'`;
}

const regels = [];
const overgeslagen = [];

for (const [veld, sleutel] of Object.entries(VELD_NAAR_SLEUTEL)) {
  if (!(veld in backup)) {
    overgeslagen.push(veld);
    continue;
  }
  const waarde = JSON.stringify(backup[veld]);
  regels.push(
    `INSERT OR REPLACE INTO opslag (sleutel, waarde) VALUES (${sqlString(sleutel)}, ${sqlString(waarde)});`
  );
}

writeFileSync(outputPath, regels.join('\n') + '\n', 'utf8');

console.log(`${regels.length} tabel(len) weggeschreven naar ${outputPath}.`);
if (overgeslagen.length) {
  console.log(`Niet gevonden in de back-up (overgeslagen, normaal als leeg/ongebruikt): ${overgeslagen.join(', ')}`);
}
console.log('');
console.log('Volgende stap, eerst lokaal testen:');
console.log(`  wrangler d1 execute bladelscreatief-db --local --file=${outputPath}`);
console.log('En daarna naar de live database:');
console.log(`  wrangler d1 execute bladelscreatief-db --remote --file=${outputPath}`);
