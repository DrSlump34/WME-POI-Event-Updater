/**
 * Banc du CHARGEMENT d'un classeur — `node tools/banc-chargement.mjs <classeur.xlsx> [xlsx.full.min.js]`
 *
 * ⭐⭐⭐⭐ IL REJOUE LE CODE DU SCRIPT, PAS UNE RECOPIE. Le corps de
 *    `reader.onload` est découpé dans le userscript entre deux ancres, puis
 *    exécuté avec de vrais bouchons. C'est tout l'intérêt : le 15/09/2026, une
 *    fiche d'essai annonçait « 6 POI » pour un onglet qui en donne 3 — la
 *    prédiction avait été écrite en RECOPIANT la logique de validation, et la
 *    recopie comptait les lignes AVANT les refus. Le rejeu l'a dit tout de suite.
 *
 * ⚠️ IL LUI FAUT SheetJS, que le dépôt ne porte pas (aucune dépendance n'y entre) :
 *    le script le charge par `@require`. Le récupérer une fois, hors du dépôt :
 *
 *      curl -o /tmp/xlsx.full.min.js \
 *        https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
 *
 *    puis passer son chemin en second argument. Par défaut : `./xlsx.full.min.js`.
 *
 * ⚠️ CE QU'IL NE VOIT PAS : le DOM, le préchargement des lieux, l'aperçu,
 *    l'application. Il éprouve ce que le fichier DEVIENT, pas ce que l'écran en
 *    montre. Cela se vérifie dans WME, et nulle part ailleurs.
 */

import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { resolve } from 'node:path';
import { source, bloc } from './extraire.mjs';

const [classeur, sheetjs = './xlsx.full.min.js'] = process.argv.slice(2);

if (!classeur) {
    console.error('Usage : node tools/banc-chargement.mjs <classeur.xlsx> [chemin/vers/xlsx.full.min.js]');
    process.exit(2);
}

const require = createRequire(import.meta.url);
let XLSX;
try {
    XLSX = require(resolve(sheetjs));
} catch {
    console.error(`✖ SheetJS introuvable : ${resolve(sheetjs)}`);
    console.error('  Voir l’en-tête de ce fichier pour le récupérer (il n’entre pas dans le dépôt).');
    process.exit(2);
}


/* Les deux fonctions dont le corps du chargement a besoin, extraites du script. */
const { mapColumns } = new Function(
    bloc('champs') + bloc('colonnes') + '\nreturn { mapColumns };'
)();

const debutVid = source.indexOf('function getVenueIdFromPermalink(url) {');
const { getVenueIdFromPermalink } = new Function(
    source.slice(debutVid, source.indexOf('\n    }', debutVid) + 6) + '\nreturn { getVenueIdFromPermalink };'
)();

/* Le corps du chargement lui-même. */
const debut = source.indexOf('const wb = XLSX.read(');
const fin = source.indexOf('if (!all.length) {');
if (debut === -1 || fin === -1 || fin < debut) {
    console.error('✖ Corps du chargement introuvable : les ancres du userscript ont changé.');
    process.exit(1);
}

const charger = new Function(
    'XLSX', 'ev', 't', 'getVenueIdFromPermalink', 'mapColumns',
    source.slice(debut, fin) + '\nreturn { all, warnings };'
);

/* `t()` rend la CLÉ du message et ses arguments : le relevé ne dépend pas de la
   langue de l'éditeur, et il nomme le message plutôt que de le traduire. */
const t = (cle, ...args) => (args.length ? `${cle}(${args.join(',')})` : cle);

const { all, warnings } = charger(
    XLSX,
    { target: { result: new Uint8Array(readFileSync(classeur)).buffer } },
    t,
    getVenueIdFromPermalink,
    mapColumns
);

const parOnglet = new Map();
all.forEach(p => parOnglet.set(p.event, (parOnglet.get(p.event) || 0) + 1));

console.log(`=== ${classeur} ===\n`);
console.log('POI RETENUS, PAR ONGLET — c’est ce que proposera le sélecteur :');
for (const [onglet, n] of parOnglet) console.log(`  ${onglet.padEnd(20)} ${n}`);
console.log(`  ${'TOTAL'.padEnd(20)} ${all.length}`);

console.log(`\nANOMALIES : ${warnings.length}`);
warnings.forEach(w => console.log(`  · ${w}`));

if (all.length === 0) {
    console.log('\n✖ Aucun POI : le script afficherait « Aucun POI valide chargé ».');
    process.exit(1);
}
