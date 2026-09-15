/**
 * Charge des blocs du userscript et les évalue — socle commun aux bancs.
 *
 * ⭐⭐ POURQUOI CE FICHIER EXISTE : les trois bancs découpaient le userscript
 *    chacun de leur côté, avec le même code. Le jour où un bloc s'est mis à
 *    dépendre d'un autre, il a fallu corriger l'ordre d'extraction TROIS fois —
 *    la troisième a suffi à dire que la cause était la copie, pas l'ordre.
 *
 * ⚠️ L'ORDRE DES BLOCS EST CELUI DE L'ÉVALUATION : `champs` avant `colonnes`,
 *    car les libellés des colonnes dérivent de `CHAMPS`.
 */

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const racine = join(dirname(fileURLToPath(import.meta.url)), '..');
export const CHEMIN_SCRIPT = join(racine, 'WME_POI_Event_Updater.user.js');
export const source = readFileSync(CHEMIN_SCRIPT, 'utf8');

/** Le texte d'un bloc balisé, sans ses marqueurs. */
export function bloc(nom) {
    const d = source.indexOf(`// ==== banc:${nom} ====`);
    const f = source.indexOf(`// ==== /banc:${nom} ====`);
    if (d === -1 || f === -1) {
        console.error(`✖ Bloc « banc:${nom} » introuvable dans le userscript.`);
        process.exit(1);
    }
    return source.slice(d, f);
}

/**
 * Évalue les blocs demandés et rend les symboles nommés.
 *
 * @param {string[]} noms   blocs à concaténer, DANS L'ORDRE D'ÉVALUATION
 * @param {string[]} symboles noms à récupérer
 */
export function charger(noms, symboles) {
    const corps = noms.map(bloc).join('\n');
    return new Function(corps + `\nreturn { ${symboles.join(', ')} };`)();
}
