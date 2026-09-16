/**
 * La feuille de style est-elle saine ?
 *     node tools/check-css.mjs
 *
 * ⭐⭐⭐⭐ UN ACCENT GRAVE DANS LE CSS CASSE TOUT LE SCRIPT. La feuille vit dans un
 *    template literal : le premier accent grave rencontre la referme, et tout
 *    ce qui suit devient du code invalide. Le script ne se charge plus — pas une
 *    ligne en moins, TOUT.
 *
 * ⚠️ ET LE CONNAITRE NE SUFFIT PAS. Ce piege est ecrit en tete du bloc, en
 *    majuscules, et il a quand meme mordu : un commentaire redige d'un trait,
 *    avec deux noms de propriete entre accents graves par simple habitude
 *    d'ecriture. La vigilance ne tient pas un piege ; un controle, si.
 *
 * ⚠️ `node --check` NE SUFFIT PAS NON PLUS pour les autres regles ci-dessous :
 *    un CSS syntaxiquement inerte passe la verification de JavaScript.
 */

import { source } from './extraire.mjs';

const ACCENT = String.fromCharCode(96);

const debut = source.indexOf('const CSS = ' + ACCENT);
if (debut === -1) {
    console.error('✖ Le bloc CSS est introuvable : rien n a ete verifie.');
    process.exit(3);
}
const fin = source.indexOf('\n    ' + ACCENT + ';', debut);
if (fin === -1) {
    console.error('✖ La fin du bloc CSS est introuvable : rien n a ete verifie.');
    process.exit(3);
}

const css = source.slice(debut + ('const CSS = ' + ACCENT).length, fin);
const soucis = [];

/* 1. Le piege du template literal. */
const accents = (css.match(new RegExp(ACCENT, 'g')) || []).length;
if (accents > 0) {
    soucis.push(`${accents} accent(s) grave(s) dans le CSS : ils referment le literal et tuent le script`);
}

/* 2. L'interpolation, qui casse aussi bien. */
if (css.includes('${')) {
    soucis.push('une interpolation ${…} dans le CSS : elle sera evaluee, pas ecrite');
}

/* 3. Les accolades doivent s'equilibrer — un bloc laisse ouvert avale la suite
      de la feuille en silence, et les regles d'apres cessent de s'appliquer. */
const ouvrantes = (css.match(/\{/g) || []).length;
const fermantes = (css.match(/\}/g) || []).length;
if (ouvrantes !== fermantes) {
    soucis.push(`${ouvrantes} accolade(s) ouvrante(s) pour ${fermantes} fermante(s)`);
}

/* 4. Chaque variable employee doit etre declaree, et porter son repli — une
      variable manquante ne colore pas en rouge, elle ne colore PAS DU TOUT. */
const declarees = new Set([...css.matchAll(/^\s*(--peu-[\w-]+)\s*:/gm)].map((m) => m[1]));
const employees = new Set([...css.matchAll(/var\((--peu-[\w-]+)/g)].map((m) => m[1]));
const inconnues = [...employees].filter((v) => !declarees.has(v));
if (inconnues.length) {
    soucis.push(`variable(s) employee(s) mais jamais declaree(s) : ${inconnues.join(', ')}`);
}

const sansRepli = [...css.matchAll(/var\((--peu-[\w-]+)\s*\)/g)].map((m) => m[1]);
const uniques = [...new Set(sansRepli)];

console.log(`CSS : ${css.split('\n').length} lignes, ${declarees.size} variables declarees, ${employees.size} employees.`);

if (soucis.length) {
    console.error('\n✖ ' + soucis.length + ' probleme(s) :');
    soucis.forEach((p) => console.error(`    ${p}`));
    process.exit(1);
}

console.log('✔ Aucun accent grave, aucune interpolation, accolades equilibrees, variables declarees.');
if (uniques.length) {
    console.log(`\n⏳ ${uniques.length} usage(s) sans repli — une variable manquante n y colorerait RIEN :`);
    console.log(`    ${uniques.join(', ')}`);
}
process.exit(0);
