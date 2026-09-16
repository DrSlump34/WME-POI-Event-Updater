/**
 * Les libelles appeles existent-ils dans les DEUX langues ?
 *     node tools/check-libelles.mjs
 *
 * ⭐⭐⭐ UN LIBELLE MANQUANT NE CASSE RIEN — il affiche « undefined » quelque part
 *    dans l'interface, souvent dans une infobulle que personne ne survole avant
 *    des semaines. C'est exactement la panne que ce projet refuse : celle qui ne
 *    se signale pas.
 *
 * ⚠️⚠️ LES DECLARATIONS NE SONT PAS EN DEBUT DE LIGNE. Plusieurs cles tiennent
 *    sur une meme ligne (`panelTitle:'…', chooseFile:'…'`). Un motif ancre sur
 *    le debut de ligne en rate les trois quarts et annonce vingt-huit libelles
 *    manquants qui existent tous — un controle qui crie au loup n'est pas moins
 *    nuisible qu'un controle muet : on cesse de le lire.
 *
 * ⇒ On isole donc CHAQUE bloc de langue, et l'on y releve les cles.
 */

import { source } from './extraire.mjs';

/**
 * VIDE LES CHAINES D'UN MORCEAU DE CODE, en gardant sa longueur et ses lignes.
 *
 * ⚠️⚠️ SANS CELA, ON COMPTE DES MOTS POUR DES CLES. « not recognised: » a
 *    l'interieur d'un libelle anglais a ete releve comme une declaration, et le
 *    controle a annonce un doublon qui n'existe pas. Un controle qui accuse a
 *    tort coute le meme temps qu'un vrai defaut, et il use la confiance qu'on
 *    lui porte — c'est pire.
 */
function sansChaines(code) {
    return code.replace(/'(?:[^'\\]|\\.)*'|"(?:[^"\\]|\\.)*"|`(?:[^`\\]|\\.)*`/g,
        (m) => m[0].repeat(m.length));
}

/** Le corps d'un bloc de langue, des accolades ouvrantes a la fermeture de meme niveau. */
function blocLangue(code) {
    const d = source.indexOf(`            ${code}: {`);
    if (d === -1) return null;
    const f = source.indexOf('\n            }', d);

    return f === -1 ? null : source.slice(d, f);
}

const langues = ['fr', 'en'];
const blocs = langues.map((c) => ({ code: c, corps: blocLangue(c) }));

const absents = blocs.filter((b) => !b.corps).map((b) => b.code);
if (absents.length) {
    console.error(`✖ Bloc(s) de langue introuvable(s) : ${absents.join(', ')}. Rien n a ete verifie.`);
    process.exit(3);
}

/* ⚠️⚠️ UNE CLE EN DOUBLE FAIT GAGNER LA MAUVAISE, EN SILENCE. En JavaScript, la
   DERNIERE declaration l emporte : un libelle refait en tete d objet est ecrase
   par son homonyme reste plus bas, et l ecran affiche l ancien texte alors que
   le nouveau est bien la, sous les yeux, dans le fichier. Vu a l usage sur le
   pied de la fenetre, qui invitait encore a « cliquer sur la coche » — un
   bouton qui n existe plus. Un Set les avale : on compte d abord. */
const doublons = [];
const clesPar = new Map();
blocs.forEach((b) => {
    const brutes = [...sansChaines(b.corps).matchAll(/(?:^|[{,\s])([a-zA-Z][a-zA-Z0-9]*)\s*:/g)].map((m) => m[1]);
    const vues = new Map();
    brutes.forEach((c) => vues.set(c, (vues.get(c) || 0) + 1));
    [...vues].filter((e) => e[1] > 1).forEach((e) => doublons.push(e[0] + ' — ' + e[1] + ' fois en ' + b.code));
    const cles = new Set(brutes);
    clesPar.set(b.code, cles);
    if (cles.size === 0) {
        console.error(`✖ Aucune cle relevee dans « ${b.code} » : le motif ne mord plus.`);
        process.exit(3);
    }
});

const appelees = new Set(
    [...source.matchAll(/\bt\('([a-zA-Z][a-zA-Z0-9]*)'/g)].map((m) => m[1])
);

const manques = [];
[...appelees].sort().forEach((cle) => {
    const sans = langues.filter((c) => !clesPar.get(c).has(cle));
    if (sans.length) manques.push(`${cle} — absent de : ${sans.join(', ')}`);
});

/* ⚠️ ET L'INVERSE : une cle declaree que plus personne n'appelle est un libelle
   mort, qu'on traduit et qu'on relit pour rien. Ce n'est pas une faute, donc on
   le SIGNALE sans echouer. */
const jamaisAppelees = [...clesPar.get('fr')].filter((c) => !appelees.has(c)).sort();

console.log(`${appelees.size} libelles appeles ; ${clesPar.get('fr').size} en fr, ${clesPar.get('en').size} en en.`);

if (doublons.length) {
    console.error('\n✖ ' + doublons.length + ' libelle(s) declare(s) DEUX FOIS — c est le DERNIER qui gagne :');
    doublons.forEach((d) => console.error('    ' + d));
}

if (manques.length) {
    console.error(`\n✖ ${manques.length} libelle(s) qui afficheraient « undefined » :`);
    manques.forEach((m) => console.error(`    ${m}`));
} else {
    console.log('✔ Tous les libelles appeles existent dans les deux langues.');
}

if (jamaisAppelees.length) {
    console.log(`\n⏳ ${jamaisAppelees.length} declare(s) que plus personne n appelle :`);
    console.log(`    ${jamaisAppelees.join(', ')}`);
}

process.exit(manques.length || doublons.length ? 1 : 0);
