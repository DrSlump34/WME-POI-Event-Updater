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

import { source, charger } from './extraire.mjs';

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
    let d = source.indexOf(`            ${code}: {`);
    if (d === -1) d = source.indexOf(`            '${code}': {`);
    if (d === -1) return null;
    const f = source.indexOf('\n            }', d);

    return f === -1 ? null : source.slice(d, f);
}

/* ⭐ LES HUIT LANGUES DE LA CHARTE COMMUNE (WCT, WJN, WRP, WDA). */
const langues = ['fr', 'en', 'de', 'es', 'it', 'pt-BR', 'pt-PT', 'he'];
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
    const cles = new Set(brutes.filter((c) => !langues.includes(c)));
    clesPar.set(b.code, cles);
    if (cles.size === 0) {
        console.error(`✖ Aucune cle relevee dans « ${b.code} » : le motif ne mord plus.`);
        process.exit(3);
    }
});

/* ⚠️⚠️ UNE CLÉ S'APPELLE AUSSI SANS `t('…')` DEVANT. `t(x ? 'a' : 'b')`,
   `montrerGuide('guideOnglet', …)`, `traduire('plusApplique')` : le motif ne
   voyait que `t('clé'`, et annonçait « morts » au moins dix-sept libellés
   vivants — dont celui qui dit qu'Appliquer RETIRE des valeurs. Un ménage fait
   d'après cette liste aurait affiché des clés brutes (audit du 25/09/2026).
   ⇒ Toute clé déclarée qui apparaît ENTRE APOSTROPHES hors du dictionnaire
     compte comme appelée. Le doute profite au libellé : un faux « vivant »
     coûte une ligne, un faux « mort » coûte un écran. */
const debutDico = source.indexOf('if (!_strings) _strings = {');
const finDico = source.indexOf('\n        };', debutDico);
const horsDico = source.slice(0, debutDico) + source.slice(finDico);
const declarees = new Set([...clesPar.values()].flatMap((s) => [...s]));

/* ⚠️⚠️ ET UNE CLÉ APPELÉE INDIRECTEMENT DOIT EXISTER. Le filtre ci-dessus ne
   retient que les clés DÉCLARÉES : une clé passée à `montrerGuide` ou choisie
   par un ternaire, puis supprimée des deux langues, disparaissait du contrôle
   au lieu de le faire échouer (mutation du 25/09/2026). D'où les formes
   d'appel indirect, relevées une par une dans le script — elles, doivent
   exister qu'elles soient déclarées ou non. */
const CLE = "'([a-zA-Z][a-zA-Z0-9]*)'";
const indirectes = [
    new RegExp('\\bt\\([^()]*?\\?\\s*' + CLE + '\\s*:\\s*' + CLE, 'g'),     // t(x ? 'a' : 'b')
    new RegExp('\\btraduire\\(' + CLE, 'g'),                                 // rendreComplements
    new RegExp('\\bmontrerGuide\\(' + CLE + ',\\s*' + CLE, 'g'),             // le guidage
    new RegExp('\\befface\\([^;]*?,\\s*' + CLE + '\\)', 'g'),                // les effacements
    new RegExp('\\b(?:applied|partial|timeout|erreur):\\s*' + CLE, 'g'),     // la table des statuts
];
const appelsIndirects = indirectes.flatMap((re) =>
    [...horsDico.matchAll(re)].flatMap((m) => m.slice(1).filter(Boolean)));
if (!appelsIndirects.length) {
    console.error('✖ Aucun appel indirect relevé : les motifs ne mordent plus.');
    process.exit(3);
}

/* ⚠️ LES CLÉS DÉRIVÉES : le libellé d'une colonne (`ch` + clé du champ), le
   motif d'un champ montré (`mo` + clé), une valeur de WME (`va` + clé WME en
   casse chameau). Elles s'appellent par concaténation — aucun motif ne les
   voit — et on les DÉDUIT donc de `CHAMPS` et de `VALEURS_WME`, les mêmes
   tables que le script lit. Une colonne ajoutée sans son libellé traduit
   fait alors échouer ce contrôle.
   ⚠️ Les catégories (traduites par WME lui-même) et les nombres de places
      (des chiffres) n'ont pas de clé. */
const { CHAMPS, VALEURS_WME } = charger(['champs'], ['CHAMPS', 'VALEURS_WME']);
const majuscule = (c) => c.charAt(0).toUpperCase() + c.slice(1);
const chameau = (k) => k.toLowerCase().split('_').map(majuscule).join('');
const derivees = [
    ...CHAMPS.map((c) => 'ch' + majuscule(c.cle)),
    ...CHAMPS.filter((c) => c.motif).map((c) => 'mo' + majuscule(c.cle)),
    ...Object.keys(VALEURS_WME).filter((r) => r !== 'categories' && r !== 'estimatedNumberOfSpots')
        .flatMap((r) => Object.keys(VALEURS_WME[r]).map((k) => 'va' + chameau(k))),
];

const appelees = new Set([
    ...[...source.matchAll(/\bt\('([a-zA-Z][a-zA-Z0-9]*)'/g)].map((m) => m[1]),
    ...appelsIndirects,
    ...derivees,
    ...[...horsDico.matchAll(/'([a-zA-Z][a-zA-Z0-9]*)'/g)].map((m) => m[1]).filter((c) => declarees.has(c)),
]);

const manques = [];
[...appelees].sort().forEach((cle) => {
    const sans = langues.filter((c) => !clesPar.get(c).has(cle));
    if (sans.length) manques.push(`${cle} — absent de : ${sans.join(', ')}`);
});

/* ⚠️ ET L'INVERSE : une cle declaree que plus personne n'appelle est un libelle
   mort, qu'on traduit et qu'on relit pour rien. Ce n'est pas une faute, donc on
   le SIGNALE sans echouer. */
const jamaisAppelees = [...clesPar.get('fr')].filter((c) => !appelees.has(c)).sort();

console.log(`${appelees.size} libelles appeles ; ` + langues.map((c) => `${clesPar.get(c).size} en ${c}`).join(', ') + '.');

/* ⚠️ LE MÊME JEU DE CLÉS PARTOUT (charte) : une clé présente en français et
   absente en hébreu s'afficherait en anglais — sans erreur, sans que personne
   ne le voie avant un éditeur israélien. */
const reference = clesPar.get('fr');
langues.slice(1).forEach((c) => {
    const ici = clesPar.get(c);
    const enMoins = [...reference].filter((k) => !ici.has(k));
    const enTrop = [...ici].filter((k) => !reference.has(k));
    if (enMoins.length) manques.push(`${c} — clés absentes par rapport au français : ${enMoins.join(', ')}`);
    if (enTrop.length) manques.push(`${c} — clés en trop par rapport au français : ${enTrop.join(', ')}`);
});

if (doublons.length) {
    console.error('\n✖ ' + doublons.length + ' libelle(s) declare(s) DEUX FOIS — c est le DERNIER qui gagne :');
    doublons.forEach((d) => console.error('    ' + d));
}

if (manques.length) {
    console.error(`\n✖ ${manques.length} libelle(s) qui afficheraient « undefined » :`);
    manques.forEach((m) => console.error(`    ${m}`));
} else {
    console.log(`✔ Tous les libelles appeles existent dans les ${langues.length} langues, avec le même jeu de clés.`);
}

if (jamaisAppelees.length) {
    console.log(`\n⏳ ${jamaisAppelees.length} declare(s) que plus personne n appelle :`);
    console.log(`    ${jamaisAppelees.join(', ')}`);
}

/* ⚠️ LA TYPOGRAPHIE FRANÇAISE, DANS LE TEXTE AFFICHÉ SEULEMENT : une espace
   fine insécable (U+202F) avant « : ; ! ? », une insécable (U+00A0) à
   l'intérieur des guillemets. Avec une espace ordinaire, le signe part seul à
   la ligne — dix-neuf cas relevés par l'audit du 25/09/2026.
   On ne lit que le CONTENU des chaînes du bloc français : le code (un
   ternaire `x ? a : b`) n'est pas du texte. */
function contenusDesChaines(code) {
    /* Les trois formes : '…', "…" et `…` (dont on retire les ${…}). */
    const contenus = [];
    code.replace(/'((?:[^'\\]|\\.)*)'|"((?:[^"\\]|\\.)*)"|`((?:[^`\\]|\\.)*)`/g, (m, simple, double, gabarit) => {
        if (simple !== undefined) contenus.push(simple);
        else if (double !== undefined) contenus.push(double);
        else contenus.push(gabarit.replace(/\$\{[^}]*\}/g, 'X'));
        return m;
    });
    return contenus;
}
const typo = [];
contenusDesChaines(blocLangue('fr')).forEach((txt) => {
    if (/ [:;!?]/.test(txt)) typo.push(`espace ordinaire avant : ; ! ? — « ${txt} »`);
    if (/« |«[^  ]/.test(txt) || / »|[^  ]»/.test(txt)) typo.push(`guillemets sans insécable — « ${txt} »`);
});
if (typo.length) {
    console.error(`\n✖ ${typo.length} faute(s) de typographie française :`);
    typo.forEach((x) => console.error('    ' + x));
} else {
    console.log('✔ Typographie française : insécables en place.');
}

process.exit(manques.length || doublons.length || typo.length ? 1 : 0);
