/**
 * Banc de la lecture des colonnes — `node tools/banc-colonnes.mjs`
 *
 * ⭐ IL ÉPROUVE LE FICHIER SERVI, PAS UNE COPIE. Le bloc `banc:colonnes` est
 *    extrait du userscript tel qu'il sera publié, puis évalué : une correction
 *    faite ici sans toucher au script ne passerait pas, et l'inverse non plus.
 *
 * Ce banc ne remplace pas les essais dans WME — il ne voit ni SheetJS, ni la
 * carte. Il tient une seule chose : quelle colonne porte quoi.
 */

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const racine = join(dirname(fileURLToPath(import.meta.url)), '..');
const source = readFileSync(join(racine, 'WME_POI_Event_Updater.user.js'), 'utf8');

const DEBUT = '// ==== banc:colonnes ====';
const FIN = '// ==== /banc:colonnes ====';
const i = source.indexOf(DEBUT);
const j = source.indexOf(FIN);
if (i === -1 || j === -1) {
    console.error(`✖ Bloc « banc:colonnes » introuvable dans le userscript.`);
    process.exit(1);
}

const { mapColumns, normalizeHeader } = new Function(
    source.slice(i + DEBUT.length, j) + '\nreturn { mapColumns, normalizeHeader };'
)();

let reussis = 0;
const echecs = [];

function verifier(intitule, attendu, obtenu) {
    const a = JSON.stringify(attendu);
    const o = JSON.stringify(obtenu);
    if (a === o) { reussis++; return; }
    echecs.push(`${intitule}\n      attendu : ${a}\n      obtenu  : ${o}`);
}

function plan(entetes) {
    const r = mapColumns(entetes);
    return { ...r.columns, byPosition: r.byPosition, usable: r.usable };
}

/* ------------------------------------------------------------------ *
 * 1. Les en-têtes réelles — gabarit officiel ET « ACO Events.xlsx »   *
 * ------------------------------------------------------------------ */
verifier(
    'Les en-têtes du gabarit officiel se lisent par leur nom',
    { perm: 0, name: 1, desc: 2, byPosition: false, usable: true },
    plan(['POI Permalink', 'POI Name', 'POI Description'])
);

verifier(
    'Le nom seul, sans « POI », est reconnu',
    { perm: 0, name: 1, desc: 2, byPosition: false, usable: true },
    plan(['Permalink', 'Name', 'Description'])
);

verifier(
    'En français aussi',
    { perm: 0, name: 1, desc: 2, byPosition: false, usable: true },
    plan(['Permalien', 'Nom', 'Description'])
);

verifier(
    'La casse et les espaces ne comptent pas',
    { perm: 0, name: 1, desc: 2, byPosition: false, usable: true },
    plan(['  PERMALIEN ', 'Nom  ', 'DESCRIPTION'])
);

/* ------------------------------------------------------------------ *
 * 2. Ce que la lecture par nom apporte                                *
 * ------------------------------------------------------------------ */
verifier(
    'Les colonnes peuvent être dans un autre ordre',
    { perm: 2, name: 0, desc: 1, byPosition: false, usable: true },
    plan(['POI Name', 'POI Description', 'POI Permalink'])
);

verifier(
    'Une colonne ajoutée à GAUCHE ne décale plus rien',
    { perm: 1, name: 2, desc: 3, byPosition: false, usable: true },
    plan(['Référence', 'POI Permalink', 'POI Name', 'POI Description'])
);

verifier(
    'Des colonnes ajoutées à DROITE sont ignorées sans bruit',
    { perm: 0, name: 1, desc: 2, byPosition: false, usable: true },
    plan(['POI Permalink', 'POI Name', 'POI Description', 'Categories', 'Services'])
);

verifier(
    'Une en-tête en double : la PREMIÈRE colonne gagne',
    { perm: 0, name: 2, desc: 3, byPosition: false, usable: true },
    plan(['POI Permalink', 'POI Permalink', 'POI Name', 'POI Description'])
);

/* ------------------------------------------------------------------ *
 * 3. Le repli — les fichiers d'hier continuent de passer              *
 * ------------------------------------------------------------------ */
verifier(
    'Trois en-têtes non reconnues : repli sur les positions A/B/C',
    { perm: 0, name: 1, desc: 2, byPosition: true, usable: true },
    plan(['Lien du lieu', 'Libellé', 'Texte'])
);

verifier(
    'Tout ou rien : une seule en-tête reconnue ne suffit pas à lire par nom',
    { perm: 0, name: 1, desc: 2, byPosition: true, usable: true },
    plan(['POI Permalink', 'Libellé', 'Texte'])
);

verifier(
    'Des en-têtes numériques restent des en-têtes non vides',
    { perm: 0, name: 1, desc: 2, byPosition: true, usable: true },
    plan([1, 2, 3])
);

/* ------------------------------------------------------------------ *
 * 4. Ce qui reste refusé, comme avant                                 *
 * ------------------------------------------------------------------ */
verifier(
    'Une troisième en-tête vide : onglet ignoré',
    { perm: 0, name: 1, desc: 2, byPosition: true, usable: false },
    plan(['POI Permalink', 'POI Name', ''])
);

verifier(
    'Une ligne d’en-tête vide : onglet ignoré',
    { perm: 0, name: 1, desc: 2, byPosition: true, usable: false },
    plan([])
);

verifier(
    'Pas de ligne d’en-tête du tout : onglet ignoré',
    { perm: 0, name: 1, desc: 2, byPosition: true, usable: false },
    plan(undefined)
);

verifier(
    'Des en-têtes faites d’espaces ne sont pas des en-têtes',
    { perm: 0, name: 1, desc: 2, byPosition: true, usable: false },
    plan(['   ', ' ', '  '])
);

/* ------------------------------------------------------------------ *
 * 5. La normalisation elle-même                                       *
 * ------------------------------------------------------------------ */
verifier('Espaces multiples réduits', 'poi name', normalizeHeader(' POI   Name '));
verifier('Une cellule absente donne une chaîne vide', '', normalizeHeader(undefined));
// ⚠️ L'espace INSÉCABLE arrive par copier-coller depuis une page web ou un
//    tableur : sans lui, « POI Name » collé d'un site ne serait pas reconnu.
verifier('Espace insécable traité comme une espace', 'poi name', normalizeHeader('POI\u00a0Name'));
/* ⚠️ Aucun libellé reconnu ne porte d'accent aujourd'hui (« permalien », « nom »,
   « description » n'en ont pas) : la normalisation des accents ne sert donc pas
   ENCORE au repérage des colonnes. Elle sera nécessaire au lot D2, dont les
   libellés en porteront (« Catégories », « Téléphone »). Ce cas l'éprouve pour
   elle-même, faute de pouvoir l'éprouver par une colonne. */
verifier('Les accents sont retirés de la clé', 'categories', normalizeHeader('Catégories'));

/* ------------------------------------------------------------------ */
if (echecs.length === 0) {
    console.log(`=== ${reussis} réussis, 0 échec ===`);
    process.exit(0);
}
echecs.forEach(e => console.error(`  [ ÉCHEC ] ${e}`));
console.error(`=== ${reussis} réussis, ${echecs.length} échec(s) ===`);
process.exit(1);
