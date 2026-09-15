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

import { charger } from './extraire.mjs';

const { mapColumns, normalizeHeader, mapChamps, CHAMPS } =
    charger(['champs', 'colonnes'], ['mapColumns', 'normalizeHeader', 'mapChamps', 'CHAMPS']);

let reussis = 0;
const echecs = [];

/* ⚠️ L'ORDRE DES CLÉS D'UN OBJET N'EST PAS UNE PROPRIÉTÉ À ÉPROUVER : `mapChamps`
   les insère dans l'ordre de `CHAMPS`, ce qui est un détail d'implémentation.
   Comparer les objets tels quels faisait échouer un contrôle pour cette seule
   raison — un banc qui signale un défaut inexistant ne vaut pas mieux qu'un
   banc aveugle. L'ordre des TABLEAUX, lui, est conservé : il porte du sens. */
const stable = (v) => {
    if (Array.isArray(v)) return v.map(stable);
    if (v && typeof v === 'object') {
        return Object.keys(v).sort().reduce((acc, k) => { acc[k] = stable(v[k]); return acc; }, {});
    }
    return v;
};

function verifier(intitule, attendu, obtenu) {
    const a = JSON.stringify(stable(attendu));
    const o = JSON.stringify(stable(obtenu));
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
verifier('Les accents sont retirés de la clé', 'categories', normalizeHeader('Catégories'));

/* ------------------------------------------------------------------ *
 * 6. LES COLONNES DU LOT D2 — au-delà des trois de base               *
 * ------------------------------------------------------------------ */
const plus = (entetes, byPosition = false) => mapChamps(entetes, byPosition);

verifier(
    'Les colonnes du parking sont repérées par leur nom',
    { perm: 0, name: 1, desc: 2, costType: 3, paymentType: 4 },
    plus(['POI Permalink', 'POI Name', 'POI Description', 'Parking Cost', 'Parking Payment'])
);

/* ⭐ C'EST ICI que la normalisation des accents sert enfin pour de vrai : ces
   libellés-là en portent, contrairement aux trois colonnes d'origine. */
verifier(
    'Une en-tête accentuée est reconnue, écrite avec ou sans ses accents',
    { phone: 0, categories: 1, operator: 2 },
    plus(['Téléphone', 'Catégories', 'Opérateur de parking'])
);
verifier(
    'La même, tapée sans accents',
    { phone: 0, categories: 1, operator: 2 },
    plus(['Telephone', 'Categories', 'Operateur de parking'])
);

verifier(
    'Une colonne inconnue n’est pas une erreur : elle est simplement absente',
    { perm: 0, name: 1 },
    plus(['POI Permalink', 'POI Name', 'Ma référence interne', 'Couleur'])
);

verifier(
    '⚠️ En REPLI, on ne repère aucune colonne de plus — on n’invente pas',
    {},
    plus(['Lien du lieu', 'Libellé', 'Texte', 'Parking Cost'], true)
);

verifier(
    'Les libellés anglais du gabarit sont reconnus',
    { services: 0, lotType: 1, spots: 2, canExit: 3 },
    plus(['Services', 'Parking Situation', 'Parking Spots', 'Parking Exit When Closed'])
);

/* La même règle que pour les trois colonnes de base : en cas de doublon, la
   PREMIÈRE gagne. Sans ce cas, rien n'empêcherait la dernière de l'emporter —
   et deux colonnes « Tarif » se liraient à l'envers, sans un mot. */
verifier(
    'Une colonne en double : la PREMIÈRE gagne',
    { costType: 0, name: 2 },
    plus(['Parking Cost', 'Parking Cost', 'POI Name'])
);

/* Aucun champ ne doit partager un libellé avec un autre : la colonne irait au
   premier déclaré, en silence. */
const tousLesLibelles = CHAMPS.flatMap(c => c.entetes);
verifier('Aucun libellé n’est revendiqué par deux champs', [],
    tousLesLibelles.filter((e, i) => tousLesLibelles.indexOf(e) !== i));

/* ------------------------------------------------------------------ */
if (echecs.length === 0) {
    console.log(`=== ${reussis} réussis, 0 échec ===`);
    process.exit(0);
}
echecs.forEach(e => console.error(`  [ ÉCHEC ] ${e}`));
console.error(`=== ${reussis} réussis, ${echecs.length} échec(s) ===`);
process.exit(1);
