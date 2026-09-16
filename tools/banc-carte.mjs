/**
 * Banc du cadrage — `node tools/banc-carte.mjs`
 *
 * ⭐⭐⭐ CE QUI SE CALCULE SE MESURE ICI ; ce qui parle à la carte, jamais. La vue
 *    d'ensemble d'un périmètre est une UNION DE BOÎTES — une opération pure, qui
 *    n'a besoin ni d'OpenLayers ni de WME. `cadrerSurLesLieux()`, qui appelle
 *    `W.map`, reste hors de portée : ce banc éprouve la règle, pas le geste.
 *
 * ⚠️ LE CAS QUI COMPTE EST LA BOÎTE INVALIDE. Un seul `NaN` contamine `Math.min`
 *    et l'englobant entier devient `NaN` : la carte se cadrerait nulle part,
 *    sans la moindre erreur. On écarte, on ne propage pas.
 */

import { charger, source } from './extraire.mjs';

const { unionDesBoites, boiteSansEtendue } =
    charger(['carte'], ['unionDesBoites', 'boiteSansEtendue']);

let reussis = 0;
const echecs = [];
const verifier = (intitule, attendu, obtenu) => {
    if (JSON.stringify(attendu) === JSON.stringify(obtenu)) { reussis++; return; }
    echecs.push(`${intitule}\n      attendu : ${JSON.stringify(attendu)}\n      obtenu  : ${JSON.stringify(obtenu)}`);
};

const boite = (left, bottom, right, top) => ({ left, bottom, right, top });

// ---------------------------------------------------------------------------
// L'englobant
// ---------------------------------------------------------------------------
verifier('aucune boîte ⇒ rien à cadrer',
    null, unionDesBoites([]));

verifier('rien du tout ⇒ rien à cadrer',
    null, unionDesBoites(null));

verifier('une seule boîte ⇒ elle-même',
    boite(1, 2, 3, 4), unionDesBoites([boite(1, 2, 3, 4)]));

verifier('deux boîtes disjointes ⇒ elles tiennent toutes deux dedans',
    boite(0, 0, 10, 10), unionDesBoites([boite(0, 0, 2, 2), boite(8, 8, 10, 10)]));

verifier('une boîte incluse dans l’autre ne l’agrandit pas',
    boite(0, 0, 10, 10), unionDesBoites([boite(0, 0, 10, 10), boite(3, 3, 4, 4)]));

verifier('l’ordre des boîtes ne change rien',
    unionDesBoites([boite(8, 8, 10, 10), boite(0, 0, 2, 2)]),
    unionDesBoites([boite(0, 0, 2, 2), boite(8, 8, 10, 10)]));

verifier('des coordonnées négatives sont respectées',
    boite(-5, -4, 1, 2), unionDesBoites([boite(-5, -4, -1, -2), boite(0, 0, 1, 2)]));

/* ⚠️⚠️ UNE BOÎTE INVALIDE S'ÉCARTE, ELLE NE SE PROPAGE PAS. */
verifier('un NaN ne contamine pas l’englobant',
    boite(0, 0, 2, 2), unionDesBoites([boite(0, 0, 2, 2), boite(NaN, 0, 1, 1)]));

verifier('un null au milieu est ignoré',
    boite(0, 0, 10, 10), unionDesBoites([boite(0, 0, 2, 2), null, boite(8, 8, 10, 10)]));

verifier('une boîte sans coordonnées est ignorée',
    boite(0, 0, 2, 2), unionDesBoites([boite(0, 0, 2, 2), {}]));

verifier('un Infinity est écarté comme un NaN',
    boite(0, 0, 2, 2), unionDesBoites([boite(0, 0, 2, 2), boite(0, 0, Infinity, 1)]));

verifier('QUE des boîtes invalides ⇒ rien à cadrer, et non une boîte NaN',
    null, unionDesBoites([{}, null, boite(NaN, NaN, NaN, NaN)]));

// ---------------------------------------------------------------------------
// La boîte sans étendue
// ---------------------------------------------------------------------------
/* ⚠️ Demander à la carte de tenir dans un point la pousse à son zoom maximal :
   on se retrouve collé au sol, sans rien voir autour. */
verifier('un seul lieu ⇒ boîte sans étendue',
    true, boiteSansEtendue(boite(5, 5, 5, 5)));

verifier('deux lieux confondus ⇒ boîte sans étendue',
    true, boiteSansEtendue(unionDesBoites([boite(5, 5, 5, 5), boite(5, 5, 5, 5)])));

verifier('rien ⇒ traité comme sans étendue',
    true, boiteSansEtendue(null));

verifier('un vrai périmètre a une étendue',
    false, boiteSansEtendue(boite(0, 0, 10, 10)));

verifier('une étendue sur un seul axe compte quand même',
    false, boiteSansEtendue(boite(0, 5, 10, 5)));

// ---------------------------------------------------------------------------
// Le câblage, faute de pouvoir l'exécuter
// ---------------------------------------------------------------------------
/* ⚠️⚠️ CE QUI SUIT LIT DU TEXTE, ET LE TEXTE NE DIT PAS CE QUE LE CODE FAIT. On
 *    ne s'en contente que parce que ces trois règles vivent dans `showOverlay()`,
 *    au milieu du DOM et de `W.map` : aucun banc ne peut les exécuter. On isole
 *    donc le CORPS de la fonction — chercher « quelque part dans le fichier »
 *    trouverait la restauration légitime de l'annulation, et conclurait au vert
 *    sur un code qui revient en arrière.
 * ⇒ Le dernier mot reste à l'essai dans WME. */
const corps = (nom) => {
    const d = source.indexOf(`function ${nom}(`);
    if (d === -1) return null;
    const f = source.indexOf('\n    }', d);

    return f === -1 ? null : source.slice(d, f);
};

/* ⚠️⚠️ LA CIBLE A CHANGÉ DE NOM, ET LE CONTRÔLE A REFUSÉ DE CONCLURE plutôt que
   de passer au vert sur une fonction disparue. C'est ce qu'on lui demande : un
   contrôle qui ne trouve plus sa cible n'a rien mesuré, et doit le dire. */
const apercu = corps('ouvrirApercu');
if (!apercu) {
    echecs.push('ouvrirApercu() est introuvable : le contrôle du câblage ne mesure plus rien');
} else {
    /* ⭐⭐⭐ LA VUE NE REVIENT PLUS EN ARRIÈRE, JAMAIS. La restauration de
       l'annulation elle-même a disparu : ce qu'on demandait, c'est que la carte
       RESTE. Après un balayage interrompu, elle reste donc sur le dernier lieu
       lu, qui est encore un lieu du fichier — et non au point de départ, qui
       n'apprend rien. */
    verifier('aucune restauration de la vue ne subsiste',
        false, /setCenter\(savedCenter/.test(apercu));

    verifier('le périmètre est cadré au terme du balayage',
        true, apercu.includes('cadrerSurLesLieux(venueMap)'));
}

/* ⚠️ ET APRÈS L'APPLICATION AUSSI, qui déplace elle aussi la carte de lieu en
   lieu. Les deux appels vivent désormais dans deux fonctions distinctes : on les
   compte sur le fichier entier, mais on EXIGE les deux. */
/* ⚠️ ON COMPTE LES APPELS, PAS UN ARGUMENT PRECIS : les deux ne passent pas
   la meme variable, et exiger le meme texte faisait echouer un code correct. */
const cadragesTotal = (source.match(/cadrerSurLesLieux\(/g) || []).length - 1;
verifier('le cadrage a lieu après le balayage ET après l’application', 2, cadragesTotal);

const application = corps('appliquerLignes');
if (!application) {
    echecs.push('appliquerLignes() est introuvable : le contrôle ne mesure plus rien');
} else {
    verifier('l’application recadre sur le périmètre en finissant',
        true, application.includes('cadrerSurLesLieux(_apercu.venueMap)'));
}

/* ⚠️ UNE SEULE FENÊTRE, PAR CONSTRUCTION : `construireOverlay` rend celle qui
   existe au lieu d'en créer une seconde. C'est ce qui remplace le nettoyage
   d'aperçus empilés — deux tableaux superposés, et l'on appliquait depuis celui
   qui ne décrivait plus le fichier chargé. */
verifier('la fenêtre ne se construit qu’une fois',
    true, /function construireOverlay\(\)[\s\S]{0,200}if \(ov\) return ov;/.test(source));

// ---------------------------------------------------------------------------
console.log(echecs.length
    ? `\n${echecs.map((e) => `  ✖ ${e}`).join('\n')}\n`
    : '');
console.log(`=== ${reussis} réussis, ${echecs.length} échec${echecs.length > 1 ? 's' : ''} ===`);
process.exit(echecs.length ? 1 : 0);
