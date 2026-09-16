/**
 * Banc de la geometrie — `node tools/banc-fenetre.mjs`
 *
 * ⭐⭐⭐ ON MESURE, ON NE SUPPOSE PAS. Un `calc(100vh - 110px)` suppose la hauteur
 *    du bandeau de WME, qui change avec la version, la langue, et la presence
 *    d'un autre script. Les bornes viennent donc de la carte elle-meme —
 *    `bornesCarte()` les releve dans le DOM et n'est pas mesurable ici. Ce que
 *    ce banc tient, c'est le CALCUL qui les respecte.
 *
 * ⚠️ CE QU'IL NE VOIT PAS : que la fenetre soit visible, que la poignee reponde,
 *    que le bouton de carte survive a un re-rendu de WME. Cela ne se voit que
 *    dans l'editeur.
 */

import { charger } from './extraire.mjs';

const { bornerFenetre, positionParDefaut } =
    charger(['geometrie'], ['bornerFenetre', 'positionParDefaut']);

let reussis = 0;
const echecs = [];
const verifier = (intitule, attendu, obtenu) => {
    if (JSON.stringify(attendu) === JSON.stringify(obtenu)) { reussis++; return; }
    echecs.push(`${intitule}\n      attendu : ${JSON.stringify(attendu)}\n      obtenu  : ${JSON.stringify(obtenu)}`);
};

/* Une carte de 1000 x 700, boutons a droite, pied en bas. */
const bornes = { gauche: 10, haut: 50, droite: 1000, bas: 700 };

// ---------------------------------------------------------------------------
// Une fenetre qui tient reste ou elle est
// ---------------------------------------------------------------------------
verifier('une fenetre qui tient n’est pas deplacee',
    { x: 400, y: 100, w: 500, h: 400, tenait: true },
    bornerFenetre({ x: 400, y: 100, w: 500, h: 400 }, bornes));

// ---------------------------------------------------------------------------
// Ce qui depasse rentre, et le dit
// ---------------------------------------------------------------------------
/* ⚠️ `tenait` EST LE POINT : rabattre une fenetre sans le dire donne
   l’impression qu’elle a disparu. L’appelant doit pouvoir faire la difference
   entre « je la repose ou elle etait » et « je l’ai ramenee de force ». */
verifier('trop a droite ⇒ ramenee, et ça se sait',
    { x: 500, y: 100, w: 500, h: 400, tenait: false },
    bornerFenetre({ x: 900, y: 100, w: 500, h: 400 }, bornes));

verifier('trop en bas ⇒ ramenee',
    { x: 400, y: 300, w: 500, h: 400, tenait: false },
    bornerFenetre({ x: 400, y: 900, w: 500, h: 400 }, bornes));

verifier('trop a gauche ⇒ collee au bord de la carte',
    { x: 10, y: 50, w: 500, h: 400, tenait: false },
    bornerFenetre({ x: -200, y: -80, w: 500, h: 400 }, bornes));

// ---------------------------------------------------------------------------
// Une fenetre plus grande que la carte
// ---------------------------------------------------------------------------
/* ⚠️ ELLE NE DEBORDE PAS : sinon son PIED sort de l’ecran, et le pied est ce
   qui porte le bouton qui applique. */
verifier('plus large que la carte ⇒ rabotee a la carte',
    { x: 10, y: 50, w: 990, h: 650, tenait: false },
    bornerFenetre({ x: 10, y: 50, w: 5000, h: 5000 }, bornes));

// ---------------------------------------------------------------------------
// Les planchers
// ---------------------------------------------------------------------------
/* ⚠️ UNE FENETRE DE 20 px N’EST PLUS UNE FENETRE : la poignee de
   redimensionnement peut la reduire a rien, et on ne sait plus la rattraper. */
verifier('une largeur derisoire remonte au plancher',
    280, bornerFenetre({ x: 10, y: 50, w: 5, h: 400 }, bornes).w);
verifier('une hauteur derisoire remonte au plancher',
    120, bornerFenetre({ x: 10, y: 50, w: 500, h: 5 }, bornes).h);

/* ⚠️ Et le plancher ne fait pas sortir la fenetre : sur une carte minuscule,
   c’est la carte qui gagne. */
const minuscule = { gauche: 0, haut: 0, droite: 100, bas: 90 };
verifier('sur une carte minuscule, la carte l’emporte sur le plancher',
    { x: 0, y: 0, w: 100, h: 90, tenait: false },
    bornerFenetre({ x: 0, y: 0, w: 500, h: 400 }, minuscule));

// ---------------------------------------------------------------------------
// La place par defaut
// ---------------------------------------------------------------------------
/* ⭐ A GAUCHE DES BOUTONS DE CARTE, jamais dessus : c’est la colonne ou vit le
   bouton qui ouvre cette fenetre. Le masquer reviendrait a cacher la poignee de
   la porte qu’on vient de franchir. */
verifier('par defaut, elle se colle au bord droit de la zone permise',
    { x: 180, y: 50, w: 820, h: 650 },
    positionParDefaut(bornes, 820));

verifier('par defaut, elle prend toute la hauteur utile',
    650, positionParDefaut(bornes, 820).h);

verifier('une largeur demandee plus grande que la carte se reduit',
    { x: 10, y: 50, w: 990, h: 650 },
    positionParDefaut(bornes, 3000));

// ---------------------------------------------------------------------------
console.log(echecs.length ? `\n${echecs.map((e) => `  ✖ ${e}`).join('\n')}\n` : '');
console.log(`=== ${reussis} réussis, ${echecs.length} échec${echecs.length > 1 ? 's' : ''} ===`);
process.exit(echecs.length ? 1 : 0);
