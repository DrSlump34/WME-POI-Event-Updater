/**
 * Banc du cochage — `node tools/banc-cochage.mjs`
 *
 * ⭐⭐⭐⭐ CE BANC EXISTE PARCE QUE LA RÈGLE N'ÉTAIT NULLE PART MESURABLE. Elle
 *    vivait dans le DOM de l'aperçu, à l'intérieur de la fonction qui construit
 *    les lignes — et les deux défauts qu'elle a connus n'ont été vus que dans
 *    l'éditeur, par l'œil :
 *      · l'aperçu annonçait « Aucune modification » sur un parking qui avait
 *        quatre champs à poser : l'indicateur ne regardait que le nom et la
 *        description. On ferme la fenêtre en confiance, et rien n'est appliqué.
 *      · une ligne qui RETIRE quelque chose était cochée d'office : une liste
 *        réduite à « Espèces » aurait supprimé « Carte de crédit » d'un parking,
 *        et le caractère destructeur n'apparaissait qu'au survol.
 *
 * ⚠️ CE QUE CE BANC NE VOIT PAS : la case elle-même. Il éprouve la RÈGLE, pas
 *    son câblage — qu'un `cb.checked` cesse d'être posé lui échapperait. Le
 *    dernier mot reste à l'essai dans WME.
 */

import { charger } from './extraire.mjs';

const { cocherDOffice, champsQuiDifferent } =
    charger(['champs', 'colonnes', 'pose'], ['cocherDOffice', 'champsQuiDifferent']);

let reussis = 0;
const echecs = [];
const verifier = (intitule, attendu, obtenu) => {
    if (JSON.stringify(attendu) === JSON.stringify(obtenu)) { reussis++; return; }
    echecs.push(`${intitule}\n      attendu : ${JSON.stringify(attendu)}\n      obtenu  : ${JSON.stringify(obtenu)}`);
};

// ---------------------------------------------------------------------------
// Combien de champs diffèrent
// ---------------------------------------------------------------------------
const lieu = {
    name: 'Parking P1',
    description: 'Ancien texte',
    categoryAttributes: { PARKING_LOT: { parkingType: 'PUBLIC', paymentType: ['CASH'] } },
};

verifier('rien à poser ⇒ aucun champ ne diffère',
    0, champsQuiDifferent(lieu, null));

verifier('une valeur identique ne compte pas',
    0, champsQuiDifferent(lieu, { 'PARKING_LOT.parkingType': 'PUBLIC' }));

verifier('une valeur différente compte',
    1, champsQuiDifferent(lieu, { 'PARKING_LOT.parkingType': 'PRIVATE' }));

/* ⚠️ LE NOM ET LA DESCRIPTION ONT LEUR PROPRE COMPARAISON — celle des deux
   colonnes de l'aperçu. Les compter ici les ferait voir deux fois. */
verifier('le nom et la description sont exclus du compte',
    0, champsQuiDifferent(lieu, { name: 'Autre nom', description: 'Autre texte' }));

verifier('le nom exclu, mais pas ce qui l’accompagne',
    1, champsQuiDifferent(lieu, { name: 'Autre nom', 'PARKING_LOT.parkingType': 'PRIVATE' }));

/* ⚠️⚠️ UN LIEU NON CHARGÉ NE SE COMPARE À RIEN : on compte TOUT ce qu'il y a à
   poser. Conclure « rien à faire » d'une absence de mesure serait le pire des
   cas — l'aperçu dirait que tout est en ordre sans avoir rien regardé. */
verifier('lieu non chargé ⇒ tous les champs comptent',
    2, champsQuiDifferent(null, { 'PARKING_LOT.parkingType': 'PUBLIC', services: ['WI_FI'] }));

verifier('lieu non chargé ⇒ le nom reste exclu',
    1, champsQuiDifferent(null, { name: 'Parking P1', services: ['WI_FI'] }));

// ---------------------------------------------------------------------------
// La règle du cochage
// ---------------------------------------------------------------------------
verifier('rien ne change ⇒ non cochée',
    false, cocherDOffice(false, false, 0, 0));

verifier('le nom change ⇒ cochée',
    true, cocherDOffice(true, false, 0, 0));

verifier('la description change ⇒ cochée',
    true, cocherDOffice(false, true, 0, 0));

/* ⭐ LE DÉFAUT VU DANS WME : seuls des champs du lot D2 changent, et l'aperçu
   annonçait « Aucune modification ». */
verifier('un champ du lot D2 seul ⇒ cochée',
    true, cocherDOffice(false, false, 4, 0));

/* ⭐⭐⭐⭐ PROPOSER UNE PERTE DEMANDE UN GESTE, JAMAIS UN DÉFAUT. */
verifier('une ligne qui RETIRE ⇒ NON cochée, même si elle change des choses',
    false, cocherDOffice(true, true, 4, 1));

verifier('perte seule, sans différence ⇒ non cochée',
    false, cocherDOffice(false, false, 0, 1));

verifier('plusieurs pertes ⇒ non cochée',
    false, cocherDOffice(false, false, 3, 2));

// ---------------------------------------------------------------------------
console.log(echecs.length
    ? `\n${echecs.map((e) => `  ✖ ${e}`).join('\n')}\n`
    : '');
console.log(`=== ${reussis} réussis, ${echecs.length} échec${echecs.length > 1 ? 's' : ''} ===`);
process.exit(echecs.length ? 1 : 0);
