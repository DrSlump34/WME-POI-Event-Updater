/**
 * Banc de la lecture d'une ligne — `node tools/banc-valeurs.mjs`
 *
 * Il éprouve `lireValeurs()` : ce qu'une ligne du classeur demande, trié en
 * trois tas — à poser, montré seulement, refusé.
 *
 * ⭐ LES DEUX RÈGLES QU'IL TIENT, et qui ne se voient dans aucun écran :
 *    · une cellule VIDE ne demande rien et n'efface rien ;
 *    · les deux colonnes de services visent le MÊME attribut de WME, et comme un
 *      tableau y remplace tout, c'est leur UNION qu'il faut envoyer.
 */

import { charger } from './extraire.mjs';

const { CHAMPS, mapChamps, lireValeurs } =
    charger(['champs', 'colonnes'], ['CHAMPS', 'mapChamps', 'lireValeurs']);

let reussis = 0;
const echecs = [];
const stable = (v) => {
    if (Array.isArray(v)) return v.map(stable);
    if (v && typeof v === 'object') {
        return Object.keys(v).sort().reduce((a, k) => { a[k] = stable(v[k]); return a; }, {});
    }
    return v;
};
const verifier = (intitule, attendu, obtenu) => {
    const a = JSON.stringify(stable(attendu)), o = JSON.stringify(stable(obtenu));
    if (a === o) { reussis++; return; }
    echecs.push(`${intitule}\n      attendu : ${a}\n      obtenu  : ${o}`);
};

/** Lit une ligne décrite par ses en-têtes, comme le ferait le script. */
function lire(entetes, cells) {
    return lireValeurs(cells, mapChamps(entetes, false));
}

/* ------------------------------------------------------------------ *
 * 1. LE CAS ORDINAIRE                                                 *
 * ------------------------------------------------------------------ */
let r = lire(
    ['POI Permalink', 'POI Name', 'POI Description', 'Parking Cost', 'Parking Type'],
    ['https://www.waze.com/editor?venues=1', 'Parking P2', 'Longue durée', 'Modéré', 'Public']
);
verifier('Les valeurs simples et converties sont à poser',
    { name: 'Parking P2', description: 'Longue durée',
      'PARKING_LOT.costType': 'MODERATE', 'PARKING_LOT.parkingType': 'PUBLIC' }, r.aPoser);
verifier('Rien à montrer, rien de refusé', [0, 0], [r.montres.length, r.refus.length]);

/* ------------------------------------------------------------------ *
 * 2. ⭐ UNE CELLULE VIDE NE DEMANDE RIEN — la règle qui protège la carte *
 * ------------------------------------------------------------------ */
r = lire(['POI Permalink', 'POI Name', 'Parking Cost', 'Services'],
         ['https://www.waze.com/editor?venues=1', 'Parking P2', '', '   ']);
verifier('Une cellule vide n’est pas posée', { name: 'Parking P2' }, r.aPoser);
verifier('Une cellule d’espaces non plus, et ce n’est pas un refus', 0, r.refus.length);

/* ------------------------------------------------------------------ *
 * 3. ⚠️ LES DEUX COLONNES DE SERVICES VISENT LE MÊME ATTRIBUT          *
 * ------------------------------------------------------------------ */
r = lire(['Services', 'Parking Services'], ['Wi-Fi ; Toilettes', 'Surveillance ; Places PMR']);
verifier('Leurs listes sont FUSIONNÉES, pas concurrentes',
    { services: ['WI_FI', 'RESTROOMS', 'SECURITY', 'DISABILITY_PARKING'] }, r.aPoser);

r = lire(['Services', 'Parking Services'], ['Réservations', 'Réservations ; Surveillance']);
verifier('Une valeur commune aux deux colonnes n’est comptée qu’une fois',
    { services: ['RESERVATIONS', 'SECURITY'] }, r.aPoser);

r = lire(['Services', 'Parking Services'], ['', 'Voiturier']);
verifier('⭐ « Voiturier » côté PARKING vaut VALET, jamais VALLET_SERVICE',
    { services: ['VALET'] }, r.aPoser);

/* ------------------------------------------------------------------ *
 * 4. CE QUI EST MONTRÉ, ET JAMAIS POSÉ                                *
 * ------------------------------------------------------------------ */
r = lire(['Horaires', 'Adresse', 'Opérateur de parking'],
         ['8h-20h sauf dimanche', '12 rue des Lilas', 'Indigo']);
verifier('Aucun champ structuré n’est posé', {}, r.aPoser);
verifier('Les trois sont montrés, avec leur libellé',
    ['Horaires', 'Adresse', 'Opérateur de parking'], r.montres.map(m => m.libelle));
verifier('Chaque valeur montrée porte son motif', [true, true, true],
    r.montres.map(m => typeof m.motif === 'string' && m.motif.length > 0));
verifier('Ce qui est montré n’est pas un refus', 0, r.refus.length);

/* ------------------------------------------------------------------ *
 * 5. LES REFUS — jamais tus, jamais posés                             *
 * ------------------------------------------------------------------ */
r = lire(['Parking Cost', 'Parking Payment'], ['Pas cher', 'Espèces ; Bitcoin']);
verifier('Une valeur simple hors référentiel est refusée, et rien n’est posé pour elle',
    [{ libelle: 'Tarif', valeurs: ['Pas cher'] }],
    r.refus.filter(x => x.libelle === 'Tarif'));
verifier('Dans une liste, seule la valeur inconnue est refusée',
    [{ libelle: 'Modes de paiement', valeurs: ['Bitcoin'] }],
    r.refus.filter(x => x.libelle === 'Modes de paiement'));
verifier('⭐ ET LE RESTE DE LA LISTE EST POSÉ QUAND MÊME',
    { 'PARKING_LOT.paymentType': ['CASH'] }, r.aPoser);

/* ------------------------------------------------------------------ *
 * 6. LES BOOLÉENS, DANS LES DEUX LANGUES                              *
 * ------------------------------------------------------------------ */
verifier('« Oui » vaut vrai',
    { 'PARKING_LOT.canExitWhileClosed': true },
    lire(['Parking Exit When Closed'], ['Oui']).aPoser);
verifier('« No » vaut faux',
    { 'PARKING_LOT.canExitWhileClosed': false },
    lire(['Parking Exit When Closed'], ['No']).aPoser);
verifier('« Peut-être » est refusé, pas interprété',
    [{ libelle: 'Sortie quand fermé', valeurs: ['Peut-être'] }],
    lire(['Parking Exit When Closed'], ['Peut-être']).refus);

/* ------------------------------------------------------------------ *
 * 7. LES LISTES LIBRES (sans référentiel)                             *
 * ------------------------------------------------------------------ */
verifier('Les noms alternatifs se découpent, sans conversion',
    { aliases: ['Gare TGV', 'Avignon TGV'] },
    lire(['Noms alternatifs'], ['Gare TGV ; Avignon TGV']).aPoser);

/* ------------------------------------------------------------------ *
 * 8. LE PERMALIEN N'EST PAS UNE VALEUR À POSER                        *
 * ------------------------------------------------------------------ */
r = lire(['POI Permalink', 'POI Name'], ['https://www.waze.com/editor?venues=1', 'X']);
verifier('Le permalien désigne le lieu, il ne s’écrit pas dedans',
    { name: 'X' }, r.aPoser);
verifier('Et il n’est pas « montré » non plus', 0, r.montres.length);

/* ------------------------------------------------------------------ *
 * 9. COHÉRENCE DE LA TABLE ELLE-MÊME                                  *
 * ------------------------------------------------------------------ */
verifier('Tout champ posé a une cible', [],
    CHAMPS.filter(c => c.pose && !c.cible).map(c => c.cle));
verifier('Tout champ montré a un motif', [],
    CHAMPS.filter(c => !c.pose && !c.identifie && !c.motif).map(c => c.cle));
verifier('Tout référentiel cité existe', [],
    CHAMPS.filter(c => c.referentiel && !['services', 'parkingServices', 'costType', 'parkingType',
        'paymentType', 'lotType', 'estimatedNumberOfSpots'].includes(c.referentiel)).map(c => c.cle));

/* ------------------------------------------------------------------ */
if (echecs.length === 0) {
    console.log(`=== ${reussis} réussis, 0 échec ===`);
    process.exit(0);
}
echecs.forEach(e => console.error(`  [ ÉCHEC ] ${e}`));
console.error(`=== ${reussis} réussis, ${echecs.length} échec(s) ===`);
process.exit(1);
