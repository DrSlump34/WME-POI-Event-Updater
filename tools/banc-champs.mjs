/**
 * Banc de la conversion vers les valeurs de WME — `node tools/banc-champs.mjs`
 *
 * ⭐⭐⭐⭐ CE BANC EST LE GARDE-FOU DU GARDE-FOU. Le SDK de WME accepte SANS ERREUR
 *    une valeur hors énumération et la POSE telle quelle (relevé du 15/09/2026,
 *    parkings de la gare TGV d'Avignon) : une conversion fausse ne provoque
 *    aucun refus, marque le lieu modifié et fait enregistrer une valeur que WME
 *    ne reconnaît pas. Rien, dans l'éditeur, ne le signalerait.
 *
 * Il éprouve donc DEUX choses :
 *   1. que chaque libellé produit par l'extranet EVIDRA trouve sa clé WME ;
 *   2. que les clés obtenues sont EXACTEMENT celles relevées dans l'éditeur.
 *
 * L'étalon des clés est figé ci-dessous : il vient du relevé, pas d'une
 * documentation. Le jour où WME en change, ce banc doit échouer — c'est son rôle.
 */

import { charger } from './extraire.mjs';

const { cleWme, clesWme, VALEURS_WME, chargerCategories } =
    charger(['champs', 'colonnes'], ['cleWme', 'clesWme', 'VALEURS_WME', 'chargerCategories']);

let reussis = 0;
const echecs = [];
const verifier = (intitule, attendu, obtenu) => {
    const a = JSON.stringify(attendu), o = JSON.stringify(obtenu);
    if (a === o) { reussis++; return; }
    echecs.push(`${intitule}\n      attendu : ${a}\n      obtenu  : ${o}`);
};

/* ------------------------------------------------------------------ *
 * 1. CE QUE L'EXTRANET IMPRIME → CE QUE WME ATTEND                    *
 *    Libellés repris de Poi::CARACTERISTIQUES, clés du relevé WME.    *
 * ------------------------------------------------------------------ */
const CORRESPONDANCES = {
    parkingType: [['Public', 'PUBLIC'], ['Privé', 'PRIVATE'], ['Restreint', 'RESTRICTED']],
    costType: [['Gratuit', 'FREE'], ['Faible', 'LOW'], ['Modéré', 'MODERATE'], ['Élevé', 'EXPENSIVE']],
    estimatedNumberOfSpots: [
        ['1-10', 'R_1_TO_10'], ['11-30', 'R_11_TO_30'], ['31-60', 'R_31_TO_60'],
        ['61-100', 'R_61_TO_100'], ['101-300', 'R_101_TO_300'], ['301-600', 'R_301_TO_600'],
        ['> 600', 'R_600_PLUS']
    ],
    lotType: [
        ['Plusieurs niveaux', 'MULTI_LEVEL'], ['Extérieur', 'STREET_LEVEL'],
        ['Extérieur couvert', 'STREET_LEVEL_COVERED'], ['Souterrain', 'UNDERGROUND']
    ],
    paymentType: [
        ['Espèces', 'CASH'], ['Chèques', 'CHECKS'], ['Carte de crédit', 'CREDIT'],
        ['Carte bancaire', 'DEBIT_CARD'], ['Portefeuille numérique', 'DIGITAL_WALLET'],
        ['Pass électronique', 'ELECTRONIC_PASS'], ['Abonnement', 'MEMBERSHIP'],
        ['Application', 'PARKING_APP'], ['Laissez-passer', 'PERMIT'],
        ['Prépaiement', 'PREPAID'], ['SMS/Appel', 'SMS_CALL']
    ],
    services: [
        ['Climatisation', 'AIR_CONDITIONING'], ['Accepte les cartes de crédit', 'CREDIT_CARDS'],
        ['Click & Collect', 'CURBSIDE_PICKUP'], ['Livraisons', 'DELIVERIES'],
        ['Drive', 'DRIVETHROUGH'], ['Terrasse extérieure', 'OUTSIDE_SEATING'],
        ['Parking client', 'PARKING_FOR_CUSTOMERS'], ['Réservations', 'RESERVATIONS'],
        ['Toilettes', 'RESTROOMS'], ['À emporter', 'TAKE_AWAY'],
        ['Service de voiturier', 'VALLET_SERVICE'],
        ['Accessible en fauteuil roulant', 'WHEELCHAIR_ACCESSIBLE'], ['Wi-Fi', 'WI_FI']
    ],
    parkingServices: [
        ['Navette aéroport', 'AIRPORT_SHUTTLE'], ['Places covoiturage', 'CARPOOL_PARKING'],
        ['Lavage auto', 'CAR_WASH'], ['Couvert', 'COVERED'], ['Places PMR', 'DISABILITY_PARKING'],
        ["Agent d'accueil", 'ON_SITE_ATTENDANT'], ['P+R', 'PARK_AND_RIDE'],
        ['Réservations', 'RESERVATIONS'], ['Surveillance', 'SECURITY'],
        ['Voiturier', 'VALET'], ['Service de voiturier', 'VALLET_SERVICE'],
        ['Bornes de charge', 'EV_CHARGING_STATION']
    ]
};

for (const [referentiel, paires] of Object.entries(CORRESPONDANCES)) {
    paires.forEach(([libelle, cle]) => {
        verifier(`${referentiel} : « ${libelle} »`, cle, cleWme(referentiel, libelle));
    });
    /* Aucune valeur du référentiel ne doit manquer, ni s'y ajouter en douce. */
    verifier(
        `${referentiel} : le référentiel a exactement les clés relevées dans WME`,
        paires.map(p => p[1]).sort(),
        Object.keys(VALEURS_WME[referentiel]).sort()
    );
}

/* ------------------------------------------------------------------ *
 * 2. LE PIÈGE : « voiturier » ne désigne pas la même valeur           *
 * ------------------------------------------------------------------ */
verifier('« Voiturier » sur un PARKING vaut VALET', 'VALET', cleWme('parkingServices', 'Voiturier'));
verifier('« Voiturier » n’existe pas dans les services d’un LIEU', null, cleWme('services', 'Voiturier'));
verifier('« Service de voiturier » vaut VALLET_SERVICE des deux côtés',
    ['VALLET_SERVICE', 'VALLET_SERVICE'],
    [cleWme('services', 'Service de voiturier'), cleWme('parkingServices', 'Service de voiturier')]);

/* ------------------------------------------------------------------ *
 * 3. CE QUI DOIT ÊTRE REFUSÉ — et jamais posé                         *
 * ------------------------------------------------------------------ */
verifier('Une clé française brute est REFUSÉE', null, cleWme('costType', 'gratuit_'));
verifier('Une valeur inventée est REFUSÉE', null, cleWme('lotType', 'Souterrain profond'));
verifier('Un référentiel inconnu rend null', null, cleWme('nExistePas', 'Public'));
verifier('Une cellule vide ne rend rien', null, cleWme('costType', ''));

/* ------------------------------------------------------------------ *
 * 4. TOLÉRANCES : casse, espaces, accents, et la clé WME elle-même    *
 * ------------------------------------------------------------------ */
verifier('La casse ne compte pas', 'CASH', cleWme('paymentType', 'ESPÈCES'));
verifier('Les accents non plus', 'MODERATE', cleWme('costType', 'modere'));
verifier('Les espaces en trop non plus', 'PUBLIC', cleWme('parkingType', '  Public  '));
verifier('La clé WME est acceptée telle quelle', 'CREDIT', cleWme('paymentType', 'CREDIT'));

/* ------------------------------------------------------------------ *
 * 5. PLUSIEURS VALEURS DANS UNE CELLULE — le format de l'export        *
 * ------------------------------------------------------------------ */
verifier('Séparées par « ; », comme l’export EVIDRA',
    { retenues: ['CASH', 'CREDIT'], refusees: [] },
    clesWme('paymentType', 'Espèces ; Carte de crédit'));
verifier('Une par ligne, comme la base',
    { retenues: ['CASH', 'CREDIT'], refusees: [] },
    clesWme('paymentType', 'Espèces\nCarte de crédit'));
verifier('Les doublons disparaissent',
    { retenues: ['CASH'], refusees: [] },
    clesWme('paymentType', 'Espèces ; Espèces ; CASH'));
verifier('⭐ UNE VALEUR NON RECONNUE EST RENDUE À PART, jamais tue',
    { retenues: ['CASH'], refusees: ['Bitcoin'] },
    clesWme('paymentType', 'Espèces ; Bitcoin'));
verifier('Une cellule vide ne rend rien du tout',
    { retenues: [], refusees: [] },
    clesWme('paymentType', ''));


/* ------------------------------------------------------------------ *
 * 6. LES CATÉGORIES — un référentiel VIVANT, relevé dans l’éditeur     *
 * ------------------------------------------------------------------ */
/* ⭐ Elles ne sont pas écrites dans le script : WME les rend déjà traduites dans
   la langue de l’utilisateur (132 le 16/09/2026). Une copie figée serait fausse
   partout ailleurs qu’en français, et périmerait au premier ajout de Waze. */

verifier('Au repos le référentiel est VIDE', 0, Object.keys(VALEURS_WME.categories).length);
verifier('⚠️ Et une catégorie est donc REFUSÉE, jamais posée à l’aveugle',
    null, cleWme('categories', 'Parking'));

const faussSdk = (liste) => ({ DataModel: { Venues: { getAllVenueCategories: () => liste } } });
verifier('Le chargement remplit le référentiel', 3, chargerCategories(faussSdk([
    { id: 'PARKING_LOT', localizedName: 'Parking' },
    { id: 'RESTAURANT', localizedName: 'Restaurant' },
    { id: 'ZOO_AQUARIUM', localizedName: 'Zoo, aquarium' }
])));
verifier('Une catégorie se reconnaît à son libellé', 'PARKING_LOT', cleWme('categories', 'Parking'));
verifier('…à sa clé WME aussi', 'ZOO_AQUARIUM', cleWme('categories', 'ZOO_AQUARIUM'));
verifier('…sans se soucier de la casse', 'RESTAURANT', cleWme('categories', 'RESTAURANT'));
verifier('Une catégorie inconnue reste refusée', null, cleWme('categories', 'Château fort'));
verifier('Plusieurs catégories, et le refus rendu à part',
    { retenues: ['PARKING_LOT', 'RESTAURANT'], refusees: ['Château fort'] },
    clesWme('categories', 'Parking ; Restaurant ; Château fort'));

/* ⭐ Le référentiel ne se recharge pas à chaque fichier : 132 entrées relues à
   chaque ouverture seraient du travail pour rien, et le SDK peut être lent. */
let appels = 0;
const sdkCompteur = { DataModel: { Venues: { getAllVenueCategories: () => { appels++; return [{ id: 'PARKING_LOT', localizedName: 'Parking' }]; } } } };
Object.keys(VALEURS_WME.categories).forEach(k => delete VALEURS_WME.categories[k]);
chargerCategories(sdkCompteur);
chargerCategories(sdkCompteur);
chargerCategories(sdkCompteur);
verifier('Trois chargements, un seul appel au SDK', 1, appels);

/* Un SDK muet ne doit rien inventer. */
Object.keys(VALEURS_WME.categories).forEach(k => delete VALEURS_WME.categories[k]);
verifier('Un SDK qui ne rend rien laisse le référentiel vide', 0,
    chargerCategories(faussSdk(null)));

/* ------------------------------------------------------------------ */
if (echecs.length === 0) {
    console.log(`=== ${reussis} réussis, 0 échec ===`);
    process.exit(0);
}
echecs.forEach(e => console.error(`  [ ÉCHEC ] ${e}`));
console.error(`=== ${reussis} réussis, ${echecs.length} échec(s) ===`);
process.exit(1);
