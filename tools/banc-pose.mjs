/**
 * Banc de l'écriture — `node tools/banc-pose.mjs`
 *
 * ⭐⭐⭐⭐ CE BANC EXISTE À CAUSE D'UN RELEVÉ, PAS D'UNE INTUITION. Le 15/09/2026,
 *    sur les parkings de la gare TGV d'Avignon :
 *      · `updateVenue({venueId, champInconnu: 1})` n'a levé AUCUNE erreur, n'a
 *        rien posé, et a quand même marqué le lieu comme modifié ;
 *      · `costType: 'gratuit'` — valeur hors énumération — a été POSÉE telle
 *        quelle, sans un mot.
 *    D'où la liste blanche (ce qu'on envoie) et la relecture (ce qui est arrivé).
 *
 * Il ne remplace pas l'essai dans WME : il ne voit ni le SDK, ni la carte.
 */

import { charger } from './extraire.mjs';

const { CHAMPS, mapChamps, lireValeurs, construireMaj, verifierPose } =
    charger(['champs', 'colonnes', 'pose'],
        ['CHAMPS', 'mapChamps', 'lireValeurs', 'construireMaj', 'verifierPose']);

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
const lire = (entetes, cells) => lireValeurs(cells, mapChamps(entetes, false));

/* ------------------------------------------------------------------ *
 * 1. CE QU'ON ENVOIE AU SDK                                           *
 * ------------------------------------------------------------------ */
let r = construireMaj({
    name: 'Parking P5', description: 'Longue durée',
    phone: '04 90 00 00 00', services: ['WI_FI', 'SECURITY'],
    'PARKING_LOT.costType': 'MODERATE', 'PARKING_LOT.paymentType': ['CASH']
});
verifier('Les champs du parking sont regroupés sous categoryAttributes.PARKING_LOT',
    { phone: '04 90 00 00 00', services: ['WI_FI', 'SECURITY'],
      categoryAttributes: { PARKING_LOT: { costType: 'MODERATE', paymentType: ['CASH'] } } },
    r.maj);
verifier('⭐ Le nom et la description NE PASSENT PAS par le SDK : ils ont leur mécanisme',
    [false, false], ['name' in r.maj, 'description' in r.maj]);
verifier('Rien n’est ignoré dans un cas normal', [], r.ignores);

verifier('Sans champ de parking, pas de categoryAttributes vide',
    false, 'categoryAttributes' in construireMaj({ phone: '01' }).maj);

/* ------------------------------------------------------------------ *
 * 2. ⚠️ LA LISTE BLANCHE — la parade à l'acceptation silencieuse       *
 * ------------------------------------------------------------------ */
r = construireMaj({ phone: '01', champInvente: 'x', 'PARKING_LOT.inconnu': 1, 'AUTRE_CAT.champ': 2 });
verifier('Un champ hors liste blanche n’est JAMAIS envoyé',
    { phone: '01' }, r.maj);
verifier('Et il est rendu, pas tu',
    ['champInvente', 'PARKING_LOT.inconnu', 'AUTRE_CAT.champ'].sort(), r.ignores.slice().sort());

/* ------------------------------------------------------------------ *
 * 3. LA RELECTURE — « appliqué » doit vouloir dire « présent »        *
 * ------------------------------------------------------------------ */
const apresEcriture = {
    name: 'Parking P5', phone: '04 90 00 00 00', services: ['SECURITY', 'WI_FI'],
    categoryAttributes: { PARKING_LOT: { costType: 'MODERATE', paymentType: ['CASH'] } }
};
r = verifierPose(apresEcriture, {
    name: 'Parking P5', phone: '04 90 00 00 00', services: ['WI_FI', 'SECURITY'],
    'PARKING_LOT.costType': 'MODERATE', 'PARKING_LOT.paymentType': ['CASH']
});
verifier('Tout ce qui est demandé est retrouvé — l’ordre d’une liste ne compte pas',
    { confirmes: ['phone', 'services', 'PARKING_LOT.costType', 'PARKING_LOT.paymentType'].sort(), manques: [] },
    { confirmes: r.confirmes.slice().sort(), manques: r.manques });

/* Le cas qui justifie tout : le SDK n'a rien posé, et n'a rien dit. */
r = verifierPose({ phone: '04 90 00 00 00', categoryAttributes: { PARKING_LOT: {} } }, {
    phone: '04 90 00 00 00', 'PARKING_LOT.costType': 'MODERATE'
});
verifier('🔴 UN CHAMP NON POSÉ EST DÉTECTÉ, même sans erreur du SDK',
    { confirmes: ['phone'], manques: ['PARKING_LOT.costType'] }, r);

r = verifierPose({ categoryAttributes: { PARKING_LOT: { costType: 'gratuit' } } },
    { 'PARKING_LOT.costType': 'FREE' });
verifier('🔴 UNE VALEUR POSÉE MAIS DIFFÉRENTE est détectée',
    ['PARKING_LOT.costType'], r.manques);

r = verifierPose({ services: ['WI_FI'] }, { services: ['WI_FI', 'RESTROOMS'] });
verifier('Une liste incomplète est un manque, pas une réussite',
    ['services'], r.manques);

r = verifierPose({}, { name: 'X', description: 'Y' });
verifier('Les champs du mécanisme hérité ne sont pas jugés ici',
    { confirmes: [], manques: [] }, r);

/* ------------------------------------------------------------------ *
 * 4. LA CHAÎNE ENTIÈRE : une ligne de classeur → l'objet envoyé       *
 * ------------------------------------------------------------------ */
const valeurs = lire(
    ['POI Permalink', 'POI Name', 'Parking Cost', 'Parking Payment', 'Parking Services',
     'Services', 'Horaires', 'Parking Spots'],
    ['https://www.waze.com/editor?venues=1', 'Parking P5', 'Modéré', 'Espèces ; Carte de crédit',
     'Surveillance', 'Wi-Fi', '8h-20h', '101-300']
);
const { maj, ignores } = construireMaj(valeurs.aPoser);
verifier('De la ligne au SDK, sans rien inventer',
    { services: ['WI_FI', 'SECURITY'],
      categoryAttributes: { PARKING_LOT: {
          costType: 'MODERATE', paymentType: ['CASH', 'CREDIT'],
          estimatedNumberOfSpots: 'R_101_TO_300' } } },
    maj);
verifier('Rien d’ignoré, et les horaires ne sont pas dans l’envoi', [], ignores);
verifier('Les horaires sont dans le tas « à montrer »', ['Horaires'], valeurs.montres.map(m => m.libelle));

/* ------------------------------------------------------------------ *
 * 5. COHÉRENCE : toute cible posable doit être constructible          *
 * ------------------------------------------------------------------ */
const toutesCibles = {};
CHAMPS.filter(c => c.pose).forEach(c => { toutesCibles[c.cible] = c.multiple ? ['X'] : 'X'; });
verifier('Aucune cible déclarée posable n’est écartée par la liste blanche',
    [], construireMaj(toutesCibles).ignores);

/* ------------------------------------------------------------------ */
if (echecs.length === 0) {
    console.log(`=== ${reussis} réussis, 0 échec ===`);
    process.exit(0);
}
echecs.forEach(e => console.error(`  [ ÉCHEC ] ${e}`));
console.error(`=== ${reussis} réussis, ${echecs.length} échec(s) ===`);
process.exit(1);
