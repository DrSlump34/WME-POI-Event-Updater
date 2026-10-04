/**
 * Banc du chemin qui ÉCRIT — `node tools/banc-application.mjs`
 *
 * ⭐⭐⭐⭐ CE BANC EXISTE PARCE QUE LE CHEMIN QUI ÉCRIT N'EN AVAIT AUCUN. L'audit du
 *    25/09/2026 a cassé la pose de quatre façons graves sans qu'un seul des
 *    treize outils passe au rouge — et ce chemin n'avait jamais été cliqué
 *    depuis la refonte de la 0.52.
 *
 * ⚠️⚠️ LES BOUCHONS NE SONT PAS COMPLAISANTS. `env` est un Proxy qui LÈVE sur
 *    tout appel qu'il ne connaît pas, et le faux SDK se comporte comme le vrai
 *    relevé le 15/09/2026 : il accepte sans erreur un champ inconnu et ne le
 *    pose pas. Un bouchon qui dit oui à tout a déjà fait dire trois fois « ça
 *    démarre » à un script mort (banc-demarrage).
 *
 * ⚠️ VU ÉCHOUER D'ABORD : sur le code du 25/09/2026, les cas marqués ⛔ tombaient
 *    (catégories permutées, nom alternatif ajouté seul, retouche sur une ligne
 *    non cochée, exception au milieu de la pose).
 *
 * Il ne remplace pas l'essai dans WME : ni `UpdateObject`, ni le vrai SDK, ni
 * un lieu verrouillé au-dessus du rang ne vivent ici.
 */

import { bloc } from './extraire.mjs';

/* `charger` rend un objet ; les deux derniers symboles peuvent ne pas exister
   encore — le banc doit alors ÉCHOUER, pas planter à l'extraction. */
const blocs = await import('./extraire.mjs');
const corps = ['champs', 'colonnes', 'pose', 'poser'].map(blocs.bloc).join('\n');
const X = new Function(corps + `
    return {
        comparerAuLieu, champsQuiDifferent, cocherDOffice, pertesDeLaPose,
        poserUnLieu, caseApresRetouche,
        poserLesLignes: typeof poserLesLignes === 'function' ? poserLesLignes : undefined,
        lieuPret: typeof lieuPret === 'function' ? lieuPret : undefined,
    };`)();

let reussis = 0;
const echecs = [];
const verifier = (intitule, attendu, obtenu) => {
    const a = JSON.stringify(attendu), o = JSON.stringify(obtenu);
    if (a === o) { reussis++; return; }
    echecs.push(`${intitule}\n      attendu : ${a}\n      obtenu  : ${o}`);
};

/* ------------------------------------------------------------------ *
 * Le faux WME                                                         *
 * ------------------------------------------------------------------ */

/** Les champs que le faux SDK sait poser — le reste est ACCEPTÉ ET IGNORÉ, comme le vrai. */
const CONNUS_DU_SDK = ['phone', 'services', 'categories', 'categoryAttributes'];

function fauxWme(lieux, options = {}) {
    const journal = [];
    const modele = new Map(lieux.map((l) => [l.attributes.id, l]));
    const methodes = {
        lieuCharge(vid) {
            journal.push('lieuCharge:' + vid);
            return modele.get(vid) || null;
        },
        async charger(perm, vid) {
            journal.push('charger:' + vid);
            return options.introuvables && options.introuvables.includes(vid) ? null : (modele.get(vid) || null);
        },
        ecrireHerite(venue, champs) {
            journal.push('herite:' + venue.attributes.id);
            if (options.refuseHerite && options.refuseHerite.includes(venue.attributes.id)) {
                throw new Error('UpdateObject refusé');
            }
            /* 0.54.00 : le nom passe par le SDK async — un refus est une PROMESSE REJETÉE. */
            if (options.heriteRejette && options.heriteRejette.includes(venue.attributes.id)) {
                return Promise.reject(new Error('updateVenue refusé'));
            }
            if (options.herite === 'muet') return;          // une suggestion : rien ne change
            ['name', 'description', 'aliases'].forEach((k) => {
                if (k in champs) venue.attributes[k] = champs[k];
            });
        },
        ecrireSdk(vid, maj) {
            journal.push('sdk:' + vid);
            /* Le SDK en mode async : un refus arrive en PROMESSE REJETÉE, pas en exception. */
            if (options.sdkRejette) return Promise.reject(new Error('SDK refusé'));
            const a = modele.get(vid).attributes;
            Object.keys(maj).forEach((k) => {
                if (!CONNUS_DU_SDK.includes(k)) return;       // accepté, ignoré, sans un mot
                if (k === 'categoryAttributes') {
                    a.categoryAttributes = a.categoryAttributes || {};
                    Object.keys(maj[k]).forEach((c) => {
                        a.categoryAttributes[c] = Object.assign({}, a.categoryAttributes[c], maj[k][c]);
                    });
                    return;
                }
                a[k] = maj[k];                                // une liste REMPLACE tout
            });
        },
        relire(vid) {
            journal.push('relire:' + vid);
            const v = modele.get(vid);
            return v ? v.attributes : null;
        },
    };
    const env = new Proxy(methodes, {
        get(cible, cle) {
            if (cle in cible) return cible[cle];
            if (cle === 'then') return undefined;             // ce n'est pas une promesse
            throw new Error('Bouchon : appel inconnu « ' + String(cle) + ' »');
        },
    });
    return { env, journal, modele };
}

const lieu = (id, attrs) => ({ attributes: Object.assign({ id: id, name: 'Lieu ' + id, description: '', aliases: [] }, attrs) });
const item = (vid, nom, desc, aPoser, extra) => Object.assign({
    vid: vid, perm: 'https://www.waze.com/editor?env=row&lat=1&lon=1&zoomLevel=17&venues=' + vid,
    nom: nom, desc: desc, valeurs: { aPoser: aPoser || {}, montres: [], refus: [] },
}, extra || {});

/* ------------------------------------------------------------------ *
 * 1. ⛔ LES ÉCARTS QUI S'ÉCRIVENT SANS SE VOIR                         *
 * ------------------------------------------------------------------ */

/* La catégorie principale est la PREMIÈRE : permuter change l'icône et le type. */
let r = X.comparerAuLieu({ categories: ['CAFE', 'RESTAURANT'] }, { categories: ['RESTAURANT', 'CAFE'] });
verifier('⛔ Catégories permutées : la catégorie principale change, c’est une DIFFÉRENCE',
    ['categories'], r.differents);
r = X.comparerAuLieu({ services: ['WI_FI', 'RESTROOMS'] }, { services: ['RESTROOMS', 'WI_FI'] });
verifier('Services permutés : l’ordre d’une liste ordinaire ne compte pas',
    [], r.differents);

/* Un nom alternatif AJOUTÉ seul ne retire rien : la ligne doit s'annoncer et se cocher. */
const avecAlias = { name: 'P1', description: '', aliases: ['Parking Est'] };
const demandeAlias = { aliases: ['Parking Est', 'P1 Est'] };
verifier('⛔ Nom alternatif ajouté seul : un champ diffère',
    1, X.champsQuiDifferent(avecAlias, demandeAlias));
verifier('⛔ … et la ligne se coche d’office (rien n’est retiré)',
    true, X.cocherDOffice(false, false, X.champsQuiDifferent(avecAlias, demandeAlias),
        Object.keys(X.pertesDeLaPose(avecAlias, demandeAlias)).length));
verifier('Un nom alternatif RETIRÉ reste une perte',
    ['aliases'], Object.keys(X.pertesDeLaPose(avecAlias, { aliases: ['P1 Est'] })));

/* ------------------------------------------------------------------ *
 * 2. ⛔ LA RETOUCHE DANS L'APERÇU                                      *
 * ------------------------------------------------------------------ */
const vue = (o) => Object.assign({ nomChange: false, descChange: false, champsDiff: 0, pertes: 0, nomVide: false }, o);
verifier('⛔ On corrige le nom d’une ligne « = » non cochée : elle se coche',
    true, X.caseApresRetouche(false, vue({ nomChange: true })));
verifier('On remet la valeur du lieu : elle se décoche',
    false, X.caseApresRetouche(true, vue({})));
verifier('⭐ Une ligne qui RETIRE ne se coche jamais seule, même retouchée',
    false, X.caseApresRetouche(false, vue({ nomChange: true, pertes: 2 })));
verifier('⭐ … et on ne décoche pas non plus ce que l’éditeur y a coché lui-même',
    true, X.caseApresRetouche(true, vue({ nomChange: true, pertes: 2 })));
verifier('⛔ On vide le nom d’une ligne cochée : elle se décoche (poser un nom vide se décide)',
    false, X.caseApresRetouche(true, vue({ nomChange: true, nomVide: true })));
verifier('Nom vide dans le classeur : la ligne n’est pas cochée d’office',
    false, X.cocherDOffice(true, false, 0, 0, true));

/* ------------------------------------------------------------------ *
 * 3. LA POSE ELLE-MÊME                                                *
 * ------------------------------------------------------------------ */
const attendre = async () => {
    /* 3.1 Le cas ordinaire : tout est posé, tout est relu. */
    let w = fauxWme([lieu('1', { phone: '' })]);
    let res = await X.poserUnLieu(item('1', 'Nouveau', 'Texte', { phone: '04' }), w.env);
    verifier('Cas ordinaire : appliqué, rien ne manque',
        ['applied', []], [res.resultat.status, res.resultat.manques]);

    /* 3.2 Le SDK accepte un champ et ne le pose pas. */
    w = fauxWme([lieu('2')]);
    res = await X.poserUnLieu(item('2', 'N', '', { url: 'https://exemple.fr', phone: '01' }), w.env);
    verifier('Le SDK a tu un refus : la relecture le voit (partiel, url manque)',
        ['partial', ['url']], [res.resultat.status, res.resultat.manques]);

    /* 3.3 ⛔ Le nom n'est pas arrivé (une suggestion, un refus muet) : ce n'est pas « appliqué ». */
    w = fauxWme([lieu('3', { name: 'Ancien' })], { herite: 'muet' });
    res = await X.poserUnLieu(item('3', 'Nouveau', '', {}), w.env);
    verifier('⛔ Nom relu après UpdateObject : absent ⇒ il MANQUE',
        ['partial', ['name']], [res.resultat.status, res.resultat.manques]);

    /* 3.4 ⛔ Un lieu déjà en mémoire ne fait pas bouger la carte. */
    w = fauxWme([lieu('4')]);
    await X.poserUnLieu(item('4', 'N', '', {}), w.env);
    verifier('⛔ Lieu déjà chargé : la carte ne se recentre pas',
        false, w.journal.includes('charger:4'));

    /* 3.5 ⛔ Une exception au milieu ne coupe pas la pose, et se dit. */
    if (typeof X.poserLesLignes !== 'function') {
        echecs.push('⛔ poserLesLignes absente : une exception au lieu k arrête la boucle sans bilan');
    } else {
        w = fauxWme([lieu('5'), lieu('6'), lieu('7')], { refuseHerite: ['6'] });
        const vus = [];
        const bilan = await X.poserLesLignes(
            [item('5', 'A', '', {}), item('6', 'B', '', {}), item('7', 'C', '', {})],
            w.env, (n) => vus.push(n));
        verifier('⛔ Trois lignes, la deuxième lève : trois résultats',
            ['applied', 'erreur', 'applied'], bilan.resultats.map((x) => x.status));
        verifier('… le message est gardé', 'UpdateObject refusé', bilan.resultats[1].erreur);
        verifier('… et la progression va au bout', [1, 2, 3], vus);
        verifier('… le lieu en erreur est rendu pour « Réessayer »', ['6'], bilan.aReprendre.map((x) => x.vid));
    }

    /* 3.5 bis ⛔ (0.54.00) Le refus du NOM arrive en promesse rejetée (SDK async) : sans
       `await` devant ecrireHerite, la ligne se dirait « posée » ou « partielle ». */
    {
        w = fauxWme([lieu('8'), lieu('9')], { heriteRejette: ['9'] });
        const bilan = await X.poserLesLignes([item('8', 'A', '', {}), item('9', 'B', '', {})], w.env);
        verifier('⛔ Nom refusé en promesse rejetée : la ligne est en ERREUR',
            ['applied', 'erreur'], bilan.resultats.map((x) => x.status));
        verifier('… avec le message du SDK', 'updateVenue refusé', bilan.resultats[1].erreur);
    }

    /* 3.6 ⛔ Un lieu sans nom est un lieu chargé. */
    if (typeof X.lieuPret !== 'function') {
        echecs.push('⛔ lieuPret absente : un lieu sans nom est déclaré introuvable');
    } else {
        verifier('⛔ Lieu sans nom : chargé', true, X.lieuPret({ attributes: { id: '8', name: '' } }));
        verifier('Rien dans le modèle : pas chargé', false, X.lieuPret(null));
        verifier('Objet sans attributs : pas chargé', false, X.lieuPret({}));
    }

    /* 3.8 ⛔ SDK async : un refus en promesse rejetée se dit, il ne s'évapore pas. */
    w = fauxWme([lieu('11')], { sdkRejette: true });
    res = await X.poserUnLieu(item('11', 'N', '', { phone: '04' }), w.env);
    verifier('⛔ SDK async qui rejette : le message est gardé', true, JSON.stringify(res).includes('SDK refusé'));

    /* 3.7 Introuvable : rien n'est écrit. */
    w = fauxWme([lieu('9')], { introuvables: ['10'] });
    res = await X.poserUnLieu(item('10', 'N', '', {}), w.env);
    verifier('Introuvable : échec, et AUCUNE écriture',
        [true, false], [res.echec, w.journal.some((j) => j.startsWith('herite:'))]);
};

await attendre().catch((e) => echecs.push('✖ Le banc a levé : ' + e.message));

echecs.forEach((e) => console.log('  ✖ ' + e));
console.log(`\n=== ${reussis} réussis, ${echecs.length} échec${echecs.length > 1 ? 's' : ''} ===`);
process.exit(echecs.length ? 1 : 0);
