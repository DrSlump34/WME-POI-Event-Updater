/**
 * Le script DEMARRE-T-IL ? — `node tools/banc-demarrage.mjs`
 *
 * ⭐⭐⭐⭐ COMPILER N'EST PAS DEMARRER. `node --check` lit la syntaxe et ne voit
 *    RIEN d'une variable utilisee avant sa declaration, d'une fonction disparue
 *    encore appelee, ou d'une propriete lue sur un objet absent. Le script
 *    passe la verification, et ne se charge pas — sans un mot dans l'interface,
 *    puisqu'il n'y a plus d'interface.
 *
 * ⇒ Ce banc EXECUTE le script dans un WME de facade. Il ne verifie pas qu'il
 *   fonctionne : il verifie qu'il DEMARRE, ce qui est la premiere chose que
 *   personne ne mesurait.
 *
 * ⚠️ LA FACADE EST VOLONTAIREMENT PAUVRE. Elle rend ce qu'il faut pour aller
 *    jusqu'au bout du chargement, et rien de plus : un stub trop complaisant
 *    masquerait justement l'erreur qu'on cherche.
 */

import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const racine = join(dirname(fileURLToPath(import.meta.url)), '..');
const source = readFileSync(join(racine, 'WME_POI_Event_Updater.user.js'), 'utf8');

/* --------------------------------------------------------------------------
   Le WME de facade — le strict minimum du chemin de chargement.
   -------------------------------------------------------------------------- */
const rien = () => {};
const elementFactice = () => {
    const el = {
        style: {}, classList: { add: rien, remove: rien, toggle: rien, contains: () => false },
        dataset: {}, children: [], attributes: [],
        appendChild: rien, prepend: rien, remove: rien, after: rien, insertCell: elementFactice,
        addEventListener: rien, removeEventListener: rien, setAttribute: rien, getAttribute: () => null,
        querySelector: () => elementFactice(), querySelectorAll: () => [],
        getBoundingClientRect: () => ({ left: 0, top: 0, right: 800, bottom: 600, width: 800, height: 600 }),
        focus: rien, click: rien, insertRow: elementFactice, createTHead: elementFactice,
        set innerHTML(v) { this._html = v; }, get innerHTML() { return this._html || ''; },
        textContent: '', value: '', hidden: false, disabled: false, checked: false,
        offsetWidth: 800, offsetHeight: 600, parentElement: null, lastElementChild: null, isConnected: true,
    };

    return el;
};

const journal = [];
globalThis.window = globalThis;
globalThis.document = {
    documentElement: { lang: 'fr' },
    head: elementFactice(),
    body: elementFactice(),
    createElement: elementFactice,
    getElementById: () => null,
    querySelector: () => null,
    querySelectorAll: () => [],
    addEventListener: rien,
    hidden: false,
};
globalThis.localStorage = {
    _d: {},
    getItem(k) { return this._d[k] || null; },
    setItem(k, v) { this._d[k] = v; },
    removeItem(k) { delete this._d[k]; },
};
globalThis.setInterval = () => 0;
globalThis.clearInterval = rien;
/* ⚠️ LE VRAI setTimeout EST CONSERVE pour le banc lui-meme : sans lui, on ne
   peut pas laisser la boucle d evenements tourner assez longtemps pour
   qu un rejet de promesse remonte. */
const setTimeoutVrai = globalThis.setTimeout;
globalThis.setTimeout = (fn) => { journal.push('setTimeout'); return 0; };
globalThis.alert = rien;
globalThis.XLSX = { read: rien, utils: {} };
/* ⭐⭐⭐⭐ UN WME SANS `W` (0.54.00). Waze retire `W` de WME le 24/11/2026 : la façade
   n'en a PAS, ni `OpenLayers`, ni `require`. Un seul accès restant, et le
   script lève ici — c'est ce qu'on veut voir avant Waze.
   ⚠️ LE FAUX SDK EST EN MODE ASYNC (seul mode au 01/01/2027) : chaque méthode
   rend une promesse, comme le vrai. Et il REFUSE ce qu'il ne connaît pas. */
const async = (v) => () => Promise.resolve(v);
const fauxSdk = {
    Sidebar: { registerScriptTab: () => { journal.push('onglet'); return Promise.resolve({ tabLabel: elementFactice(), tabPane: elementFactice() }); } },
    Settings: { getLocale: async({ localeCode: process.env.PEU_LOCALE || 'fr' }) },
    State: { getUserInfo: async({ rank: 4 }) },
    Map: { getZoomLevel: async(17), setMapCenter: async(), zoomToExtent: async(), isLayerVisible: async(true) },
    Editing: { getUnsavedChangesCount: async(0) },
    DataModel: { Venues: { getById: async(null), updateVenue: async(), ParkingLot: {} } },
};
const sdkStrict = (o, nom) => new Proxy(o, {
    get(c, k) {
        if (typeof k === 'symbol' || k === 'then') return c[k];
        if (!(k in c)) throw new Error('SDK : membre inconnu « ' + nom + '.' + String(k) + ' »');
        const v = c[k];
        return (v && typeof v === 'object') ? sdkStrict(v, nom + '.' + String(k)) : v;
    },
});
const pageGetWmeSdk = (o) => {
    if (!o || o.mode !== 'async') throw new Error('getWmeSdk doit être appelé en mode async');
    journal.push('sdk');
    return sdkStrict(fauxSdk, 'sdk');
};

/* ⭐⭐⭐ DEUX MONDES, ET LE SCRIPT DOIT DÉMARRER DANS LES DEUX.
   · la PAGE (script injecté à la main, ou `@grant none`) : W est une globale ;
   · le BAC À SABLE de Tampermonkey (`@grant GM_xmlhttpRequest`) : W n'existe
     QUE sous `unsafeWindow`. Un seul `W` lu comme globale, et le script meurt
     au premier geste dans l'éditeur — alors qu'il démarre ici en mode page.
   Le second passage se lance avec --bac-a-sable. */
const BAC = process.argv.includes('--bac-a-sable');
if (BAC) {
    globalThis.unsafeWindow = { getWmeSdk: pageGetWmeSdk, SDK_INITIALIZED: Promise.resolve() };
    globalThis.GM_xmlhttpRequest = () => { journal.push('GM_xmlhttpRequest'); };
    globalThis.GM_info = { script: { version: '0.53.00' } };
} else {
    globalThis.getWmeSdk = pageGetWmeSdk;
    globalThis.SDK_INITIALIZED = Promise.resolve();
}

/* ⚠️⚠️ initScript() EST `async` : une exception qui s y produit ne remonte PAS
   au try/catch — elle devient une promesse rejetee. C est exactement ce qui
   rend la panne invisible dans l editeur : le script ne demarre pas, et rien
   ne le dit, puisqu il n y a plus d interface pour le dire. */
let erreur = null;
process.on('unhandledRejection', (e) => { erreur = erreur || e; });
try {
    // eslint-disable-next-line no-eval
    (0, eval)(source);
} catch (e) {
    erreur = e;
}
/* On laisse la boucle d evenements tourner un tour : c est la que le rejet arrive. */
await new Promise((r) => setTimeoutVrai(r, 60));

if (erreur) {
    console.error('\n✖ LE SCRIPT NE DEMARRE PAS.\n');
    console.error(`  ${erreur.name} : ${erreur.message}`);
    const ligne = (erreur.stack || '').split('\n').find((l) => /<anonymous>:\d+/.test(l));
    if (ligne) console.error(`  ${ligne.trim()}`);
    console.error('\n  (les numeros de ligne comptent depuis le debut du fichier)\n');
    process.exit(1);
}

/* ⚠️ « AUCUNE ERREUR » NE SUFFIT PAS : un script qui n'a rien construit ne lève
   rien non plus. On exige que l'onglet ait été enregistré. */
if (!journal.includes('onglet')) {
    console.error('\n✖ Le script n’a rien construit : l’onglet n’a jamais été enregistré.\n');
    process.exit(1);
}
if (BAC && journal.filter((j) => j === 'GM_xmlhttpRequest').length !== 1) {
    console.error('\n✖ La vérification de version devait partir UNE fois : ' + journal.filter((j) => j === 'GM_xmlhttpRequest').length + '.\n');
    process.exit(1);
}

if (BAC) {
    console.log('✔ Bac à sable : le script démarre, sans W, le SDK lu par unsafeWindow.');
    process.exit(0);
}
console.log('✔ Page : le script démarre, aucune erreur levée au chargement.');
/* Le second monde, dans un processus neuf : les globales du premier ne doivent
   pas lui prêter main-forte. */
try {
    const sortie = execFileSync(process.execPath, [fileURLToPath(import.meta.url), '--bac-a-sable'], { encoding: 'utf8' });
    process.stdout.write(sortie);
} catch (e) {
    process.stdout.write(e.stdout || '');
    process.stderr.write(e.stderr || '');
    process.exit(1);
}
console.log('✔ Le script demarre : aucune erreur levee au chargement, dans les deux mondes.');
process.exit(0);
