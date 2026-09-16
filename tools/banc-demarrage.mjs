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
        offsetWidth: 800, offsetHeight: 600, parentElement: null, lastElementChild: null,
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
globalThis.require = () => function () {};
globalThis.OpenLayers = { LonLat: function () {}, Bounds: function () {}, Projection: function () {} };
/* ⚠️⚠️ `isReady` A VRAI : le chargement seul ne prouve pas grand-chose — tout
   y est declaratif. C est initScript() qui construit, et c est la que se
   trouvent les erreurs d une refonte : une fonction disparue encore appelee,
   une variable retiree encore lue. */
globalThis.W = {
    userscripts: { state: { isReady: true, locale: 'fr' }, registerSidebarTab: () => ({ tabLabel: elementFactice(), tabPane: elementFactice() }), waitForElementConnected: async () => {} },
    map: { getLayersByName: () => [], getCenter: () => ({}), getZoom: () => 17, setCenter: rien },
    model: { venues: { getObjectById: () => null }, actionManager: { add: rien, getActions: () => [] } },
    loginManager: { user: { attributes: { rank: 4 } } },
};

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

console.log('✔ Le script demarre : aucune erreur levee au chargement.');
process.exit(0);
