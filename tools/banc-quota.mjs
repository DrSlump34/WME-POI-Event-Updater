/**
 * Banc du quota — `node tools/banc-quota.mjs`
 *
 * ⭐⭐⭐ LE PLAFOND DE LA DOCTRINE FRANCE (§ 3.3), 0.55.00 : au plus le plafond de lieux
 *    modifiés par « Appliquer », par heure glissante et par éditeur, réservés à la
 *    pose, comptés à l'enregistrement, rendus sur annulation totale. Modèle de WME
 *    BAN Coverage (tools/banc-quota.js de WBC).
 *
 * Chaque « page » est une nouvelle exécution du bloc `banc:quota` (mémoire de page
 * vide, nouvelle session) sur le MÊME localStorage ; l'horloge est factice.
 *
 * ⭐⭐ LE TÉMOIN : le banc se falsifie ensuite lui-même (sceau retiré, fenêtre
 *    ignorée, rendu étendu aux autres pages, confirmation muette, plafond relevé
 *    dans le cœur…) et EXIGE que chaque mutant fasse échouer le scénario. Un
 *    mutant qui passe = un contrôle qui ment : le banc échoue.
 *
 * ⚠️ Limite assumée : qui efface le localStorage ET recharge la page repart de zéro.
 *    Il ne remplace pas l'essai dans WME : ni `wme-save-finished`, ni le compteur
 *    de modifications non enregistrées ne sont vus ici.
 */

import { bloc } from './extraire.mjs';

const H = 3600 * 1000;
const T0 = Date.UTC(2026, 9, 11, 10, 0, 0);

function scenario(code) {
    const ls = new Map();
    const localStorage = {
        getItem: (k) => (ls.has(k) ? ls.get(k) : null),
        setItem: (k, v) => ls.set(k, String(v)),
        removeItem: (k) => ls.delete(k),
    };
    const horloge = { t: T0 };
    const VraieDate = Date;
    class FausseDate extends VraieDate {
        constructor(...a) { super(...(a.length ? a : [horloge.t])); }
        static now() { return horloge.t; }
    }
    const page = (user = 'drslump34') => new Function('localStorage', 'Date',
        'let _quotaUser = ' + JSON.stringify(user) + ';\n' + code
        + '\nreturn { etatQuota, reserverLieu, confirmerQuota, rendreQuota, enAttenteIci, lireQuota, _pc, cle: _qn() };',
    )(localStorage, FausseDate);
    const res = [];
    const ok = (nom, v) => res.push([nom, !!v]);
    const reste = (p) => p.etatQuota(horloge.t).reste;

    // 1. L'auto-test et le plafond
    let a = page();
    ok('auto-test accepté sur le code livré', a._pc() === true);
    ok('plafond 150, rien d’utilisé au départ', a.etatQuota(horloge.t).max === 150 && reste(a) === 150);

    // 2. Réservation
    for (let i = 0; i < 40; i++) { horloge.t += 1000; a.reserverLieu('v' + i, horloge.t); }
    ok('40 lieux réservés ⇒ reste 110', reste(a) === 110);
    a.reserverLieu('v3', horloge.t);
    ok('un lieu déjà en attente ne compte pas deux fois', reste(a) === 110);
    ok('des réservations en attente sur cette page', a.enAttenteIci() === true);

    // 3. Rendu (annulation totale) : seulement les réservations de CETTE page
    const b = page();
    ok('une autre page voit les réservations (même stockage)', reste(b) === 110);
    ok('une autre page ne rend RIEN de cette page', b.rendreQuota() === 0 && reste(b) === 110);
    ok('annulation totale ⇒ les 40 réservations rendues', a.rendreQuota() === 40 && reste(a) === 150);
    ok('… et plus rien en attente', a.enAttenteIci() === false);

    // 4. Confirmation (enregistrement) : comptée, plus jamais rendue
    for (let i = 0; i < 30; i++) { horloge.t += 1000; a.reserverLieu('w' + i, horloge.t); }
    ok('enregistrement ⇒ 30 confirmées', a.confirmerQuota() === 30);
    ok('une confirmée n’est pas rendue', a.rendreQuota() === 0 && reste(a) === 120);
    ok('une page rechargée relit le compte', reste(page()) === 120);

    // 5. Refus au-delà du plafond
    let acceptes = 0;
    for (let i = 0; i < 200; i++) if (a.reserverLieu('x' + i, horloge.t)) acceptes++;
    ok('au-delà du plafond : 120 acceptées sur 200 demandées, les autres REFUSÉES', acceptes === 120);
    const q = a.etatQuota(horloge.t);
    ok('plein : reste 0, exactement 150 comptées', q.reste === 0 && q.utilises === 150);
    ok('heure de reprise = la plus ancienne + 1 h', q.libre === T0 + 41000 + H);

    // 6. Fenêtre glissante
    horloge.t = T0 + 41000 + H + 1;
    ok('une heure après la plus ancienne : une place se libère', reste(a) === 1);
    horloge.t = T0 + 2 * H + 100000;
    ok('deux heures plus tard : tout est libre', reste(a) === 150);

    // 7. Par éditeur, stockage illisible
    a.reserverLieu('y1', horloge.t);
    ok('un autre éditeur a son propre quota', reste(page('autre')) === 150);
    const v = ls.get(a.cle) || '';
    ok('nom de clé neutre (sans pseudo)', !/drslump/i.test(a.cle));
    const dec = Buffer.from(v.replace(/^p1\.[0-9a-z]+\./, ''), 'base64').toString('latin1');
    ok('valeur illisible, même base64 décodé', v.startsWith('p1.') && ![v, dec].some((x) => /"[tvep]":|y1/.test(x)));

    // 8. Copie retouchée ⇒ bloqué une heure
    const w = ls.get(a.cle), i = w.length - 6;
    ls.set(a.cle, w.slice(0, i) + (w[i] === 'A' ? 'B' : 'A') + w.slice(i + 1));
    ok('copie retouchée ⇒ Appliquer bloqué', reste(page()) === 0);
    horloge.t += H + 5000;
    ok('… pendant une heure seulement', reste(page()) === 150);
    ls.set(a.cle, (ls.get(a.cle) || '').replace(/^p1\.[0-9a-z]+\./, 'p1.0.'));
    ok('sceau remplacé ⇒ Appliquer bloqué', reste(page()) === 0);
    ls.clear();

    // 9. Copie effacée, MÊME page ⇒ rien de rendu (mémoire de la page)
    const c = page();
    for (let k = 0; k < 20; k++) c.reserverLieu('z' + k, horloge.t);
    c.confirmerQuota();
    ls.clear();
    ok('localStorage effacé, même page ⇒ rien de rendu', reste(c) === 130);
    return res;
}

// La constante falsifiée : l'auto-test doit REFUSER, et Appliquer se bloquer.
function falsifiee(code) {
    const ls = new Map();
    const localStorage = { getItem: (k) => ls.get(k) ?? null, setItem: (k, v) => ls.set(k, String(v)), removeItem: (k) => ls.delete(k) };
    const p = new Function('localStorage', 'let _quotaUser = "drslump34";\n' + code + '\nreturn { _pc, etatQuota };')(localStorage);
    const e = p.etatQuota(Date.now());
    return { pc: p._pc(), altere: !!e.altere, reste: e.reste };
}

const source = bloc('quota');
const echecs = [];
let reussis = 0;

const res = scenario(source);
for (const [nom, v] of res) { if (v) reussis++; else echecs.push('✖ ' + nom); }

/* L'auto-test refuse une constante falsifiée — chaque falsification seule, et la
   sonde relevée avec le cœur (la somme XOR ne suit pas). */
const FALSIFS = [
    ['plafond du cœur relevé à 1000', [/_pm = 0x96/, '_pm = 0x3e8']],
    ['fenêtre du cœur réduite à 30 min', [/_pf = 0x36ee80/, '_pf = 0x1b7740']],
    ['plafond de la sonde changé', [/_pk = \[0x96,/, '_pk = [0x3e8,']],
    ['cœur ET sonde relevés ensemble (somme XOR non suivie)', [/_pm = 0x96/, '_pm = 0x3e8'], [/_pk = \[0x96,/, '_pk = [0x3e8,']],
    ['somme de contrôle changée', [/_pq = 0x4e72ca1/, '_pq = 0x4e72ca2']],
];
for (const [nom, ...rempl] of FALSIFS) {
    let code = source;
    for (const [re, par] of rempl) {
        if (!re.test(code)) { echecs.push('✖ falsification « ' + nom + ' » : motif introuvable ' + re); continue; }
        code = code.replace(re, par);
    }
    const f = falsifiee(code);
    if (!f.pc && f.altere && f.reste === 0) reussis++;
    else echecs.push('✖ auto-test : « ' + nom + ' » ACCEPTÉ (' + JSON.stringify(f) + ')');
}
{
    const f = falsifiee(source);
    if (f.pc && !f.altere && f.reste === 150) reussis++; else echecs.push('✖ code livré refusé par l’auto-test');
}

/* ⭐⭐ TÉMOINS : chaque mutant du code DOIT faire échouer le scénario. */
const MUTANTS = [
    ['sceau ignoré', /if \(!m \|\| _qt\(m\[2\]\) !== m\[1\]\) return false;/, 'if (!m) return false;'],
    ['fenêtre glissante ignorée', /\.filter\(\(x\) => _ph\(x\) > maintenant - f\)/, '.filter((x) => _ph(x) > 0)'],
    ['rendu étendu aux autres pages', /garde = l\.filter\(\(x\) => !\(x\.e === 0 && x\.p === SESSION_PEU\)\)/, 'garde = l.filter((x) => !(x.e === 0))'],
    ['confirmation muette', /if \(x\.e !== 0 \|\| x\.p !== SESSION_PEU\) return x;/, 'return x;'],
    ['plafond relevé dans le cœur ET la sonde ET la somme', /_pm = 0x96/, '_pm = 0x3e8'],
    ['réservation sans dédoublonnage', /if \(l\.some\(\(x\) => x\.e === 0 && x\.p === SESSION_PEU && x\.v === String\(vid\)\)\) return true;/, ''],
    ['réservation sans refus au-delà du plafond', /if \(!_pc\(\) \|\| quotaPeu\(l, maintenant\)\.reste <= 0\) return false;/, ''],
    ['mémoire de la page oubliée', /listes = \[_quotaMem \|\| \[\]\]/, 'listes = [[]]'],
];
for (const [nom, re, par] of MUTANTS) {
    if (!re.test(source)) { echecs.push('✖ témoin « ' + nom + ' » : motif introuvable — le banc ne suit plus le code'); continue; }
    let code = source.replace(re, par);
    if (nom.startsWith('plafond relevé')) {
        // Le fraudeur appliqué : cœur, sonde ET somme XOR mis d'accord. L'auto-test passe ;
        // c'est le SCÉNARIO (plafond attendu 150) qui doit le voir.
        code = code.replace(/_pk = \[0x96,/, '_pk = [0x3e8,')
            .replace(/_pq = 0x[0-9a-f]+/, '_pq = 0x' + ((0x4d1c2b7 ^ 0x3e8 ^ 0x36ee80) >>> 0).toString(16));
    }
    let r;
    try { r = scenario(code); } catch (e) { r = [['exception : ' + e.message, false]]; }
    const tombes = r.filter(([, v]) => !v).map(([n]) => n);
    if (tombes.length) reussis++;
    else echecs.push('✖ TÉMOIN MUET : le mutant « ' + nom + ' » passe le scénario');
}

console.log('');
echecs.forEach((e) => console.log(e));
console.log(`=== ${reussis} réussis, ${echecs.length} échec${echecs.length > 1 ? 's' : ''} ===`);
process.exit(echecs.length ? 1 : 0);
