/**
 * Banc de l'ossature — `node tools/banc-coque.mjs`
 *
 * ⭐⭐⭐⭐ CE BANC EXISTE PARCE QUE L'INTERFACE N'ETAIT MESURABLE NULLE PART. Elle
 *    se construisait a coups de createElement au milieu du DOM et de la carte :
 *    seul l'oeil, dans l'editeur, pouvait la juger. C'est par l'oeil, tres
 *    tard, qu'on a vu que la case a cocher des lignes etait introuvable (14 px,
 *    quatrieme colonne, en-tete vide) et que le bouton qui ECRIT SUR LA CARTE
 *    etait un glyphe d'un caractere colle a celui qui ferme la fenetre.
 *
 * ⇒ L'ossature rend desormais du TEXTE, et ce texte se mesure ici.
 *
 * ⚠️ CE QU'IL NE VOIT PAS : la mise en page. Qu'un bouton soit present ne dit
 *    pas qu'il est visible, ni lisible, ni a la bonne taille. Cela se mesure
 *    dans `tools/maquette-ux.html` et, en dernier ressort, dans WME.
 */

import { charger, source } from './extraire.mjs';

/* ⚠️ LE BLOC S'EVALUE AVEC SES DEPENDANCES : `esc` et `t` vivent ailleurs dans
   le script. On les fournit ici, en les gardant INOFFENSIFS — `t` rend la clé
   demandee, ce qui permet de verifier qu'AUCUN libelle n'est ecrit en dur. */
const prelude = `
    const PEU_EMOJI = '@@';
    function esc(v) {
        return String(v === undefined || v === null ? '' : v)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }
    function t(cle, ...a) { return a.length ? cle + ':' + a.join(',') : cle; }
`;

const corps = prelude + source.slice(
    source.indexOf('// ==== banc:coque ===='),
    source.indexOf('// ==== /banc:coque ===='),
);
const { coqueOverlay, libelleAppliquer } =
    new Function(corps + '\nreturn { coqueOverlay, libelleAppliquer };')();

let reussis = 0;
const echecs = [];
const verifier = (intitule, attendu, obtenu) => {
    if (JSON.stringify(attendu) === JSON.stringify(obtenu)) { reussis++; return; }
    echecs.push(`${intitule}\n      attendu : ${JSON.stringify(attendu)}\n      obtenu  : ${JSON.stringify(obtenu)}`);
};

const html = coqueOverlay('0.52');

// ---------------------------------------------------------------------------
// Les ancres : chacune est un endroit ou le script viendra ecrire
// ---------------------------------------------------------------------------
[
    ['peu-header',        'l’en-tete, qui sert aussi de poignee de deplacement'],
    ['peu-strip',         'le bandeau d’etat : ce qui est charge, en permanence'],
    ['peu-strip-texte',   'le nom du classeur'],
    ['peu-select-onglet', 'le choix de l’onglet'],
    ['peu-strip-compte',  'le nombre de POI'],
    ['peu-btn-fichier',   'le bouton qui charge un classeur'],
    ['peu-body',          'le corps defilant'],
    ['peu-footer',        'le pied, hors defilement'],
    ['peu-footer-help',   'la phrase d’aide du pied'],
    ['peu-btn-appliquer', 'le bouton qui ecrit sur la carte'],
    ['peu-btn-export',    'le rapport'],
    ['peu-btn-replier',   'le repli'],
    ['peu-btn-fermer',    'la fermeture'],
    ['peu-resize',        'la poignee de redimensionnement'],
].forEach(([id, quoi]) => {
    verifier(`l’ancre « ${id} » existe — ${quoi}`, true, html.includes('id="' + id + '"'));
});

// ---------------------------------------------------------------------------
// Ce que la refonte corrige, et qui doit le rester
// ---------------------------------------------------------------------------
/* ⭐⭐⭐⭐ L'ACTION PRINCIPALE EST UN BOUTON PLEIN, NOMME, DANS LE PIED — et non
   plus un glyphe dans la barre de titre, a cote de la croix qui ferme. */
const pied = html.slice(html.indexOf('class="peu-footer"'));
verifier('le bouton qui ecrit est dans le PIED, pas dans l’en-tete',
    true, pied.includes('id="peu-btn-appliquer"'));
verifier('il est le seul bouton PLEIN de la fenetre',
    1, (html.match(/peu-btn-primary/g) || []).length);
verifier('il porte la classe primaire',
    true, /id="peu-btn-appliquer"[^>]*class="[^"]*peu-btn-primary|class="[^"]*peu-btn-primary[^"]*"[^>]*id="peu-btn-appliquer"/.test(html));

/* ⚠️ IL PART DESACTIVE : rien n’est charge, donc rien a poser. Un bouton actif
   qui ne ferait rien apprend a cliquer sans regarder. */
verifier('il part desactive', true, /id="peu-btn-appliquer"[^>]*disabled/.test(html));

/* ⚠️ AUCUN LIBELLE EN DUR : `t` rend ici la clé demandee, donc un texte ecrit
   en dur dans l’ossature se verrait — il ne serait pas traduit. */
verifier('aucun libelle en dur : tout passe par la traduction',
    false, />Appliquer|>Apply |>Choisir|>Choose/.test(html));

/* ⚠️ CHAQUE BOUTON PORTE UN title= : c’est le seul nom accessible de ceux qui
   n’ont qu’un glyphe, et l’infobulle de tous les autres. */
const boutons = html.match(/<button[^>]*>/g) || [];
verifier('chaque bouton porte un title', [], boutons.filter((b) => !b.includes('title=')));
verifier('il y a bien cinq boutons dans l’ossature', 5, boutons.length);

// ---------------------------------------------------------------------------
// Le libelle du bouton dit COMBIEN
// ---------------------------------------------------------------------------
/* ⭐ « Appliquer » seul ne dit pas sur quoi : on clique sans savoir si l’on
   touche un lieu ou quarante. */
verifier('rien de coche ⇒ le bouton le dit', 'btnApplyNone', libelleAppliquer(0));
verifier('une seule ligne ⇒ singulier',      'btnApplyOne',  libelleAppliquer(1));
verifier('plusieurs ⇒ le compte est dedans', 'btnApplyN:7',  libelleAppliquer(7));

// ---------------------------------------------------------------------------
// L'echappement
// ---------------------------------------------------------------------------
/* ⚠️ La version vient du script, mais la regle vaut pour tout ce qui entre. */
verifier('la version est echappee', true, coqueOverlay('0.5"><b>').includes('0.5&quot;&gt;&lt;b&gt;'));

// ---------------------------------------------------------------------------
console.log(echecs.length ? `\n${echecs.map((e) => `  ✖ ${e}`).join('\n')}\n` : '');
console.log(`=== ${reussis} réussis, ${echecs.length} échec${echecs.length > 1 ? 's' : ''} ===`);
process.exit(echecs.length ? 1 : 0);
