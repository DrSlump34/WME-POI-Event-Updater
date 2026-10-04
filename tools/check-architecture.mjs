/**
 * Le partage panneau / fenetre tient-il ?
 *     node tools/check-architecture.mjs
 *
 * ⭐⭐⭐⭐ LE PANNEAU LATERAL PORTE LES REGLAGES, LA FENETRE PORTE LE TRAVAIL. Ce
 *    n'est pas une preference : le panneau lateral de WME fait DISPARAITRE son
 *    contenu des qu'on selectionne un objet sur la carte. On ne peut pas y
 *    travailler — et c'est exactement ce que faisait ce script.
 *
 * ⚠️ CE CONTROLE LIT DU TEXTE. Il dit qu'un geste est declare au bon endroit,
 *    pas qu'il fonctionne. Le dernier mot reste a l'editeur.
 */

import { source } from './extraire.mjs';

let reussis = 0;
const echecs = [];
const verifier = (intitule, attendu, obtenu) => {
    if (JSON.stringify(attendu) === JSON.stringify(obtenu)) { reussis++; return; }
    echecs.push(`${intitule}\n      attendu : ${JSON.stringify(attendu)}\n      obtenu  : ${JSON.stringify(obtenu)}`);
};

/** Le corps d'une fonction, isole avant qu'on y cherche quoi que ce soit. */
function corps(nom) {
    const d = source.indexOf(`function ${nom}(`);
    if (d === -1) return null;
    const f = source.indexOf('\n    }', d);

    return f === -1 ? null : source.slice(d, f);
}

// ---------------------------------------------------------------------------
// Le panneau ne porte plus l'operationnel
// ---------------------------------------------------------------------------
const panneau = corps('initScript');
if (!panneau) {
    echecs.push('initScript() est introuvable : le controle ne mesure plus rien');
} else {
    /* ⚠️ AUCUN STYLE EN DUR : le panneau se construisait en style.cssText, bouton
       par bouton, et c'est pour cela que rien n'etait homogene — il n'y avait
       rien a quoi etre homogene. */
    const enDur = (panneau.match(/style\.cssText\s*=/g) || []).length;
    verifier('plus aucun style en dur dans le panneau', 0, enDur);

    /* ⭐ LE SEUL GESTE QU'ON REFAIT TOUT LE TEMPS : rouvrir la fenetre. Il est
       hors des sections, en premier — range sous un titre, il serait a trouver. */
    /* ⚠️ ON EXIGE L'AJOUT AU PANNEAU, pas la presence du nom : un bouton
       declare mais jamais pose laisse la chaine dans le fichier, et le
       controle passait au vert sur un panneau qui ne l'affichait plus. */
    verifier('le panneau porte le geste qui rouvre la fenetre',
        true, panneau.includes('container.appendChild(btnFenetre)'));

    /* ⚠️ ET IL NE PORTE PLUS LE TRAVAIL. */
    ['btnShow', 'btnChoose'].forEach((mort) => {
        verifier(`« ${mort} » a bien quitte le panneau`, false, panneau.includes(mort));
    });
}

// ---------------------------------------------------------------------------
// Un seul chemin de lecture du classeur
// ---------------------------------------------------------------------------
/* ⚠️⚠️ UN SEUL CHAMP DE FICHIER. Un second serait un second chemin de lecture :
   celui-ci valide vingt regles, tient un rapport d'anomalies et alimente
   l'historique. Deux chemins divergent, et l'un des deux le fait en silence. */
const champsFichier = (source.match(/type\s*=\s*'file'/g) || []).length;
verifier('un seul champ de fichier dans tout le script', 1, champsFichier);
verifier('la fenetre passe par CE champ', true, source.includes('_peuFileInput.click()'));

// ---------------------------------------------------------------------------
// Les garde-fous qui ne doivent pas se perdre en demenageant
// ---------------------------------------------------------------------------
const apercu = corps('ouvrirApercu');
if (!apercu) {
    echecs.push('ouvrirApercu() est introuvable : le controle ne mesure plus rien');
} else {
    /* ⚠️⚠️ LE CALQUE « LIEUX » DOIT ETRE ALLUME, SANS QUOI RIEN N'EXISTE. WME ne
       charge pas les lieux d'un calque eteint : le prechargement ne trouverait
       AUCUN POI et l'apercu annoncerait que tout est introuvable — un diagnostic
       faux, sur un fichier juste. Ce garde-fou vivait sur le bouton du panneau ;
       il a demenage avec le geste, et c'est le genre de chose qu'on perd. */
    verifier('le calque « Lieux » est allume avant de balayer',
        // 0.54.00 : la visibilité se lit par le SDK (W retiré de WME le 24/11/2026).
        true, apercu.includes("isLayerVisible({ layerName: 'venues' })") && apercu.includes('#layer-switcher-group_places'));
    verifier('et son repli parle quand la bascule est introuvable',
        true, apercu.includes("t('layerOffMsg')"));
}

/* ⚠️ LE REPLI DE LA BIBLIOTHEQUE : si XLSX n'a pas pu se charger, rien ne
   fonctionnera, et le dire vaut mieux que d'echouer au premier clic. */
verifier('l’absence de la bibliotheque se dit',
    true, source.includes("typeof XLSX === 'undefined'"));

// ---------------------------------------------------------------------------
// Le rapport d'anomalies se lit la ou l'on travaille
// ---------------------------------------------------------------------------
const rapport = corps('showValidationReport');
if (!rapport) {
    echecs.push('showValidationReport() est introuvable : le controle ne mesure plus rien');
} else {
    /* ⭐ IL DIT CE QUI N'EST PAS ENTRE. Une ligne ecartee du classeur ne se voit
       nulle part ailleurs : ni dans le tableau, qui ne montre que ce qui est
       retenu, ni sur la carte. Dans le panneau, la carte le ferait disparaitre. */
    verifier('le rapport d’anomalies s’affiche dans la fenetre',
        true, rapport.includes('#peu-body'));
}

// ---------------------------------------------------------------------------
// Le classeur se depose, et pas seulement se choisit
// ---------------------------------------------------------------------------
const depot = corps('brancherDepot');
if (!depot) {
    echecs.push('brancherDepot() est introuvable : le controle ne mesure plus rien');
} else {
    /* ⚠️⚠️ `preventDefault` SUR dragover ET SUR drop : sans le premier, le
       navigateur refuse le depot ; sans le second, il OUVRE le classeur a la
       place de la carte — et l'on perd sa session d'edition. */
    ['dragover', 'drop'].forEach((evt) => {
        verifier(`« ${evt} » est ecoute`, true, depot.includes(`'${evt}'`));
    });
    verifier('le geste par defaut du navigateur est empeche',
        true, depot.includes('e.preventDefault()'));

    /* ⚠️⚠️ LE FICHIER DEPOSE PASSE PAR LE MEME CHAMP : un second chemin de
       lecture divergerait du premier, et il le ferait en silence. */
    verifier('le fichier depose passe par le champ de fichier existant',
        true, depot.includes('_peuFileInput.files = dt.files'));
    verifier('et declenche le meme evenement',
        true, depot.includes("dispatchEvent(new Event('change'"));

    /* ⚠️ CE QUI N'EST PAS UN CLASSEUR EST REFUSE, ET LE REFUS SE DIT. */
    verifier('un fichier qui n’est pas un classeur est refuse avec un mot',
        true, depot.includes("t('dropRefus'"));
}

verifier('le depot est branche a la construction de la fenetre',
    true, source.includes('brancherDepot(ov)'));

// ---------------------------------------------------------------------------
// Aucune fonction enfermee dans initScript et appelee au-dehors
// ---------------------------------------------------------------------------
/* ⭐⭐⭐⭐ CE CONTROLE ETAIT VERT PENDANT QUE LE DEPOT D'UN FICHIER PLANTAIT.
   `showValidationReport` vivait DANS initScript ; `brancherDepot`, au-dehors,
   l'appelait pour refuser un fichier qui n'est pas un classeur — une
   ReferenceError, et rien a l'ecran. Le controle cherchait « #peu-body » dans
   son texte : le texte etait juste, la PORTEE etait fausse (audit du
   25/09/2026). On mesure donc la portee : toute fonction declaree dans
   initScript ne doit etre nommee nulle part ailleurs. */
if (panneau) {
    const internes = [...panneau.matchAll(/\n {8}function\s+([A-Za-z_$][\w$]*)\s*\(/g)].map((m) => m[1]);
    const dehors = source.slice(0, source.indexOf(panneau)) + source.slice(source.indexOf(panneau) + panneau.length);
    const fuyantes = internes.filter((n) => new RegExp('\\b' + n + '\\s*\\(').test(dehors));
    verifier('aucune fonction interne a initScript n’est appelee hors de sa portee', [], fuyantes);
    if (!internes.length && panneau.includes('function ')) {
        echecs.push('fonctions internes a initScript introuvables : le motif ne mord plus');
    }
}

// ---------------------------------------------------------------------------
// Une seule icone, celle de l'en-tete
// ---------------------------------------------------------------------------
/* ⭐ La charte veut LA MEME icone sur l'onglet, le panneau, la fenetre et le
   bouton de la carte. Il y en avait trois. La constante et l'en-tete ne
   peuvent plus diverger. */
{
    const enTete = (source.match(/^\/\/ @icon\s+(\S+)/m) || [])[1];
    const constante = (source.match(/const ICONE = '([^']+)'/) || [])[1];
    verifier('ICONE est l’icone de l’en-tete', true, !!enTete && enTete === constante);
    verifier('aucune autre image d’icone dans le script', 0,
        (source.match(/data:image\/svg\+xml;base64,/g) || []).length - 2);
}

// ---------------------------------------------------------------------------
// Aucune classe CSS orpheline
// ---------------------------------------------------------------------------
/* ⭐⭐⭐⭐ UN STYLE SANS ELEMENT A L'AIR FAIT, ET NE L'EST PAS. La zone de depot
   avait son CSS complet — bordure pointillee, survol, etat de survol du
   fichier — et rien ne la posait : elle etait dans la maquette, validee, puis
   oubliee au branchement. C'est le pire des oublis, parce qu'il se lit comme
   un travail acheve. */
const ACCENT = String.fromCharCode(96);
const dCss = source.indexOf('const CSS = ' + ACCENT);
const fCss = source.indexOf('\n    ' + ACCENT + ';', dCss);
const css = source.slice(dCss, fCss);
const apres = source.slice(fCss);

const classesCss = new Set(
    [...css.matchAll(/\.(peu-[\w-]+)/g)].map((m) => m[1])
);
/* ⚠️⚠️ ON CHERCHE LA CLASSE COMME SOUS-CHAINE, et non dans un `class="…"`
   ferme : la moitie d entre elles se construisent par concatenation —
   `class="peu-cell-old` + une variante + le guillemet plus loin. Un motif qui
   exige l attribut complet en rate les trois quarts et accuse dix-sept classes
   qui sont toutes posees. Plus permissif ici vaut mieux que criant au loup. */
const estPosee = (c) => apres.includes(c);

/* ⚠️ Les classes d'etat se construisent par concatenation (`peu-row-` + etat) :
   on les considere posees des que le prefixe l'est. */
/* ⚠️ LES CLASSES BATIES PAR CONCATENATION ne se trouvent pas telles quelles :
   `'peu-row-' + etat`, `'peu-pastille-' + classe`. On accepte donc une classe
   dont le PREFIXE est pose — sans quoi le controle accuserait dix regles
   parfaitement employees. */
const prefixesBatis = ['peu-row-', 'peu-pastille-'];
const orphelines = [...classesCss].filter((c) => {
    if (estPosee(c)) return false;
    /* ⚠️ Le prefixe est SUIVI du guillemet fermant, mais precede de ce qu on
       veut : `class = 'peu-ligne peu-row-' + etat`. On cherche donc la fin de
       la chaine, pas son debut. */
    return !prefixesBatis.some((p) => c.startsWith(p) && apres.includes(p + "'"));
});

verifier('aucune classe stylee que rien ne pose', [], orphelines.sort());
// ---------------------------------------------------------------------------
console.log(echecs.length ? `\n${echecs.map((e) => `  ✖ ${e}`).join('\n')}\n` : '');
console.log(`=== ${reussis} réussis, ${echecs.length} échec${echecs.length > 1 ? 's' : ''} ===`);
process.exit(echecs.length ? 1 : 0);
