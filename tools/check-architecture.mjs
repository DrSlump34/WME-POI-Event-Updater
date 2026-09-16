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
        true, apercu.includes("getLayersByName('venues')"));
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
console.log(echecs.length ? `\n${echecs.map((e) => `  ✖ ${e}`).join('\n')}\n` : '');
console.log(`=== ${reussis} réussis, ${echecs.length} échec${echecs.length > 1 ? 's' : ''} ===`);
process.exit(echecs.length ? 1 : 0);
