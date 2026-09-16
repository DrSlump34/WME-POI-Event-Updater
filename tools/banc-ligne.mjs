/**
 * Banc de la ligne d'apercu — `node tools/banc-ligne.mjs`
 *
 * ⭐⭐⭐⭐ C'EST ICI QUE VIVENT LES DEUX DEFAUTS QUI ONT FAIT REPRENDRE L'INTERFACE.
 *    La case a cocher etait en quatrieme et derniere colonne, large de trente
 *    pixels, sous un en-tete VIDE, a cote d'une case maitre qui lui ressemble —
 *    et le pied de page disait « decochez les lignes a exclure » en designant
 *    quelque chose que personne ne voyait. Elle est desormais en PREMIERE
 *    colonne, sous un intitule.
 *
 * ⚠️ CE BANC LIT DU TEXTE : il dit qu'une case est la, pas qu'elle se voit. La
 *    taille et la place se mesurent dans `tools/maquette-ux.html`, et le dernier
 *    mot reste a l'editeur.
 */

import { source } from './extraire.mjs';

/* `esc` et `t` vivent ailleurs : on les fournit, et `t` rend la cle demandee
   pour qu'un libelle ecrit en dur se voie. */
const prelude = `
    function esc(v) {
        return String(v === undefined || v === null ? '' : v)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }
    function t(cle, ...a) { return a.length ? cle + ':' + a.join(',') : cle; }
`;
const corps = prelude + source.slice(
    source.indexOf('// ==== banc:ligne ===='),
    source.indexOf('// ==== /banc:ligne ===='),
);
const { etatDeLaLigne, badgeDeLigne, ligneApercu, enteteApercu } =
    new Function(corps + '\nreturn { etatDeLaLigne, badgeDeLigne, ligneApercu, enteteApercu };')();

let reussis = 0;
const echecs = [];
const verifier = (intitule, attendu, obtenu) => {
    if (JSON.stringify(attendu) === JSON.stringify(obtenu)) { reussis++; return; }
    echecs.push(`${intitule}\n      attendu : ${JSON.stringify(attendu)}\n      obtenu  : ${JSON.stringify(obtenu)}`);
};

const base = {
    idx: 3, nom: 'Parking P1', desc: '', ancienNom: 'Parking P1', ancienDesc: '',
    charge: true, verrou: 'ok', pertes: 0, champsDiff: 0, nomChange: false, descChange: false,
};
const avec = (x) => Object.assign({}, base, x);

// ---------------------------------------------------------------------------
// L'etat, et son ORDRE
// ---------------------------------------------------------------------------
verifier('rien ne change ⇒ ok', 'ok', etatDeLaLigne(base));
verifier('le nom change ⇒ diff', 'diff', etatDeLaLigne(avec({ nomChange: true })));
verifier('un champ du lot D2 ⇒ diff', 'diff', etatDeLaLigne(avec({ champsDiff: 4 })));
verifier('une perte ⇒ perte', 'perte', etatDeLaLigne(avec({ pertes: 1 })));

/* ⭐⭐⭐⭐ CE QU'ON NE PEUT PAS FAIRE PASSE AVANT CE QU'ON FERAIT. Annoncer
   « 4 champs a poser » sur un lieu verrouille au-dessus du rang, ou introuvable,
   c'est promettre ce qu'on ne tiendra pas. */
verifier('un lieu introuvable l’emporte sur tout',
    'unloaded', etatDeLaLigne(avec({ charge: false, nomChange: true, champsDiff: 9, pertes: 2 })));
verifier('un verrou dur l’emporte sur la perte',
    'hard', etatDeLaLigne(avec({ verrou: 'hard', pertes: 3 })));
verifier('un SaE l’emporte sur l’ecart',
    'sae', etatDeLaLigne(avec({ verrou: 'sae', champsDiff: 2 })));
/* ⚠️ Mais la PERTE passe avant l'ecart ordinaire : elle demande un geste. */
verifier('la perte passe avant l’ecart',
    'perte', etatDeLaLigne(avec({ pertes: 1, nomChange: true, champsDiff: 5 })));

// ---------------------------------------------------------------------------
// Le badge
// ---------------------------------------------------------------------------
verifier('l’ecart compte le nom, la description ET les champs',
    '6', badgeDeLigne('diff', avec({ nomChange: true, descChange: true, champsDiff: 4 })).texte);
verifier('la perte s’annonce en negatif',
    '-2', badgeDeLigne('perte', avec({ pertes: 2 })).texte);
verifier('le verrou dur dit son niveau',
    'L5', badgeDeLigne('hard', avec({ niveau: 5 })).texte);
verifier('rien a faire ⇒ un signe egal, pas un vide',
    '=', badgeDeLigne('ok', base).texte);
/* ⚠️ Chaque badge porte une infobulle : seul, un glyphe ne dit rien. */
['unloaded', 'hard', 'sae', 'perte', 'diff', 'ok'].forEach((e) => {
    verifier(`le badge « ${e} » porte une infobulle`, true, !!badgeDeLigne(e, avec({ pertes: 1 })).titre);
});

// ---------------------------------------------------------------------------
// La ligne
// ---------------------------------------------------------------------------
const html = ligneApercu(avec({ champsDiff: 4, nomChange: true, ancienNom: 'Ancien nom' }));

/* ⭐ LA CASE EST DANS LA PREMIERE CELLULE. C'est tout l'objet de la refonte. */
const premiereCellule = html.slice(html.indexOf('<td'), html.indexOf('</td>'));
verifier('la case est dans la PREMIERE cellule', true, premiereCellule.includes('type="checkbox"'));
verifier('elle vit dans un label, donc toute la zone est cliquable', true, premiereCellule.includes('<label class="peu-check"'));
verifier('elle porte une infobulle', true, /type="checkbox"[^>]*title="/.test(premiereCellule));

/* ⚠️ RIEN N'EST COCHE ICI : `cocherDOffice` decide apres coup, et une ligne qui
   RETIRE quelque chose ne se coche jamais seule. */
verifier('aucune case n’est cochee au rendu', false, html.includes('checked'));

verifier('l’etat pose sa classe de ligne', true, html.includes('peu-row-diff'));
verifier('l’ancien nom est barre quand il change', true, html.includes('peu-cell-old changed'));
verifier('l’index de la ligne est porte', true, html.includes('data-idx="3"'));

/* ⚠️ CE QU'ON NE PEUT PAS POSER NE S'EDITE PAS : proposer un champ qui ne
   partira jamais, c'est faire perdre une saisie. */
const fige = ligneApercu(avec({ charge: false }));
verifier('un lieu introuvable : la case est desactivee', true, /type="checkbox"[^>]*disabled/.test(fige));
verifier('un lieu introuvable : les champs sont desactives', 2, (fige.match(/ disabled/g) || []).length - 1);

/* ⚠️ TOUT CE QUI VIENT DU CLASSEUR EST ECHAPPE : un nom de lieu est une donnee
   externe, et il finit dans un attribut title=. */
const vicieux = ligneApercu(avec({ nom: 'A"><script>x</script>', ancienNom: 'B" onmouseover="y' }));
verifier('le nom est echappe', false, vicieux.includes('<script>'));
verifier('l’ancien nom aussi', false, vicieux.includes('onmouseover="y'));

// ---------------------------------------------------------------------------
// L'en-tete
// ---------------------------------------------------------------------------
const tete = enteteApercu();
/* ⭐ L'EN-TETE DE LA COLONNE DE LA CASE PORTE UN INTITULE. Vide, il ne disait
   pas ce que la case decide — et c'est ce qui la rendait invisible. */
verifier('la colonne de la case a un intitule', true, tete.includes('colSelect'));
verifier('la case maitre y est', true, tete.includes('data-maitre'));
verifier('la colonne d’etat a un intitule', true, tete.includes('colEtat'));
verifier('quatre colonnes sont declarees', 4, (tete.match(/<col>/g) || []).length);
/* ⚠️ `<th[ >]` ET NON `<th` : le motif large compte aussi `<thead>`, et le
   banc annoncait cinq en-tetes pour quatre colonnes. Un motif trop large
   ment dans le sens le moins visible — il accuse un code correct. */
verifier('quatre en-tetes sont rendus', 4, (tete.match(/<th[ >]/g) || []).length);
verifier('aucun libelle en dur dans l’en-tete', false, /Poser<|État<|Nom<|Apply</.test(tete));

// ---------------------------------------------------------------------------
console.log(echecs.length ? `\n${echecs.map((e) => `  ✖ ${e}`).join('\n')}\n` : '');
console.log(`=== ${reussis} réussis, ${echecs.length} échec${echecs.length > 1 ? 's' : ''} ===`);
process.exit(echecs.length ? 1 : 0);
