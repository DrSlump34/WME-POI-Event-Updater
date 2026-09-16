# Où l'on en est — WME POI Event Updater

> ⚠️ **Ce fichier dit l'état À LA DATE OÙ IL EST ÉCRIT.** Ce qui reste à faire se
> re-mesure (`node tools/banc-*.mjs`, `node tools/check-*.mjs`), jamais ne se lit
> ici de confiance.

---

# 🆕 16/09/2026 — LA REFONTE DE L'INTERFACE — v0.52 LOCALE, ⏳ NON PUBLIÉE

## ⏳ Ce qui n'est PAS publié

| Version | Où | État |
|---|---|---|
| **0.50** | GreasyFork 578776, GitHub, Discuss | en ligne |
| 0.51 | locale | la carte reste sur le périmètre — **jamais publiée** |
| **0.52** | locale | la refonte complète — **jamais publiée** |

⛔ **Le dépôt GitHub est resté à la 0.50.** Rien n'est poussé ni publié sans le
geste de l'auteur.

## 🔴 LE SEUL CHEMIN NON ÉPROUVÉ, ET C'EST LE PLUS SENSIBLE

**« Appliquer » n'a jamais été cliqué dans la nouvelle interface.** C'est le
chemin qui ÉCRIT sur la carte, et il a été remanié : la boucle de pose est
sortie de l'ancien aperçu (`poserUnLieu`), et les lignes à poser se lisent
autrement (`lignesCochees`).

⚠️ **Le code de pose lui-même n'a pas été réécrit, il a été EXTRAIT** — mot pour
mot, à une chose près : il recevait des éléments du DOM (`inpName.value`), il
reçoit des chaînes. Ses deux garde-fous sont intacts : on relit après avoir
écrit, et un champ à poser jamais envoyé est un MANQUE, pas un succès.

⇒ **À faire avant toute publication** : charger un classeur d'essai, cocher une
ligne, appliquer, et vérifier que le bilan dit le bon nombre de modifications en
attente — puis annuler dans WME.

## ✅ Ce qui est éprouvé dans l'éditeur

Le script démarre · le bouton se docke sans recouvrir ceux de WME · la fenêtre
s'ouvre à leur gauche, **se déplace et se redimensionne à la souris**, **se
replie** · le glisser-déposer lit le classeur · le premier onglet s'ouvre seul ·
la carte se cadre sur le périmètre (zoom plafonné à 19) · le tableau distingue
les états, la case est en tête de ligne, la ligne qui RETIRE arrive orange et
non cochée, et le bouton annonce le nombre exact de lignes cochées.

## 🧰 Les outils, et ce que chacun NE voit pas

| Outil | Ce qu'il tient | Son angle mort |
|---|---|---|
| `banc-demarrage.mjs` | le script démarre dans un WME de façade | ses bouchons sont **complaisants** — ils ont dit trois fois « ça démarre » sur un script mort |
| `banc-demarrage.html` | le même, dans un **vrai navigateur** | il faut le servir (`php -S 127.0.0.1:8131 -t .`) |
| `banc-coque.mjs` · `banc-ligne.mjs` | l'ossature et les lignes, rendues en TEXTE | qu'un bouton existe ne dit pas qu'il se voit |
| `banc-tableau.html` · `maquette-ux.html` | le rendu réel, mesuré | ni le comportement, ni la carte |
| `banc-fenetre.mjs` | la géométrie (bornes, planchers, plafond) | pas que la souris réponde |
| `banc-carte.mjs` | le cadrage, et son câblage lu dans la source | le SDK et la carte |
| `banc-cochage.mjs` | la règle du cochage | pas la case elle-même |
| `check-css.mjs` | accent grave, interpolation, accolades, variables, `[hidden]` | la mise en page |
| `check-libelles.mjs` | libellés manquants **et en double** | les appels dynamiques (`t(x ? 'a' : 'b')`) |
| `check-architecture.mjs` | panneau/fenêtre, un seul champ de fichier, garde-fous, **classes orphelines** | lit du texte, pas du comportement |

## ⏳ Reste

1. **Appliquer**, une fois, dans l'éditeur (ci-dessus).
2. Publier quand l'auteur le décidera : GitHub, GreasyFork 578776, Discuss
   404593. ⚠️ Deux versions non publiées se suivent — l'annonce doit parler de
   la **0.52** et dire ce que la 0.51 apportait.
3. Les libellés morts de l'ancienne interface (une quarantaine) : ils ne gênent
   rien, `check-libelles.mjs` les liste en ⏳.
