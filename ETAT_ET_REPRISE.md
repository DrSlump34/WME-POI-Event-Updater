# Où l'on en est — WME POI Event Updater

> ⚠️ **Ce fichier dit l'état À LA DATE OÙ IL EST ÉCRIT.** Ce qui reste à faire se
> re-mesure (`node tools/banc-*.mjs`, `node tools/check-*.mjs`), jamais ne se lit
> ici de confiance.

---

# ✅ 25/09/2026 — AUDIT TRAITÉ — v0.53.00 PUBLIÉE

| Où | Vérifié |
|---|---|
| **GreasyFork 578776** | code SERVI identique au local (hors les deux lignes réécrites) ; descriptions EN/FR en Markdown ; **une seule capture**, la 0.53 |
| **GitHub** | `master` poussé (capture 0.53, README, descriptions) |
| **Discuss 404593** | post 1 réécrit (version ×2, capture 0.53, liens GF + GitHub dans chaque section) ; annonce bilingue (post 8) |

⏳ Reste : un lieu SaE et le téléchargement du rapport, jamais essayés dans WME ; les deux lignes de liens FR du post 1 ont encore une espace ordinaire avant « : ».

L'audit du 25/09/2026 (11 dimensions, un réfuteur chacune, un arbitre ; rapport
local `AUDIT-2026-09-25.md`, **exclu du dépôt**) a été traité en entier, dans
l'ordre qu'il proposait : les filets d'abord, puis ce qui écrit, puis ce qui
trompe, puis la charte.

## ✅ ESSAI DANS WME DU 25/09/2026 (Tampermonkey, bac à sable) — SANS ENREGISTRER

Installée par Tampermonkey (pas injectée) : démarrage sans erreur ; onglet
Scripts mesuré (`getComputedStyle`) conforme à la charte — titre « Waze Boing
Medium » 13 px #2196f3, version 11 px #9e9e9e, intro 11 px #566372, pilule
#1976d2 3px 10px, aide repliée ; les 15 px sous le h2 viennent de WME et valent
pour les six scripts. Fenêtre à gauche des boutons de la carte.
`Essai 0.50 - champs du lot D2.xlsx` : 18 lieux, tous trouvés. Retouche d'une
ligne « = » ⇒ cochée, badge 1 ; retouche effacée ⇒ décochée ; nom vidé ⇒
décochée, « le nom du lieu serait effacé ». Onglet Essai D2 : P1 orange -1 non
cochée, parking Moto 4 champs cochée. **Appliquer 1 ligne** ⇒ « ✔ 1 lieu posé,
1 modification ajoutée à la pile », ligne ✔ décochée, bouton « Rien de coché »,
historique « appliqué » daté ; modèle relu : FREE, CASH, STREET_LEVEL, PMR +
Surveillance. **Une annulation ⇒ pile 0, lieu revenu à l'avant (UNKNOWN, [],
[], []), Enregistrer grisé.**
Défaut vu : « (remplace : UNKNOWN) » brut ⇒ corrigé (`948736f`).

**Non essayé** : un lieu SaE (aucun verrouillé au-dessus du rang 6 dans
l'essai), le téléchargement du rapport, l'hébreu.

## Ce qui avait été prévu pour l'essai

**Le chemin qui écrit n'a toujours pas été cliqué dans l'éditeur** — ni en 0.52,
ni en 0.53. Il est désormais couvert par un banc (`banc-application`), mais un
banc ne voit ni `UpdateObject`, ni le vrai SDK, ni un lieu verrouillé.

Et **la 0.53 change le monde dans lequel le script tourne** : `@grant
GM_xmlhttpRequest` le place dans le bac à sable de Tampermonkey, où `W`,
`OpenLayers`, `require` et `getWmeSdk` ne se lisent que par `unsafeWindow`.
`banc-demarrage` le vérifie au démarrage, pas au premier geste.

⇒ **L'essai, dans l'ordre** (installer la 0.53.00 dans Tampermonkey, pas par
injection — l'injection tourne en mode page et ne prouve rien du bac à sable) :

1. le script démarre ; l'onglet Scripts montre l'icône, `v0.53.00`, l'aide
   repliée, le pied de liens ; la console ne dit rien ;
2. charger `Essai 0.50 - champs du lot D2.xlsx` ; les anomalies s'affichent et
   **restent** au-dessus du tableau ;
3. retoucher le nom d'une ligne « = » : elle se **coche** ; le vider : elle se
   décoche ;
4. cocher **une** ligne, dont un lieu verrouillé au-dessus du rang si possible
   (SaE), puis Appliquer ; lire le bilan (actions ajoutées, SaE), la ligne ✔, le
   rapport exporté (colonne « Champs non posés ») ;
5. **Annuler dans WME jusqu'à ce que la pile soit vide.** Enregistrer reste grisé.

Puis refaire les captures (l'interface a changé de couleurs) et publier selon
`peu-publication-workflow` : la description est prête dans
`Descr. GreasyFork 0.53.md`.

## ✅ Ce que la 0.53 corrige (audit du 25/09/2026)

| # | Défaut | Correction |
|---|---|---|
| A2a | L'ordre des catégories n'était pas comparé : la catégorie principale changeait en silence | `LISTES_ORDONNEES = ['categories']` |
| A1 | Une retouche dans l'aperçu était jetée en silence (régression 0.52) | `caseApresRetouche` + écouteur `input` |
| A2b | Noms alternatifs ni comparés, ni montrés, ni relus | comparés comme les autres champs |
| A5 | Repli A/B/C appliqué alors que des en-têtes reconnues étaient ailleurs | onglet refusé (`sheetHeaderConflit`) |
| A4 | L'infobulle promettait qu'une description vide n'efface rien | texte corrigé, effacement dit dans le champ, **sans décocher** |
| A4b | Nom vide coché d'office | décoché |
| A6 | Pose sans `try` : une exception arrêtait tout, sans bilan | `poserLesLignes`, un `try` par lieu |
| A7 | Aucun état « occupé » : double pose d'un clic | `occuper()` ; lignes posées décochées |
| A9 | Lieu sans nom déclaré introuvable | `lieuPret` |
| B1 | Rapport sans permalien ni champs non posés | 8 colonnes, statuts traduits |
| B3 | Anomalies effacées au rendu du tableau | `_anomalies`, redessinées à chaque rendu |
| E1 | Valeur remplacée visible seulement en infobulle | « (remplace : …) » en clair |
| — | Nom et description jamais relus après `UpdateObject` | la relecture fait foi pour tout |
| — | Recentrage à chaque pose | seulement si le lieu manque |
| — | `recordFileApplied` jamais appelée | rétablie |
| — | « 0 en attente » sur une erreur | ajouts de CETTE pose, ou rien |
| — | Lieux en échec non nommés, « Réessayer » disparu | nommés, restent cochés, Réessayer |
| — | Lignes cochées masquées par le filtre | le pied le dit |
| — | `showValidationReport` hors de portée (dépôt d'un non-classeur) | sortie d'`initScript` |
| — | Calque des lieux allumé sans relecture, `alert()` | relu, message dans la fenêtre |
| — | Cellules lues brutes | `String(…).trim()` |
| — | Deux colonnes pour un champ, `venues=` multiples, autre `env` | signalés |
| — | Tri promis mais absent | tri par nom / description, au clavier aussi |
| — | Replier puis déplacer ⇒ 120 px | `memoriserGeometrie` |
| — | SheetJS 0.18.5 (2 CVE), deux `@require` sans empreinte | 0.20.3, `cdn.sheetjs.com`, `#sha256=` |
| — | Gabarit publié rejeté en entier | permaliens complétés (0/0) |
| — | Accessibilité : focus, rôles, contrastes, `aria-label`, clavier | voir SPEC § 7 |
| — | Code mort (`makeDraggable`, `GEOM_KEY`, `champsParCible`…) | retiré |
| — | Documents périmés (SPEC, README, description GreasyFork) | réécrits |

**Charte commune** : en-tête (`0.53.00`, `@homepageURL`, `@supportURL`,
`@connect`), une seule icône, panneau aux valeurs relevées dans WME, aide
repliable, pied de liens, pastille de nouvelle version, 8 langues, insécables,
un seul bouton plein. ⚖️ Pilules pleines en **#1976d2** (décision de l'auteur).
⚖️ **`@namespace` gardé** (`tampermonkey.net`) : le changer peut faire voir un
nouveau script aux gestionnaires, donc des doublons chez les installés.

## 🧰 Les outils, et ce que chacun NE voit pas

| Outil | Ce qu'il tient | Son angle mort |
|---|---|---|
| `banc-application.mjs` | **le chemin qui écrit** : pose, relecture, exception au milieu, retouche, catégories permutées, alias ajouté ; bouchons qui REFUSENT l'inconnu | `UpdateObject`, le vrai SDK, un lieu SaE |
| `banc-demarrage.mjs` | le script démarre, **en mode page ET en bac à sable** (W par `unsafeWindow` seulement), l'onglet est construit | ce qui se passe au premier geste |
| `banc-demarrage.html` | le même, dans un **vrai navigateur** | il faut le servir (`php -S 127.0.0.1:8131 -t .`) |
| `banc-coque.mjs` · `banc-ligne.mjs` | l'ossature et les lignes, rendues en TEXTE ; `esc()` **extrait** du script | qu'un bouton existe ne dit pas qu'il se voit |
| `banc-tableau.html` · `maquette-ux.html` | le rendu réel, mesuré | ni le comportement, ni la carte |
| `banc-fenetre.mjs` | la géométrie (bornes, planchers, plafond) | pas que la souris réponde |
| `banc-carte.mjs` | le cadrage, et son câblage lu dans la source | le SDK et la carte |
| `banc-cochage.mjs` | la règle du cochage | pas la case elle-même |
| `banc-colonnes.mjs` | la lecture des en-têtes, le repli et son refus, les colonnes en double | le DOM |
| `banc-champs.mjs` · `banc-valeurs.mjs` · `banc-pose.mjs` | conversion, liste blanche, relecture | le SDK |
| `banc-chargement.mjs` | rejoue la lecture d'un vrai classeur (SheetJS **0.20.3** à fournir) | l'écran |
| `check-css.mjs` | accent grave, interpolation, accolades, variables, `[hidden]` | la mise en page |
| `check-libelles.mjs` | **8 langues**, même jeu de clés, appels directs ET indirects, clés dérivées (`ch`/`mo`/`va`), insécables françaises | la justesse d'une traduction |
| `check-architecture.mjs` | panneau/fenêtre, un seul champ de fichier, **portée des fonctions d'`initScript`**, icône unique, classes orphelines | lit du texte, pas du comportement |

Chaque contrôle ajouté ou réparé le 25/09/2026 a été **vu échouer** d'abord
(sur la 0.52 ou par mutation) : `banc-application` (11 échecs sur la 0.52),
`banc-colonnes` (3), `check-libelles` (5 mutations, 38 fautes de typographie),
`check-architecture` (portée, icône), `banc-demarrage` (W lu comme globale),
`banc-ligne`/`banc-coque` (`esc()` cassé).

## ⏳ Reste

1. **L'essai dans WME** ci-dessus, puis captures et publication.
2. La **lecture** du classeur ne reconnaît que les libellés français et les clés
   WME (les catégories exceptées) : un classeur rempli dans une autre langue
   voit ses valeurs refusées — signalées, jamais posées.
3. Les liens du pied et les titres en #2196f3 font 3,12:1 sur blanc (petit
   texte) : c'est la charte, à arbitrer pour les quatre scripts comme l'ont été
   les pilules.
