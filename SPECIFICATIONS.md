# PEU — WME POI Event Updater · Dossier de spécifications

> **Version du code décrite ici : 0.53.00** (lue dans le bloc `==UserScript==` de
> `WME_POI_Event_Updater.user.js`). Elle traite l'audit du 25/09/2026
> (`AUDIT-2026-09-25.md`, local, hors dépôt public).
> Diffusé sur **GreasyFork 578776**, dépôt `github.com/DrSlump34/WME-POI-Event-Updater`,
> fil Discuss **404593**.

---

## 0. À qui s'adresse ce dossier

Dossier de reprise du projet. Les bancs et contrôles de `tools/` (quinze, dont `banc-application`
pour le chemin qui écrit) se listent dans `ETAT_ET_REPRISE.md`, avec l'angle mort de chacun.
⚠️ **Aucun ne voit la carte** : tout ce qui touche à WME ne se vérifie que dans l'éditeur, et les
deux défauts les plus graves de la 0.50 y ont été trouvés (§ 3.2).

| Document | Rôle |
|---|---|
| `README.md` | Vitrine anglaise : à quoi ça sert, le format du fichier, l'installation |
| **`SPECIFICATIONS.md`** (ce fichier) | **Normatif** : le contrat, le modèle de données, les invariants |
| `WME_POI_Event_Updater_Template.xlsx` | Le **gabarit vierge** du fichier d'entrée — c'est lui qui fait foi sur les colonnes |
| `Descr. GreasyFork 0.53.md` | Les descriptions GreasyFork à jour (EN et FR), à recopier à la publication |

⚠️ **`ACO Events.xlsx` contient des données d'événement réelles** et est exclu du dépôt public
(`.gitignore`). Ne jamais le publier, ne jamais le joindre à un rapport de bug — le README le dit
explicitement aux utilisateurs.

---

## 1. Contexte et enjeu

### 1.1 Le besoin

Certains éditeurs **re-libellent le même jeu de lieux à chaque édition d'un événement récurrent** :
parkings, entrées, arrêts de navette, campings. À la main, cela veut dire ouvrir chaque lieu, taper
le nouveau nom, taper la nouvelle description, recommencer — des dizaines de fois, deux fois par an.

PEU renverse le geste : **on remplit un tableur une fois, le script l'applique**.

### 1.2 L'enjeu

Le script **écrit sur la carte**. Il pose des noms et des descriptions sur des lieux réels, en lot,
et l'éditeur enregistre ensuite. Deux exigences en découlent :

1. **Rien ne s'applique sans avoir été montré.** Le fichier est d'abord **affiché** ligne à ligne,
   avec l'ancien et le nouveau contenu, éditable à l'écran, avant tout bouton « Appliquer ».
2. **Ce qui n'a pas pu être fait doit être dit.** Un lieu que le script n'a pas réussi à charger est
   compté `timeout` et figure **dans le rapport**, jamais tu.

---

## 2. Périmètre

### 2.1 Ce que PEU fait

- Lire un **fichier Excel** (une ligne par lieu) contenant un permalink, un nom et une description.
- **Précharger** chaque lieu en recadrant la carte dessus, puis afficher le tableau des changements.
- Permettre de **filtrer, relire et retoucher** chaque valeur avant application.
- **Appliquer** les changements dans WME (sans enregistrer — c'est l'éditeur qui enregistre).
- Produire un **rapport Excel** de ce qui a été fait, avec l'avant et l'après.
- Garder un **historique des cinq derniers fichiers**, avec la date de chargement et la date
  d'application.
- Fonctionner dans les **huit langues de la charte commune** : français, anglais, allemand,
  espagnol, italien, portugais (Brésil et Portugal), hébreu (de droite à gauche).

### 2.2 Ce que PEU ne fait pas

- **Il ne crée pas de lieux** et n'en supprime aucun : il met à jour les champs d'un lieu existant.
- **Les noms alternatifs ne sont touchés que si le classeur en porte** (colonne `Alternative
  Names`, depuis la 0.50) ; sinon ils sont relus et repassés tels quels. Depuis la 0.53 ils se
  **comparent** comme les autres champs : un ajout seul s'annonce et se coche, un retrait est une
  perte (§ 3.2).
- **Il n'enregistre pas.** Les modifications sont posées dans la pile d'annulation de WME.
- **Un seul appel réseau propre au script** : la vérification de nouvelle version, au plus une
  fois par 24 h, par `GM_xmlhttpRequest` vers `update.greasyfork.org` (§ 10.1). Rien d'autre ne sort.

---

## 3. Le fichier d'entrée

**Une ligne par lieu, un fichier par événement.** Gabarit vierge :
`WME_POI_Event_Updater_Template.xlsx`.

| Colonne | Contenu |
|---|---|
| `POI Permalink` | Permalink WME pointant le lieu (`env`, `lat`, `lon`, `zoomLevel`, `venues`) |
| `POI Name` | Nom à poser |
| `POI Description` | Description à poser — **une cellule vide EFFACE la description** |

⚠️⚠️ **LE NOM ET LA DESCRIPTION SONT TOUJOURS ENVOYÉS, VIDES COMPRIS.** Une description vide
efface celle du lieu, et c'est voulu : l'onglet « Hors Evenement » remet les lieux à nu (251
descriptions vides sur 342 dans les données ACO). L'aperçu le dit **dans le champ** (« la
description du lieu sera effacée »), sans décocher. Un **nom** vide, lui, n'est jamais coché
d'office. La règle « une cellule vide ne demande rien » ne vaut **que pour les champs du lot D2**.

Lecture par **SheetJS 0.20.3**, un seul `@require` depuis `cdn.sheetjs.com` (accepté par
GreasyFork), avec son empreinte `#sha256=`. ⚠️ **La 0.18.5 des CDN publics avait deux failles
connues** (CVE-2023-30533, pollution de prototype ; CVE-2024-22363, ReDoS) — SheetJS ne publie plus
sur npm, d'où l'hôte. Il n'y a **pas de repli** : les deux `@require` d'avant étaient **exécutés
tous les deux**, ce n'était pas un repli mais une surface doublée.

### 3.0 Les colonnes se lisent par leur EN-TÊTE — 0.49

Jusqu'en 0.48, les colonnes étaient lues **par position** (`header:['perm','name','desc']`,
`range:1`) : les trois en-têtes devaient seulement exister et n'être pas vides, leur libellé
n'était jamais lu. L'ordre des colonnes était donc un **contrat tacite**, et une colonne insérée à
gauche décalait tout **en silence**.

`mapColumns(entetes)` repère désormais chaque colonne par son libellé (casse, espaces — insécable
compris — et accents ignorés) :

| Colonne | Libellés reconnus |
|---|---|
| permalien | `POI Permalink`, `Permalink`, `Permalien`, `POI Permalien` |
| nom | `POI Name`, `Name`, `Nom`, `POI Nom`, `Nom du POI` |
| description | `POI Description`, `Description`, `Desc`, `POI Desc` |

⭐ **Tout ou rien.** Les trois en-têtes reconnues → lecture par nom ; sinon **repli sur A/B/C**,
comme en 0.48, et une anomalie le dit (`sheetHeaderFallback`). Un repli partiel attribuerait une
colonne au hasard ; la règle, elle, se dit en une phrase. Un onglet dont les trois premières
en-têtes sont vides reste **ignoré**, comme avant.

⚠️⚠️ **0.53 — LE REPLI SE REFUSE QUAND UNE EN-TÊTE RECONNUE EST AILLEURS QU'À SA PLACE A/B/C.**
`Lien WME | Description | Nom` : le repli lisait la description comme le nom, sur tout l'onglet,
chaque ligne cochée d'office. L'onglet est désormais ignoré (`sheetHeaderConflit`). Cette garde ne
joue **qu'en repli** : un classeur lu par ses en-têtes n'est jamais concerné.

⚠️ **Deux colonnes pour un même champ** (« Site » et « Website ») : la première est lue, et
depuis la 0.53 **cela se dit** (`colonneDoublon`, lettres des colonnes à l'appui).

⇒ Les colonnes peuvent être **réordonnées**, et des colonnes **supplémentaires sont ignorées** :
c'est ce qui ouvre l'enrichissement du format (lot D2 de l'extranet EVIDRA).

⚠️ **Le numéro de ligne des anomalies est celui du tableur**, calculé depuis `!ref` : une feuille
dont la plage ne commence pas en A1 ne renvoie plus à une ligne introuvable.

✅ **Éprouvé** : `node tools/banc-colonnes.mjs` — 19 cas, **7 mutations sur 7 mordent**. Le banc
**extrait le bloc du userscript lui-même**, il n'en garde pas de copie. Et la lecture réelle a été
rejouée hors WME sur `ACO Events.xlsx` (**342 lignes, 6 onglets**) et sur le gabarit : lignes et
numéros de ligne **identiques** à ceux de 0.48.

✅ **Et éprouvé DANS WME le 15/09/2026**, sur un classeur d'essai à sept onglets : « 35 POI
chargés — 6 ⚠️ », les **six anomalies au mot près**, le sélecteur limité aux **cinq** onglets
exploitables (`Config` et l'onglet sans en-têtes absents), l'onglet aux colonnes permutées avec
deux colonnes en plus rendant ses **8 POI** et l'aperçu « ✅ À jour » — celui que 0.48 laissait
vide. Rien n'a été appliqué ; « Enregistrer » est resté grisé.

✅ **Tranché en 0.53 : une ligne au nom vide arrive DÉCOCHÉE** (`cocherDOffice`, cinquième
argument), avec « le nom du lieu serait effacé » dans le champ. La cocher reste possible.

⚠️ **Les cellules se lisent en texte, sans espace autour** (`String(…).trim()`) : un nom gardait
son espace final et le posait tel quel, et un nombre ne se comparait jamais égal au texte du lieu.

## 3.1 Le permalink, et ce qu'on en tire

`getVenueIdFromPermalink(url)` lit `venues=` : l'identifiant peut être **numérique (ancien) ou un
GUID alphanumérique (nouveau)**, et il peut y en avoir plusieurs séparés par des virgules —
**on prend le premier**, et depuis la 0.53 **l'anomalie le dit** (`urlPlusieursLieux`). Un
permalien d'un **autre serveur** que celui de l'éditeur (`env=`) est gardé mais signalé
(`urlAutreEnv`) : « introuvable » n'en donnerait pas la cause.

`parseLatLon(url)` lit `lat` / `lon` **en gérant les valeurs négatives** (hémisphère sud, ouest de
Greenwich) : c'est ce qui rend le script utilisable ailleurs qu'en Europe.

---

### 3.2 Les champs du lot D2 — 0.50

Au-delà du permalien, du nom et de la description, le classeur peut porter **22
colonnes**, reconnues par leur en-tête (§ 3.0). **14 sont POSÉES** dans WME, **8
sont MONTRÉES** sans être posées — horaires, adresse, points d'entrée, opérateur
de parking et les champs Google : les poser demanderait d'interpréter
une phrase ou de deviner un identifiant, et une interprétation fausse s'écrit sur
la carte sans que rien ne la signale.

**Ce qui protège l'écriture**, et chaque point vient d'un relevé, pas d'une
intuition :

1. une **table de conversion** relevée dans l'éditeur — le SDK accepte une valeur
   hors énumération et la POSE telle quelle ;
2. une **liste blanche** des champs — un nom de champ inconnu est accepté sans
   erreur, ne pose rien, et marque quand même le lieu modifié ;
3. une **relecture du lieu après écriture** (`comparerAuLieu`) — sans elle,
   « appliqué » ne veut dire que « l'appel n'a pas levé d'exception » ;
4. un **refus de compter pour un succès** un champ à poser qui n'a jamais été
   envoyé (voir ci-dessous).

⚠️ **Un TABLEAU remplace tout son contenu** dans WME : la pastille d'une valeur
posée dit **en clair** ce qu'elle **remplace** (« (remplace : …) » — en infobulle
seulement jusqu'en 0.52, inatteignable au clavier et au doigt). Vérifié sur un
cas réel : un classeur demandant « Espèces » sur un parking qui portait « Carte
de crédit, Espèces » supprime la carte de crédit.

⭐⭐⭐ **L'ORDRE DES CATÉGORIES COMPTE, CELUI DES AUTRES LISTES NON.** La première
catégorie est la catégorie principale (icône, type du lieu). Jusqu'en 0.52,
`comparerAuLieu` triait les deux listes : une permutation passait pour égale
avant la pose ET à la relecture, alors que le SDK recevait l'ordre du classeur.
`LISTES_ORDONNEES = ['categories']` le corrige ; services et modes de paiement
restent comparés sans ordre.

⭐⭐ **LES NOMS ALTERNATIFS SE COMPARENT** depuis la 0.53 : exclus de la
comparaison (`CIBLES_HERITEES`), un ajout seul arrivait « = », non coché, et
n'était jamais posé. Ils s'écrivent toujours par `UpdateObject`.

### Les catégories sont POSÉES — et leur référentiel est VIVANT (16/09/2026)

Elles étaient d'abord rangées parmi les champs montrés, faute de référentiel. Le
relevé dans l'éditeur a levé les deux objections :

* **`sdk.DataModel.Venues.getAllVenueCategories()` rend les 132 catégories déjà
  TRADUITES dans la langue de l'utilisateur.** Le script ne fige donc aucune
  copie : `chargerCategories()` remplit la table au premier fichier ouvert. Une
  copie écrite en dur serait fausse partout ailleurs qu'en français, et périmée
  au premier ajout de Waze.
* ⭐⭐ **Les catégories sont le SEUL champ que le SDK valide vraiment** : une
  valeur inconnue est refusée (« categories[0] must match the configured type »),
  et un parking ne peut en porter qu'une (« Parking lot and Charging station
  can't have more than one category »). Partout ailleurs il se tait.

⚠️ **Si le référentiel n'a pas pu être chargé, il reste VIDE et toute catégorie du
classeur est REFUSÉE — donc signalée.** Poser une catégorie qu'on n'a pas pu
vérifier serait pire que de ne pas la poser : c'est elle qui commande les champs
disponibles du lieu.
⚠️ Retirer une catégorie d'un lieu qui en porte plusieurs est une **perte**, donc
une pastille orange et une ligne non cochée (voir plus bas).

⭐ **L'aperçu ne montre en vert que ce qui CHANGE**, et compte le reste (« déjà
conforme : 3 ») : un classeur portant l'état complet d'un parc affichait sept
pastilles par lieu là où rien n'était à faire.

### 🔴 Les deux défauts que SEUL l'essai dans WME a trouvés (16/09/2026)

**Aucun banc ne pouvait les voir**, et ils étaient tous deux silencieux :

1. l'indicateur « à jour » ne regardait que le nom et la description : l'aperçu
   annonçait « ✅ Aucune modification » sur un parking qui avait quatre champs à
   poser. On ferme la fenêtre en confiance, et rien n'est appliqué ;
2. la liste à appliquer se construisait **sans les valeurs** : le SDK n'était
   jamais appelé, aucune action n'entrait dans la pile, et l'écran annonçait
   « ✔ 1 POI appliqué avec succès » sur un lieu intact.

⭐⭐⭐⭐ **UNE CHAÎNE QUI SE COUPE ENTRE L'APERÇU ET L'APPLICATION NE SE VOIT QUE SUR
LA CARTE.** L'aperçu, lui, était juste — et c'est lui qu'on regarde.

✅ **Éprouvé dans WME le 16/09** sur les parkings de la gare TGV d'Avignon :
18 POI chargés, l'aperçu signalant 2 lieux à modifier sur 9, application sur un
seul — les quatre champs vides retrouvés dans le lieu après écriture
(`UNKNOWN → FREE`, `[] → CASH`, `[] → STREET_LEVEL`, services `[] → PMR +
Surveillance`), puis **annulation vérifiée** : lieu identique à l'avant, pile
d'actions revenue à son étalon. Rien n'a été enregistré.

---

## 4. 🔴 Le préchargement — le point délicat du script

Un lieu ne peut être modifié que s'il est **chargé dans le modèle** de WME. Le script recadre donc
la carte sur chaque permalink avant d'agir (`centerAndLoad`).

### 4.1 Le zoom est volontairement large : 16 à 17

**Le lat/lon d'un permalink cadre souvent la CARTE, pas le POI.** Cas mesuré : un lieu à **361 m**
du point du permalink. À zoom 19, un POI décalé tombe **hors des tuiles chargées** et n'est jamais
trouvé — il ressort « non chargé », alors qu'il existe.

Le script suit donc le `zoomLevel` du permalink **en le bornant à [16, 17]**.

**Ne pas resserrer ces bornes** sans refaire la mesure : c'est le réglage qui décide du taux
d'échec du script.

### 4.2 « Chargé » ne veut pas dire « modifiable »

Le sondage attend un lieu **présent dans le modèle** (`lieuPret` : `v && v.attributes`), avec ou
sans nom — jusqu'en 0.52 il exigeait un nom, et un lieu sans nom était déclaré introuvable alors
que lui en donner un est l'usage même de la colonne Nom. Il **n'exige pas `isEditable()`** : un
lieu verrouillé au-dessus du rang de l'éditeur **est bien chargé**, il sera simplement proposé en
*Suggest an Edit*.

⚠️ **Ce que WME fait d'un `UpdateObject` sur un lieu SaE n'est PAS MESURÉ.** La relecture après
pose dira ce qui a réellement changé (§ 5), et le bilan compte ces lieux à part.

`getLockStatus(venue)` rend trois valeurs :

| Valeur | Sens |
|---|---|
| `ok` | Édition directe possible |
| `sae` | Verrou supérieur au rang de l'éditeur → *Suggest an Edit* |
| `hard` | Verrou de niveau 7 → **staff Waze uniquement** |

(`lockRank` : 0 = L1 … 5 = L6, 6 = L7 staff ; comparé à
`W.loginManager.user.attributes.rank`.)

### 4.3 Le sondage

Scrutation toutes les **80 ms**, délai maximal **4 000 ms** au préchargement, **3 000 ms** à
l'application. Un lieu déjà en mémoire n'est **pas** recadré — ni au préchargement, ni à
l'application depuis la 0.53. `preloadVenues` accepte un `cancelRef.cancelled` pour **interrompre
proprement** une longue boucle, et signale sa progression.

⭐ **La vue n'est plus restaurée** (depuis la 0.51) : après un balayage ou une pose, la carte est
**cadrée sur le périmètre des lieux** (`cadrerSurLesLieux`, zoom borné à [12, 19]). Revenir au
point de départ faisait balayer le terrain pour ne rien montrer de ce qu'on venait de charger.

⚠️ **Le calque « Lieux » est allumé s'il est éteint, puis RELU** : un clic qui n'a rien allumé ne
vaut pas un calque allumé, et le message le dit dans la fenêtre (plus d'`alert()`).

---

## 5. L'application

`appliquerLignes()` → `poserLesLignes(items, env)` → `poserUnLieu(item, env)`, dans le bloc
`banc:poser` que `tools/banc-application.mjs` extrait. Tout ce qui touche WME passe par `env`
(`environnementDePose()` dans l'éditeur, des bouchons non complaisants dans le banc).

1. les lignes cochées sont lues **une fois**, en chaînes (`lignesCochees`) : un redessin ne peut
   plus changer ce qui est posé ; l'aperçu est **figé** au clic (`const apercu = _apercu`) ;
2. le lieu est pris dans le modèle, ou chargé s'il n'y est pas (`centerAndLoad`, 3 s) ;
3. relecture de l'**ancien** nom et de l'**ancienne** description — c'est ce qui alimente le rapport ;
4. une action `UpdateObject` (`require('Waze/Action/UpdateObject')`) avec `id`, `name`,
   `description` et `aliases` (ceux du classeur, sinon ceux du lieu) ; puis les champs du lot D2
   par `sdk.DataModel.Venues.updateVenue`, filtrés par la liste blanche ;
5. ⭐⭐ **LA RELECTURE FAIT FOI, POUR TOUT** — nom et description compris depuis la 0.53. Ce qui
   n'est pas retrouvé dans le lieu est un **manque** ; la ligne est alors `partial`.

⚠️⚠️ **UNE EXCEPTION N'ARRÊTE PLUS LA POSE** (0.53) : chaque lieu a son `try`, et un lieu qui lève
devient `erreur`, avec son message. Jusqu'en 0.52, une exception au lieu k laissait k-1 actions
dans la pile, la barre figée, et aucun bilan.

⚠️⚠️ **RIEN NE SE FAIT DEUX FOIS** : pendant un balayage ou une pose, le bouton, les cases, le
menu d'onglet, le bouton de fichier et le dépôt sont bloqués (`occuper`). Après la pose, les
lignes posées sont **décochées** et relues sur le lieu vivant (✔ si tout est posé, l'écart restant
sinon) ; les lignes en échec restent cochées, et **Réessayer** ne repose qu'elles.

Statuts : `applied`, `partial`, `timeout` (introuvable), `erreur`. Le bilan compte chacun, nomme
les lieux en échec, compte à part les lieux **SaE**, et dit combien d'actions **cette pose** a
ajoutées à la pile (rien quand on ne peut pas le lire — plus jamais « 0 » sur une erreur).
L'historique note « appliqué » dès qu'au moins un lieu a été posé.

⚠️ Le script passe par `require('Waze/Action/UpdateObject')` et par `W.model` — **pas par le SDK**.
C'est le point d'attache le plus fragile du script : c'est là qu'il cassera le jour où WME changera.

---

## 6. Le rapport Excel

`exportReport(eventName, results)` produit un classeur d'une feuille, **dans la langue du script** :

| Colonne | Contenu |
|---|---|
| 1 | Permalien |
| 2 | Nom avant |
| 3 | Nom après |
| 4 | Description avant |
| 5 | Description après |
| 6 | Statut (appliqué / partiel / introuvable / erreur, « (SaE) » s'il y a lieu) |
| 7 | **Champs non posés** — ce qu'il faut reprendre avant d'enregistrer |
| 8 | Message d'erreur |

Nom de feuille : `AAAA-MM-JJ <événement>`, caractères interdits par Excel remplacés, **tronqué à 31
caractères**. Nom de fichier : `POI_Report_<événement>_<AAAA-MM-JJ>.xlsx`.

⚠️ **XLSX en version *lite* ne gère pas les styles** : l'en-tête n'est pas coloré, et c'est
volontaire — ne pas ajouter de dépendance pour cela.

---

## 7. Interface

**Le panneau porte les réglages, la fenêtre porte le travail** (depuis la 0.52) : le panneau
latéral de WME vide son contenu dès qu'on sélectionne un objet, on ne peut pas y travailler.

**L'onglet Scripts suit la charte commune** (WCT, WJN, WRP, WDA — valeurs relevées dans WME le
25/09/2026) : l'icône du script (la **même** que `@icon`, sur l'onglet, en tête du panneau, dans la
fenêtre et sur le bouton de la carte), `POI Event Updater vX.YY.ZZ`, la pastille de nouvelle
version, une phrase d'introduction, le bouton qui ouvre la fenêtre, les fichiers récents, l'aide
**repliable** (trois volets fermés), et au pied 💬 Discuss · 🔗 GreasyFork · GitHub puis « ✍️
Appliquer écrit dans l'éditeur ; le script n'enregistre jamais » — **pas** le « 🔒 ne modifie
jamais la carte » des scripts voisins : celui-ci écrit.

Couleurs : **#2196f3** pour les titres et les accents, **#1976d2** pour ce qui est plein avec du
texte blanc (4,60:1 — décision de l'auteur, le #2196f3 n'y donne que 3,12:1). Un seul bouton plein
par écran : Appliquer ; le filtre « ≠ Écarts » actif se marque d'un contour.

La **fenêtre** (bouton dans `.overlay-buttons-container`, `order:100`) porte le bandeau du fichier
et de l'onglet, le tableau (case en première colonne, état en trois signes : liseré, fond, badge),
le filtre, le tri par nom ou description, et le pied avec Appliquer. Elle est **non modale**
(`role="dialog"`) ; le focus y entre à l'ouverture et revient au bouton de la carte à la fermeture.

⭐⭐⭐ **UNE RETOUCHE DANS L'APERÇU RE-DÉCIDE LA CASE** (`caseApresRetouche`) : corriger le nom
d'une ligne « = » la coche ; revenir à la valeur du lieu la décoche ; vider le nom la décoche ; une
ligne qui RETIRE n'est jamais touchée. La 0.52 l'avait perdu : la retouche était jetée en silence.

Les **anomalies du classeur** restent visibles au-dessus du tableau (compte toujours visible,
liste repliée au-delà de trois) — jusqu'en 0.52, le rendu du tableau les effaçait.

### 7.1 La géométrie de la fenêtre

⭐ **Bornée à gauche de la colonne des boutons de la carte, mesurée à chaque geste**
(`bornesCarte` à l'ouverture, au déplacement, au redimensionnement, au redimensionnement du
navigateur). Déplacer par l'en-tête, redimensionner par le coin — **à la souris ou au clavier**
(flèches, Maj pour un grand pas, quand l'en-tête ou le coin a le focus : WCAG 2.5.7). Double-clic
sur l'en-tête : retour à la place par défaut.

⚠️ **Repliée, la fenêtre ne mesure que son en-tête** : on garde alors la hauteur d'avant
(`memoriserGeometrie`), sans quoi elle se rouvrait à 120 px.

---

## 8. Persistance

Tout vit dans le `localStorage`, sans bibliothèque :

| Clé | Contenu |
|---|---|
| `peu_file_history` | Les **5 derniers fichiers** (`HISTORY_MAX = 5`) : nom, date de chargement, date d'application |
| `peu_ui_geom` | Position et taille de la fenêtre (`{x, y, w, h}`) |
| `peu_maj` | Dernière vérification de version : `{t, v}` (au plus une par 24 h) |

`recordFileLoaded` / `recordFileApplied` distinguent explicitement **« chargé »** et
**« appliqué »** : un fichier chargé mais jamais appliqué s'affiche « Jamais appliqué ». ⚠️
`recordFileApplied` n'était plus appelée depuis la refonte de la 0.52 : l'historique disait
« jamais appliqué » de tout. Rétabli en 0.53.

---

## 9. Internationalisation

**Huit langues** (charte commune) : `fr`, `en`, `de`, `es`, `it`, `pt-BR`, `pt-PT`, `he`.
Détection sur `W.userscripts.state.locale`, puis `document.documentElement.lang`, puis
`navigator.language` ; le portugais se décide par le pays ; toute autre langue → anglais. L'hébreu
passe la fenêtre et le panneau en `dir="rtl"` (le CSS est en propriétés logiques), les champs de
saisie sont en `dir="auto"`.

Le dictionnaire est **mémoïsé** (`_strings`), construit au premier appel de `t()`. Il porte aussi
les libellés des colonnes (`ch…`), les motifs des champs montrés (`mo…`) et les valeurs de WME
(`va…`) : `CHAMPS.libelle` et `VALEURS_WME` restent **la donnée**, en français, et le repli quand
aucune traduction n'est fournie. ⚠️ La **lecture** du classeur, elle, reconnaît les libellés
français et les clés WME, pas ceux des autres langues — seules les catégories sont lues dans la
langue de l'éditeur.

Dates et nombres dans la langue du script (`localeDuScript`) ; noms de fichiers en `AAAA-MM-JJ`.
Français : espace fine insécable avant `: ; ! ?`, insécable dans « » — `check-libelles` le vérifie.

⚠️ `_peuLang` est initialisé **dans `initScript`, avant tout appel à `t()`**.

---

## 10. Contraintes non négociables

1. **`@grant GM_xmlhttpRequest` et `unsafeWindow`, `@connect update.greasyfork.org` seulement** —
   pour la pastille de nouvelle version (la politique de sécurité de WME interdit d'appeler
   GreasyFork depuis la page). ⚠️⚠️ **Accorder une permission place le script dans un bac à
   sable** : `W`, `OpenLayers`, `require` et `getWmeSdk` ne s'y lisent que par `unsafeWindow`.
   Ces quatre noms sont donc **déclarés dans le script** et posés au démarrage (`_peuInit`) ;
   `banc-demarrage` fait tourner le script dans les deux mondes (page et bac à sable).
2. **Ne jamais appliquer sans avoir montré.** Le tableau des changements est une étape obligatoire.
3. **Ne jamais taire un échec.** Un `timeout` figure dans le rapport et dans le pied d'échec.
4. **Ne pas resserrer les bornes de zoom [16, 17]** sans refaire la mesure (§ 4.1).
5. **Ne pas exiger `isEditable()`** pour considérer un lieu chargé (§ 4.2).
6. **Repasser les `aliases`** tels quels dans chaque `UpdateObject` quand le classeur n'en porte pas.
7. **Cadrer la carte sur le périmètre des lieux** après un balayage ou une pose (§ 4.3) — la
   restauration de la vue a été retirée en 0.51, ne pas la réintroduire.
8. **Ne jamais publier de fichier de données réelles** (`ACO Events.xlsx` et assimilés), **ni un
   rapport d'audit** (`AUDIT-*.md`, exclus par `.gitignore`).
9. **Ne jamais cesser d'envoyer une description vide**, ni décocher d'office son effacement : le
   retour à l'onglet ordinaire en dépend (§ 3).

---

## 11. Ce qui reste ouvert et fragile

- **Le harnais ne voit pas la carte.** Quinze bancs et contrôles (liste dans `ETAT_ET_REPRISE.md`),
  dont `banc-application` pour le chemin qui écrit, avec des bouchons qui refusent l'inconnu.
  **Ni `UpdateObject`, ni le vrai SDK, ni un lieu SaE** ne vivent dans un banc : l'essai d'Appliquer
  dans WME, sans enregistrer, reste le dernier mot.
- **L'attache à `W.model` et à `require('Waze/Action/UpdateObject')`** n'est pas du SDK : c'est le
  point qui cassera en premier lors d'une évolution de WME. Une migration vers
  `sdk.DataModel.Venues` serait le chantier naturel, et devrait conserver la sémantique du § 4.2.
- **Le format du fichier n'est pas versionné** : le gabarit `.xlsx` est la seule référence. Depuis
  0.49 (§ 3.0), une colonne **ajoutée** est ignorée par les versions qui ne la connaissent pas, et
  l'ordre n'est plus un contrat — mais **rien ne dit à l'utilisateur qu'une colonne est ignorée
  faute d'être comprise** : un fichier plus riche que le script reste muet.
- **Ce que WME fait d'une pose sur un lieu SaE n'est pas mesuré** : la relecture dira ce qui a
  changé, et le bilan comme le rapport marquent ces lieux « (SaE) ».
- **La lecture du classeur reconnaît les valeurs en français et les clés WME** : un classeur
  rempli en allemand avec les libellés allemands de WME verrait ses valeurs refusées (et signalées).

### Annexes du dépôt

| Fichier | Contenu |
|---|---|
| `WME_POI_Event_Updater_Template.xlsx` | Gabarit vierge — **la référence du format** |
| `tools/` | Bancs et contrôles — la liste et l'angle mort de chacun sont dans `ETAT_ET_REPRISE.md` |
| `tools/banc-chargement.mjs` | Banc du chargement d’un classeur — rejoue le code du script (SheetJS **0.20.3** à fournir) |
| `Capture 0.*.png` | Captures par version, publiées avec les annonces |
| `Descr. GreasyFork 0.53.md` | Descriptions GreasyFork EN et FR, à recopier à la publication |
| `Archives/` | Anciennes versions (ignoré par git) |
| `ACO Events.xlsx` | **Données réelles — ignoré par git, ne pas publier** |
