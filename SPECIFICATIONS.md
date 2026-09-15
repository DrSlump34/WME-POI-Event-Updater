# PEU — WME POI Event Updater · Dossier de spécifications

> **Version du code décrite ici : 0.49** (lue dans le bloc `==UserScript==` de
> `WME_POI_Event_Updater.user.js`).
> Diffusé sur **GreasyFork 578776**, dépôt `github.com/DrSlump34/WME-POI-Event-Updater`,
> fil Discuss **404593**.

---

## 0. À qui s'adresse ce dossier

Dossier de reprise du projet. Le script fait **1 529 lignes** et n'a qu'**un seul banc**
(`tools/banc-colonnes.mjs`, la lecture des colonnes) : pour tout le reste, ce document et les
commentaires du code sont la seule mémoire de ses choix.

| Document | Rôle |
|---|---|
| `README.md` | Vitrine anglaise : à quoi ça sert, le format du fichier, l'installation |
| **`SPECIFICATIONS.md`** (ce fichier) | **Normatif** : le contrat, le modèle de données, les invariants |
| `WME_POI_Event_Updater_Template.xlsx` | Le **gabarit vierge** du fichier d'entrée — c'est lui qui fait foi sur les colonnes |
| `Descr. HTML GreasyFork 0.23.txt` | La description publiée, en HTML |

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
- Fonctionner en **français et en anglais**.

### 2.2 Ce que PEU ne fait pas

- **Il ne crée pas de lieux** et n'en supprime aucun : il met à jour `name` et `description`.
- **Il ne touche pas aux alias** : ils sont relus et repassés tels quels.
- **Il n'enregistre pas.** Les modifications sont posées dans la pile d'annulation de WME.
- **Il n'écrit rien hors du navigateur** : `@grant none`, aucun appel réseau propre au script.

---

## 3. Le fichier d'entrée

**Une ligne par lieu, un fichier par événement.** Gabarit vierge :
`WME_POI_Event_Updater_Template.xlsx`.

| Colonne | Contenu |
|---|---|
| `POI Permalink` | Permalink WME pointant le lieu |
| `POI Name` | Nom à poser |
| `POI Description` | Description à poser |

Lecture par **SheetJS (`xlsx` 0.18.5)**, chargé par `@require`. ⚠️ **Deux `@require` pointent la
même bibliothèque** — cdnjs **puis** jsDelivr : c'est un **repli** volontaire, si un CDN est
inaccessible l'autre sert.

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

⇒ Les colonnes peuvent être **réordonnées**, et des colonnes **supplémentaires sont ignorées** :
c'est ce qui ouvre l'enrichissement du format (lot D2 de l'extranet EVIDRA).

⚠️ **Le numéro de ligne des anomalies est celui du tableur**, calculé depuis `!ref` : une feuille
dont la plage ne commence pas en A1 ne renvoie plus à une ligne introuvable.

✅ **Éprouvé** : `node tools/banc-colonnes.mjs` — 19 cas, **7 mutations sur 7 mordent**. Le banc
**extrait le bloc du userscript lui-même**, il n'en garde pas de copie. Et la lecture réelle a été
rejouée hors WME sur `ACO Events.xlsx` (**342 lignes, 6 onglets**) et sur le gabarit : lignes et
numéros de ligne **identiques** à ceux de 0.48.

### 3.1 Le permalink, et ce qu'on en tire

`getVenueIdFromPermalink(url)` lit `venues=` : l'identifiant peut être **numérique (ancien) ou un
GUID alphanumérique (nouveau)**, et il peut y en avoir plusieurs séparés par des virgules —
**on prend le premier**.

`parseLatLon(url)` lit `lat` / `lon` **en gérant les valeurs négatives** (hémisphère sud, ouest de
Greenwich) : c'est ce qui rend le script utilisable ailleurs qu'en Europe.

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

Le sondage attend un lieu **présent avec un nom**. Il **n'exige pas `isEditable()`** : un lieu
verrouillé au-dessus du rang de l'éditeur **est bien chargé**, il sera simplement proposé en
*Suggest an Edit*.

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
l'application. Un lieu déjà en mémoire n'est **pas** recadré. `preloadVenues` accepte un
`cancelRef.cancelled` pour **interrompre proprement** une longue boucle, et signale sa progression.

⚠️ **La vue de l'utilisateur est restaurée** après l'opération (`savedCenter` / `savedZoom`) : le
script emprunte la carte, il ne la garde pas.

---

## 5. L'application

`runApply(items, allResults)` :

1. recharge le lieu (`centerAndLoad`, 3 s) ;
2. relit l'**ancien** nom et l'**ancienne** description — c'est ce qui alimente le rapport ;
3. pose une action `UpdateObject` (`require('Waze/Action/UpdateObject')`) dans
   `W.model.actionManager`, avec `id`, `name`, `description`, et **`aliases` repassés tels quels** ;
4. avance la barre de progression ;
5. à la fin, **restaure la vue** puis affiche le pied de page d'export.

Chaque ligne produit un résultat de statut `applied` ou `timeout`. **Les échecs sont regroupés dans
un rapport d'échec** affiché à l'écran, et l'historique n'enregistre « appliqué » que si l'opération
s'est terminée sans échec.

⚠️ Le script passe par `require('Waze/Action/UpdateObject')` et par `W.model` — **pas par le SDK**.
C'est le point d'attache le plus fragile du script : c'est là qu'il cassera le jour où WME changera.

---

## 6. Le rapport Excel

`exportReport(eventName, results)` produit un classeur d'une feuille :

| Colonne | Contenu |
|---|---|
| 1 | Ancien nom |
| 2 | Nouveau nom |
| 3 | Ancienne description |
| 4 | Nouvelle description |
| 5 | Statut (`appliqué` / `délai dépassé`) |

Largeurs fixées (30, 30, 45, 45, 15). Nom de feuille : `JJ-MM-AAAA <événement>`, **tronqué à 31
caractères** (limite Excel). Nom de fichier : `POI_Report_<événement>_<date>.xlsx`.

⚠️ **XLSX en version *lite* ne gère pas les styles** : l'en-tête n'est pas coloré, et c'est
volontaire — ne pas ajouter de dépendance pour cela.

---

## 7. Interface

Le script s'installe dans le **panneau latéral** par `W.userscripts.registerSidebarTab(scriptId)`,
avec une **icône de pin à la place du nom** (le nom reste en infobulle).

L'écran de travail est un **overlay déplaçable** (`makeDraggable`), avec :

- le choix du fichier et la liste des **fichiers récents** ;
- un **filtre par nom** ;
- le tableau : recherche, nom, description — **éditables** ;
- une barre de progression pendant le préchargement puis pendant l'application ;
- un pied de page d'export une fois l'application terminée.

Deux points d'accessibilité tenus dans le code : chaque bouton porte un `title` **recopié en
`aria-label`**.

### 7.1 La géométrie de l'overlay — un défaut à ne pas refaire

Historique de la 0.48, en deux temps :

1. l'overlay se retrouvait **collé en haut à gauche** parce qu'une géométrie **sauvée à zéro** était
   relue telle quelle ;
2. la correction a été de **ne plus mémoriser la hauteur** : elle **s'adapte au contenu**.

⚠️ `makeDraggable` ne persiste la position qu'**après** que la position initiale a été posée
(`geomReady`) — sans ce drapeau, on réécrit un zéro.

---

## 8. Persistance

Tout vit dans le `localStorage`, sans bibliothèque :

| Clé | Contenu |
|---|---|
| `peu_file_history` | Les **5 derniers fichiers** (`HISTORY_MAX = 5`) : nom, date de chargement, date d'application |
| `peu_overlay_geom` | Position de l'overlay (**pas la hauteur**, § 7.1) |

`recordFileLoaded` / `recordFileApplied` distinguent explicitement **« chargé »** et
**« appliqué »** : un fichier chargé mais jamais appliqué s'affiche « ✔ Jamais appliqué ». C'est
une information, pas un vide.

---

## 9. Internationalisation

**Deux langues** : français et anglais. Détection sur `W.userscripts.state.locale`, puis
`document.documentElement.lang`, puis `navigator.language` — **tout ce qui ne commence pas par `fr`
est traité comme anglais**.

Le dictionnaire est **mémoïsé** (`_strings`), construit au premier appel de `t()` : sans cela,
l'objet entier serait reconstruit à chaque appel.

⚠️ `_peuLang` est initialisé **dans `initScript`, avant tout appel à `t()`**.

---

## 10. Contraintes non négociables

1. **`@grant none`** — le script ne fait aucun appel réseau propre ; seuls les deux `@require`
   chargent SheetJS.
2. **Ne jamais appliquer sans avoir montré.** Le tableau des changements est une étape obligatoire.
3. **Ne jamais taire un échec.** Un `timeout` figure dans le rapport et dans le pied d'échec.
4. **Ne pas resserrer les bornes de zoom [16, 17]** sans refaire la mesure (§ 4.1).
5. **Ne pas exiger `isEditable()`** pour considérer un lieu chargé (§ 4.2).
6. **Repasser les `aliases`** tels quels dans chaque `UpdateObject`.
7. **Restaurer la vue de l'utilisateur** après toute opération qui déplace la carte.
8. **Ne jamais publier de fichier de données réelles** (`ACO Events.xlsx` et assimilés).

---

## 11. Ce qui reste ouvert et fragile

- **Presque aucun harnais de test.** `tools/banc-colonnes.mjs` (0.49) tient la lecture des
  colonnes ; **tout le reste** — préchargement, aperçu, application, rapport — ne se vérifie que
  dans WME.
- **L'attache à `W.model` et à `require('Waze/Action/UpdateObject')`** n'est pas du SDK : c'est le
  point qui cassera en premier lors d'une évolution de WME. Une migration vers
  `sdk.DataModel.Venues` serait le chantier naturel, et devrait conserver la sémantique du § 4.2.
- **Le format du fichier n'est pas versionné** : le gabarit `.xlsx` est la seule référence. Depuis
  0.49 (§ 3.0), une colonne **ajoutée** est ignorée par les versions qui ne la connaissent pas, et
  l'ordre n'est plus un contrat — mais **rien ne dit à l'utilisateur qu'une colonne est ignorée
  faute d'être comprise** : un fichier plus riche que le script reste muet.
- **Le rapport ne distingue pas `ok` / `sae` / `hard`** : `getLockStatus` est calculé et affiché à
  l'écran, mais le statut du rapport se limite à `applied` / `timeout`.

### Annexes du dépôt

| Fichier | Contenu |
|---|---|
| `WME_POI_Event_Updater_Template.xlsx` | Gabarit vierge — **la référence du format** |
| `tools/banc-colonnes.mjs` | Banc de la lecture des colonnes — `node tools/banc-colonnes.mjs` |
| `Capture 0.*.png` | Captures par version, publiées avec les annonces |
| `Descr. HTML GreasyFork 0.23.txt` | Description publiée |
| `Archives/` | Anciennes versions (ignoré par git) |
| `ACO Events.xlsx` | **Données réelles — ignoré par git, ne pas publier** |
