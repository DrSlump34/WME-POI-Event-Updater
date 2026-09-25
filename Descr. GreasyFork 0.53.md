# Descriptions GreasyFork — 0.53.00

> À recopier dans le formulaire de nouvelle version (`#script-version-additional-info-0` = EN,
> `#script-version-additional-info-1` = FR), en **Markdown**. Remplacent celles publiées jusqu'à la
> 0.52, qui annonçaient encore le retour de la carte à sa position d'origine, le bouton
> « Réessayer » disparu, une interface bilingue et des valeurs « dans la langue de votre éditeur ».
> ⚠️ La capture (champ à part du formulaire) est à refaire  l'interface a changé de couleurs et
> le panneau suit la charte commune — cocher « Supprimer » sur l'ancienne.

---

## EN

## Description

**WME POI Event Updater** is a script for Waze editors who manage places (POIs) whose name, description and fields change with the events held there — races, festivals, fairs.

Instead of editing each place by hand in WME, you prepare an Excel workbook once, one sheet per event, and the script applies it: it shows every change first, you tick what to write, and it writes into the editor. **It never saves** — you review on the map, then click Save yourself.

### Features

- One **sheet per event**; an off-event sheet puts places back to their ordinary state
- **Automatic preloading** of every place, with a progress bar; the Places layer is switched on if needed; the map then frames **all the places of the sheet**
- A **before / after** table, with the state of each row in three signs (border, background, badge): fields to write, values that would be **removed** (orange, never ticked by default), lock above your rank (SaE), staff lock (L7), place not found
- **Edit a name or a description in the preview**: the row's box is decided again
- The value a field **replaces** is written in clear next to the new one
- Filter by name, show only the rows that change, sort by name or description
- **Apply** says how many rows it will write — including those hidden by the filter
- After Apply: written rows are unticked and marked ✔, failed places are **named** and stay ticked for **Retry**, and the summary tells how many changes were added to the WME undo stack
- **Excel report**: permalink, before / after, status, **fields not written**, error message
- Workbook checks on load: URL format and parameters, empty names, duplicate permalinks, headers out of place, two columns for one field — all listed, nothing dropped silently
- The window stays left of the map buttons; move and resize it with the mouse **or the keyboard**
- **8 languages**: English, French, German, Spanish, Italian, Portuguese (Brazil and Portugal), Hebrew
- A notice in the Scripts tab when a new version is published

### Excel file format

Each sheet represents one event. Row 1 holds the headers; **columns are matched by header name**, so you may reorder them and add columns of your own.

| POI Permalink | POI Name | POI Description |
|---|---|---|
| Full WME permalink of the place | Name to write | Description to write — **an empty cell erases the place's description** |

The workbook may also carry the place's fields: **categories** (in your editor's language — the first one is the main category), alternative names, phone, website, services, and the parking fields (type, cost, payment, services, situation, spots, exit when closed). Write values as **WME's French labels** or its **internal keys** (`FREE`, `CASH`, `WI_FI`…); separate several values with `;`. **For these fields, a blank cell asks for nothing and erases nothing**, and anything not recognised is shown rather than applied.

Opening hours, address, entry points, parking operator and the Google fields are **read and displayed, but never written** — writing them would mean guessing.

A blank template, `WME_POI_Event_Updater_Template.xlsx`, is in the GitHub repository.

### Installation

1. Install **Tampermonkey** (or a compatible userscript manager).
2. Click **Install** on this page.
3. Open WME: a **POI Event Updater** button appears in the map button column and opens the work window. The script's tab in the Scripts panel holds the recent files, the help and the links.

The script asks for one permission, `GM_xmlhttpRequest` on `update.greasyfork.org`, used only to check for a new version (at most once a day). The Excel library is SheetJS 0.20.3.

**Source code and issues:** https://github.com/DrSlump34/WME-POI-Event-Updater
**Discussion:** https://www.waze.com/discuss/t/script-wme-poi-event-updater/404593

---

## FR

## Description

**WME POI Event Updater** est un script pour les éditeurs Waze qui gèrent des lieux (POI) dont le nom, la description et les champs changent avec les événements qui s'y tiennent — courses, festivals, salons.

Plutôt que de modifier chaque lieu à la main dans WME, vous préparez une fois un classeur Excel, un onglet par événement, et le script l'applique  il montre d'abord chaque changement, vous cochez ce qu'il faut écrire, et il l'écrit dans l'éditeur. **Il n'enregistre jamais** — vous relisez sur la carte, puis vous cliquez vous-même sur Enregistrer.

### Fonctionnalités

- Un **onglet par événement**  un onglet « hors événement » remet les lieux dans leur état ordinaire
- **Préchargement automatique** de chaque lieu, avec une barre de progression  le calque Lieux est allumé s'il le faut  la carte cadre ensuite **tous les lieux de l'onglet**
- Un tableau **avant / après**, l'état de chaque ligne en trois signes (liseré, fond, badge)  champs à poser, valeurs qui seraient **retirées** (orange, jamais cochées d'office), verrou au-dessus de votre rang (SaE), verrou staff (L7), lieu introuvable
- **Retouchez un nom ou une description dans l'aperçu**  la case de la ligne est redécidée
- La valeur qu'un champ **remplace** est écrite en clair à côté de la nouvelle
- Filtre par nom, affichage des seules lignes qui changent, tri par nom ou description
- **Appliquer** dit combien de lignes il va écrire — y compris celles que le filtre masque
- Après Appliquer  les lignes posées se décochent et portent ✔, les lieux en échec sont **nommés** et restent cochés pour **Réessayer**, et le bilan dit combien de modifications ont rejoint la pile d'annulation de WME
- **Rapport Excel**  permalien, avant / après, statut, **champs non posés**, message d'erreur
- Contrôle du classeur au chargement  format et paramètres des URL, noms vides, permaliens en double, en-têtes mal placées, deux colonnes pour un même champ — tout est listé, rien n'est écarté en silence
- La fenêtre reste à gauche des boutons de la carte  déplacement et redimensionnement à la souris **ou au clavier**
- **8 langues**  français, anglais, allemand, espagnol, italien, portugais (Brésil et Portugal), hébreu
- Un avis dans l'onglet Scripts quand une nouvelle version est publiée

### Format du fichier Excel

Chaque onglet représente un événement. La ligne 1 porte les en-têtes  **les colonnes sont repérées par leur en-tête**, vous pouvez donc les réordonner et ajouter les vôtres.

| POI Permalink | POI Name | POI Description |
|---|---|---|
| Permalien WME complet du lieu | Nom à poser | Description à poser — **une cellule vide efface la description du lieu** |

Le classeur peut aussi porter les champs du lieu  **catégories** (dans la langue de votre éditeur — la première est la catégorie principale), noms alternatifs, téléphone, site web, services, et les champs du parking (type, tarif, paiements, services, situation, nombre de places, sortie quand il est fermé). Écrivez les valeurs avec **les libellés français de WME** ou ses **clés internes** (`FREE`, `CASH`, `WI_FI`…)  séparez plusieurs valeurs par `;`. **Pour ces champs, une cellule vide ne demande rien et n'efface rien**, et ce qui n'est pas reconnu est montré au lieu d'être posé.

Horaires, adresse, points d'entrée, opérateur de parking et champs Google sont **lus et affichés, jamais écrits** — les écrire reviendrait à deviner.

Un gabarit vierge, `WME_POI_Event_Updater_Template.xlsx`, est dans le dépôt GitHub.

### Installation

1. Installez **Tampermonkey** (ou un gestionnaire de scripts compatible).
2. Cliquez sur **Installer** sur cette page.
3. Ouvrez WME  un bouton **POI Event Updater** apparaît dans la colonne des boutons de la carte et ouvre la fenêtre de travail. L'onglet du script, dans le panneau Scripts, garde les fichiers récents, l'aide et les liens.

Le script demande une seule permission, `GM_xmlhttpRequest` vers `update.greasyfork.org`, pour vérifier seulement s'il existe une nouvelle version (au plus une fois par jour). La bibliothèque Excel est SheetJS 0.20.3.

**Code source et anomalies :** https://github.com/DrSlump34/WME-POI-Event-Updater
**Discussion :** https://www.waze.com/discuss/t/script-wme-poi-event-updater/404593
