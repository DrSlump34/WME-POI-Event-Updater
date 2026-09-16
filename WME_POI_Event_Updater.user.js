// ==UserScript==
// @name         WME POI Event Updater
// @name:fr      WME POI Event Updater
// @namespace    http://tampermonkey.net/
// @version      0.52
// @description  Bulk-update WME POI names and descriptions per event via Excel file
// @description:fr Mise à jour en masse des POI WME par événement via un fichier Excel
// @author       DrSlump34
// @copyright    DrSlump34 2025
// @license      MIT
// @match        https://www.waze.com/fr/editor?env=row*
// @match        https://www.waze.com/*/editor*
// @match        https://www.waze.com/editor*
// @match        https://waze.com/editor*
// @match        https://waze.com/*/editor*
// @match        https://beta.waze.com/*/editor*
// @match        https://beta.waze.com/fr/editor?env=row*
// @grant        none
// @icon         data:image/svg+xml;base64,PHN2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHdpZHRoPSc2NCcgaGVpZ2h0PSc2NCcgdmlld0JveD0nMCAwIDEyOCAxMjgnPgogIDxyZWN0IHdpZHRoPScxMjgnIGhlaWdodD0nMTI4JyByeD0nMjQnIGZpbGw9JyMyQzZFRDUnLz4KICA8ZyB0cmFuc2Zvcm09J3JvdGF0ZSgtMTggNjQgNjQpJz4KICAgIDxwYXRoIGQ9J000MCA0MCBMODYgNDAgTDEwMiA2NCBMODYgODggTDQwIDg4IFonIGZpbGw9J3doaXRlJy8+CiAgICA8Y2lyY2xlIGN4PSc1MicgY3k9JzY0JyByPSc2JyBmaWxsPScjMkM2RUQ1Jy8+CiAgPC9nPgogIDxnIGZpbGw9J25vbmUnIHN0cm9rZT0nI0ZGQzQwMCcgc3Ryb2tlLXdpZHRoPSc2JyBzdHJva2UtbGluZWNhcD0ncm91bmQnPgogICAgPHBhdGggZD0nTTQ0IDEwMCBBMjIgMjIgMCAwIDEgODQgOTInLz4KICAgIDxwYXRoIGQ9J004NCAyOCBBMjIgMjIgMCAwIDEgNDQgMzYnLz4KICA8L2c+CiAgPHBvbHlnb24gcG9pbnRzPSc4NCw4NCA5Miw5NCA3OCw5OCcgZmlsbD0nI0ZGQzQwMCcvPgogIDxwb2x5Z29uIHBvaW50cz0nNDQsNDQgMzYsMzQgNTAsMzAnIGZpbGw9JyNGRkM0MDAnLz4KPC9zdmc+
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @downloadURL  https://update.greasyfork.org/scripts/578776/WME%20POI%20Event%20Updater.user.js
// @updateURL    https://update.greasyfork.org/scripts/578776/WME%20POI%20Event%20Updater.meta.js
// ==/UserScript==

(function() {
    'use strict';

    const scriptId   = 'poi-event-updater';
    const HISTORY_KEY = 'peu_file_history'; // clé localStorage
    const GEO_KEY     = 'peu_ui_geom';      // position et taille de la fenêtre
    /* ⚠️ LA VERSION EST LUE DANS L’EN-TÊTE DU SCRIPT, jamais recopiée : deux
       exemplaires d’un numéro de version divergent le jour où l’on bumpe. */
    /* ⚠️⚠️ `typeof` ET NON UN SIMPLE TEST : avec `@grant none`, rien ne garantit
       que `GM_info` existe, et une ReferenceError ici ne casserait pas une
       ligne — elle tuerait le script entier au chargement, sans un mot. */
    const PEU_VERSION = (typeof GM_info !== 'undefined' && GM_info.script && GM_info.script.version)
        ? GM_info.script.version : '?';
    const HISTORY_MAX = 5;
    const GEOM_KEY = 'peu_overlay_geom';    // taille + position mémorisées de l'overlay
    // Icône de l'onglet : pin de localisation (= POI), détouré, affiché à la place du nom
    const TAB_ICON = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHZpZXdCb3g9JzAgMCAyNCAyNCcgd2lkdGg9JzI0JyBoZWlnaHQ9JzI0Jz48cGF0aCBmaWxsPScjMkM2RUQ1JyBkPSdNMTIgMkM4LjEzIDIgNSA1LjEzIDUgOWMwIDUuMjUgNyAxMyA3IDEzczctNy43NSA3LTEzYzAtMy44Ny0zLjEzLTctNy03eicvPjxjaXJjbGUgY3g9JzEyJyBjeT0nOScgcj0nMi42JyBmaWxsPScjZmZmZmZmJy8+PC9zdmc+';
    let poiData = [];
    /* ⚠️ LA POIGNEE VERS LE CHAMP DE FICHIER, pas une seconde lecture : le
       chemin qui lit le classeur valide vingt règles, tient un rapport
       d’anomalies et alimente l’historique. Le déplacer aurait été échanger
       une interface contre un risque. */
    let _peuFileInput = null;
    let _peuLang = 'en'; // initialisé dans initScript avant tout appel à t()

    // Détection langue
    function detectLang() {
        try {
            const l = W?.userscripts?.state?.locale || document.documentElement.lang || navigator.language || 'en';
            return l.toLowerCase().startsWith('fr') ? 'fr' : 'en';
        } catch { return 'en'; }
    }

    // Dictionnaire i18n construit une seule fois (mémoïsé) au 1er appel, puis réutilisé
    // — évite de reconstruire tout l'objet à chaque appel de t().
    let _strings = null;
    function t(key, ...args) {
        if (!_strings) _strings = {
            fr: {
                tabTitle:'POI Events', tabTooltip:'Mise à jour POI via Excel',
                panelTitle:'POI Event Updater', chooseFile:'📂 Choisir un fichier',
                chooseFileTitle:'Charger un classeur .xlsx depuis votre disque',
                selectSheet:'L’onglet du classeur à poser',
                fabTitle:'POI Event Updater — afficher la fenêtre',
                fabTitleOn:'POI Event Updater — masquer la fenêtre',
                btnApplyNone:'Rien de coché', btnApplyOne:'Appliquer 1 ligne',
                btnApplyN:(n)=>`Appliquer les ${n} lignes cochées`,
                btnApplyTitle:'Poser les valeurs des lignes cochées dans l’éditeur. Rien n’est enregistré : vous relirez dans WME.',
                btnExportTitle:'Enregistrer le rapport de cet aperçu',
                footerHelpVide:'Rien n’est écrit sur la carte tant que vous n’avez pas cliqué sur Appliquer.',
                guideFichier:'Choisissez le classeur de l’événement.',
                guideFichierSuite:'Ensuite vous choisirez l’onglet, puis vous relirez chaque ligne avant d’appliquer.',
                guideOnglet:'Choisissez l’onglet à poser.',
                guideOngletSuite:'Un onglet par événement. Celui « Hors Evenement » remet les lieux dans leur état ordinaire.',
                colSelect:'Poser', colEtat:'État',
                dropLigne1:'📄 Déposez un classeur ici',
                dropLigne2:'ou cliquez pour le choisir',
                dropTitre:'Déposez un fichier .xlsx n’importe où sur cette fenêtre, ou cliquez pour le choisir',
                dropRefus:(nom)=>`« ${nom} » n’est pas un classeur Excel : seuls les fichiers .xlsx et .xls se chargent ici.`,
                sbOuvrir:'Afficher la fenêtre', sbOuvrirTitre:'Ouvrir la fenêtre de travail — c’est là qu’on charge un classeur et qu’on relit avant d’appliquer',
                sbReglages:'Réglages', sbReglagesNote:'Langue, densité d’affichage, comportement de la carte — à venir. L’emplacement est réservé.',
                cancelTitle:'Interrompre : ce qui est déjà lu est conservé',
                footerHelp:'Décochez ce que vous ne voulez pas poser. Les lignes orange RETIRENT des valeurs : elles ne sont jamais cochées d’office.',
                bilanPartiel:(n)=>`⚠️ ${n} lieu(x) n’ont reçu qu’une partie des valeurs — voir le rapport.`,
                bilanEchec:(n)=>`⚠️ ${n} lieu(x) n’ont pas pu être chargés : rien n’y a été posé.`,
                bilanNonEnregistre:(n)=>`${n} modification(s) en attente dans WME — RIEN N’EST ENREGISTRÉ : relisez, puis cliquez sur Enregistrer dans l’éditeur.`,
                cbTitre:'Poser cette ligne dans l’éditeur',
                cbFige:'Cette ligne ne peut pas être posée — voir le badge d’état',
                colNameTitle:'Le nom à poser. Modifiable avant d’appliquer.',
                colDescTitle:'La description à poser. Une cellule vide ne demande rien et n’efface rien.',
                triTitre:'Trier sur cette colonne',
                badgePerteTitle:(n)=>`Appliquer RETIRERAIT ${n} valeur(s) au lieu`,
                badgeDiffTitle:(n)=>`${n} champ(s) à poser`,
                badgeRienTitle:'Rien à poser : le lieu porte déjà ces valeurs',
                badgeOffTitle:'Lieu introuvable dans l’éditeur : rien ne peut être posé',
                noFile:'Aucun fichier choisi', showBtn:'▶ Afficher modifications',
                historyTitle:'Fichiers récents', histLoaded:'📂 Chargé :', histApplied:'✔ Appliqué :',
                histNeverApplied:'✔ Jamais appliqué', filterPlaceholder:'🔍 Filtrer par nom…',
                draggable:'✥ déplaçable',
                colSearch:'🔍', colName:'Nom', colDesc:'Description',
                btnMinimize:'Réduire / Restaurer', btnRestore:'Restaurer',
                btnApply:'Appliquer', btnClose:'Fermer',
                btnDiffActive:'≠ Diff', btnDiffAll:'≡ Tout', btnUpToDate:'✅ À jour',
                tooltipDiffOn:'Afficher tous les POI', tooltipDiffOff:'Afficher uniquement les POI modifiés',
                footerSae:'⚠️ = Suggest an Edit (lock supérieur à votre niveau).',
                footerHard:'🔒 = Verrou L7 staff, non modifiable.',
                emptyMsg:'✅ Aucune modification à apporter pour cet événement.',
                emptyShowAll:'Voir tous les POI',
                lockSaeTitle:(lock,user)=>`Verrou L${lock+1} — votre niveau : L${user+1} → Suggest an Edit`,
                lockHardTitle:'Verrou niveau 7 (staff Waze) — édition impossible',
                unloadedLabel:'⚠ non chargé', unloadedTitle:"Ce POI n'a pas pu être chargé lors du préchargement",
                diffTitle:"Valeur différente de l'état actuel",
                loading:(n,total)=>`Chargement des POI… ${n} / ${total}`,
                applying:(n,total)=>`Application… ${n} / ${total}`,
                poisLoaded:(n)=>`✔ ${n} POI chargés`, noPoisLoaded:'✖ Aucun POI valide chargé',
                anomalies:(n)=>`⚠️ ${n} anomalie${n>1?'s':''} détectée${n>1?'s':''}`,
                timeoutReport:(n)=>`⚠️ ${n} POI non appliqué${n>1?'s':''} (timeout) :`,
                btnRetry:(n)=>`🔄 Réessayer (${n} POI)`,
                successMsg:(n)=>`✔ ${n} POI appliqués avec succès`,
                btnExport:'📥 Exporter le rapport',
                sheetHeaderErr:'En-têtes de colonnes absentes ou non reconnues — onglet ignoré',
                sheetHeaderFallback:'en-têtes non reconnues : colonnes lues par position (A = permalien, B = nom, C = description)',
                plusApplique:'✔ appliqué :', plusMain:'✋ à poser à la main :', plusRefus:'⚠ non reconnu :',
                plusConforme:'· déjà conforme :', plusAvant:'remplace :', plusRetire:'RETIRE',
                urlInvalid:'URL invalide',
                urlBadHost:'URL non reconnue (doit être waze.com ou beta.waze.com/…/editor)',
                urlNoEnv:'paramètre env= manquant', urlBadLat:'lat= absent ou invalide',
                urlBadLon:'lon= absent ou invalide', urlBadZoom:'zoomLevel= absent ou invalide',
                urlNoVenues:'venues= absent', urlNoVid:"impossible d'extraire l'ID du venue",
                nameEmpty:'nom vide', dupRow:(a,b)=>`Lignes ${a} et ${b} : permalink en double`,
                reportHeaders:['POI (Avant Nom)','Après Nom','Avant Desc','Après Desc','Statut'],
                statusApplied:'✔ Appliqué', statusTimeout:'✖ Timeout',
                diffCount:(n)=>`${n} diff`, upToDateTitle:'À jour', noPoi:'Aucun POI trouvé.',
                rowLabel:'Ligne', loadingPois:(n,total)=>`Chargement des POI… ${n} / ${total}`,
                masterCbTitle:'Tout cocher / décocher', searchClearTitle:'Effacer',
                btnReduce:'Réduire', poiCountDiff:(n)=>`${n} POI (diff)`, poiCount:(n)=>`${n} POI`,
                badgeSae:(n)=>`${n} ⚠️ SaE`, badgeSaeTitle:'Ces POI seront soumis en Suggest an Edit',
                badgeHardTitle:'Ces POI sont verrouillés niveau 7 — édition impossible',
                layerOffMsg:'⚠️ Le calque "Lieux" est désactivé dans WME.\n\nActivez-le (menu Calques > Lieux) avant de lancer le script.',
                cancelBtn:'Annuler', preloadCancelled:'Préchargement annulé.', clearHistoryTitle:"Effacer l'historique",
                xlsxMissing:'⚠️ Librairie Excel (XLSX) non chargée. Vérifiez votre connexion ou autorisez cdnjs.cloudflare.com / jsdelivr.net, puis rechargez la page (F5).',
                locateTitle:'Recentrer la carte sur ce POI',
            },
            en: {
                tabTitle:'POI Events', tabTooltip:'Bulk-update POIs via Excel',
                panelTitle:'POI Event Updater', chooseFile:'📂 Choose a file',
                chooseFileTitle:'Load an .xlsx workbook from your disk',
                selectSheet:'Which sheet to apply',
                fabTitle:'POI Event Updater — show the window',
                fabTitleOn:'POI Event Updater — hide the window',
                btnApplyNone:'Nothing ticked', btnApplyOne:'Apply 1 row',
                btnApplyN:(n)=>`Apply the ${n} ticked rows`,
                btnApplyTitle:'Write the ticked rows into the editor. Nothing is saved: you will review in WME.',
                btnExportTitle:'Save the report of this preview',
                footerHelpVide:'Nothing is written to the map until you click Apply.',
                guideFichier:'Choose the event workbook.',
                guideFichierSuite:'Then pick the sheet, and review every row before applying.',
                guideOnglet:'Choose the sheet to apply.',
                guideOngletSuite:'One sheet per event. The « Hors Evenement » one puts places back to their ordinary state.',
                colSelect:'Apply', colEtat:'State',
                dropLigne1:'📄 Drop a workbook here',
                dropLigne2:'or click to pick one',
                dropTitre:'Drop an .xlsx file anywhere on this window, or click to pick one',
                dropRefus:(nom)=>`« ${nom} » is not an Excel workbook: only .xlsx and .xls files load here.`,
                sbOuvrir:'Show the window', sbOuvrirTitre:'Open the work window — that is where you load a workbook and review before applying',
                sbReglages:'Settings', sbReglagesNote:'Language, display density, map behaviour — to come. The place is reserved.',
                cancelTitle:'Stop: what is already loaded is kept',
                footerHelp:'Untick what you do not want to write. Orange rows REMOVE values: they are never ticked by default.',
                bilanPartiel:(n)=>`⚠️ ${n} place(s) only received part of the values — see the report.`,
                bilanEchec:(n)=>`⚠️ ${n} place(s) could not be loaded: nothing was written there.`,
                bilanNonEnregistre:(n)=>`${n} pending change(s) in WME — NOTHING IS SAVED: review, then click Save in the editor.`,
                cbTitre:'Write this row into the editor',
                cbFige:'This row cannot be written — see the state badge',
                colNameTitle:'The name to write. Editable before applying.',
                colDescTitle:'The description to write. An empty cell asks for nothing and erases nothing.',
                triTitre:'Sort on this column',
                badgePerteTitle:(n)=>`Applying would REMOVE ${n} value(s) from the place`,
                badgeDiffTitle:(n)=>`${n} field(s) to write`,
                badgeRienTitle:'Nothing to write: the place already carries these values',
                badgeOffTitle:'Place not found in the editor: nothing can be written',
                noFile:'No file chosen', showBtn:'▶ Show changes',
                historyTitle:'Recent files', histLoaded:'📂 Loaded:', histApplied:'✔ Applied:',
                histNeverApplied:'✔ Never applied', filterPlaceholder:'🔍 Filter by name…',
                draggable:'✥ draggable',
                colSearch:'🔍', colName:'Name', colDesc:'Description',
                btnMinimize:'Minimize / Restore', btnRestore:'Restore',
                btnApply:'Apply', btnClose:'Close',
                btnDiffActive:'≠ Diff', btnDiffAll:'≡ All', btnUpToDate:'✅ Up to date',
                tooltipDiffOn:'Show all POIs', tooltipDiffOff:'Show only modified POIs',
                footerSae:'⚠️ = Suggest an Edit (lock level above yours).',
                footerHard:'🔒 = L7 staff lock, cannot be edited.',
                emptyMsg:'✅ No changes to apply for this event.',
                emptyShowAll:'Show all POIs',
                lockSaeTitle:(lock,user)=>`Lock L${lock+1} — your level: L${user+1} → Suggest an Edit`,
                lockHardTitle:'Level 7 lock (Waze staff) — editing not possible',
                unloadedLabel:'⚠ not loaded', unloadedTitle:'This POI could not be loaded during preloading',
                diffTitle:'Value differs from current state',
                loading:(n,total)=>`Loading POIs… ${n} / ${total}`,
                applying:(n,total)=>`Applying… ${n} / ${total}`,
                poisLoaded:(n)=>`✔ ${n} POIs loaded`, noPoisLoaded:'✖ No valid POI loaded',
                anomalies:(n)=>`⚠️ ${n} anomal${n>1?'ies':'y'} detected`,
                timeoutReport:(n)=>`⚠️ ${n} POI${n>1?'s':''} not applied (timeout):`,
                btnRetry:(n)=>`🔄 Retry (${n} POI${n>1?'s':''})`,
                successMsg:(n)=>`✔ ${n} POI${n>1?'s':''} applied successfully`,
                btnExport:'📥 Export report',
                sheetHeaderErr:'Column headers missing or unrecognised — sheet ignored',
                sheetHeaderFallback:'headers not recognised: columns read by position (A = permalink, B = name, C = description)',
                plusApplique:'✔ applied:', plusMain:'✋ to set by hand:', plusRefus:'⚠ not recognised:',
                plusConforme:'· already correct:', plusAvant:'replaces:', plusRetire:'REMOVES',
                urlInvalid:'Invalid URL',
                urlBadHost:'Unrecognised URL (must be waze.com or beta.waze.com/…/editor)',
                urlNoEnv:'missing env= parameter', urlBadLat:'lat= missing or invalid',
                urlBadLon:'lon= missing or invalid', urlBadZoom:'zoomLevel= missing or invalid',
                urlNoVenues:'venues= missing', urlNoVid:'unable to extract venue ID',
                nameEmpty:'empty name', dupRow:(a,b)=>`Rows ${a} and ${b}: duplicate permalink`,
                reportHeaders:['POI (Before Name)','After Name','Before Desc','After Desc','Status'],
                statusApplied:'✔ Applied', statusTimeout:'✖ Timeout',
                diffCount:(n)=>`${n} diff`, upToDateTitle:'Up to date', noPoi:'No POI found.',
                rowLabel:'Row', loadingPois:(n,total)=>`Loading POIs… ${n} / ${total}`,
                masterCbTitle:'Check / uncheck all', searchClearTitle:'Clear',
                btnReduce:'Minimize', poiCountDiff:(n)=>`${n} POI (diff)`, poiCount:(n)=>`${n} POIs`,
                badgeSae:(n)=>`${n} ⚠️ SaE`, badgeSaeTitle:'These POIs will be submitted as Suggest an Edit',
                badgeHardTitle:'These POIs are level-7 locked — editing not possible',
                layerOffMsg:'⚠️ The "Places" layer is disabled in WME.\n\nPlease enable it (Layers menu > Places) before running the script.',
                cancelBtn:'Cancel', preloadCancelled:'Preloading cancelled.', clearHistoryTitle:'Clear history',
                xlsxMissing:'⚠️ Excel library (XLSX) not loaded. Check your connection or allow cdnjs.cloudflare.com / jsdelivr.net, then reload the page (F5).',
                locateTitle:'Center the map on this POI',
            }
        };
        const val = _strings[_peuLang]?.[key] ?? _strings.en[key] ?? key;
        return typeof val === 'function' ? val(...args) : val;
    }

    // ── Historique ──────────────────────────────────────────────────────────
    function getHistory() {
        try { return JSON.parse(localStorage.getItem(HISTORY_KEY)) || []; }
        catch { return []; }
    }
    function saveHistory(history) {
        try { localStorage.setItem(HISTORY_KEY, JSON.stringify(history)); } catch {}
    }
    function getOverlayGeom() {
        try { return JSON.parse(localStorage.getItem(GEOM_KEY)) || null; }
        catch { return null; }
    }
    function saveOverlayGeom(g) {
        try { localStorage.setItem(GEOM_KEY, JSON.stringify(g)); } catch {}
    }
    function formatDateTime(iso) {
        if (!iso) return '—';
        const d = new Date(iso);
        const loc = _peuLang === 'fr' ? 'fr-FR' : 'en-GB';
        return d.toLocaleDateString(loc) + ' ' + d.toLocaleTimeString(loc, {hour:'2-digit', minute:'2-digit'});
    }
    function recordFileLoaded(fileName) {
        const history = getHistory().filter(h => h.name !== fileName); // déduplique
        history.unshift({ name: fileName, loaded: new Date().toISOString(), applied: null });
        saveHistory(history.slice(0, HISTORY_MAX));
    }
    function recordFileApplied(fileName) {
        const history = getHistory();
        const entry = history.find(h => h.name === fileName);
        if (entry) { entry.applied = new Date().toISOString(); saveHistory(history); }
    }
    // ────────────────────────────────────────────────────────────────────────

    const PEU_EMOJI = '📍';

    /**
     * ECHAPPE UNE DONNEE AVANT DE L'INSERER DANS DU HTML.
     *
     * ⚠️⚠️ LES CINQ CARACTERES, PAS TROIS. Oublier l'apostrophe et le chevron
     *    fermant suffit a faire sortir une valeur de son attribut : un nom de
     *    lieu venu d'un classeur est une donnee EXTERNE, et ce script en pose
     *    dans des title= a chaque ligne.
     */
    function esc(v) {
        return String(v === undefined || v === null ? '' : v)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }

    /* ======================================================================
       LA FEUILLE DE STYLE — une seule, injectee une fois, prefixe peu- partout.

       ⚠️⚠️ AUCUN ACCENT GRAVE DANS CE BLOC. Il vit dans un template literal :
          un seul accent grave le referme et casse tout le script. Le piege est
          connu des deux autres scripts de la famille, qui le rappellent chacun
          six fois.

       ⭐ TOUT EST EN em, A PARTIR DE --peu-fs-base : changer cette seule valeur
          redimensionne l'interface entiere. C'est ce qui rend une densite
          reglable possible sans repeindre trente classes.

       ⚠️ CHAQUE USAGE PORTE SON REPLI : var(--peu-blue, #2C6ED5). Une variable
          manquante ne doit jamais faire DISPARAITRE une couleur.
       ====================================================================== */
    const CSS = `
:root {
    --peu-blue:    #2C6ED5;
    --peu-blue-dk: #1a4fa0;
    --peu-green:   #43a047;
    --peu-red:     #e53935;
    --peu-orange:  #f57c00;
    --peu-grey:    #9e9e9e;
    --peu-warn:    #f9a825;
    --peu-surface: #ffffff;
    --peu-bg:      #f5f7f9;
    --peu-border:  #dde3ea;
    --peu-text:    #2d3748;
    --peu-text2:   #566372;
    --peu-radius:  8px;
    --peu-shadow:  0 8px 32px rgba(0,0,0,.22), 0 2px 8px rgba(0,0,0,.12);
    --peu-fs-base: 12px;
}

/* ⚠️ box-sizing SUR TOUT CE QUI EST A NOUS : sans lui, un width:100% sort du
   panneau lateral, qui n'offre que 315 px utiles. */
#peu-overlay, #peu-overlay *, .peu-container, .peu-container *,
#peu-fab-wrap, #peu-fab-wrap * { box-sizing: border-box; }

/* ⚠️⚠️ L ATTRIBUT hidden NE RESISTE PAS A UN display EXPLICITE. Nos boutons
   sont en display:inline-flex : pose sur eux, hidden ne masque RIEN, et le
   bouton du rapport s affichait alors qu aucun apercu n existait. Le piege
   vaut pour tout composant a qui l on donne un display — c est-a-dire presque
   tous. */
#peu-overlay [hidden], .peu-container [hidden] { display: none !important; }

/* ----------------------------------------------------------------------
   LE BOUTON DE CARTE
   ⚠️ NI position NI z-index : il est docke dans le conteneur natif des
      boutons de WME, dont il herite le contexte d'empilement. Pose en
      position:fixed, il passerait PAR-DESSUS le panneau des calques.
   ⚠️ order:100 — le conteneur est une grille ; WNA prend 99, WCT reinsere
      le sien en dernier. 100 nous range derriere les deux, et l'ordre ne
      depend plus de qui a demarre le premier.
   ---------------------------------------------------------------------- */
#peu-fab-wrap { width: 40px; height: 40px; order: 100; }
#peu-fab-btn {
    width: 40px; height: 40px; padding: 0; margin: 0; border: none; border-radius: 50%;
    background: #fff; box-shadow: 0 2px 6px rgba(0,0,0,.3);
    cursor: pointer; position: relative;
    display: flex; align-items: center; justify-content: center;
    transition: box-shadow .15s;
}
#peu-fab-btn:hover  { box-shadow: 0 3px 10px rgba(0,0,0,.4); }
#peu-fab-btn.peu-fab-on { box-shadow: 0 0 0 2px var(--peu-blue, #2C6ED5), 0 2px 6px rgba(0,0,0,.3); }
.peu-fab-badge {
    position: absolute; top: -4px; right: -4px;
    background: var(--peu-green, #43a047); color: #fff;
    border-radius: 50px; font: 700 10px/1 'Rubik','Open Sans',sans-serif;
    padding: 2px 4px; border: 2px solid #fff;
    pointer-events: none; white-space: nowrap; display: none;
}
#peu-fab-btn.peu-has-file .peu-fab-badge { display: block; }

/* ----------------------------------------------------------------------
   LA FENETRE — le TRAVAIL. Le panneau lateral, lui, porte les REGLAGES.
   ---------------------------------------------------------------------- */
#peu-overlay {
    position: fixed; z-index: 9200;
    width: min(820px, calc(100vw - 24px));
    background: var(--peu-surface, #fff);
    border: 1px solid var(--peu-border, #dde3ea); border-radius: 12px;
    box-shadow: var(--peu-shadow, 0 8px 32px rgba(0,0,0,.22), 0 2px 8px rgba(0,0,0,.12));
    display: none; flex-direction: column;
    max-height: calc(100vh - 110px);
    font-family: 'Rubik','Open Sans',sans-serif;
    font-size: var(--peu-fs-base, 12px); color: var(--peu-text, #2d3748);
    overflow: hidden;
}
#peu-overlay.peu-open { display: flex; }
#peu-overlay.peu-replie #peu-body,
#peu-overlay.peu-replie .peu-footer,
#peu-overlay.peu-replie #peu-strip { display: none; }
#peu-overlay.peu-replie { height: auto !important; max-height: none !important; resize: none; }

.peu-header {
    background: linear-gradient(135deg, #3d84e8 0%, var(--peu-blue-dk, #1a4fa0) 100%);
    color: #fff; padding: 9px 12px;
    display: flex; align-items: center; justify-content: space-between;
    cursor: move; user-select: none;
    border-radius: 11px 11px 0 0; flex-shrink: 0;
}
#peu-overlay.peu-replie .peu-header { border-radius: 11px; }
.peu-header-left {
    font-size: 1.083em; font-weight: 700;
    display: flex; align-items: center; gap: 7px;
    min-width: 0; white-space: nowrap; overflow: hidden;
}
.peu-header-version { font-size: .833em; opacity: .6; flex-shrink: 0; }
.peu-header-btns { display: flex; gap: 5px; flex-shrink: 0; }
.peu-btn-icon {
    background: rgba(255,255,255,.18); border: none; color: #fff;
    width: 24px; height: 24px; min-height: 0; border-radius: 50%;
    cursor: pointer; padding: 0;
    display: flex; align-items: center; justify-content: center;
    font-family: inherit; font-size: 1.083em; line-height: 1;
    transition: background .15s;
}
.peu-btn-icon:hover { background: rgba(255,255,255,.35); }
.peu-btn-icon:disabled { opacity: .45; cursor: default; }

/* Le bandeau d'etat : ce qui est charge, sous les yeux en permanence. */
#peu-strip {
    display: flex; align-items: center; gap: 7px; flex-wrap: wrap;
    padding: 6px 12px; background: var(--peu-bg, #f5f7f9);
    border-bottom: 1px solid var(--peu-border, #dde3ea);
    font-size: .917em; flex-shrink: 0;
}
.peu-strip-dot { width: 8px; height: 8px; border-radius: 50%; background: var(--peu-grey, #9e9e9e); flex-shrink: 0; }
#peu-strip.peu-has-file .peu-strip-dot { background: var(--peu-green, #43a047); }
.peu-strip-file { font-weight: 700; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 40%; }
.peu-strip-sep { color: var(--peu-border, #dde3ea); }
.peu-strip-info { color: var(--peu-text2, #566372); }

#peu-body { flex: 1; overflow-y: auto; min-height: 0; }


/* ----------------------------------------------------------------------
   LES BOUTONS
   ⚠️⚠️ height:auto ET min-height : WME impose height:32px a TOUT bouton par
      sa feuille globale. Sans cette parade, un libelle sur deux lignes est
      coupe net.
   ⚠️ :not(:disabled) SUR CHAQUE VARIANTE : sans lui, .peu-btn:hover (plus
      specifique) l'emporte sur .peu-btn-primary et repeint le fond en clair
      SOUS un texte reste blanc. Regle generale : ne jamais poser un FOND
      sans poser la COULEUR DE TEXTE qui va avec.
   ---------------------------------------------------------------------- */
.peu-btn {
    display: inline-flex; align-items: center; justify-content: center; gap: 5px;
    padding: .417em 1em; border: none; border-radius: 50px;
    font-family: inherit; font-size: .917em; font-weight: 600;
    height: auto; min-height: 28px; line-height: 1.35;
    cursor: pointer; white-space: nowrap;
    transition: filter .15s, transform .1s, background .15s;
}
.peu-btn:active:not(:disabled) { transform: scale(.97); }
.peu-btn-primary { background: var(--peu-blue, #2C6ED5);   color: #fff; }
.peu-btn-neutral { background: var(--peu-border, #dde3ea); color: var(--peu-text, #2d3748); }
.peu-btn-primary:hover:not(:disabled) { background: var(--peu-blue-dk, #1a4fa0); color: #fff; }
.peu-btn-neutral:hover:not(:disabled) { filter: brightness(.95); }
.peu-btn:disabled { opacity: .45; cursor: not-allowed; }
.peu-btn-sm { padding: .25em .75em; font-size: .833em; min-height: 24px; }
.peu-btn-full { width: 100%; }

/* Bouton discret de ligne (recentrage). Pas de fond, pas de bordure. */
.peu-btn-center {
    background: transparent; border: none; padding: 1px 3px; margin: 0;
    cursor: pointer; font-size: 1.25em; line-height: 1;
    height: auto; min-height: 0; transition: transform .1s;
}
.peu-btn-center:hover { transform: scale(1.18); }
.peu-btn-center:active { transform: scale(.9); }

/* ----------------------------------------------------------------------
   CHAMPS
   ---------------------------------------------------------------------- */
/* ⚠️⚠️ BOX-SIZING SUR LE COMPOSANT LUI-MEME, et pas seulement herite
   d un ancetre : un champ en largeur 100 pour cent sans lui deborde de sa
   colonne de la largeur de son padding et de sa bordure. La regle globale
   plus haut ne couvre que ce qui vit DANS la fenetre — un composant sorti
   de la pour un banc, ou pose ailleurs demain, perdrait la regle sans que
   rien ne le dise. */
.peu-input, .peu-textarea, .peu-search, .peu-select {
    box-sizing: border-box; width: 100%; padding: .25em .45em;
    border: 1px solid var(--peu-border, #dde3ea); border-radius: var(--peu-radius, 8px);
    font-family: inherit; font-size: 1em;
    background: #fff; color: var(--peu-text, #2d3748);
    transition: border-color .15s;
}
.peu-textarea { resize: vertical; min-height: 2.2em; line-height: 1.4; }
.peu-input:focus, .peu-textarea:focus, .peu-search:focus, .peu-select:focus {
    outline: none; border-color: var(--peu-blue, #2C6ED5);
    box-shadow: 0 0 0 3px rgba(44,110,213,.15);
}
.peu-input:disabled, .peu-textarea:disabled { background: #f1f3f6; color: var(--peu-grey, #9e9e9e); }
.peu-select { width: auto; padding: .2em .4em; }

/* La case vit DANS un label : toute la zone devient cliquable. */
.peu-check { display: inline-flex; align-items: center; cursor: pointer; }
.peu-check input, .peu-checkbox {
    width: 15px; height: 15px; margin: 0;
    cursor: pointer; accent-color: var(--peu-blue, #2C6ED5);
}
.peu-check input:disabled, .peu-checkbox:disabled { cursor: not-allowed; }

.peu-toolbar { display: flex; align-items: center; gap: 6px; padding: 8px 12px; flex-wrap: wrap; }
.peu-search { flex: 1; min-width: 140px; }
.peu-search-count { font-size: .833em; color: var(--peu-text2, #566372); white-space: nowrap; }

/* ----------------------------------------------------------------------
   LE TABLEAU
   ⭐⭐⭐⭐ LA CASE EST EN PREMIERE COLONNE, avec un intitule. Elle vivait en
      quatrieme et derniere, large de 30 px, sous un en-tete vide : le pied
      de page disait de decocher des lignes en designant quelque chose que
      personne ne voyait.
   ---------------------------------------------------------------------- */
.peu-table {
    width: 100%; border-collapse: collapse; table-layout: fixed;
    font-size: .958em; color: var(--peu-text, #2d3748);
}
.peu-table colgroup col:nth-child(1) { width: 44px; }
.peu-table colgroup col:nth-child(2) { width: 62px; }
.peu-table colgroup col:nth-child(3) { width: 40%; }
.peu-table colgroup col:nth-child(4) { width: auto; }
.peu-table thead th {
    background: var(--peu-bg, #f5f7f9); color: var(--peu-blue, #2C6ED5);
    font-size: .833em; font-weight: 700; text-transform: uppercase; letter-spacing: .04em;
    padding: 7px 6px; border-bottom: 2px solid var(--peu-blue, #2C6ED5);
    position: sticky; top: 0; z-index: 2; text-align: start;
}
.peu-table thead th.center { text-align: center; }
.peu-table thead th.sortable { cursor: pointer; user-select: none; }
.peu-table thead th.sortable:hover { background: #e3ecfa; }
.peu-table thead th .peu-sort-icon { display: inline-block; margin-left: 4px; opacity: .35; font-style: normal; }
.peu-table thead th.sort-asc .peu-sort-icon,
.peu-table thead th.sort-desc .peu-sort-icon { opacity: 1; }
.peu-table td { padding: 6px; border-bottom: 1px solid var(--peu-border, #dde3ea); vertical-align: top; word-break: break-word; }
.peu-table td.center { text-align: center; vertical-align: middle; }
.peu-table tbody tr.peu-ligne:hover > td { background: #e8f0fd; }

/* ⚠️⚠️ TROIS SIGNES POUR UN ETAT, JAMAIS UN SEUL : un lisere a gauche, un
   fond, un badge. Le fond seul disparait sous le survol, et le badge seul se
   rate. Le lisere, lui, tient dans les trois cas. */
.peu-ligne > td:first-child { border-inline-start: 3px solid transparent; }
.peu-row-diff     > td:first-child { border-inline-start-color: var(--peu-green, #43a047); }
.peu-row-perte    > td:first-child { border-inline-start-color: var(--peu-orange, #f57c00); }
.peu-row-perte    > td { background: #fff8ec; }
.peu-row-hard     > td:first-child { border-inline-start-color: var(--peu-red, #e53935); }
.peu-row-hard     > td { background: #fdf0f0; }
.peu-row-sae      > td:first-child { border-inline-start-color: var(--peu-warn, #f9a825); }
.peu-row-sae      > td { background: #fffdf0; }
.peu-row-unloaded > td:first-child { border-inline-start-color: var(--peu-grey, #9e9e9e); }
.peu-row-unloaded > td { background: #f5f5f5; color: var(--peu-grey, #9e9e9e); }

.peu-cell-old {
    color: #8a8a8a; font-size: .909em; line-height: 1.4; min-height: 14px;
    margin-bottom: 3px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.peu-cell-old.changed {
    color: var(--peu-red, #e53935); text-decoration: line-through;
    text-decoration-color: rgba(229,57,53,.5);
}

/* Badges : fond pastel, texte de la meme teinte en fonce. cursor:help des
   qu'ils portent une explication qu'on ne lit qu'au survol. */
.peu-badge {
    display: inline-flex; align-items: center; justify-content: center;
    padding: .083em .5em; border-radius: 50px;
    font-size: .833em; font-weight: 700; white-space: nowrap;
}
.peu-badge-diff  { background: #e8f5e9; color: #1b5e20; cursor: help; }
.peu-badge-perte { background: #fff3e0; color: #e65100; cursor: help; }
.peu-badge-lock  { background: #ffebee; color: #c62828; cursor: help; }
.peu-badge-sae   { background: #fff8e1; color: #f57f17; cursor: help; }
.peu-badge-off   { background: #eceff1; color: #37474f; cursor: help; }
.peu-badge-ok    { background: #eceff1; color: #607d8b; cursor: help; }

/* Les champs du lot D2, sous la ligne du lieu. */
.peu-comp > td { background: rgba(0,0,0,.015); padding-top: 0; }
/* ⚠️⚠️ CES PASTILLES AVAIENT DISPARU A LA REECRITURE DE LA FEUILLE, et elles ne
   sont PAS decoratives : chaque teinte dit ce que le script fera de la valeur.
   Sans elles, les champs du lot D2 s affichaient tous pareil — on ne voyait plus
   la difference entre ce qui sera ecrit, ce qui est seulement montre, ce qui est
   refuse et ce qui sera RETIRE. */
.peu-plus {
    display: flex; flex-wrap: wrap; gap: 3px 5px; align-items: baseline;
    margin-top: 4px; font-size: .875em; line-height: 1.5; color: var(--peu-text2, #566372);
}
.peu-pastille {
    border-radius: 3px; padding: 1px 6px; white-space: nowrap;
    border: 1px solid transparent;
}
/* Le vert n est PAS decoratif : il dit « le script ecrira ceci ». */
.peu-pastille-pose   { background: #eaf6ec; border-color: #bfe0c6; color: #1e6b2f; }
/* Le gris dit « a toi de le poser » — ni succes, ni alerte. */
.peu-pastille-montre { background: #f2f2f2; border-color: #ddd;    color: #555; }
.peu-pastille-refus  { background: #fdeceb; border-color: #f5c6c2; color: #a3281e; }
/* L orange dit « j enleve » : ni un succes, ni une erreur — une perte. */
.peu-pastille-perte  { background: #fdf3e3; border-color: #f0d3a0; color: #8a5a00; }
.peu-pastille b { font-weight: 600; }
.peu-pastille-titre { color: var(--peu-text2, #566372); padding: 1px 0; font-size: .875em; }

/* ----------------------------------------------------------------------
   LE PIED D'ACTION — hors defilement.
   ⭐ UN SEUL BOUTON PLEIN : l'ETAPE SUIVANTE. Appliquer etait un glyphe de
      un caractere dans la barre de titre, colle a la croix qui ferme : deux
      signes de meme taille, l'un ecrit sur la carte, l'autre abandonne.
   ---------------------------------------------------------------------- */
.peu-footer {
    flex-shrink: 0; border-top: 1px solid var(--peu-border, #dde3ea);
    background: var(--peu-bg, #f5f7f9); padding: 8px 12px;
    display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
}
.peu-footer-help { font-size: .833em; color: var(--peu-text2, #566372); line-height: 1.4; flex: 1; min-width: 150px; }
.peu-footer-actions { display: flex; gap: 6px; align-items: center; margin-inline-start: auto; }
.peu-alert ul { margin: 3px 0 0; padding-inline-start: 16px; }
.peu-alert ul li { margin-bottom: 2px; }
.peu-error-title { font-weight: 700; display: block; margin-bottom: 4px; }

/* ----------------------------------------------------------------------
   BANDEAUX — trois familles, et leur sens ne se melange pas :
   bleu = une information, vert = un resultat, orange = UN GESTE A FAIRE.
   ---------------------------------------------------------------------- */
.peu-alert  { border-radius: var(--peu-radius, 8px); padding: 7px 10px; margin: 8px 12px;
              font-size: .833em; line-height: 1.5; }
.peu-alert-ok   { background: #e8f5e9; border: 1px solid #a5d6a7; color: #2e7d32; }
.peu-alert-warn { background: #fff3e0; border: 1px solid #ffb74d; color: #a34a00; }

/* Le guidage : il dit TOUJOURS le geste suivant. Un ecran vide avec un
   bouton grise n'explique rien. */
.peu-guide {
    display: flex; align-items: flex-start; gap: 8px;
    background: #e3f2fd; border: 1px solid #90caf9; color: #0d47a1;
    border-radius: var(--peu-radius, 8px); padding: 8px 10px; margin: 10px 12px;
    font-size: .917em; line-height: 1.45;
}
.peu-guide-n {
    flex: 0 0 auto; min-width: 17px; height: 17px; line-height: 17px; text-align: center;
    border-radius: 50%; background: var(--peu-blue, #2C6ED5); color: #fff;
    font-size: .833em; font-weight: 700;
}
.peu-guide-suite { margin-top: 2px; opacity: .85; font-size: .909em; }
.peu-dropzone {
    border: 2px dashed var(--peu-border, #dde3ea); border-radius: var(--peu-radius, 8px);
    padding: 16px 12px; margin: 0 12px 12px; text-align: center;
    font-size: .917em; color: var(--peu-text2, #566372); line-height: 1.6;
    cursor: pointer; transition: border-color .15s, color .15s;
}
.peu-dropzone:hover, .peu-dropzone.peu-drop-hover {
    border-color: var(--peu-blue, #2C6ED5); color: var(--peu-blue, #2C6ED5);
}

/* ----------------------------------------------------------------------
   LA PROGRESSION — bleu franc : c'est du TRAVAIL EN COURS, ni une alerte
   (orange) ni un resultat (vert). Chiffres tabulaires, sans quoi ils
   dansent d'un rafraichissement a l'autre.
   ---------------------------------------------------------------------- */
.peu-prog {
    background: #e3f2fd; border: 1px solid #90caf9; color: #0d47a1;
    border-radius: var(--peu-radius, 8px); padding: 7px 9px; margin: 10px 12px;
    font-size: .917em;
}
.peu-prog-t { display: flex; align-items: center; gap: 6px; }
.peu-progress-label {
    flex: 1 1 auto; min-width: 0; font-weight: 600;
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.peu-prog-pct { flex: 0 0 auto; font-variant-numeric: tabular-nums; }
.peu-progress-bar-bg { height: 6px; background: #bbdefb; border-radius: 3px; overflow: hidden; margin: 5px 0 4px; }
.peu-progress-bar { display: block; height: 100%; width: 0; background: var(--peu-blue, #2C6ED5);
                    border-radius: 3px; transition: width .15s linear; }
.peu-prog-b { display: flex; align-items: center; gap: 6px; }
.peu-prog-d { flex: 1 1 auto; font-size: .909em; font-variant-numeric: tabular-nums; }

/* ----------------------------------------------------------------------
   LE PANNEAU LATERAL — les REGLAGES et l'HISTORIQUE, jamais le travail.
   ⚠️ Le panneau fait disparaitre son contenu des qu'on selectionne un objet
      sur la carte : on ne peut pas y travailler.
   ---------------------------------------------------------------------- */
.peu-container {
    padding: 10px; font-family: 'Rubik','Open Sans',sans-serif;
    font-size: var(--peu-fs-base, 12px); color: var(--peu-text, #2d3748);
}
.peu-container h3 { margin: 0 0 8px; font-size: 1.083em; color: var(--peu-blue, #2C6ED5); }
.peu-side-sect {
    font-size: .833em; font-weight: 700; text-transform: uppercase; letter-spacing: .06em;
    color: var(--peu-blue, #2C6ED5); border-bottom: 1px solid var(--peu-border, #dde3ea);
    margin: 12px 0 6px; padding-bottom: 3px;
    display: flex; align-items: center; gap: 5px;
}
.peu-hist-row {
    margin-bottom: 5px; padding: 5px 7px; background: var(--peu-bg, #f5f7f9);
    border-radius: 4px; border-inline-start: 3px solid var(--peu-blue, #2C6ED5);
    font-size: .833em;
}
.peu-hist-name { font-weight: 700; color: var(--peu-blue, #2C6ED5);
                 white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.peu-hist-meta { color: var(--peu-text2, #566372); margin-top: 1px; }

/* La poignee de redimensionnement, en bas a droite. */
#peu-resize {
    position: absolute; right: 0; bottom: 0; width: 18px; height: 18px;
    cursor: nwse-resize; z-index: 30; opacity: .45;
    background:
        linear-gradient(135deg, transparent 46%, var(--peu-grey, #9e9e9e) 46%, var(--peu-grey, #9e9e9e) 54%, transparent 54%),
        linear-gradient(135deg, transparent 70%, var(--peu-grey, #9e9e9e) 70%, var(--peu-grey, #9e9e9e) 78%, transparent 78%);
}
#peu-resize:hover { opacity: .9; }
#peu-overlay.peu-replie #peu-resize { display: none; }

/* ⚠️ FOCUS VISIBLE SUR TOUT CE QUI EST ATTEIGNABLE AU CLAVIER : nos champs
   posent outline:none, et sans cette regle on tabule a l'aveugle. */
#peu-overlay :focus-visible, #peu-fab-btn:focus-visible, .peu-container :focus-visible {
    outline: 2px solid var(--peu-blue, #2C6ED5); outline-offset: 1px; border-radius: 3px;
}

/* ⚠️ LA BALISE kbd A SON PROPRE STYLE DANS WME, en texte BLANC : sans cette
   regle les touches s'affichent blanc sur fond clair, donc vides. Ne jamais
   compter sur l'heritage pour une balise semantique, la page hote a le sien. */
#peu-overlay kbd {
    display: inline-block; background: var(--peu-bg, #f5f7f9); color: var(--peu-text, #2d3748);
    border: 1px solid var(--peu-border, #dde3ea); border-bottom-width: 2px; border-radius: 3px;
    padding: 0 5px; font-family: ui-monospace, Menlo, Consolas, monospace;
    font-size: .917em; line-height: 1.5;
}

@media (max-height: 820px) { #peu-overlay { --peu-fs-base: 11px; max-height: calc(100vh - 68px); } }
@media (max-height: 680px) { #peu-overlay { max-height: calc(100vh - 44px); } }
    `;

    function injectCSS() {
        if (document.getElementById('peu-style')) return;
        const style = document.createElement('style');
        style.id = 'peu-style';
        style.textContent = CSS;
        document.head.appendChild(style);
    }

    // ==== banc:champs ====
    // Extrait tel quel par tools/banc-champs.mjs : aucune dépendance (ni DOM, ni SDK).

    /* LES VALEURS DE WME, RELEVÉES DANS L'ÉDITEUR LE 15/09/2026 — jamais écrites
       de mémoire. Chaque entrée associe la clé que WME attend aux libellés qui
       peuvent arriver dans le classeur.

       ⚠️⚠️ LA CONVERSION EST LE GARDE-FOU, PAS UNE COMMODITÉ. Le SDK accepte SANS
          ERREUR une valeur hors énumération et la POSE telle quelle : envoyer
          « gratuit » au lieu de « FREE » ne provoque aucun refus, marque le lieu
          modifié, et fait enregistrer une valeur que WME ne reconnaît pas.

       ⚠️ Les libellés français sont ceux de l'éditeur, et ce sont aussi ceux que
          l'extranet EVIDRA imprime dans son export — la conversion par libellé est
          donc le chemin normal, la clé WME n'étant qu'un raccourci pour qui la
          connaît. */
    const VALEURS_WME = {
        parkingType: {
            PUBLIC: ['Public'], PRIVATE: ['Privé'], RESTRICTED: ['Restreint']
        },
        costType: {
            FREE: ['Gratuit'], LOW: ['Faible'], MODERATE: ['Modéré'], EXPENSIVE: ['Élevé']
        },
        estimatedNumberOfSpots: {
            R_1_TO_10: ['1-10'], R_11_TO_30: ['11-30'], R_31_TO_60: ['31-60'],
            R_61_TO_100: ['61-100'], R_101_TO_300: ['101-300'], R_301_TO_600: ['301-600'],
            R_600_PLUS: ['> 600', '>600', 'plus-600']
        },
        lotType: {
            MULTI_LEVEL: ['Plusieurs niveaux'], STREET_LEVEL: ['Extérieur'],
            STREET_LEVEL_COVERED: ['Extérieur couvert'], UNDERGROUND: ['Souterrain']
        },
        paymentType: {
            CASH: ['Espèces'], CHECKS: ['Chèques'], CREDIT: ['Carte de crédit'],
            DEBIT_CARD: ['Carte bancaire'], DIGITAL_WALLET: ['Portefeuille numérique'],
            ELECTRONIC_PASS: ['Pass électronique'], MEMBERSHIP: ['Abonnement'],
            PARKING_APP: ['Application'], PERMIT: ['Laissez-passer'],
            PREPAID: ['Prépaiement'], SMS_CALL: ['SMS/Appel', 'sms appel']
        },
        /* ⚠️⚠️ « services » EST LE MÊME CHAMP POUR UN LIEU ET POUR UN PARKING :
              c'est la CATÉGORIE du lieu qui décide des valeurs proposées. D'où deux
              tables, et une clé qui ne désigne pas la même chose dans les deux —
              « Voiturier » vaut VALET sur un parking, « Service de voiturier »
              vaut VALLET_SERVICE partout. Une table unique poserait l'un pour
              l'autre, sans erreur visible.
           📌 « VALLET » est la graphie de Waze, faute de frappe comprise : on la
              recopie, on ne la corrige pas. */
        services: {
            AIR_CONDITIONING: ['Climatisation'],
            CREDIT_CARDS: ['Accepte les cartes de crédit'],
            CURBSIDE_PICKUP: ['Click & Collect', 'click and collect'],
            DELIVERIES: ['Livraisons'], DRIVETHROUGH: ['Drive'],
            OUTSIDE_SEATING: ['Terrasse extérieure'],
            PARKING_FOR_CUSTOMERS: ['Parking client'], RESERVATIONS: ['Réservations'],
            RESTROOMS: ['Toilettes'], TAKE_AWAY: ['À emporter'],
            VALLET_SERVICE: ['Service de voiturier'],
            WHEELCHAIR_ACCESSIBLE: ['Accessible en fauteuil roulant'], WI_FI: ['Wi-Fi', 'wifi']
        },
        /* ⭐⭐⭐⭐ LES CATÉGORIES NE SONT PAS ÉCRITES ICI : elles sont RELEVÉES DANS
           L'ÉDITEUR au chargement (`chargerCategories`), parce que WME les rend
           déjà traduites dans la langue de l'utilisateur — 132 le 16/09/2026.
           Une copie figée serait fausse dans toute autre langue que le français,
           et périmerait au premier ajout de Waze. */
        categories: {},
        parkingServices: {
            AIRPORT_SHUTTLE: ['Navette aéroport'],
            CARPOOL_PARKING: ['Places covoiturage'], CAR_WASH: ['Lavage auto'],
            COVERED: ['Couvert'], DISABILITY_PARKING: ['Places PMR'],
            ON_SITE_ATTENDANT: ['Agent d’accueil', "agent d'accueil"],
            PARK_AND_RIDE: ['P+R'], RESERVATIONS: ['Réservations'],
            SECURITY: ['Surveillance'], VALET: ['Voiturier'],
            VALLET_SERVICE: ['Service de voiturier'],
            EV_CHARGING_STATION: ['Bornes de charge']
        }
    };

    /**
     * La clé WME d'une valeur lue dans le classeur, ou `null` si elle n'est pas
     * reconnue — ce qui doit TOUJOURS se signaler, jamais se taire.
     */
    /**
     * Remplit le référentiel des catégories depuis le SDK.
     *
     * ⚠️ S'il échoue, la table reste VIDE et toute catégorie du classeur sera
     *    REFUSÉE — donc signalée. C'est voulu : poser une catégorie qu'on n'a pas
     *    pu vérifier serait pire que de ne pas la poser du tout.
     * ⭐ Le SDK valide les catégories, lui : une valeur inconnue est refusée avec
     *    une vraie erreur (« categories[0] must match the configured type »), et
     *    un parking ne peut en porter qu'une seule. C'est le seul champ où il ne
     *    se tait pas.
     */
    function chargerCategories(sdk) {
        const table = VALEURS_WME.categories;
        if (Object.keys(table).length) return Object.keys(table).length;
        (sdk.DataModel.Venues.getAllVenueCategories() || []).forEach(c => {
            if (c && c.id) table[c.id] = [c.localizedName || c.id];
        });
        return Object.keys(table).length;
    }

    function cleWme(referentiel, saisie) {
        const table = VALEURS_WME[referentiel];
        if (!table) return null;
        const v = normalizeHeader(saisie);
        if (v === '') return null;
        /* La clé de WME est reconnue par la boucle elle-même (« CREDIT » se
           normalise en « credit ») : pas de test séparé, il serait mort. */
        for (const cle of Object.keys(table)) {
            if (normalizeHeader(cle) === v) return cle;
            if (table[cle].some(lib => normalizeHeader(lib) === v)) return cle;
        }
        return null;
    }

    /**
     * Une cellule qui porte plusieurs valeurs : une par ligne, ou séparées par
     * « ; » — c'est ce que produit l'export de l'extranet.
     *
     * ⚠️ Rend les clés RETENUES et les valeurs REFUSÉES : un appelant qui ignore
     *    les secondes construit une panne silencieuse.
     */
    function clesWme(referentiel, cellule) {
        const retenues = [], refusees = [];
        String(cellule === undefined || cellule === null ? '' : cellule)
            .split(/\r?\n|;/)
            .map(m => m.trim())
            .filter(Boolean)
            .forEach(morceau => {
                const cle = cleWme(referentiel, morceau);
                if (cle === null) refusees.push(morceau);
                else if (!retenues.includes(cle)) retenues.push(cle);
            });
        return { retenues, refusees };
    }
    /* Un booléen écrit en toutes lettres, dans les deux langues. */
    function booleenWme(saisie) {
        const v = normalizeHeader(saisie);
        if (v === '') return null;
        if (['oui', 'yes', 'true', '1', 'x'].includes(v)) return true;
        if (['non', 'no', 'false', '0'].includes(v)) return false;
        return null;
    }

    /* LES COLONNES DU CLASSEUR.
       `pose: true` — WPEU l'écrit dans WME.
       `pose: false` — WPEU l'AFFICHE seulement : la valeur demande une
       interprétation qu'aucun script ne peut faire sans se tromper en silence
       (horaires en toutes lettres, rue à retrouver dans le modèle, points sur la
       carte), ou sa liste de valeurs n'a pas été relevée.
       ⚠️ « Montré » est un engagement, pas un repli : une demande du client qui
          n'apparaîtrait nulle part serait perdue sans trace. */
    const CHAMPS = [
        /* Le permalien n'est pas une valeur à poser : c'est lui qui DÉSIGNE le
           lieu. Il figure ici pour que les libellés des colonnes aient une seule
           source, et `pose: false` sans motif le distingue des champs montrés. */
        { cle: 'perm',        entetes: ['poi permalink', 'permalink', 'permalien', 'poi permalien'],
          pose: false,        identifie: true, libelle: 'Permalien' },
        { cle: 'name',        entetes: ['poi name', 'name', 'nom', 'poi nom', 'nom du poi'],
          cible: 'name',      pose: true,  libelle: 'Nom' },
        { cle: 'desc',        entetes: ['poi description', 'description', 'desc', 'poi desc'],
          cible: 'description', pose: true, libelle: 'Description' },
        { cle: 'aliases',     entetes: ['alternative names', 'alternate names', 'noms alternatifs', 'nom alternatif'],
          cible: 'aliases',   pose: true,  multiple: true, libelle: 'Noms alternatifs' },
        { cle: 'phone',       entetes: ['phone', 'telephone', 'téléphone'],
          cible: 'phone',     pose: true,  libelle: 'Téléphone' },
        { cle: 'url',         entetes: ['website', 'site web', 'site'],
          cible: 'url',       pose: true,  libelle: 'Site web' },
        { cle: 'services',    entetes: ['services', 'services du lieu'],
          cible: 'services',  pose: true,  multiple: true, referentiel: 'services', libelle: 'Services' },

        { cle: 'parkingType', entetes: ['parking type', 'type de parking'],
          cible: 'PARKING_LOT.parkingType', pose: true, referentiel: 'parkingType', libelle: 'Type de parking' },
        { cle: 'hasTBR',      entetes: ['parking type varies', 'type variable'],
          cible: 'PARKING_LOT.hasTBR', pose: true, booleen: true, libelle: 'Type variable' },
        { cle: 'costType',    entetes: ['parking cost', 'tarif', 'tarif du parking'],
          cible: 'PARKING_LOT.costType', pose: true, referentiel: 'costType', libelle: 'Tarif' },
        { cle: 'paymentType', entetes: ['parking payment', 'modes de paiement', 'paiements'],
          cible: 'PARKING_LOT.paymentType', pose: true, multiple: true, referentiel: 'paymentType', libelle: 'Modes de paiement' },
        { cle: 'parkingServices', entetes: ['parking services', 'services du parking'],
          cible: 'services',  pose: true,  multiple: true, referentiel: 'parkingServices', libelle: 'Services du parking' },
        { cle: 'lotType',     entetes: ['parking situation', 'situation'],
          cible: 'PARKING_LOT.lotType', pose: true, multiple: true, referentiel: 'lotType', libelle: 'Situation' },
        { cle: 'spots',       entetes: ['parking spots', 'nombre de places', 'places'],
          cible: 'PARKING_LOT.estimatedNumberOfSpots', pose: true, referentiel: 'estimatedNumberOfSpots', libelle: 'Nombre de places' },
        { cle: 'canExit',     entetes: ['parking exit when closed', 'sortie parking ferme', 'sortie parking fermé'],
          cible: 'PARKING_LOT.canExitWhileClosed', pose: true, booleen: true, libelle: 'Sortie quand fermé' },

        /* ---- Montrés seulement ---- */
        { cle: 'categories',  entetes: ['categories', 'catégories', 'category', 'catégorie'],
          cible: 'categories', pose: true, multiple: true, referentiel: 'categories',
          libelle: 'Catégories' },
        { cle: 'hours',       entetes: ['opening hours', 'horaires'],
          pose: false, libelle: 'Horaires', motif: 'WME attend des créneaux, pas une phrase' },
        { cle: 'address',     entetes: ['address', 'adresse'],
          pose: false, libelle: 'Adresse', motif: 'WME attend un numéro et une rue de son propre modèle' },
        { cle: 'entryPoints', entetes: ['entry points', 'points d’entree', "points d'entree", 'points d’entrée', "points d'entrée"],
          pose: false, libelle: 'Points d’entrée', motif: 'ce sont des points sur la carte, pas du texte' },
        { cle: 'operator',    entetes: ['parking operator', 'operateur de parking', 'opérateur de parking', 'opérateur'],
          pose: false, libelle: 'Opérateur de parking', motif: 'liste fermée chez WME, dont les clés ne sont pas relevées' },
        { cle: 'googleName',  entetes: ['google name', 'nom google'],
          pose: false, libelle: 'Nom Google', motif: 'ne relève pas de WME' },
        { cle: 'googleCategory', entetes: ['google category', 'catégorie google', 'categorie google'],
          pose: false, libelle: 'Catégorie Google', motif: 'ne relève pas de WME' },
        { cle: 'googlePosition', entetes: ['google position', 'position google', 'position'],
          pose: false, libelle: 'Position Google', motif: 'ne relève pas de WME' }
    ];

    /** Les champs qui visent le MÊME attribut de WME, pour les fusionner. */
    function champsParCible(cible) {
        return CHAMPS.filter(c => c.pose && c.cible === cible);
    }

    /**
     * Ce qu'une ligne du classeur demande, trié en trois tas.
     *
     * ⭐⭐⭐⭐ UNE CELLULE VIDE NE DEMANDE RIEN, ET N'EFFACE RIEN. C'est la règle la
     *    plus importante de cette fonction : dans un classeur, une case laissée
     *    blanche veut dire « je n'ai pas de consigne », jamais « efface ce qui
     *    est sur la carte ». Sans elle, exporter un parc où le client n'a
     *    renseigné que les noms viderait tous les autres champs de WME.
     *
     * ⚠️⚠️ DEUX COLONNES VISENT `services` — celles du lieu et celles du parking.
     *    Comme un tableau REMPLACE tout son contenu dans WME, il faut envoyer
     *    leur UNION : poser l'une sans l'autre effacerait l'autre.
     *
     * @return {{aPoser: Object, montres: Array, refus: Array}}
     */
    function lireValeurs(cells, idxParChamp) {
        const aPoser = {}, montres = [], refus = [];
        const listesParCible = {};

        const cellule = (i) => {
            const v = cells[i];
            return v === undefined || v === null ? '' : String(v).trim();
        };

        CHAMPS.forEach(champ => {
            const idx = idxParChamp[champ.cle];
            if (idx === undefined || champ.identifie) return;
            const brute = cellule(idx);
            if (brute === '') return;               // rien demandé : on ne touche pas

            if (!champ.pose) {
                montres.push({ cle: champ.cle, libelle: champ.libelle, valeur: brute, motif: champ.motif });
                return;
            }

            if (champ.booleen) {
                const b = booleenWme(brute);
                if (b === null) refus.push({ libelle: champ.libelle, valeurs: [brute] });
                else aPoser[champ.cible] = b;
                return;
            }

            if (champ.referentiel) {
                if (champ.multiple) {
                    const { retenues, refusees } = clesWme(champ.referentiel, brute);
                    if (refusees.length) refus.push({ libelle: champ.libelle, valeurs: refusees });
                    if (retenues.length) {
                        (listesParCible[champ.cible] ||= []).push(...retenues);
                    }
                } else {
                    const cle = cleWme(champ.referentiel, brute);
                    if (cle === null) refus.push({ libelle: champ.libelle, valeurs: [brute] });
                    else aPoser[champ.cible] = cle;
                }
                return;
            }

            if (champ.multiple) {
                const morceaux = brute.split(/\r?\n|;/).map(m => m.trim()).filter(Boolean);
                if (morceaux.length) (listesParCible[champ.cible] ||= []).push(...morceaux);
                return;
            }

            aPoser[champ.cible] = brute;
        });

        /* L'union des colonnes qui visent la même liste, doublons retirés. */
        Object.keys(listesParCible).forEach(cible => {
            aPoser[cible] = listesParCible[cible].filter((v, i, t) => t.indexOf(v) === i);
        });

        return { aPoser, montres, refus };
    }
    // ==== /banc:champs ====

    // ==== banc:colonnes ====
    // Ce bloc est extrait par tools/banc-colonnes.mjs, AVEC celui des champs :
    // les libellés des colonnes en dérivent. Ni DOM, ni XLSX, ni t().

    /* ⭐ LES LIBELLÉS NE SONT PAS RÉÉCRITS ICI : ils DÉRIVENT de `CHAMPS`, seule
       source. Deux listes des mêmes en-têtes divergeraient le jour où l'une
       s'enrichit — et la colonne cesserait d'être reconnue d'un côté seulement.
       ⚠️ AUCUN LIBELLÉ GÉNÉRIQUE dans `CHAMPS` pour le permalien — ni « lien »,
          ni « url » : le fichier porte aussi le site web d'un lieu, dont la
          colonne s'appellerait ainsi, et elle serait lue comme le permalien. */
    const COLONNES_ATTENDUES = {
        perm: libellesDe('perm'),
        name: libellesDe('name'),
        desc: libellesDe('desc')
    };

    function libellesDe(cle) {
        const champ = CHAMPS.find(c => c.cle === cle);
        return champ ? champ.entetes : [];
    }

    function normalizeHeader(valeur) {
        return String(valeur === undefined || valeur === null ? '' : valeur)
            .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
            .toLowerCase().replace(/\s+/g, ' ').trim();
    }

    /* Quelle colonne porte quoi, d'après la ligne d'en-tête.
       Tout ou rien : les trois en-têtes reconnues, ou repli sur les positions
       A/B/C — la règle se dit en une phrase, et un repli partiel attribuerait
       une colonne au hasard. Le repli exige, comme avant, trois en-têtes non
       vides : un onglet sans en-tête reste ignoré. */
    function mapColumns(entetes) {
        const ligne = Array.isArray(entetes) ? entetes : [];
        const parNom = {};
        Object.keys(COLONNES_ATTENDUES).forEach(cle => {
            const idx = ligne.findIndex(e => COLONNES_ATTENDUES[cle].indexOf(normalizeHeader(e)) !== -1);
            if (idx !== -1) parNom[cle] = idx;
        });

        if (parNom.perm !== undefined && parNom.name !== undefined && parNom.desc !== undefined) {
            return { columns: parNom, byPosition: false, usable: true };
        }

        const troisEnTetes = [0, 1, 2].every(i => normalizeHeader(ligne[i]) !== '');
        return { columns: { perm: 0, name: 1, desc: 2 }, byPosition: true, usable: troisEnTetes };
    }

    /**
     * Toutes les colonnes reconnues, au-delà des trois de base : clé du champ →
     * index de colonne.
     *
     * ⚠️ UNE COLONNE INCONNUE N'EST PAS UNE ERREUR — le fichier peut porter les
     *    repères de son auteur (une référence interne, une couleur). Elle est
     *    simplement absente du résultat.
     * ⚠️ En REPLI (en-têtes non reconnues), on ne repère rien de plus : les
     *    positions A/B/C sont un contrat d'un autre âge, y ajouter des colonnes
     *    devinées reviendrait à inventer.
     */
    function mapChamps(entetes, byPosition) {
        const trouves = {};
        if (byPosition) return trouves;
        const ligne = Array.isArray(entetes) ? entetes : [];
        CHAMPS.forEach(champ => {
            const idx = ligne.findIndex(e => champ.entetes.indexOf(normalizeHeader(e)) !== -1);
            if (idx !== -1) trouves[champ.cle] = idx;
        });
        return trouves;
    }
    // ==== /banc:colonnes ====


    // ==== banc:pose ====
    // Logique pure de l'écriture : ce qu'on envoie, et ce qu'on vérifie après.
    // L'appel au SDK, lui, tient en trois lignes dans runApply.

    /* Ce que l'ANCIEN mécanisme écrit déjà (`UpdateObject`, depuis la 0.1) :
       on n'y touche pas. Le reste passe par le SDK.
       ⚠️ Deux mécanismes cohabitent donc, et c'est délibéré : le nom et la
          description sont le cœur du script, éprouvé sur le terrain depuis des
          mois. Les migrer « tant qu'on y est » aurait mis en jeu ce qui marche
          pour gagner une élégance que personne ne verrait. */
    const CIBLES_HERITEES = ['name', 'description', 'aliases'];

    /**
     * L'objet à passer à `updateVenue`, et ce qui a été écarté.
     *
     * ⚠️⚠️ LA LISTE BLANCHE N'EST PAS UNE PRÉCAUTION DE STYLE. Le SDK accepte SANS
     *    ERREUR un nom de champ qu'il ne connaît pas : il ne pose rien, ne dit
     *    rien, et marque quand même le lieu comme modifié. Un objet construit
     *    dynamiquement à partir du classeur ouvrirait donc la porte à une
     *    modification vide, impossible à voir autrement qu'en relisant le lieu.
     */
    function construireMaj(aPoser) {
        const autorisees = CHAMPS.filter(c => c.pose).map(c => c.cible);
        const maj = {};
        const parking = {};
        const ignores = [];

        Object.keys(aPoser).forEach(cible => {
            if (CIBLES_HERITEES.includes(cible)) return;          // écrit ailleurs
            if (!autorisees.includes(cible)) { ignores.push(cible); return; }
            const point = cible.indexOf('.');
            if (point === -1) { maj[cible] = aPoser[cible]; return; }
            const [categorie, champ] = [cible.slice(0, point), cible.slice(point + 1)];
            if (categorie !== 'PARKING_LOT') { ignores.push(cible); return; }
            parking[champ] = aPoser[cible];
        });

        if (Object.keys(parking).length) maj.categoryAttributes = { PARKING_LOT: parking };
        return { maj, ignores };
    }

    /**
     * Ce que le lieu porte DÉJÀ, comparé à ce que le classeur demande.
     *
     * ⭐⭐⭐⭐ LA MÊME COMPARAISON RÉPOND À DEUX QUESTIONS, ET C'EST POURQUOI ELLE NE
     *    S'APPELLE PLUS « vérifierPose » :
     *      · AVANT d'écrire — que reste-t-il à appliquer ? `differents` le dit, et
     *        c'est ce qui fait qu'un lieu déjà conforme ne s'annonce pas à modifier ;
     *      · APRÈS avoir écrit — qu'est-ce qui a vraiment été posé ? Le SDK ne
     *        signale ni un champ inconnu, ni une valeur hors énumération : sans
     *        cette relecture, « appliqué » ne voudrait dire que « l'appel n'a pas
     *        levé d'exception ».
     *
     * @param attributs les attributs du lieu, lus dans le modèle de WME
     */
    /** La valeur que le lieu porte pour une cible, plate ou sous categoryAttributes. */
    function valeurDuLieu(attributs, cible) {
        const point = cible.indexOf('.');
        return point === -1
            ? attributs[cible]
            : ((attributs.categoryAttributes || {})[cible.slice(0, point)] || {})[cible.slice(point + 1)];
    }

    /**
     * Ce que la pose FERAIT DISPARAÎTRE du lieu : cible → valeurs perdues.
     *
     * ⭐⭐⭐⭐ POSER N'EST PAS TOUJOURS AJOUTER. Dans WME, un tableau REMPLACE tout
     *    son contenu : un classeur qui demande « Espèces » sur un parking qui
     *    porte « Carte de crédit, Espèces » **supprime la carte de crédit**. La
     *    cause est banale : une valeur refusée réduit la liste demandée, et rien
     *    ne dit que le reste partira avec elle.
     *
     * ⇒ Une ligne qui RETIRE quelque chose ne se coche pas d'office, et sa
     *   pastille ne se peint pas en vert : le vert dit « j'ajoute », et l'œil ne
     *   lit pas une infobulle qu'il ne soupçonne pas.
     *
     * ⚠️ Seules les LISTES peuvent perdre. Un champ simple qu'on remplace est un
     *    changement voulu, et une cellule vide ne pose rien (elle n'efface donc
     *    jamais) — la règle est ailleurs, dans `lireValeurs`.
     */
    function pertesDeLaPose(attributs, aPoser) {
        const pertes = {};
        if (!attributs) return pertes;
        Object.keys(aPoser).forEach(cible => {
            const voulu = aPoser[cible];
            if (!Array.isArray(voulu)) return;
            const actuel = valeurDuLieu(attributs, cible);
            if (!Array.isArray(actuel)) return;
            const disparus = actuel.filter(v => voulu.indexOf(v) === -1);
            if (disparus.length) pertes[cible] = disparus;
        });
        return pertes;
    }

    function comparerAuLieu(attributs, aPoser) {
        const identiques = [], differents = [];
        const memeValeur = (a, b) => {
            if (Array.isArray(a) || Array.isArray(b)) {
                const x = (a || []).slice().sort(), y = (b || []).slice().sort();
                return x.length === y.length && x.every((v, i) => v === y[i]);
            }
            return a === b;
        };

        Object.keys(aPoser).forEach(cible => {
            if (CIBLES_HERITEES.includes(cible)) return;
            (memeValeur(valeurDuLieu(attributs, cible), aPoser[cible]) ? identiques : differents).push(cible);
        });

        return { identiques, differents };
    }

    /**
     * COMBIEN DE CHAMPS DU LOT D2 DIFFÈRENT DE CE QUE PORTE LE LIEU.
     *
     * ⚠️ `name` et `description` SONT EXCLUS : ils ont leur propre comparaison,
     *    celle des deux colonnes de l'aperçu, et les compter deux fois ferait
     *    paraître une différence là où l'écran en montre déjà une.
     *
     * ⚠️⚠️ UN LIEU NON CHARGÉ NE SE COMPARE À RIEN. On compte alors TOUS les
     *    champs à poser, plutôt que de conclure « rien à faire » d'une absence
     *    de mesure — une absence de signal n'est pas un signal d'absence.
     */
    function champsQuiDifferent(attributsDuLieu, aPoser) {
        if (!aPoser) return 0;
        const cibles = attributsDuLieu
            ? comparerAuLieu(attributsDuLieu, aPoser).differents
            : Object.keys(aPoser);

        return cibles.filter(c => c !== 'name' && c !== 'description').length;
    }

    /**
     * CETTE LIGNE SE COCHE-T-ELLE D'OFFICE ?
     *
     * ⭐⭐⭐⭐ DEUX DÉFAUTS TROUVÉS DANS L'ÉDITEUR VIVENT ICI, et aucun banc ne
     *    pouvait les voir tant que la règle habitait le DOM :
     *      · l'aperçu annonçait « Aucune modification » sur un parking qui avait
     *        quatre champs à poser — l'indicateur ne regardait que le nom et la
     *        description ;
     *      · une ligne qui RETIRE quelque chose était cochée par défaut : une
     *        liste réduite à « Espèces » aurait supprimé « Carte de crédit » d'un
     *        parking, et le caractère destructeur n'apparaissait qu'au survol.
     *
     * ⇒ Proposer une perte demande un geste, jamais un défaut.
     */
    function cocherDOffice(nomChange, descriptionChange, champsDifferents, pertes) {
        const differe = nomChange || descriptionChange || champsDifferents > 0;

        return differe && !pertes;
    }
    // ==== /banc:pose ====

    // ==== banc:apercu ====
    // Rendu pur : il ne touche qu'au document qu'on lui donne, pour être
    // éprouvable hors de WME (tools/banc-apercu.html).

    /** Le libellé d'un champ posé, retrouvé par sa cible. */
    function libelleDeCible(cible) {
        const champ = CHAMPS.find(c => c.cible === cible && c.pose);
        return champ ? champ.libelle : cible;
    }

    /**
     * Le libellé lisible d'une clé de WME, cherché dans TOUS les référentiels.
     *
     * ⚠️⚠️ NE JAMAIS AFFICHER LA CLÉ BRUTE À L'OPÉRATEUR. Le premier rendu
     *    montrait « Tarif MODERATE » et « Nombre de places R_101_TO_300 » : c'est
     *    le vocabulaire de la machine, et il oblige à traduire de tête au moment
     *    où l'on décide d'appliquer ou non.
     * ⚠️ La recherche passe par tous les référentiels parce qu'une valeur peut
     *    venir de l'un ou l'autre — `VALET` n'existe que du côté parking, alors
     *    que les deux colonnes visent le même attribut.
     */
    function libelleDeValeur(cle) {
        for (const table of Object.values(VALEURS_WME)) {
            if (table[cle] && table[cle].length) {
                return table[cle][0];
            }
        }
        return cle;
    }

    /** Ce qu'une valeur posée donne à lire : les listes se NOMMENT, elles ne se comptent pas. */
    function valeurLisible(v) {
        if (v === true) return 'oui';
        if (v === false) return 'non';
        if (Array.isArray(v)) {
            const noms = v.map(libelleDeValeur);
            const texte = noms.join(', ');
            /* Au-delà de trois, on nomme les deux premières et on compte le reste :
               une pastille qui déborde ne se lit plus. */
            return texte.length <= 44 ? texte : noms.slice(0, 2).join(', ') + ' +' + (noms.length - 2);
        }
        const texte = libelleDeValeur(String(v));
        return texte.length > 28 ? texte.slice(0, 27) + '…' : texte;
    }

    /**
     * La ligne des champs supplémentaires — ou `null` s'il n'y a rien à dire.
     *
     * ⭐⭐⭐ ELLE N'EXISTE QUE SI LE CLASSEUR PORTE AUTRE CHOSE QUE NOM ET
     *    DESCRIPTION. Un fichier d'hier garde donc exactement l'aspect d'hier :
     *    la nouveauté ne coûte rien à qui ne s'en sert pas.
     *
     * ⚠️ « Montré » n'est pas un repli mais un ENGAGEMENT : une demande du client
     *    que le script ne sait pas poser doit se VOIR, sinon elle est perdue sans
     *    trace — et personne ne saura qu'il fallait la poser à la main.
     */
    function rendreComplements(doc, valeurs, traduire, attributs) {
        /* ⭐⭐⭐⭐ ON NE MONTRE EN VERT QUE CE QUI CHANGE. Sans l'état du lieu, l'écran
           annonçait « appliqué » sur des valeurs que la carte portait déjà : sept
           pastilles pour un parking où rien n'était à faire, et l'œil finit par
           ne plus les lire. Le classeur d'un parc entier porte l'état COMPLET de
           chaque lieu — la plupart des valeurs y sont donc conformes.
           ⚠️ Sans `attributs` (lieu non chargé), on montre tout : ne rien montrer
              faute d'avoir pu comparer serait le pire des deux. */
        const ecarts = attributs ? comparerAuLieu(attributs, valeurs.aPoser) : null;
        const pertes = pertesDeLaPose(attributs, valeurs.aPoser);
        const aChanger = ecarts ? ecarts.differents : Object.keys(valeurs.aPoser);
        const poses = aChanger.filter(c => c !== 'name' && c !== 'description');
        const dejaConformes = ecarts
            ? ecarts.identiques.filter(c => c !== 'name' && c !== 'description').length : 0;

        if (!poses.length && !valeurs.montres.length && !valeurs.refus.length && !dejaConformes) return null;

        /* ⚠️⚠️ CE BLOC VIT DANS LA LIGNE DU LIEU, PAS DANS UNE LIGNE À PART.
           Deux raisons, et la première suffit : le tableau se TRIE par colonne —
           une ligne supplémentaire serait détachée de son lieu au premier clic
           sur un en-tête, et l'on lirait les champs d'un parking sous le nom
           d'un autre. La seconde : sept boucles parcourent `tbody` en supposant
           que chaque ligne porte ses champs de saisie. */
        const zone = doc.createElement('div');
        zone.className = 'peu-plus';

        /* ⚠️ Un libellé long déborde la pastille et chasse les suivantes hors de
           vue : il est raccourci ICI, et le complet reste en infobulle. */
        const pastille = (classe, titre, texte) => {
            const el = doc.createElement('span');
            el.className = 'peu-pastille peu-pastille-' + classe;
            const b = doc.createElement('b');
            b.textContent = titre.length > 24 ? titre.slice(0, 23) + '…' : titre;
            if (b.textContent !== titre) el.title = titre;
            el.appendChild(b);
            if (texte) el.appendChild(doc.createTextNode(' ' + texte));
            return el;
        };
        const titre = (texte) => {
            const el = doc.createElement('span');
            el.className = 'peu-pastille-titre';
            el.textContent = texte;
            return el;
        };

        if (poses.length) {
            zone.appendChild(titre(traduire('plusApplique')));
            poses.forEach(cible => {
                const perdu = pertes[cible];
                const el = pastille(perdu ? 'perte' : 'pose', libelleDeCible(cible),
                    valeurLisible(valeurs.aPoser[cible])
                    + (perdu ? ' — ' + traduire('plusRetire') + ' ' + valeurLisible(perdu) : ''));
                /* ⚠️ CE QUE LA VALEUR REMPLACE, EN INFOBULLE : un tableau écrase tout
                   son contenu dans WME, et sans cela on efface sans le savoir ce
                   qu'un autre éditeur avait renseigné. */
                const avant = attributs ? valeurDuLieu(attributs, cible) : undefined;
                if (avant !== undefined && avant !== null && String(avant) !== '') {
                    el.title = traduire('plusAvant') + ' ' + valeurLisible(avant);
                }
                zone.appendChild(el);
            });
        }

        /* Ce qui est déjà conforme se compte, il ne s'énumère pas : c'est le cas
           ordinaire quand le classeur porte l'état complet du parc. */
        if (dejaConformes) {
            zone.appendChild(titre(traduire('plusConforme') + ' ' + dejaConformes));
        }
        if (valeurs.montres.length) {
            zone.appendChild(titre(traduire('plusMain')));
            valeurs.montres.forEach(m => {
                const el = pastille('montre', m.libelle, valeurLisible(m.valeur));
                if (m.motif) el.title = m.motif;
                zone.appendChild(el);
            });
        }
        if (valeurs.refus.length) {
            zone.appendChild(titre(traduire('plusRefus')));
            valeurs.refus.forEach(r => {
                zone.appendChild(pastille('refus', r.libelle, '« ' + r.valeurs.join(' », « ') + ' »'));
            });
        }

        return zone;
    }
    // ==== /banc:apercu ====

    /* Le SDK, obtenu une seule fois. `@grant none` : il vient de la page.
       ⚠️ S'il manque (version de WME plus ancienne), les champs du lot D2 ne se
          posent pas — et cela DOIT se voir dans le rapport, pas se taire. */
    let _sdk = null;
    function obtenirSdk() {
        if (_sdk) return _sdk;
        if (typeof window.getWmeSdk !== 'function') {
            throw new Error('SDK de WME indisponible');
        }
        _sdk = window.getWmeSdk({ scriptId: 'poi-event-updater', scriptName: 'WME POI Event Updater' });
        return _sdk;
    }

    function getVenueIdFromPermalink(url) {
        // venues= peut contenir un ID numérique (ancien) ou un GUID alphanumérique
        // (nouveau), éventuellement plusieurs séparés par des virgules → on prend le 1er.
        const m = url.match(/venues=([^&,]+)/);
        return m ? decodeURIComponent(m[1]) : null;
    }

    // Extrait lat/lon d'un permalink en gérant les valeurs négatives
    // (hémisphère sud pour lat, ouest de Greenwich pour lon) → indispensable
    // pour un fonctionnement mondial. Retourne {lat, lon} ou null si absent/invalide.
    function parseLatLon(url) {
        try {
            const p = new URL(url).searchParams;
            const lat = parseFloat(p.get('lat'));
            const lon = parseFloat(p.get('lon'));
            if (isNaN(lat) || isNaN(lon)) return null;
            return { lat, lon };
        } catch { return null; }
    }

    // Retourne le statut de lock d'un venue par rapport au rang de l'éditeur connecté
    // 'ok'   : édition directe possible
    // 'sae'  : lock supérieur au rang → Suggest an Edit
    // 'hard' : lock niveau 7 (staff Waze uniquement)
    function getLockStatus(venue) {
        if (!venue) return 'ok';
        const lockRank = venue.attributes.lockRank ?? 0; // 0=L1 … 5=L6, 6=L7 staff
        const userRank = W?.loginManager?.user?.attributes?.rank ?? 0;
        if (lockRank >= 6) return 'hard';  // niveau 7 staff
        if (lockRank > userRank) return 'sae';
        return 'ok';
    }

    function makeDraggable(box, handle) {
        let startX, startY, startLeft, startTop;
        let geomReady = false; // passe à true une fois la position initiale posée

        const persist = () => {
            if (!geomReady) return;                            // pas avant le placement initial
            if (box.classList.contains('minimized')) return;   // ne pas mémoriser l'état réduit
            if (!box.offsetWidth || !box.offsetHeight) return; // box détachée/masquée → ignorer
            // On ne mémorise QUE la largeur et la position : la hauteur reste automatique
            // (elle s'adapte au nombre de POI de chaque événement, sinon elle resterait
            // figée sur une petite taille d'un événement précédent → tableau illisible).
            saveOverlayGeom({
                left: box.offsetLeft, top: box.offsetTop,
                width: box.offsetWidth
            });
        };

        // Géométrie initiale : restaure la taille/position mémorisées SI valides,
        // sinon largeur par défaut du CSS, ancrée en haut à droite (laisse voir la carte).
        const applyInitialGeom = () => {
            const saved = getOverlayGeom();
            const valid = saved && saved.width > 0
                          && Number.isFinite(saved.left) && Number.isFinite(saved.top);
            const vw = window.innerWidth, vh = window.innerHeight;
            // Largeur restaurée si valide ; hauteur JAMAIS forcée → s'adapte au contenu.
            if (valid) box.style.width = Math.min(saved.width, vw - 20) + 'px';
            const bw = box.offsetWidth, bh = box.offsetHeight;
            const left = valid ? saved.left : (vw - bw - 20); // défaut : coin haut-droit
            const top  = valid ? saved.top  : 64;              // sous le header WME
            // Clamp dans le viewport (la fenêtre a pu changer de taille depuis)
            box.style.left = Math.max(0, Math.min(vw - bw, left)) + 'px';
            box.style.top  = Math.max(0, Math.min(vh - bh, top)) + 'px';
            geomReady = true;
        };
        requestAnimationFrame(applyInitialGeom);

        // Mémorise la taille quand l'utilisateur redimensionne (poignée native, débounce léger)
        let saveTO = null;
        const ro = new ResizeObserver(() => {
            clearTimeout(saveTO);
            saveTO = setTimeout(persist, 300);
        });
        ro.observe(box);

        handle.addEventListener('mousedown', e => {
            // Ne pas déclencher sur les boutons
            if (e.target.closest('.peu-btn-icon')) return;
            e.preventDefault();
            startX = e.clientX; startY = e.clientY;
            startLeft = box.offsetLeft; startTop = box.offsetTop;
            document.addEventListener('mousemove', onMove);
            document.addEventListener('mouseup', onUp);
        });

        function onMove(e) {
            const dx = e.clientX - startX, dy = e.clientY - startY;
            const newLeft = Math.max(0, Math.min(window.innerWidth  - box.offsetWidth,  startLeft + dx));
            const newTop  = Math.max(0, Math.min(window.innerHeight - box.offsetHeight, startTop  + dy));
            box.style.left = newLeft + 'px';
            box.style.top  = newTop  + 'px';
        }

        function onUp() {
            document.removeEventListener('mousemove', onMove);
            document.removeEventListener('mouseup', onUp);
            persist();
        }
    }

    // ── Export rapport ───────────────────────────────────────────────────────
    function exportReport(eventName, results) {
        // results = [{name, oldName, newName, oldDesc, newDesc, status}]
        const wb = XLSX.utils.book_new();
        const headers = t('reportHeaders');
        const statusLabel = s => s === 'applied' ? t('statusApplied') : s === 'timeout' ? t('statusTimeout') : s;
        const rows = results.map(r => [r.oldName, r.newName, r.oldDesc, r.newDesc, statusLabel(r.status)]);
        const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);

        // Largeurs colonnes
        ws['!cols'] = [30, 30, 45, 45, 15].map(w => ({wch: w}));

        // Style en-tête (fond bleu, texte blanc) — XLSX.js lite ne supporte pas les styles
        // On préfixe le nom de l'onglet avec la date
        const dateStr = new Date().toLocaleDateString('fr-FR').replace(/\//g, '-');
        const sheetName = `${dateStr} ${eventName}`.slice(0, 31); // max 31 chars Excel
        XLSX.utils.book_append_sheet(wb, ws, sheetName);

        // Téléchargement
        const fileName = `POI_Report_${eventName.replace(/\s+/g, '_')}_${dateStr}.xlsx`;
        XLSX.writeFile(wb, fileName);
    }
    // ────────────────────────────────────────────────────────────────────────
    function centerAndLoad(permalink, vid, timeoutMs = 4000) {
        return new Promise(resolve => {
            const coords = parseLatLon(permalink);
            if (!coords) return resolve(null);
            const { lat, lon } = coords;

            // Zoom de préchargement volontairement large (16–17) : le lat/lon d'un
            // permalink cadre souvent la CARTE, pas le POI (ex. Fresnes 1 : venue à
            // 361 m du point du permalink). À zoom 19, un POI décalé tombe hors des
            // tuiles chargées et n'est jamais trouvé → « non chargé ». On suit le
            // zoomLevel du permalink en le plafonnant à 17. Au terme du balayage,
            // la carte est cadrée sur le PÉRIMÈTRE des lieux (cadrerSurLesLieux),
            // et non ramenée là où elle était avant.
            let z = parseInt(new URL(permalink).searchParams.get('zoomLevel'), 10);
            if (isNaN(z)) z = 17;
            z = Math.max(16, Math.min(z, 17));

            const lonlat = new OpenLayers.LonLat(lon, lat).transform(
                new OpenLayers.Projection('EPSG:4326'),
                W.map.getProjectionObject()
            );
            W.map.setCenter(lonlat, z);

            const t0 = Date.now();
            const poll = setInterval(() => {
                const v = W.model.venues.getObjectById(vid);
                // « Chargé » = venue présent avec un nom. On n'exige PAS isEditable() :
                // un POI verrouillé au-dessus du rang de l'éditeur est bien chargé,
                // il sera simplement proposé en Suggest an Edit (cf. getLockStatus).
                const ready = v && v.attributes && v.attributes.name;
                if (ready || Date.now() - t0 > timeoutMs) {
                    clearInterval(poll);
                    resolve(ready ? v : null);
                }
            }, 80);
        });
    }

    // ==== banc:carte ====
    // Extrait tel quel par tools/banc-carte.mjs : aucune dépendance (ni DOM, ni carte).

    /**
     * L'ENGLOBANT DE PLUSIEURS BOÎTES — la vue d'ensemble d'un périmètre.
     *
     * ⚠️ UNE BOÎTE SANS COORDONNÉES FINIES EST ÉCARTÉE, pas propagée : un seul
     *    `NaN` contaminerait `Math.min` et l'englobant entier deviendrait `NaN`,
     *    ce qui cadrerait la carte nulle part — sans la moindre erreur.
     *
     * @param {{left:number,bottom:number,right:number,top:number}[]} boites
     * @returns {?{left:number,bottom:number,right:number,top:number}}
     */
    function unionDesBoites(boites) {
        const fini = (v) => typeof v === 'number' && isFinite(v);
        const valides = (boites || []).filter(
            (b) => b && fini(b.left) && fini(b.bottom) && fini(b.right) && fini(b.top)
        );
        if (!valides.length) return null;

        return valides.reduce((a, b) => ({
            left:   Math.min(a.left, b.left),
            bottom: Math.min(a.bottom, b.bottom),
            right:  Math.max(a.right, b.right),
            top:    Math.max(a.top, b.top),
        }), {
            left: valides[0].left, bottom: valides[0].bottom,
            right: valides[0].right, top: valides[0].top,
        });
    }

    /**
     * UNE BOÎTE SANS ÉTENDUE — un seul lieu, ou plusieurs confondus.
     *
     * ⚠️ ELLE NE SE CADRE PAS : demander à la carte de tenir dans un point la
     *    pousse à son zoom maximal, et l'on se retrouve collé au sol sans rien
     *    voir autour. On centre alors, en gardant un zoom lisible.
     */
    function boiteSansEtendue(boite) {
        return !boite || (boite.right - boite.left === 0 && boite.top - boite.bottom === 0);
    }
    // ==== /banc:carte ====

    // Précharge tous les venues d'un événement en parcourant les permalinks
    // Appelle onProgress(loaded, total) à chaque étape
    // cancelRef.cancelled : si passé à true, interrompt proprement la boucle
    async function preloadVenues(pois, onProgress, cancelRef = {cancelled:false}) {
        const results = {};
        for (let i = 0; i < pois.length; i++) {
            if (cancelRef.cancelled) break;
            const p = pois[i];
            const vid = getVenueIdFromPermalink(p.perm);
            // Si déjà en mémoire, pas besoin de centrer
            let venue = W.model.venues.getObjectById(vid);
            if (!venue || !venue.attributes.name) {
                venue = await centerAndLoad(p.perm, vid);
            }
            results[vid] = venue;
            onProgress(i + 1, pois.length);
        }
        return results;
    }


    /**
     * CADRE LA CARTE SUR TOUS LES LIEUX DU FICHIER.
     *
     * ⭐⭐⭐⭐ LA VUE NE REVIENT PLUS EN ARRIÈRE. Le préchargement parcourt les
     *    lieux un par un, puis la vue était remise là où elle était avant — on
     *    voyait la carte balayer le terrain pour finir exactement au point de
     *    départ, sans rien montrer de ce qu'on venait de charger. Ce qu'on veut
     *    voir après un balayage, c'est le PÉRIMÈTRE qu'on vient de parcourir.
     *
     * ⚠️ SI LE CADRAGE ÉCHOUE, ON NE RESTAURE RIEN : la carte reste où le
     *    balayage l'a laissée, c'est-à-dire sur le dernier lieu. C'est le
     *    moindre des deux maux, et c'est encore un lieu du fichier.
     *
     * ⚠️ PLANCHER À ZOOM 12 : sous ce seuil WME décharge les objets, et l'on
     *    perdrait les venues qu'on vient de précharger.
     */
    function cadrerSurLesLieux(venues) {
        const boites = Object.keys(venues || {}).map((cle) => {
            try {
                const g = venues[cle] && venues[cle].getOLGeometry && venues[cle].getOLGeometry();
                return g && g.getBounds ? g.getBounds() : null;
            } catch (e) {
                return null;
            }
        });

        const u = unionDesBoites(boites);
        if (!u) return false;

        if (boiteSansEtendue(u)) {
            W.map.setCenter(
                new OpenLayers.LonLat((u.left + u.right) / 2, (u.bottom + u.top) / 2),
                Math.min(Math.max(W.map.getZoom(), 17), 19)
            );

            return true;
        }

        const etendue = new OpenLayers.Bounds(u.left, u.bottom, u.right, u.top);
        if (typeof W.map.zoomToExtent === 'function') {
            W.map.zoomToExtent(etendue);
            /* ⚠️⚠️ UN PLAFOND, ET PAS SEULEMENT UN PLANCHER. Cadrer sur UN SEUL lieu
               colle la carte au sol : l'echelle tombe a deux metres, on ne voit
               plus ni la rue, ni les lieux voisins, ni ou l'on est. Le plancher
               protege des lieux disperses, le plafond du lieu unique — et le
               second se rencontre bien plus souvent que le premier. */
            const z = W.map.getZoom();
            if (z < 12) W.map.setCenter(etendue.getCenterLonLat(), 12);
            else if (z > 19) W.map.setCenter(etendue.getCenterLonLat(), 19);

            return true;
        }
        W.map.setCenter(etendue.getCenterLonLat());

        return true;
    }

    // ==== banc:coque ====
    // Extrait tel quel par tools/banc-coque.mjs : rend du texte, ne touche a rien.

    /**
     * L'OSSATURE DE LA FENETRE, rendue en HTML.
     *
     * ⭐⭐⭐ ELLE REND DU TEXTE, ET C'EST CE QUI LA REND MESURABLE. Construite a
     *    coups de createElement au milieu du DOM et de la carte, l'ancienne
     *    interface n'etait verifiable que par l'oeil, dans l'editeur — et c'est
     *    par l'oeil, tres tard, qu'on a vu que la case a cocher des lignes
     *    etait introuvable et que le bouton qui ECRIT SUR LA CARTE etait un
     *    glyphe colle a celui qui ferme la fenetre.
     *
     * ⚠️ AUCUNE DONNEE EXTERNE ICI : rien a echapper. Le nom du fichier, les
     *    noms de lieux et tout ce qui vient du classeur entrent plus tard, par
     *    textContent.
     */
    function coqueOverlay(version) {
        return ''
            + '<div class="peu-header" id="peu-header">'
            +   '<div class="peu-header-left">'
            +     '<span aria-hidden="true">' + PEU_EMOJI + '</span>'
            +     '<span>' + esc(t('panelTitle')) + '</span>'
            +     '<span class="peu-header-version">v' + esc(version) + '</span>'
            +   '</div>'
            +   '<div class="peu-header-btns">'
            +     '<button type="button" class="peu-btn-icon" id="peu-btn-replier" title="' + esc(t('btnReduce')) + '">-</button>'
            +     '<button type="button" class="peu-btn-icon" id="peu-btn-fermer" title="' + esc(t('btnClose')) + '">X</button>'
            +   '</div>'
            + '</div>'
            + '<div id="peu-strip">'
            +   '<span class="peu-strip-dot"></span>'
            +   '<span class="peu-strip-file" id="peu-strip-texte">' + esc(t('noFile')) + '</span>'
            +   '<span class="peu-strip-sep" id="peu-strip-sep1" hidden>&middot;</span>'
            +   '<select class="peu-select" id="peu-select-onglet" hidden title="' + esc(t('selectSheet')) + '"></select>'
            +   '<span class="peu-strip-sep" id="peu-strip-sep2" hidden>&middot;</span>'
            +   '<span class="peu-strip-info" id="peu-strip-compte" hidden></span>'
            +   '<button type="button" class="peu-btn peu-btn-neutral peu-btn-sm" id="peu-btn-fichier"'
            +          ' title="' + esc(t('chooseFileTitle')) + '" style="margin-inline-start:auto">'
            +     esc(t('chooseFile')) + '</button>'
            + '</div>'
            + '<div id="peu-body"></div>'
            + '<div class="peu-footer" id="peu-footer">'
            +   '<div class="peu-footer-help" id="peu-footer-help">' + esc(t('footerHelpVide')) + '</div>'
            +   '<div class="peu-footer-actions">'
            +     '<button type="button" class="peu-btn peu-btn-neutral peu-btn-sm" id="peu-btn-export" hidden'
            +            ' title="' + esc(t('btnExportTitle')) + '">' + esc(t('btnExport')) + '</button>'
            +     '<button type="button" class="peu-btn peu-btn-primary" id="peu-btn-appliquer" disabled'
            +            ' title="' + esc(t('btnApplyTitle')) + '">' + esc(t('btnApply')) + '</button>'
            +   '</div>'
            + '</div>'
            + '<div id="peu-resize" aria-hidden="true"></div>';
    }

    /**
     * LE LIBELLE DU BOUTON QUI ECRIT SUR LA CARTE.
     *
     * ⭐⭐⭐⭐ IL DIT COMBIEN DE LIGNES IL VA POSER. « Appliquer » seul ne dit pas
     *    sur quoi : on clique sans savoir si l'on touche un lieu ou quarante.
     *    Le compte est la seule chose qui distingue un geste anodin d'un geste
     *    qu'on veut relire avant.
     *
     * ⚠️ ZERO EST UN RESULTAT : on ne cache pas le bouton, on le desactive en
     *    disant qu'il n'y a rien de coche. Un bouton qui disparait se lit comme
     *    une panne.
     */
    function libelleAppliquer(nbCoche) {
        return nbCoche === 0 ? t('btnApplyNone')
             : nbCoche === 1 ? t('btnApplyOne')
             : t('btnApplyN', nbCoche);
    }
    // ==== /banc:coque ====

    // ==== banc:ligne ====
    // Extrait tel quel par tools/banc-ligne.mjs : rend du texte, ne touche a rien.

    /**
     * L'ETAT D'UNE LIGNE, EN UN MOT.
     *
     * ⭐⭐⭐⭐ IL Y A UN ORDRE, ET IL N'EST PAS ARBITRAIRE. Ce qu'on ne peut PAS
     *    faire passe avant ce qu'on ferait : un lieu introuvable ou verrouille
     *    au-dessus du rang ne sera pas pose, quoi qu'il ait a changer, et
     *    l'annoncer « 4 champs a poser » serait une promesse qu'on ne tient pas.
     *    Vient ensuite la PERTE, parce qu'elle demande un geste ; puis
     *    l'ecart ordinaire ; puis rien.
     */
    function etatDeLaLigne(infos) {
        if (!infos.charge)      return 'unloaded';
        if (infos.verrou === 'hard') return 'hard';
        if (infos.verrou === 'sae')  return 'sae';
        if (infos.pertes > 0)   return 'perte';
        if (infos.nomChange || infos.descChange || infos.champsDiff > 0) return 'diff';

        return 'ok';
    }

    /**
     * LE BADGE QUI DIT CET ETAT — son texte, sa classe, son infobulle.
     *
     * ⚠️ UN BADGE NE REMPLACE PAS LE LISERE : il se rate, et il disparait sous
     *    le survol quand on ne l'accompagne pas d'un fond. Les trois signes vont
     *    ensemble, c'est la regle de la refonte.
     */
    function badgeDeLigne(etat, infos) {
        if (etat === 'unloaded') return { classe: 'peu-badge-off',   texte: '?',  titre: t('badgeOffTitle') };
        if (etat === 'hard')     return { classe: 'peu-badge-lock',  texte: 'L' + (infos.niveau || ''), titre: t('lockHardTitle') };
        if (etat === 'sae')      return { classe: 'peu-badge-sae',   texte: 'SaE', titre: t('badgeSaeTitle') };
        if (etat === 'perte')    return { classe: 'peu-badge-perte', texte: '-' + infos.pertes, titre: t('badgePerteTitle', infos.pertes) };
        if (etat === 'diff') {
            const n = (infos.nomChange ? 1 : 0) + (infos.descChange ? 1 : 0) + infos.champsDiff;
            return { classe: 'peu-badge-diff', texte: String(n), titre: t('badgeDiffTitle', n) };
        }

        return { classe: 'peu-badge-ok', texte: '=', titre: t('badgeRienTitle') };
    }

    /**
     * UNE LIGNE DE L'APERCU, EN HTML.
     *
     * ⭐⭐⭐⭐ LA CASE EST EN PREMIERE COLONNE. Elle vivait en quatrieme et
     *    derniere, large de trente pixels, sous un en-tete VIDE, a cote d'une
     *    case maitre qui lui ressemble — et le pied de page disait « decochez
     *    les lignes a exclure » en designant quelque chose que personne ne
     *    voyait. On la cherche en tete de ligne : c'est le reflexe de tous les
     *    tableaux, et c'est la qu'elle doit etre.
     *
     * ⚠️ RIEN N'EST COCHE D'OFFICE ICI : c'est `cocherDOffice` qui decide, apres
     *    coup, et le banc du cochage la tient. Une ligne qui RETIRE quelque
     *    chose ne se coche jamais seule.
     *
     * ⚠️ TOUT CE QUI VIENT DU CLASSEUR PASSE PAR esc() : un nom de lieu est une
     *    donnee externe, et il finit dans un attribut title=.
     */
    function ligneApercu(vue) {
        const etat = etatDeLaLigne(vue);
        const badge = badgeDeLigne(etat, vue);
        const fige = etat === 'unloaded' || etat === 'hard';
        const classes = 'peu-ligne peu-row-' + etat;

        const vieux = (val, change) => '<div class="peu-cell-old' + (change ? ' changed' : '') + '"'
            + (val ? ' title="' + esc(val) + '"' : '') + '>' + (val ? esc(val) : '&mdash;') + '</div>';

        return '<tr class="' + classes + '" data-idx="' + Number(vue.idx) + '">'
            + '<td class="center">'
            +   '<label class="peu-check"><input type="checkbox" class="peu-checkbox"'
            +     (fige ? ' disabled' : '') + ' title="' + esc(t(fige ? 'cbFige' : 'cbTitre')) + '"></label>'
            + '</td>'
            + '<td class="center">'
            +   '<button type="button" class="peu-btn-center" data-centrer'
            +     ' title="' + esc(t('locateTitle')) + '" aria-label="' + esc(t('locateTitle')) + '">&#127919;</button>'
            +   '<span class="peu-badge ' + badge.classe + '" title="' + esc(badge.titre) + '">' + esc(badge.texte) + '</span>'
            + '</td>'
            + '<td>' + vieux(vue.ancienNom, vue.nomChange)
            +   '<input type="text" class="peu-input" data-nom' + (fige ? ' disabled' : '')
            +     ' value="' + esc(vue.nom) + '" title="' + esc(t('colNameTitle')) + '"></td>'
            + '<td>' + vieux(vue.ancienDesc, vue.descChange)
            +   '<textarea class="peu-textarea" data-desc rows="1"' + (fige ? ' disabled' : '')
            +     ' title="' + esc(t('colDescTitle')) + '">' + esc(vue.desc) + '</textarea></td>'
            + '</tr>';
    }

    /**
     * L'EN-TETE DU TABLEAU.
     *
     * ⚠️ LA COLONNE DE LA CASE PORTE UN INTITULE. Vide, elle ne disait pas ce
     *    que la case decide — et c'est ce qui la rendait invisible.
     */
    function enteteApercu() {
        return '<colgroup><col><col><col><col></colgroup>'
            + '<thead><tr>'
            + '<th class="center" title="' + esc(t('masterCbTitle')) + '">'
            +   '<label class="peu-check"><input type="checkbox" class="peu-checkbox" data-maitre'
            +     ' title="' + esc(t('masterCbTitle')) + '"></label>'
            +   '<div style="font-size:.833em;font-weight:700;margin-top:2px">' + esc(t('colSelect')) + '</div>'
            + '</th>'
            + '<th class="center">' + esc(t('colEtat')) + '</th>'
            + '<th class="sortable" data-tri="nom" title="' + esc(t('triTitre')) + '">'
            +   esc(t('colName')) + '<i class="peu-sort-icon">&#9650;</i></th>'
            + '<th class="sortable" data-tri="desc" title="' + esc(t('triTitre')) + '">'
            +   esc(t('colDesc')) + '<i class="peu-sort-icon">&#9650;</i></th>'
            + '</tr></thead>';
    }
    // ==== /banc:ligne ====

    // ==== banc:geometrie ====
    // Extrait tel quel par tools/banc-fenetre.mjs : du calcul, pas de DOM.

    /**
     * RAMENE UNE FENETRE DANS LES BORNES DE LA CARTE.
     *
     * ⭐⭐⭐ ON MESURE, ON NE SUPPOSE PAS. Un `calc(100vh - 110px)` suppose la
     *    hauteur du bandeau de WME, qui change avec la version, la langue et la
     *    presence d'un autre script. Les bornes viennent donc de la carte
     *    elle-meme, relevees dans le DOM, et ce calcul-ci ne fait que les
     *    respecter.
     *
     * ⚠️ UNE POSITION HORS BORNES EST REFUSEE, PAS RABOTEE quand elle vient
     *    d'une session precedente : l'ecran a pu changer de taille, et rabattre
     *    une fenetre dans un coin sans le dire donne l'impression qu'elle a
     *    disparu. Ici on rabat — mais l'appelant, lui, sait distinguer les deux
     *    cas par `tenait`.
     */
    function bornerFenetre(geo, bornes) {
        /* ⚠️⚠️ LE PLANCHER NE PASSE JAMAIS DEVANT LA CARTE. Un plancher pose en
           dernier (`Math.max(280, …)`) l'emporte sur la mesure et fait SORTIR la
           fenetre d'une carte etroite — donc son pied, donc le bouton qui
           applique. Le plancher protege d'une fenetre reduite a rien par la
           poignee ; il ne decide pas de la place disponible. */
        const dispoL = bornes.droite - bornes.gauche;
        const dispoH = bornes.bas - bornes.haut;
        const largeur = Math.min(Math.max(Math.min(geo.w, dispoL), 280), dispoL);
        const hauteur = Math.min(Math.max(Math.min(geo.h, dispoH), 120), dispoH);
        const x = Math.max(bornes.gauche, Math.min(geo.x, bornes.droite - largeur));
        const y = Math.max(bornes.haut,   Math.min(geo.y, bornes.bas - hauteur));

        return {
            x: x, y: y, w: largeur, h: hauteur,
            tenait: x === geo.x && y === geo.y && largeur === geo.w && hauteur === geo.h,
        };
    }

    /**
     * OU SE POSE LA FENETRE QUAND ELLE N'A PAS DE POSITION MEMORISEE.
     *
     * ⚠️ A GAUCHE DES BOUTONS DE CARTE, jamais dessus : c'est la colonne ou vit
     *    le bouton qui ouvre cette fenetre, et le masquer reviendrait a cacher
     *    la poignee de la porte qu'on vient de franchir.
     */
    function positionParDefaut(bornes, largeur) {
        const l = Math.min(largeur, bornes.droite - bornes.gauche);

        return {
            x: Math.max(bornes.gauche, bornes.droite - l),
            y: bornes.haut,
            w: l,
            h: bornes.bas - bornes.haut,
        };
    }
    // ==== /banc:geometrie ====

    /* ----------------------------------------------------------------------
       LE BOUTON DE CARTE
       ---------------------------------------------------------------------- */

    /** Le conteneur natif des boutons de carte, ou rien s'il n'est pas encore la. */
    function conteneurBoutonsCarte() {
        return document.querySelector('.overlay-buttons-container.top')
            || document.querySelector('.overlay-buttons-container');
    }

    /**
     * POSE LE BOUTON DANS LA COLONNE DES BOUTONS DE CARTE.
     *
     * ⭐⭐⭐⭐ DOCKE, JAMAIS EN position:fixed. Docke, il suit le zoom et la
     *    resolution, et partage le contexte d'empilement des boutons natifs —
     *    donc il passe DERRIERE le panneau des calques comme eux. En fixed, il
     *    passerait par-dessus, et il faudrait une bagarre de z-index sans fin.
     *
     * ⚠️ IL EST LE SEUL ACCES AU SCRIPT : s'il disparait, tout disparait. WME
     *    re-rend cette colonne, et le bouton part avec. D'ou le filet de
     *    `installerFab`.
     */
    function poserFab() {
        const cont = conteneurBoutonsCarte();
        if (!cont) return false;
        if (cont.querySelector('#peu-fab-wrap')) return true;

        // Un exemplaire detache par un re-rendu precedent ne doit pas rester.
        const vieux = document.querySelector('#peu-fab-wrap');
        if (vieux) vieux.remove();

        const wrap = document.createElement('div');
        wrap.id = 'peu-fab-wrap';
        wrap.innerHTML = '<button type="button" id="peu-fab-btn" title="' + esc(t('fabTitle')) + '">'
            + '<img src="' + TAB_ICON + '" alt="" width="22" height="22" style="display:block">'
            + '<span class="peu-fab-badge" id="peu-fab-badge"></span>'
            + '</button>';
        cont.appendChild(wrap);
        wrap.querySelector('button').addEventListener('click', basculerOverlay);
        majFab();

        return true;
    }

    /**
     * ⚠️ UN INTERVALLE, ET NON UN MutationObserver. L'observateur a ete essaye
     *    dans le script voisin : il ne reposait pas le bouton en direct et
     *    coutait bien plus cher. Deux secondes suffisent — personne ne remarque
     *    un bouton qui revient en deux secondes, tout le monde remarque un
     *    bouton qui ne revient jamais.
     */
    function installerFab() {
        poserFab();
        setInterval(poserFab, 2000);
        document.addEventListener('visibilitychange', () => { if (!document.hidden) poserFab(); });
    }

    /** Le bouton dit ce qu'il fera au clic, et ce qui est charge. */
    function majFab() {
        const btn = document.getElementById('peu-fab-btn');
        if (!btn) return;
        const ouvert = !!document.querySelector('#peu-overlay.peu-open');
        btn.classList.toggle('peu-fab-on', ouvert);
        btn.title = t(ouvert ? 'fabTitleOn' : 'fabTitle');

        const badge = document.getElementById('peu-fab-badge');
        const nb = (poiData || []).length;
        btn.classList.toggle('peu-has-file', nb > 0);
        if (badge) badge.textContent = nb ? String(nb) : '';
    }

    /* ----------------------------------------------------------------------
       LA FENETRE
       ---------------------------------------------------------------------- */

    /**
     * LES BORNES DE LA CARTE, MESUREES.
     *
     * ⚠️ TROIS MESURES, ET AUCUNE SUPPOSITION : le bord de la carte, son pied
     *    de page, et la colonne des boutons. Un repli prudent si l'un manque —
     *    mieux vaut une fenetre un peu large qu'une fenetre introuvable.
     */
    function bornesCarte() {
        const carte = document.getElementById('WazeMap');
        const pied  = document.querySelector('.wz-map-ol-footer');
        const btns  = conteneurBoutonsCarte();
        const r = carte ? carte.getBoundingClientRect() : { left: 0, top: 40, right: window.innerWidth, bottom: window.innerHeight };
        const rb = btns ? btns.getBoundingClientRect() : null;

        return {
            gauche: Math.round(r.left) + 6,
            haut:   Math.round(r.top) + 6,
            droite: Math.round(rb ? rb.left - 8 : r.right - 8),
            bas:    Math.round(pied ? pied.getBoundingClientRect().top - 6 : r.bottom - 6),
        };
    }

    /** La geometrie mise de cote, ou rien tant que l'editeur n'a rien touche. */
    function lireGeometrie() {
        try {
            const brut = localStorage.getItem(GEO_KEY);
            const g = brut ? JSON.parse(brut) : null;

            return g && ['x', 'y', 'w', 'h'].every(k => typeof g[k] === 'number' && isFinite(g[k])) ? g : null;
        } catch (e) { return null; }
    }

    function ecrireGeometrie(g) {
        try { localStorage.setItem(GEO_KEY, JSON.stringify(g)); } catch (e) { /* stockage refuse : tant pis */ }
    }

    /** Pose la fenetre : sa position memorisee si elle tient encore, sinon la place par defaut. */
    function placerFenetre(ov) {
        const bornes = bornesCarte();
        const memo = lireGeometrie();
        const voulu = memo || positionParDefaut(bornes, ov.offsetWidth || 820);
        const g = bornerFenetre(voulu, bornes);

        ov.style.left = g.x + 'px';
        ov.style.top = g.y + 'px';
        ov.style.right = 'auto';
        ov.style.width = g.w + 'px';
        ov.style.maxHeight = (bornes.bas - g.y) + 'px';
        if (memo) ov.style.height = g.h + 'px';
    }

    /**
     * DEPLACEMENT PAR L'EN-TETE.
     *
     * ⚠️ ON N'ENREGISTRE QU'AU RELACHEMENT, jamais a chaque mouvement : ecrire
     *    dans le stockage soixante fois par seconde fait sauter le glissement.
     * ⚠️ LE maxHeight SUIT LA POSITION, sinon descendre la fenetre fait sortir
     *    son pied — celui qui porte le bouton qui applique — hors de l'ecran.
     */
    function rendreDeplacable(ov, poignee) {
        let ox = 0, oy = 0, actif = false;

        poignee.addEventListener('mousedown', (e) => {
            if (e.target.closest('button')) return;
            actif = true;
            const r = ov.getBoundingClientRect();
            ox = e.clientX - r.left;
            oy = e.clientY - r.top;
            e.preventDefault();
        });

        document.addEventListener('mousemove', (e) => {
            if (!actif) return;
            const bornes = bornesCarte();
            const g = bornerFenetre(
                { x: e.clientX - ox, y: e.clientY - oy, w: ov.offsetWidth, h: ov.offsetHeight },
                bornes
            );
            ov.style.left = g.x + 'px';
            ov.style.top = g.y + 'px';
            ov.style.right = 'auto';
            ov.style.maxHeight = Math.max(120, bornes.bas - g.y) + 'px';
        });

        document.addEventListener('mouseup', () => {
            if (!actif) return;
            actif = false;
            const r = ov.getBoundingClientRect();
            ecrireGeometrie({ x: Math.round(r.left), y: Math.round(r.top), w: Math.round(r.width), h: Math.round(r.height) });
        });

        /* ⭐ DOUBLE-CLIC SUR L'EN-TETE : retour au dimensionnement automatique.
           Une fenetre qu'on a malmenee doit pouvoir revenir sans qu'on cherche
           ou est le reglage. */
        poignee.addEventListener('dblclick', (e) => {
            if (e.target.closest('button')) return;
            try { localStorage.removeItem(GEO_KEY); } catch (err) { /* rien a oublier */ }
            ov.style.height = '';
            placerFenetre(ov);
        });
    }

    /**
     * REDIMENSIONNEMENT PAR LA POIGNEE DU COIN.
     *
     * ⚠️ LE COIN HAUT-GAUCHE SE FIGE AU PREMIER GESTE : tant que la fenetre est
     *    posee par `right`, l'elargir la fait fuir sous le curseur.
     */
    function rendreRedimensionnable(ov, poignee) {
        let actif = false;

        poignee.addEventListener('mousedown', (e) => {
            actif = true;
            const r = ov.getBoundingClientRect();
            ov.style.left = Math.round(r.left) + 'px';
            ov.style.top = Math.round(r.top) + 'px';
            ov.style.right = 'auto';
            e.preventDefault();
            e.stopPropagation();
        });

        document.addEventListener('mousemove', (e) => {
            if (!actif) return;
            const r = ov.getBoundingClientRect();
            const bornes = bornesCarte();
            const g = bornerFenetre(
                { x: r.left, y: r.top, w: e.clientX - r.left, h: e.clientY - r.top },
                bornes
            );
            ov.style.width = g.w + 'px';
            ov.style.height = g.h + 'px';
            ov.style.maxHeight = Math.max(120, bornes.bas - g.y) + 'px';
        });

        document.addEventListener('mouseup', () => {
            if (!actif) return;
            actif = false;
            const r = ov.getBoundingClientRect();
            ecrireGeometrie({ x: Math.round(r.left), y: Math.round(r.top), w: Math.round(r.width), h: Math.round(r.height) });
        });
    }

    /**
     * CONSTRUIT LA FENETRE, UNE FOIS.
     *
     * ⚠️⚠️ LES EVENEMENTS DE LA SOURIS S'ARRETENT A LA FENETRE. Sans cela, la
     *    molette zoome la carte pendant qu'on fait defiler la liste, et un clic
     *    dans un champ desselectionne ce qu'on regardait.
     */
    function construireOverlay() {
        let ov = document.getElementById('peu-overlay');
        if (ov) return ov;

        ov = document.createElement('div');
        ov.id = 'peu-overlay';
        ov.innerHTML = coqueOverlay(PEU_VERSION);
        document.body.appendChild(ov);

        ['wheel', 'mousedown', 'dblclick', 'contextmenu'].forEach(
            (evt) => ov.addEventListener(evt, (e) => e.stopPropagation())
        );

        brancherDepot(ov);
        rendreDeplacable(ov, ov.querySelector('#peu-header'));
        rendreRedimensionnable(ov, ov.querySelector('#peu-resize'));

        ov.querySelector('#peu-btn-fermer').addEventListener('click', fermerOverlay);

        /* ⭐ LE SEUL BOUTON PLEIN DE LA FENETRE : l'etape suivante. Il etait un
           glyphe d'un caractere dans la barre de titre, colle a celui qui ferme. */
        ov.querySelector('#peu-btn-appliquer').addEventListener('click', () => { appliquerLignes(); });

        ov.querySelector('#peu-btn-export').addEventListener('click', () => {
            if (_apercu && _apercu.resultats) exportReport(_apercu.eventName, _apercu.resultats);
        });

        /* ⚠️ LE CHOIX DU FICHIER VIT DANS LA FENETRE, pas dans le panneau lateral :
           c'est le premier geste du travail, et le travail est ici. Mais il passe
           par LE MEME champ que le panneau — un second champ serait un second
           chemin de lecture, et les deux divergeraient. */
        ov.querySelector('#peu-btn-fichier').addEventListener('click', () => {
            if (_peuFileInput) _peuFileInput.click();
        });

        /* ⚠️ ON NE RELANCE PAS UN BALAYAGE POUR RIEN : rechoisir l'onglet deja
           affiche relancerait la lecture de tous les lieux sans rien changer. */
        ov.querySelector('#peu-select-onglet').addEventListener('change', (e) => {
            if (e.target.value && (!_apercu || _apercu.eventName !== e.target.value)) {
                ouvrirApercu(e.target.value);
            }
        });
        ov.querySelector('#peu-btn-replier').addEventListener('click', () => {
            const replie = ov.classList.toggle('peu-replie');
            const b = ov.querySelector('#peu-btn-replier');
            b.textContent = replie ? '+' : '-';
            b.title = t(replie ? 'btnRestore' : 'btnReduce');
        });

        window.addEventListener('resize', () => {
            if (ov.classList.contains('peu-open')) placerFenetre(ov);
        });

        return ov;
    }

    function ouvrirOverlay() {
        const ov = construireOverlay();
        ov.classList.add('peu-open');
        placerFenetre(ov);
        /* ⭐ RIEN DE CHARGE ⇒ ON DIT PAR OU COMMENCER. Une fenetre vide avec un
           bouton grise n'explique rien. */
        if (!ov.querySelector('#peu-body').children.length) {
            montrerGuide(poiData.length ? 'guideOnglet' : 'guideFichier',
                poiData.length ? 'guideOngletSuite' : 'guideFichierSuite');
        }
        majFab();
    }

    function fermerOverlay() {
        const ov = document.getElementById('peu-overlay');
        if (ov) ov.classList.remove('peu-open');
        majFab();
    }

    function basculerOverlay() {
        const ov = document.getElementById('peu-overlay');
        if (ov && ov.classList.contains('peu-open')) fermerOverlay(); else ouvrirOverlay();
    }

    async function initScript() {
        _peuLang = detectLang();
        injectCSS();
        const { tabLabel, tabPane } = W.userscripts.registerSidebarTab(scriptId);
        // Icône (pin) à la place du nom, nom conservé en infobulle.
        tabLabel.textContent = '';
        tabLabel.style.display = 'flex';
        tabLabel.style.alignItems = 'center';
        tabLabel.style.justifyContent = 'center';
        tabLabel.style.height = '100%';
        const tabIcon = document.createElement('img');
        tabIcon.src = TAB_ICON;
        tabIcon.alt = t('tabTitle');
        tabIcon.width = 18;
        tabIcon.height = 18;
        tabIcon.style.display = 'block';
        tabLabel.appendChild(tabIcon);
        tabLabel.title = t('tabTitle');
        await W.userscripts.waitForElementConnected(tabPane);
        // Le conteneur d'onglet (<a> parent) est étiré sur toute la hauteur de
        // l'onglet ; on le centre aussi pour placer l'icône au milieu vertical
        // (sinon elle se colle en haut). Vérifié en direct dans WME.
        const tabLink = tabLabel.parentElement;
        if (tabLink) {
            tabLink.style.display = 'flex';
            tabLink.style.alignItems = 'center';
            tabLink.style.justifyContent = 'center';
        }

        /* ⭐⭐⭐⭐ LE PANNEAU PORTE LES REGLAGES, LA FENETRE PORTE LE TRAVAIL. Ce
           panneau fait disparaitre son contenu des qu'on selectionne un objet sur
           la carte : on ne peut pas y travailler. Tout l'operationnel — choisir
           un classeur, choisir l'onglet, relire, appliquer — a donc demenage dans
           la fenetre, qui ne disparait pas.

           ⚠️ PLUS UNE LIGNE DE STYLE EN DUR ICI : le panneau se construisait en
              style.cssText, bouton par bouton, et c'est pour cela que rien
              n'etait homogene — il n'y avait rien a quoi etre homogene. */
        const container = document.createElement('div');
        container.className = 'peu-container';

        const title = document.createElement('h3');
        title.textContent = PEU_EMOJI + ' ' + t('panelTitle');
        container.appendChild(title);

        /* ⚠️ HORS DES SECTIONS, ET EN PREMIER : « afficher la fenetre » est ce
           qu'on vient chercher ici neuf fois sur dix. Range sous un titre, il
           serait a trouver. */
        const btnFenetre = document.createElement('button');
        btnFenetre.type = 'button';
        btnFenetre.className = 'peu-btn peu-btn-primary peu-btn-full';
        btnFenetre.textContent = t('sbOuvrir');
        btnFenetre.title = t('sbOuvrirTitre');
        btnFenetre.addEventListener('click', ouvrirOverlay);
        container.appendChild(btnFenetre);

        const fileInput = document.createElement('input');
        fileInput.type = 'file'; fileInput.accept = '.xlsx,.xls'; fileInput.style.display = 'none';
        _peuFileInput = fileInput;
        container.appendChild(fileInput);

        /* ⚠️ LE REPLI DE LA BIBLIOTHEQUE RESTE, ET IL DOIT SE VOIR. Si XLSX n'a pas
           pu se charger (reseau coupe, CDN bloque), rien ne fonctionnera — et le
           dire ici vaut mieux que de laisser le script echouer au premier clic. */
        if (typeof XLSX === 'undefined') {
            const alerte = document.createElement('div');
            alerte.className = 'peu-alert peu-alert-warn';
            alerte.style.margin = '8px 0';
            alerte.textContent = t('xlsxMissing');
            container.appendChild(alerte);
            btnFenetre.disabled = true;
        }

        /* Le titre de section, et la zone d'historique en dessous. */
        const titreHist = document.createElement('div');
        titreHist.className = 'peu-side-sect';
        titreHist.textContent = t('historyTitle');
        container.appendChild(titreHist);

        const historyDiv = document.createElement('div');
        container.appendChild(historyDiv);

        /* ⚠️ L'EMPLACEMENT DES REGLAGES EST RESERVE, ET IL LE DIT. Une section
           vide sans un mot se lit comme un defaut d'affichage. */
        const titreReg = document.createElement('div');
        titreReg.className = 'peu-side-sect';
        titreReg.textContent = t('sbReglages');
        container.appendChild(titreReg);

        const noteReg = document.createElement('div');
        noteReg.className = 'peu-hist-meta';
        noteReg.textContent = t('sbReglagesNote');
        container.appendChild(noteReg);

        function renderHistory() {
            historyDiv.innerHTML = '';
            const history = getHistory();
            if (!history.length) return;

            // En-tête avec bouton RAZ
            const hheader = document.createElement('div');
            hheader.className = 'peu-side-sect';
            const htitle = document.createElement('div');
            htitle.style.flex = '1';
            htitle.textContent = t('historyTitle');
            const btnRaz = document.createElement('button');
            btnRaz.textContent = '🗑';
            btnRaz.title = t('clearHistoryTitle');
            btnRaz.className = 'peu-btn-center';
            btnRaz.onmouseenter = () => btnRaz.style.color = '#c0392b';
            btnRaz.onmouseleave = () => btnRaz.style.color = '#aaa';
            btnRaz.onclick = () => {
                saveHistory([]);
                renderHistory();
            };
            hheader.appendChild(htitle); hheader.appendChild(btnRaz);
            historyDiv.appendChild(hheader);

            history.forEach(h => {
                const row = document.createElement('div');
                row.className = 'peu-hist-row';
                const name = document.createElement('div');
                name.className = 'peu-hist-name';
                name.textContent = h.name; name.title = h.name;
                const loaded = document.createElement('div');
                loaded.className = 'peu-hist-meta';
                loaded.textContent = `${t('histLoaded')} ${formatDateTime(h.loaded)}`;
                const applied = document.createElement('div');
                applied.className = 'peu-hist-meta';
                applied.textContent = h.applied ? `${t('histApplied')} ${formatDateTime(h.applied)}` : t('histNeverApplied');
                row.appendChild(name); row.appendChild(loaded); row.appendChild(applied);
                historyDiv.appendChild(row);
            });
        }

        tabPane.appendChild(container);

        /* ⭐ LE BOUTON DE CARTE EST LE SEUL ACCÈS AU TRAVAIL. Le panneau latéral
           garde les réglages et l’historique : il fait disparaître son contenu
           dès qu’on sélectionne un objet sur la carte, on ne peut pas y
           travailler. */
        installerFab();
        renderHistory();

        /**
         * LE RAPPORT D'ANOMALIES — dans la FENETRE, pas dans le panneau.
         *
         * ⭐ IL DIT CE QUI N'EST PAS ENTRE. Une ligne ecartee du classeur ne se
         *    voit nulle part ailleurs : ni dans le tableau, qui ne montre que ce
         *    qui est retenu, ni sur la carte. Sans ce rapport, le fichier parait
         *    complet et il ne l'est pas.
         */
        function showValidationReport(warnings) {
            const ov = document.getElementById('peu-overlay');
            const corps = ov ? ov.querySelector('#peu-body') : null;
            if (!corps) return;
            const vieux = corps.querySelector('[data-rapport]');
            if (vieux) vieux.remove();
            if (!warnings.length) return;
            const validationReport = document.createElement('div');
            validationReport.className = 'peu-alert peu-alert-warn';
            validationReport.setAttribute('data-rapport', '');
            corps.prepend(validationReport);
            const title = document.createElement('div');
            title.className = 'peu-error-title';
            title.textContent = t('anomalies', warnings.length) + ' :';
            validationReport.appendChild(title);
            const ul = document.createElement('ul');
            warnings.forEach(w => {
                const li = document.createElement('li'); li.style.marginBottom = '2px';
                li.textContent = w; ul.appendChild(li);
            });
            validationReport.appendChild(ul);
        }

        fileInput.addEventListener('change', e => {
            const file = e.target.files[0];
            /* ⚠️ ON REPART DE ZERO A CHAQUE FICHIER : laisser le bandeau de
               l'ancien classeur pendant qu'on en lit un autre, c'est afficher
               deux verites a la fois. */
            poiData = [];
            majStrip(null, [], 0);
            montrerGuide('guideFichier', 'guideFichierSuite');
            if (!file) return;
            const reader = new FileReader();
            reader.onload = ev => {
                try {
                    /* Le référentiel des catégories vient de l'éditeur, dans SA langue.
                       Un échec laisse la table vide : les catégories seront alors
                       refusées et signalées, jamais posées à l'aveugle. */
                    try { chargerCategories(obtenirSdk()); } catch (e) { /* signalé à la ligne */ }
                    const wb = XLSX.read(new Uint8Array(ev.target.result), {type:'array'});
                    const all = [];
                    const warnings = [];

                    wb.SheetNames.filter(n => n !== 'Config').forEach(sheet => {
                        const sh = wb.Sheets[sheet];
                        // Les colonnes se lisent par leur EN-TÊTE, avec repli sur les
                        // positions A/B/C : l'ordre des colonnes cesse d'être un contrat
                        // tacite, et une colonne ajoutée à droite ne décale plus rien.
                        const rows = XLSX.utils.sheet_to_json(sh, {header:1, defval:''});
                        const plan = mapColumns(rows[0]);
                        if (!plan.usable) {
                            warnings.push(`[${sheet}] ${t('sheetHeaderErr')}`);
                            return;
                        }
                        if (plan.byPosition) warnings.push(`[${sheet}] ${t('sheetHeaderFallback')}`);
                        // Le numéro de ligne affiché est celui du tableur, même si la
                        // feuille ne commence pas en A1 : sans cela, l'anomalie renvoie
                        // à une ligne que personne ne retrouve.
                        const premiere = (XLSX.utils.decode_range(sh['!ref'] || 'A1:C1').s.r || 0) + 1;
                        const champsIdx = mapChamps(rows[0], plan.byPosition);
                        const sheetPerms = new Map();
                        rows.slice(1).forEach((cells, idx) => {
                            const rowNum = premiere + 1 + idx;
                            const valeur = (i) => {
                                const v = cells[i];
                                return v === undefined || v === null ? '' : v;
                            };
                            const r = {
                                perm: String(valeur(plan.columns.perm)).trim(),
                                name: valeur(plan.columns.name),
                                desc: valeur(plan.columns.desc)
                            };
                            if (!r.perm) return; // ligne vide ignorée silencieusement

                            // 1. URL syntaxiquement valide ?
                            let parsedUrl;
                            try { parsedUrl = new URL(r.perm); } catch {
                                warnings.push(`[${sheet}] ${t('rowLabel')} ${rowNum}: ${t('urlInvalid')}`);
                                return;
                            }

                            // 2. Domaine et chemin WME valides ?
                            const validHost = ['www.waze.com', 'waze.com', 'beta.waze.com'].includes(parsedUrl.hostname);
                            const validPath = parsedUrl.pathname.includes('/editor');
                            if (!validHost || !validPath) {
                                warnings.push(`[${sheet}] ${t('urlBadHost')} (${t('rowLabel')} ${rowNum})`);
                                return;
                            }

                            // 3. Paramètres obligatoires présents et cohérents ?
                            const params   = parsedUrl.searchParams;
                            const lat      = parseFloat(params.get('lat'));
                            const lon      = parseFloat(params.get('lon'));
                            const zoom     = params.get('zoomLevel');
                            const venues   = params.get('venues');

                            if (!params.get('env')) {
                                warnings.push(`[${sheet}] ${t('urlNoEnv')} (${t('rowLabel')} ${rowNum})`);
                                return;
                            }
                            if (isNaN(lat) || lat < -90 || lat > 90) {
                                warnings.push(`[${sheet}] ${t('urlBadLat')} (${t('rowLabel')} ${rowNum})`);
                                return;
                            }
                            if (isNaN(lon) || lon < -180 || lon > 180) {
                                warnings.push(`[${sheet}] ${t('urlBadLon')} (${t('rowLabel')} ${rowNum})`);
                                return;
                            }
                            if (!zoom || isNaN(parseInt(zoom))) {
                                warnings.push(`[${sheet}] ${t('urlBadZoom')} (${t('rowLabel')} ${rowNum})`);
                                return;
                            }
                            if (!venues) {
                                warnings.push(`[${sheet}] ${t('urlNoVenues')} (${t('rowLabel')} ${rowNum})`);
                                return;
                            }

                            // 4. Extraire l'ID du venue
                            const vid = getVenueIdFromPermalink(r.perm);
                            if (!vid) {
                                warnings.push(`[${sheet}] ${t('urlNoVid')} (${t('rowLabel')} ${rowNum})`);
                                return;
                            }

                            // 5. Nom vide ?
                            if (!r.name || !r.name.toString().trim()) {
                                warnings.push(`[${sheet}] ${t('nameEmpty')} (${t('rowLabel')} ${rowNum})`);
                                // on garde quand même la ligne
                            }

                            // 6. Doublon dans le même onglet ?
                            if (sheetPerms.has(vid)) {
                                const firstRow = sheetPerms.get(vid);
                                warnings.push(`[${sheet}] ${t('dupRow', firstRow, rowNum)}`);
                                return;
                            }
                            sheetPerms.set(vid, rowNum);

                            all.push({event:sheet, perm:r.perm, name:r.name, desc:r.desc,
                                      valeurs: lireValeurs(cells, champsIdx)});
                        });
                    });

                    if (!all.length) {
                        // Rien de valide — on affiche quand même le rapport d'anomalies
                        montrerGuide('guideFichier', 'guideFichierSuite');
                        showValidationReport(warnings.length ? warnings : [t('noPoisLoaded')]);
                        return;
                    }
                    poiData = all;
                    recordFileLoaded(file.name);
                    renderHistory();
                    /* ⭐ LE BANDEAU DE LA FENETRE PORTE TOUT : le nom du classeur, la
                       liste des onglets et le nombre de POI. Un seul endroit le dit,
                       donc il n'y a plus deux comptes a garder d'accord. */
                    ouvrirOverlay();
                    majStrip(file.name, Array.from(new Set(poiData.map((p) => p.event))), all.length);
                    showValidationReport(warnings);
                } catch (err) {
                    /* ⚠️ UNE LECTURE QUI ECHOUE SE DIT, ET SE DIT LA OU L'ON REGARDE.
                       Un message pose dans un panneau que la carte fait disparaitre
                       ne serait lu par personne. */
                    ouvrirOverlay();
                    showValidationReport(['✖ ' + err.message]);
                }
                fileInput.value = '';
            };
            reader.readAsArrayBuffer(file);
        });
    }

    /* ======================================================================
       L'APERCU — il vit dans la fenetre, plus dans une boite a lui.
       ====================================================================== */

    /** L'etat courant de l'apercu : ce que l'ecran montre et sur quoi on agit. */
    let _apercu = null;

    /**
     * LE BANDEAU D'ETAT — ce qui est charge, sous les yeux en permanence.
     *
     * ⭐ IL PORTE LE CHOIX DE L'ONGLET, parce que c'est la QUESTION qui suit le
     *    chargement : quel evenement pose-t-on ? La reponse commande tout le
     *    reste de l'ecran, elle ne se range pas dans un menu.
     */
    function majStrip(nomFichier, onglets, nbPoi) {
        const ov = construireOverlay();
        const strip = ov.querySelector('#peu-strip');
        strip.classList.toggle('peu-has-file', !!nomFichier);
        ov.querySelector('#peu-strip-texte').textContent = nomFichier || t('noFile');

        const sel = ov.querySelector('#peu-select-onglet');
        sel.innerHTML = '';
        (onglets || []).forEach((o) => {
            const opt = document.createElement('option');
            opt.value = o;
            opt.textContent = o;
            sel.appendChild(opt);
        });
        const aDesOnglets = !!(onglets && onglets.length);
        sel.hidden = !aDesOnglets;
        ov.querySelector('#peu-strip-sep1').hidden = !aDesOnglets;
        ov.querySelector('#peu-strip-sep2').hidden = !aDesOnglets;
        const compte = ov.querySelector('#peu-strip-compte');
        compte.hidden = !aDesOnglets;
        compte.textContent = t('poiCount', nbPoi || 0);

        majFab();

        /* ⭐⭐⭐⭐ LE PREMIER ONGLET SE CHARGE TOUT SEUL, parce que le menu l'affiche
           DEJA. Remplir un `<select>` par programme ne declenche pas `change` :
           l'ecran montrait « Hors Evenement » en tete du menu tout en demandant
           de choisir un onglet, et il fallait en prendre un autre puis revenir
           au premier pour que quelque chose se passe. Deux choses se
           contredisaient a l'ecran, et c'est celle qui ne fait rien qui gagnait.

           ⚠️ CE N'EST PAS UN RACCOURCI, C'EST LA SUITE DU GESTE : on a choisi un
              classeur pour le regarder. Le balayage qui suit est interruptible,
              et le menu reste la pour changer d'onglet. */
        if (aDesOnglets) ouvrirApercu(onglets[0]);
    }

    /** Le corps de la fenetre, vide. */
    function corpsFenetre() {
        const ov = construireOverlay();
        const corps = ov.querySelector('#peu-body');
        corps.innerHTML = '';

        return corps;
    }

    /**
     * LE GUIDAGE — il dit TOUJOURS le geste suivant.
     *
     * ⭐⭐⭐ UN ECRAN VIDE AVEC UN BOUTON GRISE N'EXPLIQUE RIEN. Quand rien n'est
     *    charge, la fenetre ne doit pas se contenter de ne rien montrer : elle
     *    doit dire par ou commencer.
     */
    function montrerGuide(cle, suite, extra) {
        const corps = corpsFenetre();
        /* ⭐ LA ZONE DE DEPOT N'APPARAIT QUE QUAND C'EST UN FICHIER QU'ON ATTEND.
           Proposee au moment de choisir un onglet, elle inviterait a un geste qui
           ne mene nulle part. */
        const attendUnFichier = cle === 'guideFichier';

        corps.innerHTML = '<div class="peu-guide"><span class="peu-guide-n">1</span>'
            + '<div><b>' + esc(t(cle)) + '</b>'
            + '<div class="peu-guide-suite">' + esc(t(suite)) + '</div></div></div>'
            + (attendUnFichier
                ? '<div class="peu-dropzone" id="peu-dropzone" title="' + esc(t('dropTitre')) + '">'
                    + esc(t('dropLigne1')) + '<br><span style="font-size:.833em">'
                    + esc(t('dropLigne2')) + '</span></div>'
                : '')
            + (extra || '');

        const zone = corps.querySelector('#peu-dropzone');
        if (zone) zone.addEventListener('click', () => { if (_peuFileInput) _peuFileInput.click(); });
        majPied();
    }

    /**
     * LE CLASSEUR SE DEPOSE SUR LA FENETRE, PAS SEULEMENT SUR LA ZONE.
     *
     * ⭐⭐⭐ VISER UNE ZONE DE 60 PIXELS AVEC UN FICHIER AU BOUT DU CURSEUR EST UN
     *    EXERCICE. Toute la fenetre accepte donc le depot — la zone dessinee dit
     *    OU l'on peut lacher, elle ne dit pas que c'est le seul endroit.
     *
     * ⚠️⚠️ ET IL PASSE PAR LE MEME CHAMP DE FICHIER. Le chemin de lecture valide
     *    vingt regles, tient un rapport d'anomalies et alimente l'historique :
     *    un second chemin pour un fichier depose divergerait du premier, et il le
     *    ferait en silence. On remplit donc le champ, et l'on declenche son
     *    evenement — c'est le meme code qui lit, quelle que soit la main.
     *
     * ⚠️ `preventDefault` SUR dragover ET SUR drop : sans le premier, le navigateur
     *    refuse le depot ; sans le second, il OUVRE le classeur a la place de la
     *    carte, et l'on perd sa session d'edition.
     */
    function brancherDepot(ov) {
        const surviens = (e) => {
            e.preventDefault();
            e.stopPropagation();
        };

        ['dragenter', 'dragover'].forEach((evt) => ov.addEventListener(evt, (e) => {
            surviens(e);
            const zone = ov.querySelector('#peu-dropzone');
            if (zone) zone.classList.add('peu-drop-hover');
        }));

        ['dragleave', 'dragend'].forEach((evt) => ov.addEventListener(evt, (e) => {
            surviens(e);
            /* ⚠️ On ne retire la marque QUE si le curseur a vraiment quitte la
               fenetre : `dragleave` se declenche aussi en passant d'un enfant a
               l'autre, et la zone clignoterait tout du long. */
            if (e.relatedTarget && ov.contains(e.relatedTarget)) return;
            const zone = ov.querySelector('#peu-dropzone');
            if (zone) zone.classList.remove('peu-drop-hover');
        }));

        ov.addEventListener('drop', (e) => {
            surviens(e);
            const zone = ov.querySelector('#peu-dropzone');
            if (zone) zone.classList.remove('peu-drop-hover');

            const fichier = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
            if (!fichier || !_peuFileInput) return;

            /* ⚠️ ON REFUSE ICI CE QUI N'EST PAS UN CLASSEUR, et on le DIT : laisser
               le lecteur s'en charger donnerait une erreur de format la ou une
               phrase suffit. */
            if (!/\.xlsx?$/i.test(fichier.name)) {
                ouvrirOverlay();
                showValidationReport([t('dropRefus', fichier.name)]);
                return;
            }

            const dt = new DataTransfer();
            dt.items.add(fichier);
            _peuFileInput.files = dt.files;
            _peuFileInput.dispatchEvent(new Event('change', { bubbles: true }));
        });
    }

    /**
     * LE PIED — il dit ce que le bouton va faire, et sur combien de lignes.
     *
     * ⚠️ ZERO EST UN RESULTAT : le bouton reste, desactive, en disant qu'il n'y
     *    a rien de coche. Un bouton qui disparait se lit comme une panne.
     */
    function majPied() {
        const ov = document.getElementById('peu-overlay');
        if (!ov) return;
        const btn = ov.querySelector('#peu-btn-appliquer');
        const exp = ov.querySelector('#peu-btn-export');
        const aide = ov.querySelector('#peu-footer-help');
        const cases = ov.querySelectorAll('#peu-body .peu-ligne .peu-checkbox:checked');
        const nb = cases.length;

        btn.textContent = libelleAppliquer(nb);
        btn.disabled = nb === 0;
        if (exp) exp.hidden = !_apercu || !_apercu.resultats;
        if (aide) aide.textContent = t(_apercu ? 'footerHelp' : 'footerHelpVide');
    }

    /**
     * LIT DANS LA FENETRE LES LIGNES QUE L'ON VA POSER.
     *
     * ⚠️⚠️ ON LIT LES CHAMPS ICI, UNE FOIS, ET L'ON PASSE DES CHAINES. La pose
     *    ne doit jamais dependre d'un element du DOM : un redessin entre la
     *    lecture et l'ecriture, et l'on pose ce qu'affichait l'ecran d'avant.
     */
    function lignesCochees() {
        const ov = document.getElementById('peu-overlay');
        if (!ov || !_apercu) return [];

        return [...ov.querySelectorAll('#peu-body .peu-ligne')].filter((tr) => {
            const cb = tr.querySelector('.peu-checkbox');

            return cb && cb.checked && !cb.disabled;
        }).map((tr) => {
            const p = _apercu.pois[Number(tr.dataset.idx)];

            return {
                vid: getVenueIdFromPermalink(p.perm),
                perm: p.perm,
                nom: tr.querySelector('[data-nom]').value,
                desc: tr.querySelector('[data-desc]').value,
                valeurs: p.valeurs,
            };
        });
    }

    /**
     * POSE UN LIEU — extrait tel quel de l'ancienne boucle, aux chaines pres.
     *
     * ⚠️⚠️ ON RELIT APRES AVOIR ECRIT, ET C'EST OBLIGATOIRE : le SDK n'eleve
     *    aucune erreur devant un champ qu'il ne connait pas, ni devant une
     *    valeur hors enumeration. Sans cette relecture, « applique » ne voudrait
     *    dire que « l'appel n'a pas plante ».
     *
     * ⚠️⚠️ UN CHAMP A POSER QUI N'EST JAMAIS ENVOYE EST UN MANQUE, PAS UN SUCCES.
     *    Si la liste se construit sans les valeurs, la mise a jour sort vide, le
     *    SDK n'est pas appele, et l'ecran annonce « applique avec succes » sur un
     *    lieu que personne n'a touche. Le compte ci-dessous le refuse.
     */
    async function poserUnLieu(item) {
        const UpdateObject = require('Waze/Action/UpdateObject');
        const venue = await centerAndLoad(item.perm, item.vid, 3000);
        if (!venue) {
            return { echec: true, resultat: { oldName: '', newName: item.nom, oldDesc: '', newDesc: item.desc, status: 'timeout' } };
        }

        const oldName = venue.attributes.name || '';
        const oldDesc = venue.attributes.description || '';
        /* ⚠️ Les noms alternatifs viennent du classeur s'il en porte, sinon on
           REPASSE ceux du lieu tels quels : ne jamais les perdre au passage. */
        const aPoser = (item.valeurs && item.valeurs.aPoser) || {};
        W.model.actionManager.add(new UpdateObject(venue, {
            id: venue.attributes.id,
            name: item.nom,
            description: item.desc,
            aliases: aPoser.aliases || venue.attributes.aliases || [],
        }));

        let poseSdk = null;
        const { maj, ignores } = construireMaj(aPoser);

        const attendus = Object.keys(aPoser).filter((c) => !CIBLES_HERITEES.includes(c));
        if (attendus.length && !Object.keys(maj).length) {
            poseSdk = { identiques: [], differents: attendus };
        }

        if (Object.keys(maj).length) {
            try {
                const sdk = obtenirSdk();
                sdk.DataModel.Venues.updateVenue(Object.assign({ venueId: item.vid }, maj));
                const relu = W.model.venues.getObjectById(item.vid);
                poseSdk = comparerAuLieu(relu ? relu.attributes : {}, aPoser);
            } catch (e) {
                poseSdk = { identiques: [], differents: Object.keys(maj), erreur: e.message };
            }
        }
        if (ignores.length) {
            poseSdk = poseSdk || { identiques: [], differents: [] };
            poseSdk.differents = poseSdk.differents.concat(ignores);
        }

        return {
            echec: false,
            resultat: {
                oldName: oldName, newName: item.nom,
                oldDesc: oldDesc, newDesc: item.desc,
                status: poseSdk && poseSdk.differents.length ? 'partial' : 'applied',
                poses: poseSdk ? poseSdk.identiques : [],
                manques: poseSdk ? poseSdk.differents : [],
            },
        };
    }

    /**
     * LA BARRE DE PROGRESSION, dans le corps de la fenetre.
     *
     * ⚠️ BLEU FRANC : c'est du TRAVAIL EN COURS, ni une alerte (orange) ni un
     *    resultat (vert). Les chiffres sont tabulaires, sans quoi ils dansent
     *    d'un rafraichissement a l'autre.
     */
    function poserProgression(corps, libelle, total, surAnnulation, enTete) {
        const div = document.createElement('div');
        div.className = 'peu-prog';
        div.innerHTML = '<div class="peu-prog-t">'
            + '<span class="peu-progress-label">' + esc(libelle) + '</span>'
            + '<span class="peu-prog-pct">0 %</span></div>'
            + '<div class="peu-progress-bar-bg"><i class="peu-progress-bar"></i></div>'
            + '<div class="peu-prog-b"><span class="peu-prog-d"></span>'
            + (surAnnulation ? '<button type="button" class="peu-btn peu-btn-neutral peu-btn-sm" data-annuler'
                + ' title="' + esc(t('cancelTitle')) + '">' + esc(t('cancelBtn')) + '</button>' : '')
            + '</div>';
        if (enTete) corps.prepend(div); else corps.appendChild(div);

        if (surAnnulation) div.querySelector('[data-annuler]').addEventListener('click', surAnnulation);

        return {
            avance(n, sur) {
                const t2 = sur || total;
                const p = t2 ? Math.round((n / t2) * 100) : 0;
                div.querySelector('.peu-progress-bar').style.width = p + '%';
                div.querySelector('.peu-prog-pct').textContent = p + ' %';
                div.querySelector('.peu-prog-d').textContent = n + ' / ' + t2;
            },
            libelle(txt) { div.querySelector('.peu-progress-label').textContent = txt; },
            retirer() { div.remove(); },
        };
    }

    /**
     * OUVRE L'APERCU D'UN ONGLET : precharge, cadre, puis montre le tableau.
     */
    async function ouvrirApercu(eventName) {
        const pois = poiData.filter((p) => p.event === eventName);
        if (!pois.length) { montrerGuide('guideOnglet', 'guideOngletSuite'); return; }

        /* ⚠️⚠️ LE CALQUE « LIEUX » DOIT ETRE ALLUME, SANS QUOI RIEN N'EXISTE. WME ne
           charge pas les lieux d'un calque eteint : le prechargement ne trouverait
           AUCUN POI et l'apercu annoncerait que tout est introuvable — un diagnostic
           faux, sur un fichier juste. On l'allume donc, et l'on attend que WME
           serve les lieux avant de balayer. */
        const calque = W.map.getLayersByName('venues')[0];
        if (calque && !calque.getVisibility()) {
            const bascule = document.querySelector('#layer-switcher-group_places');
            if (!bascule) { alert(t('layerOffMsg')); return; }
            bascule.click();
            await new Promise((r) => setTimeout(r, 1500));
        }

        ouvrirOverlay();
        const corps = corpsFenetre();
        const annule = { cancelled: false };
        const prog = poserProgression(corps, t('loadingPois', 0, pois.length), pois.length,
            () => { annule.cancelled = true; });

        const venueMap = await preloadVenues(pois, (n, total) => prog.avance(n, total), annule);
        prog.retirer();
        if (annule.cancelled) { montrerGuide('guideOnglet', 'guideOngletSuite'); return; }

        // ⭐ ON RESTE SUR LE PERIMETRE qu'on vient de parcourir.
        cadrerSurLesLieux(venueMap);

        _apercu = { eventName: eventName, pois: pois, venueMap: venueMap, resultats: null };
        /* ⚠️ LE MENU DIT CE QUI EST OUVERT. Ouvert par un autre chemin que lui,
           il afficherait autre chose que ce que le tableau montre. */
        const menu = document.getElementById('peu-select-onglet');
        if (menu && menu.value !== eventName) menu.value = eventName;
        rendreTableau();
    }

    /** Ce qu'il y a a dire d'une ligne, avant de la rendre. */
    function vueDeLaLigne(p, idx, venueMap) {
        const vid = getVenueIdFromPermalink(p.perm);
        const venue = venueMap[vid];
        const attributs = venue ? venue.attributes : null;
        const aPoser = (p.valeurs && p.valeurs.aPoser) || null;
        const verrou = venue ? getLockStatus(venue) : 'ok';

        return {
            idx: idx,
            nom: p.name, desc: p.desc,
            ancienNom: attributs ? (attributs.name || '') : '',
            ancienDesc: attributs ? (attributs.description || '') : '',
            nomChange: !!attributs && p.name !== (attributs.name || ''),
            descChange: !!attributs && p.desc !== (attributs.description || ''),
            charge: !!venue,
            verrou: verrou,
            niveau: venue && venue.attributes ? (venue.attributes.lockRank || 0) + 1 : 0,
            pertes: attributs && aPoser ? Object.keys(pertesDeLaPose(attributs, aPoser)).length : 0,
            champsDiff: champsQuiDifferent(attributs, aPoser),
        };
    }

    /** Rend le tableau complet dans le corps de la fenetre, et le branche. */
    function rendreTableau() {
        const corps = corpsFenetre();
        const vues = _apercu.pois.map((p, i) => vueDeLaLigne(p, i, _apercu.venueMap));

        const barre = document.createElement('div');
        barre.className = 'peu-toolbar';
        barre.innerHTML = '<input type="text" class="peu-search" data-filtre'
            + ' placeholder="' + esc(t('filterPlaceholder')) + '" title="' + esc(t('filterPlaceholder')) + '">'
            + '<button type="button" class="peu-btn peu-btn-neutral peu-btn-sm" data-ecarts'
            + ' title="' + esc(t('tooltipDiffOn')) + '">' + esc(t('btnDiffActive')) + '</button>'
            + '<span class="peu-search-count"></span>';
        corps.appendChild(barre);

        const table = document.createElement('table');
        table.className = 'peu-table';
        table.innerHTML = enteteApercu() + '<tbody>'
            + vues.map((v) => ligneApercu(v)).join('') + '</tbody>';
        corps.appendChild(table);

        /* Les champs du lot D2, sous chaque ligne — le rendu existant, inchange. */
        const tbody = table.querySelector('tbody');
        [...tbody.querySelectorAll('.peu-ligne')].forEach((tr, i) => {
            const p = _apercu.pois[i];
            if (!p.valeurs) return;
            const venue = _apercu.venueMap[getVenueIdFromPermalink(p.perm)];
            const bloc = rendreComplements(document, p.valeurs, t, venue ? venue.attributes : null);
            if (!bloc || !bloc.childNodes.length) return;
            const trc = document.createElement('tr');
            trc.className = 'peu-ligne peu-comp peu-row-' + (tr.className.match(/peu-row-(\w+)/) || [])[1];
            trc.innerHTML = '<td></td><td></td><td colspan="2"></td>';
            trc.lastElementChild.appendChild(bloc);
            tr.after(trc);
        });

        /* ⚠️ LE COCHAGE EST DECIDE PAR LA REGLE, PAS PAR LE RENDU : une ligne qui
           RETIRE quelque chose ne se coche jamais d'office. */
        [...tbody.querySelectorAll('.peu-ligne:not(.peu-comp)')].forEach((tr, i) => {
            const v = vues[i];
            const cb = tr.querySelector('.peu-checkbox');
            if (cb && !cb.disabled) {
                cb.checked = cocherDOffice(v.nomChange, v.descChange, v.champsDiff, v.pertes);
            }
        });

        brancherTableau(table, barre, vues);
        majPied();
    }

    /** Branche les gestes du tableau : cases, filtre, recentrage, edition. */
    function brancherTableau(table, barre, vues) {
        const maitre = table.querySelector('[data-maitre]');

        table.addEventListener('change', (e) => {
            if (e.target.classList.contains('peu-checkbox') && e.target !== maitre) majPied();
        });

        maitre.addEventListener('change', () => {
            table.querySelectorAll('tbody .peu-ligne:not(.peu-comp)').forEach((tr) => {
                if (tr.style.display === 'none') return;
                const cb = tr.querySelector('.peu-checkbox');
                if (cb && !cb.disabled) cb.checked = maitre.checked;
            });
            majPied();
        });

        table.addEventListener('click', (e) => {
            const cible = e.target.closest('[data-centrer]');
            if (!cible) return;
            const tr = cible.closest('.peu-ligne');
            const p = _apercu.pois[Number(tr.dataset.idx)];
            centerAndLoad(p.perm, getVenueIdFromPermalink(p.perm), 1500);
        });

        const filtre = barre.querySelector('[data-filtre]');
        const btnEcarts = barre.querySelector('[data-ecarts]');
        let ecartsSeuls = false;

        const appliquerFiltres = () => {
            const mot = filtre.value.trim().toLowerCase();
            let vus = 0;
            table.querySelectorAll('tbody .peu-ligne:not(.peu-comp)').forEach((tr, i) => {
                const v = vues[i];
                const texte = (tr.querySelector('[data-nom]').value + ' ' + tr.querySelector('[data-desc]').value).toLowerCase();
                const garde = (!mot || texte.includes(mot))
                    && (!ecartsSeuls || v.champsDiff > 0 || v.nomChange || v.descChange || v.pertes > 0);
                tr.style.display = garde ? '' : 'none';
                const comp = tr.nextElementSibling;
                if (comp && comp.classList.contains('peu-comp')) comp.style.display = garde ? '' : 'none';
                if (garde) vus++;
            });
            barre.querySelector('.peu-search-count').textContent = t('poiCount', vus);
        };

        filtre.addEventListener('input', appliquerFiltres);
        btnEcarts.addEventListener('click', () => {
            ecartsSeuls = !ecartsSeuls;
            btnEcarts.classList.toggle('peu-btn-primary', ecartsSeuls);
            btnEcarts.classList.toggle('peu-btn-neutral', !ecartsSeuls);
            btnEcarts.textContent = t(ecartsSeuls ? 'btnDiffAll' : 'btnDiffActive');
            btnEcarts.title = t(ecartsSeuls ? 'tooltipDiffOff' : 'tooltipDiffOn');
            appliquerFiltres();
        });

        appliquerFiltres();
    }

    /**
     * APPLIQUE LES LIGNES COCHEES.
     *
     * ⛔ LE SCRIPT N'ENREGISTRE JAMAIS. Il pose des modifications dans la pile de
     *    WME ; c'est l'editeur qui enregistre, apres avoir relu.
     */
    async function appliquerLignes() {
        const items = lignesCochees();
        if (!items.length) return;

        const ov = document.getElementById('peu-overlay');
        const btn = ov.querySelector('#peu-btn-appliquer');
        btn.disabled = true;

        const corps = ov.querySelector('#peu-body');
        /* ⚠️ EN TETE DU CORPS : la liste peut etre longue, et une barre posee en
           bas d une zone defilante travaille hors de vue. */
        const prog = poserProgression(corps, t('applying', 0, items.length), items.length, null, true);

        const resultats = [];
        const echecs = [];
        for (let i = 0; i < items.length; i++) {
            const r = await poserUnLieu(items[i]);
            resultats.push(r.resultat);
            if (r.echec) echecs.push(items[i]);
            prog.avance(i + 1);
        }
        prog.retirer();

        cadrerSurLesLieux(_apercu.venueMap);
        _apercu.resultats = resultats;
        montrerBilan(corps, resultats, echecs);
        majPied();
    }

    /**
     * LE BILAN — il dit ce qui est pose, ce qui manque, et que RIEN N'EST
     * ENREGISTRE.
     *
     * ⚠️ « Applique » ne veut pas dire « enregistre » : la confusion coute une
     *    session de travail perdue, et elle ne se voit qu'au rechargement.
     */
    function montrerBilan(corps, resultats, echecs) {
        const poses = resultats.filter((r) => r.status === 'applied').length;
        const partiels = resultats.filter((r) => r.status === 'partial').length;
        const div = document.createElement('div');
        div.className = 'peu-alert ' + (echecs.length || partiels ? 'peu-alert-warn' : 'peu-alert-ok');
        div.innerHTML = '<b>' + esc(t('successMsg', poses)) + '</b>'
            + (partiels ? '<br>' + esc(t('bilanPartiel', partiels)) : '')
            + (echecs.length ? '<br>' + esc(t('bilanEchec', echecs.length)) : '')
            + '<br>' + esc(t('bilanNonEnregistre', nbModifsEnAttente()));
        corps.prepend(div);
    }

    /** Ce que WME a en attente — le chiffre qui dit qu'il reste a enregistrer. */
    function nbModifsEnAttente() {
        try { return W.model.actionManager.getActions().length; } catch (e) { return 0; }
    }

    let _peuInited = false;
    function _peuInit() { if (_peuInited) return; _peuInited = true; initScript(); }

    /* ⚠️⚠️ `typeof` ET NON `W?.` : l’optional chaining protège d’un objet NUL,
       pas d’une variable JAMAIS DÉCLARÉE. Si le script s’exécute avant que WME
       ait posé son `W`, `W?.x` lève une ReferenceError — et elle survient AVANT
       que la moindre ligne d’interface soit construite. Le script meurt alors
       en entier, sans rien poser : ni feuille de style, ni bouton, ni onglet.
       C’est exactement le symptôme d’un script « qui n’a pas chargé ». */
    const wmePret = () => typeof W !== 'undefined' && W?.userscripts?.state?.isReady;

    if (wmePret()) {
        _peuInit();
    } else {
        document.addEventListener('wme-ready', _peuInit, {once:true});
        const fallback = setInterval(() => {
            if (wmePret()) { clearInterval(fallback); _peuInit(); }
        }, 500);
        setTimeout(() => clearInterval(fallback), 30000);
    }
})();
