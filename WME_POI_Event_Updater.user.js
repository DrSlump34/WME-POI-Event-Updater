// ==UserScript==
// @name         WME POI Event Updater
// @name:fr      WME POI Event Updater
// @namespace    http://tampermonkey.net/
// @version      0.51
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
    const HISTORY_MAX = 5;
    const GEOM_KEY = 'peu_overlay_geom';    // taille + position mémorisées de l'overlay
    // Icône de l'onglet : pin de localisation (= POI), détouré, affiché à la place du nom
    const TAB_ICON = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHZpZXdCb3g9JzAgMCAyNCAyNCcgd2lkdGg9JzI0JyBoZWlnaHQ9JzI0Jz48cGF0aCBmaWxsPScjMkM2RUQ1JyBkPSdNMTIgMkM4LjEzIDIgNSA1LjEzIDUgOWMwIDUuMjUgNyAxMyA3IDEzczctNy43NSA3LTEzYzAtMy44Ny0zLjEzLTctNy03eicvPjxjaXJjbGUgY3g9JzEyJyBjeT0nOScgcj0nMi42JyBmaWxsPScjZmZmZmZmJy8+PC9zdmc+';
    let poiData = [];
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
                footerHelp:'Décochez les lignes à exclure, éditez si besoin, puis cliquez ✔ pour appliquer.',
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
                footerHelp:'Uncheck rows to exclude, edit if needed, then click ✔ to apply.',
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
    box-shadow: var(--peu-shadow);
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
.peu-scroll { overflow: visible; }

.peu-section {
    font-size: .833em; font-weight: 700; text-transform: uppercase;
    letter-spacing: .07em; color: var(--peu-blue, #2C6ED5);
    border-bottom: 1px solid var(--peu-border, #dde3ea);
    margin: 0; padding: 9px 12px 5px;
    display: flex; align-items: center; gap: 6px; flex-wrap: wrap;
}

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
.peu-btn-success { background: var(--peu-green, #43a047);  color: #fff; }
.peu-btn-danger  { background: var(--peu-red, #e53935);    color: #fff; }
.peu-btn-neutral { background: var(--peu-border, #dde3ea); color: var(--peu-text, #2d3748); }
.peu-btn-primary:hover:not(:disabled) { background: var(--peu-blue-dk, #1a4fa0); color: #fff; }
.peu-btn-success:hover:not(:disabled),
.peu-btn-danger:hover:not(:disabled)  { filter: brightness(1.1); }
.peu-btn-neutral:hover:not(:disabled) { filter: brightness(.95); }
.peu-btn:disabled { opacity: .45; cursor: not-allowed; }
.peu-btn-sm { padding: .25em .75em; font-size: .833em; min-height: 24px; }
.peu-btn-full { width: 100%; }

/* Bouton discret de ligne (recentrage). Pas de fond, pas de bordure. */
.peu-btn-center, .peu-btn-retry {
    background: transparent; border: none; padding: 1px 3px; margin: 0;
    cursor: pointer; font-size: 1.25em; line-height: 1;
    height: auto; min-height: 0; transition: transform .1s;
}
.peu-btn-center:hover, .peu-btn-retry:hover { transform: scale(1.18); }
.peu-btn-center:active, .peu-btn-retry:active { transform: scale(.9); }

/* ----------------------------------------------------------------------
   CHAMPS
   ---------------------------------------------------------------------- */
.peu-input, .peu-textarea, .peu-search, .peu-select {
    width: 100%; padding: .25em .45em;
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
.peu-search-clear {
    background: none; border: none; cursor: pointer; color: var(--peu-grey, #9e9e9e);
    font-size: 1.1em; padding: 0 2px; line-height: 1; display: none;
    height: auto; min-height: 0;
}
.peu-search-clear:hover { color: var(--peu-red, #e53935); }
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
.peu-lock-badge, .peu-diff-dot, .peu-unloaded-icon { display: none; }

/* Les champs du lot D2, sous la ligne du lieu. */
.peu-comp > td { background: rgba(0,0,0,.015); padding-top: 0; }
.peu-comp-titre {
    font-size: .833em; text-transform: uppercase; letter-spacing: .05em;
    color: var(--peu-text2, #566372); font-weight: 700; margin-bottom: 3px;
}
.peu-comp-liste { display: flex; flex-wrap: wrap; gap: 3px 14px; font-size: .909em; color: var(--peu-text2, #566372); }
.peu-comp-liste b { color: var(--peu-text, #2d3748); font-weight: 600; }
.peu-pastille, .peu-pastille-titre, .peu-plus { font-size: .909em; }
.peu-plus { color: var(--peu-text2, #566372); }

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
.peu-footer-error {
    background: #fff0f0; border: 1px solid var(--peu-red, #e53935); color: #8a1c14;
    border-radius: var(--peu-radius, 8px); padding: 6px 9px; margin: 8px 12px;
    font-size: .833em; line-height: 1.45;
}
.peu-footer-error .peu-error-title { font-weight: 700; display: block; margin-bottom: 3px; }
.peu-footer-error ul { margin: 3px 0 0; padding-inline-start: 18px; }
.peu-footer-error ul li { margin-bottom: 2px; }

/* ----------------------------------------------------------------------
   BANDEAUX — trois familles, et leur sens ne se melange pas :
   bleu = une information, vert = un resultat, orange = UN GESTE A FAIRE.
   ---------------------------------------------------------------------- */
.peu-alert  { border-radius: var(--peu-radius, 8px); padding: 7px 10px; margin: 8px 12px;
              font-size: .833em; line-height: 1.5; }
.peu-alert-info { background: #e3f2fd; border: 1px solid #90caf9; color: #0d47a1; }
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
.peu-hist-meta.peu-jamais { color: var(--peu-orange, #f57c00); font-weight: 600; }
.peu-drag-hint { font-size: .833em; color: var(--peu-text2, #566372); }

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
                Math.max(W.map.getZoom(), 17)
            );

            return true;
        }

        const etendue = new OpenLayers.Bounds(u.left, u.bottom, u.right, u.top);
        if (typeof W.map.zoomToExtent === 'function') {
            W.map.zoomToExtent(etendue);
            if (W.map.getZoom() < 12) W.map.setCenter(etendue.getCenterLonLat(), 12);

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
            +   '<span class="peu-strip-info" id="peu-strip-texte">' + esc(t('noFile')) + '</span>'
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
        tabIcon.style.cssText = 'width:18px;height:18px;display:block;';
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

        const container = document.createElement('div');
        container.className = 'peu-container';
        container.style.cssText = 'padding:10px;font-family:Segoe UI,Arial,sans-serif;font-size:12px;';

        const title = document.createElement('h3');
        title.textContent = t('panelTitle');
        title.style.cssText = 'margin:0 0 10px;font-size:14px;color:#2C6ED5;';
        container.appendChild(title);

        const fileInput = document.createElement('input');
        fileInput.type = 'file'; fileInput.accept = '.xlsx,.xls'; fileInput.style.display = 'none';
        container.appendChild(fileInput);

        const btnChoose = document.createElement('button');
        btnChoose.textContent = t('chooseFile');
        btnChoose.style.cssText = 'background:#2C6ED5;color:#fff;border:none;border-radius:5px;padding:5px 10px;font-size:12px;cursor:pointer;margin-bottom:6px;';
        btnChoose.onclick = () => fileInput.click();
        container.appendChild(btnChoose);

        const status = document.createElement('div');
        status.textContent = t('noFile');
        status.style.cssText = 'margin:6px 0;font-size:11px;color:#888;';
        container.appendChild(status);

        // Repli : si la librairie XLSX n'a pas pu se charger (CDN bloqué / hors ligne),
        // on informe clairement au lieu de laisser le script planter en silence.
        if (typeof XLSX === 'undefined') {
            status.textContent = t('xlsxMissing');
            status.style.color = '#c0392b';
            btnChoose.disabled = true;
            btnChoose.style.opacity = '0.5';
            btnChoose.style.cursor = 'not-allowed';
        }

        const select = document.createElement('select');
        select.style.cssText = 'display:none;width:100%;margin:6px 0;font-size:12px;border:1px solid #c5d3e8;border-radius:5px;padding:4px 6px;';
        container.appendChild(select);

        const btnShow = document.createElement('button');
        btnShow.textContent = t('showBtn');
        btnShow.style.cssText = 'background:#2C6ED5;color:#fff;border:none;border-radius:5px;padding:5px 10px;font-size:12px;cursor:not-allowed;margin-top:2px;width:100%;opacity:0.5;transition:opacity 0.15s;';
        btnShow.disabled = true; btnShow.style.opacity = '0.5'; btnShow.style.cursor = 'not-allowed';
        btnShow.onclick = async () => {
            // Vérifier que le calque Lieux est actif — si non, l'activer automatiquement
            const vLayer = W.map.getLayersByName('venues')[0];
            if (vLayer && !vLayer.getVisibility()) {
                const toggle = document.querySelector('#layer-switcher-group_places');
                if (toggle) {
                    toggle.click();
                    // Attendre que WME charge les venues
                    await new Promise(r => setTimeout(r, 1500));
                } else {
                    // Fallback si le toggle n'est pas trouvé
                    alert(t('layerOffMsg'));
                    return;
                }
            }
            showOverlay(select.value);
        };
        container.appendChild(btnShow);

        // Zone rapport de validation (sous le bouton Afficher)
        const validationReport = document.createElement('div');
        validationReport.style.cssText = 'margin-top:8px;font-size:10.5px;display:none;';
        container.appendChild(validationReport);

        // Zone historique
        const historyDiv = document.createElement('div');
        historyDiv.style.cssText = 'margin-top:10px;font-size:10.5px;';
        container.appendChild(historyDiv);

        function renderHistory() {
            historyDiv.innerHTML = '';
            const history = getHistory();
            if (!history.length) return;

            // En-tête avec bouton RAZ
            const hheader = document.createElement('div');
            hheader.style.cssText = 'display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;';
            const htitle = document.createElement('div');
            htitle.style.cssText = 'color:#888;font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:0.3px;';
            htitle.textContent = t('historyTitle');
            const btnRaz = document.createElement('button');
            btnRaz.textContent = '🗑';
            btnRaz.title = t('clearHistoryTitle');
            btnRaz.style.cssText = 'background:none;border:none;cursor:pointer;font-size:12px;color:#aaa;padding:0;line-height:1;';
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
                row.style.cssText = 'margin-bottom:5px;padding:4px 6px;background:#f0f4fb;border-radius:4px;border-left:3px solid #2C6ED5;';
                const name = document.createElement('div');
                name.style.cssText = 'font-weight:600;color:#2C6ED5;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;';
                name.textContent = h.name; name.title = h.name;
                const loaded = document.createElement('div');
                loaded.style.cssText = 'color:#888;margin-top:1px;';
                loaded.textContent = `${t('histLoaded')} ${formatDateTime(h.loaded)}`;
                const applied = document.createElement('div');
                applied.style.cssText = 'color:#888;margin-top:1px;';
                applied.textContent = h.applied ? `${t('histApplied')} ${formatDateTime(h.applied)}` : t('histNeverApplied');
                row.appendChild(name); row.appendChild(loaded); row.appendChild(applied);
                historyDiv.appendChild(row);
            });
        }

        tabPane.appendChild(container);
        renderHistory();

        function showValidationReport(warnings) {
            if (!warnings.length) { validationReport.style.display = 'none'; return; }
            validationReport.style.display = 'block';
            validationReport.innerHTML = '';
            const title = document.createElement('div');
            title.style.cssText = 'color:#e67e22;font-weight:700;margin-bottom:4px;';
            title.textContent = t('anomalies', warnings.length) + ' :';
            validationReport.appendChild(title);
            const ul = document.createElement('ul');
            ul.style.cssText = 'margin:0 0 0 14px;padding:0;color:#555;';
            warnings.forEach(w => {
                const li = document.createElement('li'); li.style.marginBottom = '2px';
                li.textContent = w; ul.appendChild(li);
            });
            validationReport.appendChild(ul);
        }

        fileInput.addEventListener('change', e => {
            const file = e.target.files[0];
            poiData = []; select.style.display = 'none'; btnShow.disabled = true; btnShow.style.opacity = '0.5'; btnShow.style.cursor = 'not-allowed';
            status.textContent = t('noFile'); status.style.color = '#888';
            validationReport.style.display = 'none';
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
                        status.textContent = t('noPoisLoaded');
                        status.style.color = '#c0392b';
                        showValidationReport(warnings);
                        return;
                    }
                    poiData = all;
                    select.innerHTML = '';
                    Array.from(new Set(poiData.map(p => p.event))).forEach(evt => {
                        const o = document.createElement('option'); o.value = evt; o.textContent = evt; select.appendChild(o);
                    });
                    select.style.display = 'block'; btnShow.disabled = false; btnShow.style.opacity = '1'; btnShow.style.cursor = 'pointer';
                    const warnTxt = warnings.length ? ` — ${warnings.length} ⚠️` : '';
                    status.textContent = t('poisLoaded', all.length) + warnTxt;
                    status.style.color = warnings.length ? '#e67e22' : '#27ae60';
                    recordFileLoaded(file.name);
                    renderHistory();
                    showValidationReport(warnings);
                } catch(err) {
                    status.textContent = '✖ ' + err.message; status.style.color = '#c0392b';
                }
                fileInput.value = '';
            };
            reader.readAsArrayBuffer(file);
        });
    }

    async function showOverlay(eventName) {
        const pois = poiData.filter(p => p.event === eventName);
        if (!pois.length) return alert(t('noPoi'));

        /* ⚠️⚠️ UN SEUL APERÇU À LA FOIS. Rien ne retirait le précédent : charger
           un second fichier sans fermer le premier empilait deux tableaux, dont
           celui du dessous restait atteignable au clavier — et l'on pouvait
           appliquer depuis un aperçu qui ne décrivait plus le fichier chargé. */
        document.querySelectorAll('.peu-overlay').forEach((o) => o.remove());

        // Où l'on était avant le balayage — utile seulement si l'on ANNULE.
        const savedCenter = W.map.getCenter();
        const savedZoom   = W.map.getZoom();

        // --- Phase 1 : écran de chargement dans le panneau latéral ---
        const loadingDiv = document.createElement('div');
        loadingDiv.style.cssText = 'margin-top:10px;';

        const loadLabel = document.createElement('div');
        loadLabel.className = 'peu-progress-label';
        loadLabel.textContent = t('loadingPois', 0, pois.length);

        const barBg = document.createElement('div'); barBg.className = 'peu-progress-bar-bg';
        const bar   = document.createElement('div'); bar.className   = 'peu-progress-bar';
        barBg.appendChild(bar);
        loadingDiv.appendChild(loadLabel);
        loadingDiv.appendChild(barBg);

        // Bouton Annuler le préchargement
        const cancelRef = { cancelled: false };
        const btnCancel = document.createElement('button');
        btnCancel.textContent = t('cancelBtn');
        btnCancel.style.cssText = 'margin-top:6px;width:100%;background:#888;color:#fff;border:none;border-radius:5px;padding:4px 10px;font-size:11px;cursor:pointer;';
        btnCancel.onclick = () => { cancelRef.cancelled = true; btnCancel.disabled = true; btnCancel.style.opacity = '0.5'; };
        loadingDiv.appendChild(btnCancel);

        const container = document.querySelector('.peu-container');
        container && container.appendChild(loadingDiv);

        // Preload avec progression (annulable)
        const venueMap = await preloadVenues(pois, (loaded, total) => {
            bar.style.width = Math.round(loaded / total * 100) + '%';
            loadLabel.textContent = t('loadingPois', loaded, total);
        }, cancelRef);

        /* Si annulé : on revient d'où l'on vient. Annuler, c'est dire « pas
           ça » — y compris le déplacement qu'on vient de subir. */
        if (cancelRef.cancelled) {
            W.map.setCenter(savedCenter, savedZoom);
            loadingDiv.remove();
            return;
        }

        // ⭐ ON RESTE SUR LE PÉRIMÈTRE qu'on vient de parcourir.
        cadrerSurLesLieux(venueMap);
        loadingDiv.remove();

        // --- Phase 2 : overlay tableau ---
        const overlay = document.createElement('div'); overlay.className = 'peu-overlay';
        const box = document.createElement('div'); box.className = 'peu-box';

        // Header
        const header = document.createElement('div'); header.className = 'peu-header';
        const headerLeft = document.createElement('div'); headerLeft.className = 'peu-header-left';
        const headerTitle = document.createElement('span');
        headerTitle.textContent = `${eventName} — ${t('poiCount', pois.length)}`;
        const dragHint = document.createElement('span'); dragHint.className = 'peu-drag-hint';
        dragHint.textContent = t('draggable');
        headerLeft.appendChild(headerTitle); headerLeft.appendChild(dragHint);

        const headerBtns = document.createElement('div'); headerBtns.className = 'peu-header-btns';
        const btnMin = document.createElement('button'); btnMin.className = 'peu-btn-icon minimize'; btnMin.title = t('btnMinimize'); btnMin.textContent = '▾';
        let minimized = false;
        btnMin.onclick = () => {
            minimized = !minimized;
            box.classList.toggle('minimized', minimized);
            btnMin.textContent = minimized ? '▴' : '▾';
            btnMin.title = minimized ? t('btnRestore') : t('btnReduce');
        };

        // Bouton diff only
        const btnDiff = document.createElement('button'); btnDiff.className = 'peu-btn-icon diffonly active'; btnDiff.title = t('tooltipDiffOff');
        let diffOnly = true;
        btnDiff.textContent = t('btnDiffActive');
        btnDiff.onclick = () => {
            diffOnly = !diffOnly;
            btnDiff.classList.toggle('active', diffOnly);
            btnDiff.title = diffOnly ? t('tooltipDiffOff') : t('tooltipDiffOn');
            applyFilters();
            updateMasterCb();
        };
        const btnV = document.createElement('button'); btnV.className = 'peu-btn-icon apply'; btnV.title = t('btnApply'); btnV.textContent = '✔'; btnV.onclick = () => applyChanges();
        const btnX = document.createElement('button'); btnX.className = 'peu-btn-icon cancel'; btnX.title = t('btnClose'); btnX.textContent = '✕'; btnX.onclick = () => overlay.remove();
        headerBtns.appendChild(btnMin); headerBtns.appendChild(btnDiff); headerBtns.appendChild(btnV); headerBtns.appendChild(btnX);
        // a11y : pour ces boutons purement iconographiques, le nom accessible = l'infobulle
        headerBtns.querySelectorAll('button').forEach(b => b.title && b.setAttribute('aria-label', b.title));
        header.appendChild(headerLeft); header.appendChild(headerBtns);
        box.appendChild(header);

        // Tableau
        // Toolbar de recherche
        const toolbar = document.createElement('div'); toolbar.className = 'peu-toolbar';
        const searchInput = document.createElement('input'); searchInput.className = 'peu-search'; searchInput.placeholder = t('filterPlaceholder'); searchInput.type = 'text';
        const searchClear = document.createElement('button'); searchClear.className = 'peu-search-clear'; searchClear.textContent = '✕'; searchClear.title = t('searchClearTitle');
        const searchCount = document.createElement('span'); searchCount.className = 'peu-search-count';
        toolbar.appendChild(searchInput); toolbar.appendChild(searchClear); toolbar.appendChild(searchCount);
        box.appendChild(toolbar);

        // Masquer toolbar quand minimisé
        box.classList.contains('minimized') && (toolbar.style.display = 'none');

        let searchTerm = '';
        searchInput.addEventListener('input', () => {
            searchTerm = searchInput.value.trim().toLowerCase();
            searchClear.style.display = searchTerm ? 'block' : 'none';
            applyFilters();
        });
        searchClear.addEventListener('click', () => {
            searchInput.value = ''; searchTerm = '';
            searchClear.style.display = 'none';
            searchInput.focus();
            applyFilters();
        });

        const scroll = document.createElement('div'); scroll.className = 'peu-scroll';
        const table = document.createElement('table'); table.className = 'peu-table';
        const cg = document.createElement('colgroup');
        [null,null,null,null].forEach(() => cg.appendChild(document.createElement('col')));
        table.appendChild(cg);
        const thead = table.createTHead(); const trh = thead.insertRow();

        // Colonnes : [label, sortable, extractFn]
        // extractFn sera définie après construction du tbody (accès à tr._inputs)
        const colDefs = [
            { label: t('colSearch'), sortable: false },
            { label: t('colName'),   sortable: true,  key: 'newName' },
            { label: t('colDesc'),   sortable: true,  key: 'newDesc' },
            { label: null,           sortable: false, master: true },
        ];

        const ths = colDefs.map(def => {
            const th = document.createElement('th');
            if (def.sortable) {
                th.className = 'sortable';
                const labelSpan = document.createElement('span'); labelSpan.textContent = def.label;
                const sortIcon  = document.createElement('i'); sortIcon.className = 'peu-sort-icon'; sortIcon.textContent = '▲';
                th.appendChild(labelSpan); th.appendChild(sortIcon);
                th.dataset.sortKey = def.key;
                th.dataset.sortDir = 'none';
            } else if (def.master) {
                // Case à cocher maître
                const masterCb = document.createElement('input'); masterCb.type = 'checkbox';
                masterCb.className = 'peu-checkbox'; masterCb.checked = true;
                masterCb.title = t('masterCbTitle');
                masterCb.addEventListener('change', () => {
                    // N'agit que sur les lignes visibles et non désactivées
                    tbody.querySelectorAll('tr').forEach(tr => {
                        if (tr.style.display === 'none') return;
                        const {cb} = tr._inputs;
                        if (!cb.disabled) { cb.checked = masterCb.checked; tr._cbUserSet = true; }
                    });
                });
                th.appendChild(masterCb);
                // Mettre à jour la case maître quand une case individuelle change
                th._masterCb = masterCb;
            } else {
                th.textContent = def.label;
            }
            trh.appendChild(th);
            return th;
        });

        // Référence à la case maître pour mise à jour
        const masterTh = ths.find(t => t._masterCb);
        function updateMasterCb() {
            if (!masterTh) return;
            const visibleCbs = Array.from(tbody.querySelectorAll('tr'))
                .filter(tr => tr.style.display !== 'none')
                .map(tr => tr._inputs?.cb)
                .filter(cb => cb && !cb.disabled);
            const allChecked  = visibleCbs.every(cb => cb.checked);
            const noneChecked = visibleCbs.every(cb => !cb.checked);
            masterTh._masterCb.checked = allChecked;
            masterTh._masterCb.indeterminate = !allChecked && !noneChecked;
        }

        // Etat du tri courant
        let currentSortTh = null;

        function sortTable(th) {
            const key = th.dataset.sortKey;
            const dir = th.dataset.sortDir;
            // Cycle : none/desc → asc, asc → desc, desc → none (ordre original)
            const nextDir = dir === 'asc' ? 'desc' : dir === 'desc' ? 'none' : 'asc';

            // Reset tous les th
            ths.forEach(t => {
                if (!t.dataset.sortKey) return;
                t.dataset.sortDir = 'none';
                t.classList.remove('sort-asc', 'sort-desc');
                t.querySelector('.peu-sort-icon').textContent = '▲';
            });

            if (nextDir === 'none') {
                // Restaurer l'ordre original
                const rows = Array.from(tbody.querySelectorAll('tr'));
                rows.sort((a, b) => parseInt(a.dataset.origIdx) - parseInt(b.dataset.origIdx));
                rows.forEach(r => tbody.appendChild(r));
                currentSortTh = null;
                return;
            }

            th.dataset.sortDir = nextDir;
            th.classList.add(nextDir === 'asc' ? 'sort-asc' : 'sort-desc');
            th.querySelector('.peu-sort-icon').textContent = nextDir === 'asc' ? '▲' : '▼';
            currentSortTh = th;

            const rows = Array.from(tbody.querySelectorAll('tr'));
            rows.sort((a, b) => {
                const va = (a.dataset[key] || '').toLowerCase();
                const vb = (b.dataset[key] || '').toLowerCase();
                return nextDir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
            });
            rows.forEach(r => tbody.appendChild(r));
        }

        const tbody = table.createTBody();
        let cntSae = 0, cntHard = 0;

        pois.forEach((p, idx) => {
            const tr = tbody.insertRow();
            tr.dataset.origIdx = idx; // pour restaurer l'ordre original
            const vid = getVenueIdFromPermalink(p.perm);
            const venue = venueMap[vid];
            const oldName = venue?.attributes?.name || '';
            const oldDesc = venue?.attributes?.description || '';
            const venueLoaded = !!venue; // false si préchargement échoué

            /* Ce que le lieu porte déjà, et combien de champs du lot D2 en diffèrent.
               ⚠️ Un lieu non chargé ne se compare à rien : on compte alors TOUS les
                  champs comme à appliquer, plutôt que de conclure « rien à faire »
                  d'une absence de mesure. */
            const attributsDuLieu = venueLoaded ? venue.attributes : null;
            const pertesDuLieu = p.valeurs && attributsDuLieu
                ? Object.keys(pertesDeLaPose(attributsDuLieu, p.valeurs.aPoser)).length : 0;
            const champsDiff = champsQuiDifferent(attributsDuLieu, p.valeurs && p.valeurs.aPoser);

            // Données de tri/filtre stockées sur la ligne (mises à jour si l'utilisateur édite)
            tr.dataset.oldName = oldName;
            tr.dataset.oldDesc = oldDesc;
            tr.dataset.newName = p.name;
            tr.dataset.newDesc = p.desc;
            tr._cbUserSet = false; // passe à true dès que l'utilisateur (dé)coche lui-même

            const lockStatus = getLockStatus(venue);
            const lockRank = venue?.attributes?.lockRank ?? 0;
            const userRank = W?.loginManager?.user?.attributes?.rank ?? 0;

            if (!venueLoaded) {
                tr.classList.add('peu-row-unloaded');
            } else if (lockStatus === 'sae')  { cntSae++;  tr.classList.add('peu-row-sae'); }
            if (lockStatus === 'hard') { cntHard++; tr.classList.add('peu-row-hard'); }

            // Met à jour l'indicateur diff (à la construction + à chaque frappe)
            function updateDiff() {
                const nameChanged = tr.dataset.newName !== tr.dataset.oldName;
                const descChanged = tr.dataset.newDesc !== tr.dataset.oldDesc;
                /* ⚠️ LES CHAMPS DU LOT D2 COMPTENT AUSSI : sans eux, l'aperçu
                   annonce « Aucune modification » sur un lieu qui a des champs à
                   poser, et l'on ferme la fenêtre en confiance. */
                const isDiff = nameChanged || descChanged || champsDiff > 0;
                tr.dataset.hasDiff = isDiff ? 'true' : 'false';
                // Ne pas écraser hard/sae
                if (lockStatus === 'ok') tr.classList.toggle('peu-row-diff', isDiff);
                diffDot.style.display = isDiff ? 'block' : 'none';
                // Barrer l'ancienne valeur uniquement si elle change réellement
                if (oldNameEl) oldNameEl.classList.toggle('changed', nameChanged);
                if (oldDescEl) oldDescEl.classList.toggle('changed', descChanged);
                /* ⚠️ LA RÈGLE DU COCHAGE VIT DANS `cocherDOffice`, éprouvée par
                   `tools/banc-cochage.mjs` : elle n'existait que dans ce DOM, où
                   rien ne pouvait la mesurer. */
                if (!cb.disabled && !tr._cbUserSet) {
                    cb.checked = cocherDOffice(nameChanged, descChanged, champsDiff, pertesDuLieu);
                }
            }

            // ── Colonne 🎯 (recentrage) + indicateurs (diff / non chargé / lock) ──
            const td0 = tr.insertCell(); td0.className = 'center';
            const b0 = document.createElement('button'); b0.className = 'peu-btn-center'; b0.textContent = '🎯';
            b0.title = t('locateTitle'); b0.setAttribute('aria-label', t('locateTitle'));
            b0.onclick = () => {
                const v = W.model.venues.getObjectById(vid);
                const bds = v?.getOLGeometry()?.getBounds();
                if (bds) {
                    const center = bds.getCenterLonLat();
                    if (W.map.getZoom() < 17) W.map.setCenter(center, 17);
                    else W.map.setCenter(center);
                    return;
                }
                // Fallback : venue plus en mémoire → recentrer via le permalink
                const coords = parseLatLon(p.perm);
                if (coords) {
                    const ll = new OpenLayers.LonLat(coords.lon, coords.lat).transform(
                        new OpenLayers.Projection('EPSG:4326'),
                        W.map.getProjectionObject()
                    );
                    W.map.setCenter(ll, 17);
                }
            };
            td0.appendChild(b0);

            const diffDot = document.createElement('span'); diffDot.className = 'peu-diff-dot';
            diffDot.title = t('diffTitle'); diffDot.style.display = 'none';
            td0.appendChild(diffDot);

            if (!venueLoaded) {
                const unloadedIcon = document.createElement('span');
                unloadedIcon.className = 'peu-unloaded-icon';
                unloadedIcon.textContent = t('unloadedLabel');
                unloadedIcon.title = t('unloadedTitle');
                td0.appendChild(unloadedIcon);
            }
            if (lockStatus !== 'ok') {
                const lockIcon = document.createElement('div');
                lockIcon.className = lockStatus === 'hard' ? 'peu-lock-hard' : 'peu-lock-sae';
                lockIcon.textContent = lockStatus === 'hard' ? '🔒' : '⚠️';
                lockIcon.title = lockStatus === 'hard'
                    ? t('lockHardTitle')
                    : t('lockSaeTitle', lockRank, userRank);
                td0.appendChild(lockIcon);
            }

            // La ligne « ancienne valeur » est réservée dans les DEUX colonnes dès que
            // l'une des deux a une valeur, afin que les champs restent alignés.
            const hasOld = !!oldName || !!oldDesc;
            const makeOldEl = (val) => {
                const el = document.createElement('div');
                el.className = 'peu-cell-old';
                el.textContent = val;
                if (val) el.title = val; // texte complet en infobulle (affichage tronqué)
                return el;
            };

            // ── Colonne NOM (ancien au-dessus, barré si modifié, + champ éditable) ──
            const tdName = tr.insertCell();
            let oldNameEl = null;
            if (hasOld) { oldNameEl = makeOldEl(oldName); tdName.appendChild(oldNameEl); }
            const inpName = document.createElement('input'); inpName.className = 'peu-input'; inpName.value = p.name;
            if (lockStatus === 'hard') inpName.disabled = true;
            inpName.addEventListener('input', () => { tr.dataset.newName = inpName.value; updateDiff(); applyFilters(); });
            tdName.appendChild(inpName);

            // ── Colonne DESCRIPTION (ancienne au-dessus, barrée si modifiée, + champ) ──
            const tdDesc = tr.insertCell();
            let oldDescEl = null;
            if (hasOld) { oldDescEl = makeOldEl(oldDesc); tdDesc.appendChild(oldDescEl); }
            const txtArea = document.createElement('textarea'); txtArea.className = 'peu-textarea'; txtArea.value = p.desc;
            if (lockStatus === 'hard') txtArea.disabled = true;
            txtArea.addEventListener('input', () => { tr.dataset.newDesc = txtArea.value; updateDiff(); applyFilters(); });
            tdDesc.appendChild(txtArea);

            /* Les champs du lot D2, s'il y en a : dans LA MÊME cellule, pour survivre au tri.
               ⚠️ Le lieu est passé au rendu pour qu'il ne montre que ce qui CHANGE. */
            if (p.valeurs) {
                const complements = rendreComplements(document, p.valeurs, t, attributsDuLieu);
                if (complements) tdDesc.appendChild(complements);
            }

            // ── Colonne case à cocher ──
            const tdCB = tr.insertCell(); tdCB.className = 'center';
            const cb = document.createElement('input'); cb.type = 'checkbox'; cb.className = 'peu-checkbox';
            cb.checked = false;
            if (lockStatus === 'hard' || !venueLoaded) cb.disabled = true;
            cb.addEventListener('change', () => { tr._cbUserSet = true; updateMasterCb(); });
            tdCB.appendChild(cb);

            tr._inputs = {vid, inpName, txtArea, cb, valeurs: p.valeurs};
            tr.dataset.champsDiff = String(champsDiff);

            // Initialiser l'indicateur diff au chargement
            updateDiff();
        });

        // Message "rien à modifier" — inséré dans le scroll, affiché si besoin
        const emptyMsg = document.createElement('div');
        emptyMsg.style.cssText = 'padding:24px;text-align:center;color:#888;font-size:12px;';
        emptyMsg.innerHTML = t('emptyMsg') + '<br><span style="color:#2C6ED5;cursor:pointer;text-decoration:underline;" id="peu-show-all">' + t('emptyShowAll') + '</span>';
        scroll.appendChild(emptyMsg);
        emptyMsg.style.display = 'none';
        emptyMsg.querySelector('#peu-show-all').addEventListener('click', () => {
            diffOnly = false;
            btnDiff.classList.remove('active');
            btnDiff.textContent = t('btnDiffAll');
            btnDiff.title = t('tooltipDiffOff');
            applyFilters();
        });

        // Filtre combiné : diff only + recherche par nom
        function applyFilters() {
            let visible = 0;
            let totalDiff = 0;
            tbody.querySelectorAll('tr').forEach(tr => {
                if (tr.dataset.hasDiff === 'true') totalDiff++;
                const hasDiff    = tr.dataset.hasDiff === 'true';
                const isUnloaded = tr.classList.contains('peu-row-unloaded');
                const oldName    = (tr.dataset.oldName || '').toLowerCase();
                const newName    = (tr.dataset.newName || '').toLowerCase();
                const matchDiff  = !diffOnly || hasDiff || isUnloaded;
                const matchSearch = !searchTerm || oldName.includes(searchTerm) || newName.includes(searchTerm);
                const show = matchDiff && matchSearch;
                tr.style.display = show ? '' : 'none';
                if (show) visible++;
            });
            const total = pois.length;

            // Message vide si mode diff et aucun diff
            const noChanges = diffOnly && totalDiff === 0;
            emptyMsg.style.display = noChanges ? 'block' : 'none';
            table.style.display    = noChanges ? 'none'  : '';

            // Label bouton Diff adaptatif
            if (diffOnly) {
                btnDiff.classList.add('active');
                btnDiff.textContent = totalDiff === 0 ? t('btnUpToDate') : t('btnDiffActive');
                btnDiff.title = t('tooltipDiffOn');
            } else {
                btnDiff.classList.remove('active');
                btnDiff.textContent = t('btnDiffAll');
                btnDiff.title = t('tooltipDiffOff');
            }

            // Compteur recherche
            searchCount.textContent = searchTerm
                ? `${visible} / ${total}`
                : diffOnly && totalDiff > 0 ? t('diffCount', totalDiff) : '';

            // Titre
            headerTitle.childNodes[0].textContent = diffOnly
                ? `${eventName} — ${totalDiff > 0 ? t('poiCountDiff', totalDiff) : t('upToDateTitle')}`
                : `${eventName} — ${t('poiCount', total)}`;

            updateMasterCb();
        }

        table.appendChild(tbody); scroll.appendChild(table); box.appendChild(scroll);

        // Appliquer le filtre initial (diff only par défaut)
        // Tri par défaut : diff en premier, puis alpha dans chaque groupe
        function applyDefaultSort() {
            const rows = Array.from(tbody.querySelectorAll('tr'));
            rows.sort((a, b) => {
                const aDiff = a.dataset.hasDiff === 'true' ? 0 : 1;
                const bDiff = b.dataset.hasDiff === 'true' ? 0 : 1;
                if (aDiff !== bDiff) return aDiff - bDiff;
                const va = (a.dataset.oldName || a.dataset.newName || '').toLowerCase();
                const vb = (b.dataset.oldName || b.dataset.newName || '').toLowerCase();
                return va.localeCompare(vb);
            });
            rows.forEach(r => tbody.appendChild(r));
        }

        // Si aucun tri volontaire actif, on maintient le tri alpha par défaut
        // (le tri volontaire positionne currentSortTh, le retour à "none" le remet à null)
        const origSortTable = sortTable;
        function sortTableWithDefaultFallback(th) {
            origSortTable(th);
            // Si on revient à "none" (ordre original), on réapplique le tri alpha par défaut
            if (!currentSortTh) applyDefaultSort();
        }
        ths.forEach(th => {
            if (th.dataset.sortKey) {
                th._sortHandler = () => sortTableWithDefaultFallback(th);
                th.addEventListener('click', th._sortHandler);
            }
        });

        applyDefaultSort();
        applyFilters();
        updateMasterCb();

        // Compteurs lock dans le titre
        if (cntSae > 0 || cntHard > 0) {
            if (cntSae > 0) {
                const b = document.createElement('span'); b.className = 'peu-lock-badge sae';
                b.textContent = t('badgeSae', cntSae); b.title = t('badgeSaeTitle');
                headerTitle.appendChild(b);
            }
            if (cntHard > 0) {
                const b = document.createElement('span'); b.className = 'peu-lock-badge hard';
                b.textContent = `${cntHard} 🔒`; b.title = t('badgeHardTitle');
                headerTitle.appendChild(b);
            }
        }

        const footer = document.createElement('div'); footer.className = 'peu-footer';
        let footerTxt = t('footerHelp');
        if (cntSae > 0)  footerTxt += '  ' + t('footerSae');
        if (cntHard > 0) footerTxt += '  ' + t('footerHard');
        footer.textContent = footerTxt;
        box.appendChild(footer);

        overlay.appendChild(box);
        document.body.appendChild(overlay);
        makeDraggable(box, header);

        function applyChanges() {
            // Collecter les lignes cochées
            const toApply = [];
            tbody.querySelectorAll('tr').forEach(tr => {
                const {vid, inpName, txtArea, cb, valeurs} = tr._inputs;
                if (!cb.checked) return;
                const poi = pois.find(p => getVenueIdFromPermalink(p.perm) === vid);
                /* 🔴 `valeurs` A MANQUÉ ICI, ET LE SCRIPT A DIT « SUCCÈS ». Sans lui,
                   `runApply` recevait un objet sans champs du lot D2 : rien n'était
                   envoyé au SDK, aucune action n'entrait dans la pile, et l'écran
                   annonçait « 1 POI appliqué avec succès » sur un lieu intact.
                   ⇒ Une chaîne qui se coupe entre l'aperçu et l'application ne se
                     voit QUE sur la carte : l'aperçu, lui, était juste. */
                if (poi) toApply.push({vid, perm: poi.perm, inpName, txtArea, name: inpName.value, valeurs});
            });

            if (!toApply.length) { overlay.remove(); return; }

            // Désactiver les boutons pendant l'application
            btnV.disabled = true; btnX.disabled = true; btnMin.disabled = true;
            btnV.style.opacity = '0.5'; btnX.style.opacity = '0.5';

            // Masquer le footer normal et créer un div de progression dédié
            footer.style.display = 'none';
            const progressDiv = document.createElement('div');
            progressDiv.style.cssText = 'padding:8px 14px;background:#f0f4fb;border-top:1px solid #dde3ee;flex-shrink:0;';
            const applyLabel = document.createElement('div');
            applyLabel.className = 'peu-progress-label';
            applyLabel.style.marginBottom = '4px';
            applyLabel.textContent = t('applying', 0, toApply.length);
            const applyBarBg = document.createElement('div'); applyBarBg.className = 'peu-progress-bar-bg';
            const applyBar   = document.createElement('div'); applyBar.className   = 'peu-progress-bar';
            applyBarBg.appendChild(applyBar);
            progressDiv.appendChild(applyLabel);
            progressDiv.appendChild(applyBarBg);
            box.appendChild(progressDiv);

            function showExportFooter(results) {
                // Retirer le div de progression
                if (progressDiv.parentNode) progressDiv.parentNode.removeChild(progressDiv);
                // Créer le footer export et l'ajouter à la box
                const exportFooter = document.createElement('div');
                exportFooter.style.cssText = 'padding:10px 14px;background:#f0f4fb;border-top:2px solid #27ae60;font-size:11px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;';
                const appliedCount = results.filter(r => r.status === 'applied').length;
                const msg = document.createElement('span');
                msg.style.color = '#27ae60'; msg.style.fontWeight = '600';
                msg.textContent = t('successMsg', appliedCount);
                const btnExport = document.createElement('button');
                btnExport.style.cssText = 'background:#2C6ED5;color:#fff;border:none;border-radius:4px;padding:5px 12px;font-size:11px;cursor:pointer;font-weight:600;';
                btnExport.textContent = t('btnExport');
                btnExport.onclick = () => exportReport(eventName, results);
                const btnClose = document.createElement('button');
                btnClose.style.cssText = 'background:#888;color:#fff;border:none;border-radius:4px;padding:5px 12px;font-size:11px;cursor:pointer;margin-left:6px;';
                btnClose.textContent = t('btnClose');
                btnClose.onclick = () => overlay.remove();
                const btns = document.createElement('div'); btns.style.display = 'flex';
                btns.appendChild(btnExport); btns.appendChild(btnClose);
                exportFooter.appendChild(msg); exportFooter.appendChild(btns);
                box.appendChild(exportFooter);
            }

            async function runApply(items, allResults = []) {
                const UpdateObject = require('Waze/Action/UpdateObject');
                const failed = [];

                for (let i = 0; i < items.length; i++) {
                    const item = items[i];
                    const {vid, perm, inpName, txtArea} = item;
                    const venue = await centerAndLoad(perm, vid, 3000);
                    if (venue) {
                        const oldName = venue.attributes.name || '';
                        const oldDesc = venue.attributes.description || '';
                        /* ⚠️ Les noms alternatifs viennent du classeur s'il en porte,
                           sinon on REPASSE ceux du lieu tels quels — règle d'origine
                           du script : ne jamais les perdre au passage. */
                        const aPoser = (item.valeurs && item.valeurs.aPoser) || {};
                        W.model.actionManager.add(new UpdateObject(venue, {
                            id: venue.attributes.id,
                            name: inpName.value,
                            description: txtArea.value,
                            aliases: aPoser.aliases || venue.attributes.aliases || []
                        }));

                        /* ── Les champs du lot D2, par le SDK ──
                           ⚠️⚠️ ON RELIT APRÈS AVOIR ÉCRIT, ET C'EST OBLIGATOIRE : le SDK
                              n'élève aucune erreur devant un champ qu'il ne connaît pas
                              ni devant une valeur hors énumération. Sans cette relecture,
                              « appliqué » ne voudrait dire que « l'appel n'a pas planté ». */
                        let poseSdk = null;
                        const { maj, ignores } = construireMaj(aPoser);

                        /* ⚠️⚠️ UN CHAMP À POSER QUI N'EST JAMAIS ENVOYÉ EST UN MANQUE, PAS UN
                           SUCCÈS. Si la liste à appliquer se construit sans les valeurs,
                           `maj` sort vide, le SDK n'est pas appelé, et l'écran annonce
                           « appliqué avec succès » sur un lieu que personne n'a touché.
                           Le compte ci-dessous le refuse. */
                        const attendus = Object.keys(aPoser).filter(c => !CIBLES_HERITEES.includes(c));
                        if (attendus.length && !Object.keys(maj).length) {
                            poseSdk = { identiques: [], differents: attendus };
                        }

                        if (Object.keys(maj).length) {
                            try {
                                const sdk = obtenirSdk();
                                sdk.DataModel.Venues.updateVenue({ venueId: vid, ...maj });
                                const relu = W.model.venues.getObjectById(vid);
                                poseSdk = comparerAuLieu(relu ? relu.attributes : {}, aPoser);
                            } catch (e) {
                                poseSdk = { identiques: [], differents: Object.keys(maj), erreur: e.message };
                            }
                        }
                        if (ignores.length) {
                            poseSdk = poseSdk || { identiques: [], differents: [] };
                            poseSdk.differents = poseSdk.differents.concat(ignores);
                        }

                        allResults.push({
                            oldName, newName: inpName.value,
                            oldDesc, newDesc: txtArea.value,
                            status: poseSdk && poseSdk.differents.length ? 'partial' : 'applied',
                            poses: poseSdk ? poseSdk.identiques : [],
                            manques: poseSdk ? poseSdk.differents : []
                        });
                    } else {
                        failed.push(item);
                        allResults.push({
                            oldName: '', newName: inpName.value,
                            oldDesc: '', newDesc: txtArea.value,
                            status: 'timeout'
                        });
                    }
                    applyBar.style.width = Math.round((i + 1) / items.length * 100) + '%';
                    applyLabel.textContent = t('applying', i+1, items.length);
                }

                // ⭐ ON RESTE SUR LE PÉRIMÈTRE, comme après le préchargement :
                //    l'application déplace la carte de lieu en lieu, et revenir
                //    au point de départ masquerait ce qu'on vient de poser.
                cadrerSurLesLieux(venueMap);

                if (failed.length === 0) {
                    await new Promise(r => setTimeout(r, 100));
                    showExportFooter(allResults);
                    btnX.disabled = false; btnX.style.opacity = '1';
                    // Enregistrer dans l'historique avec délai pour éviter interférence DOM
                    const currentFile = getHistory()[0]?.name;
                    if (currentFile) {
                        recordFileApplied(currentFile);
                        setTimeout(() => renderHistory(), 500);
                    }
                    return;
                }

                // Rapport d'échec
                box.removeChild(progressDiv);
                const errFooter = document.createElement('div'); errFooter.className = 'peu-footer-error';
                const errTitle = document.createElement('div'); errTitle.className = 'peu-error-title';
                errTitle.textContent = t('timeoutReport', failed.length);
                errFooter.appendChild(errTitle);
                const ul = document.createElement('ul');
                failed.forEach(f => {
                    const li = document.createElement('li');
                    li.textContent = f.inpName.value || f.vid;
                    ul.appendChild(li);
                });
                errFooter.appendChild(ul);

                const btnRetry = document.createElement('button'); btnRetry.className = 'peu-btn-retry';
                btnRetry.textContent = t('btnRetry', failed.length);
                btnRetry.onclick = () => {
                    box.removeChild(errFooter);
                    box.appendChild(progressDiv);
                    applyLabel.textContent = t('applying', 0, failed.length);
                    applyBar.style.width = '0%';
                    runApply(failed, allResults);
                };
                errFooter.appendChild(btnRetry);
                box.appendChild(errFooter);

                // Réactiver fermeture uniquement
                btnX.disabled = false; btnX.style.opacity = '1';
                btnX.onclick = () => overlay.remove();
                btnMin.disabled = false;
            }

            runApply(toApply);
        }
    }

    let _peuInited = false;
    function _peuInit() { if (_peuInited) return; _peuInited = true; initScript(); }

    if (W?.userscripts?.state?.isReady) {
        _peuInit();
    } else {
        document.addEventListener('wme-ready', _peuInit, {once:true});
        const fallback = setInterval(() => {
            if (W?.userscripts?.state?.isReady) { clearInterval(fallback); _peuInit(); }
        }, 500);
        setTimeout(() => clearInterval(fallback), 30000);
    }
})();
