// ==UserScript==
// @name         WME POI Event Updater
// @name:fr      WME POI Event Updater
// @namespace    http://tampermonkey.net/
// @version      0.48
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
                sheetHeaderErr:'En-têtes manquants en A1/B1/C1 — onglet ignoré',
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
                sheetHeaderErr:'Missing headers in A1/B1/C1 — sheet ignored',
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

    const CSS = `
        .peu-overlay {
            position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
            background: transparent; z-index: 10000;
            font-family: 'Segoe UI', Arial, sans-serif;
            pointer-events: none;
        }
        .peu-box {
            position: absolute;
            background: #f9f9f9; width: 480px; min-width: 340px;
            max-width: 96vw; max-height: 88vh; min-height: 120px; border-radius: 10px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.35); display: flex; flex-direction: column;
            overflow: hidden; pointer-events: all; resize: both;
            transition: max-height 0.2s ease, box-shadow 0.2s;
        }
        .peu-box.minimized {
            max-height: 44px !important;
            box-shadow: 0 4px 16px rgba(0,0,0,0.25);
        }
        .peu-box.minimized .peu-scroll,
        .peu-box.minimized .peu-footer,
        .peu-box.minimized .peu-toolbar { display: none; }
        .peu-scroll { overflow: auto; flex: 1; }
        .peu-header {
            display: flex; align-items: center; justify-content: space-between;
            background: #2C6ED5; color: #fff; padding: 8px 14px; flex-shrink: 0;
            cursor: grab; user-select: none;
        }
        .peu-header:active { cursor: grabbing; }
        .peu-header-left { display: flex; align-items: center; gap: 8px; }
        .peu-header-left span { font-size: 13px; font-weight: 600; letter-spacing: 0.3px; }
        .peu-drag-hint { font-size: 10px; opacity: 0.6; letter-spacing: 0.2px; }
        .peu-header-btns { display: flex; gap: 6px; }
        .peu-btn-icon {
            background: rgba(255,255,255,0.15); border: none; border-radius: 5px;
            color: #fff; font-size: 14px; width: 28px; height: 28px; cursor: pointer;
            display: flex; align-items: center; justify-content: center;
            transition: background 0.15s; flex-shrink: 0;
        }
        .peu-btn-icon:hover { background: rgba(255,255,255,0.30); }
        .peu-btn-icon.apply:hover { background: #27ae60; }
        .peu-btn-icon.cancel:hover { background: #c0392b; }
        .peu-btn-icon.minimize { font-size: 16px; }
        .peu-btn-icon.diffonly { font-size: 11px; width: auto; padding: 0 7px; letter-spacing: 0.2px; }
        .peu-btn-icon.diffonly.active { background: rgba(255,255,255,0.35); }
        .peu-toolbar {
            display: flex; align-items: center; gap: 8px;
            padding: 6px 10px; background: #eef2fa;
            border-bottom: 1px solid #d0ddf5; flex-shrink: 0;
        }
        .peu-search {
            flex: 1; font-size: 11.5px; font-family: inherit;
            border: 1px solid #c5d3e8; border-radius: 4px; padding: 3px 8px;
            background: #fff; color: #222; outline: none;
            transition: border 0.15s;
        }
        .peu-search:focus { border-color: #2C6ED5; background: #f0f6ff; }
        .peu-search-clear {
            background: none; border: none; cursor: pointer; color: #aaa;
            font-size: 14px; padding: 0 2px; line-height: 1;
            display: none;
        }
        .peu-search-clear:hover { color: #c0392b; }
        .peu-search-count { font-size: 11px; color: #888; white-space: nowrap; }
        .peu-table {
            width: 100%; border-collapse: collapse; table-layout: fixed;
            font-size: 11.5px; color: #222;
        }
        .peu-table colgroup col:nth-child(1) { width: 34px; }
        .peu-table colgroup col:nth-child(2) { width: 42%; }
        .peu-table colgroup col:nth-child(3) { width: 46%; }
        .peu-table colgroup col:nth-child(4) { width: 30px; }
        .peu-table thead th {
            background: #e8eef8; color: #2C6ED5; font-size: 11px; font-weight: 700;
            text-transform: uppercase; letter-spacing: 0.4px; padding: 6px 5px;
            border-bottom: 2px solid #2C6ED5; position: sticky; top: 0; z-index: 2;
            text-align: center;
        }
        .peu-table thead th.sortable {
            cursor: pointer; user-select: none;
        }
        .peu-table thead th.sortable:hover { background: #d0ddf5; }
        .peu-table thead th .peu-sort-icon {
            display: inline-block; margin-left: 4px; opacity: 0.35; font-style: normal;
        }
        .peu-table thead th.sort-asc .peu-sort-icon,
        .peu-table thead th.sort-desc .peu-sort-icon { opacity: 1; color: #1a4fa0; }
        .peu-table tbody tr { transition: background 0.1s; }
        .peu-table tbody tr:nth-child(even) { background: #f0f4fb; }
        .peu-table tbody tr:nth-child(odd)  { background: #ffffff; }
        .peu-table tbody tr:hover { background: #dce8fb; }
        .peu-table td {
            padding: 4px 5px; border-bottom: 1px solid #e0e0e0;
            vertical-align: top; word-break: break-word;
        }
        .peu-table td.center { text-align: center; vertical-align: middle; }
        /* Vue fusionnée : ancienne valeur (1 ligne, tronquée) au-dessus du champ.
           Hauteur fixe identique dans les 2 colonnes → les champs restent alignés. */
        .peu-cell-old {
            color: #8a8a8a; font-size: 10px; line-height: 1.4; height: 14px;
            margin-bottom: 3px; white-space: nowrap; overflow: hidden;
            text-overflow: ellipsis;
        }
        .peu-cell-old.changed {
            color: #c0392b; text-decoration: line-through;
            text-decoration-color: rgba(192,57,43,0.5);
        }
        .peu-lock-ok   { font-size: 13px; cursor: default; }
        .peu-lock-sae  { font-size: 13px; cursor: default; color: #e67e22; }
        .peu-lock-hard { font-size: 13px; cursor: default; color: #c0392b; }
        .peu-table tr.peu-row-sae  { background: #fff8f0 !important; }
        .peu-table tr.peu-row-hard { background: #fff0f0 !important; opacity: 0.7; }
        .peu-table tr.peu-row-diff { background: #eef4ff !important; }
        .peu-table tr.peu-row-diff:hover { background: #dce8fb !important; }
        .peu-table tr.peu-row-unloaded { background: #fafafa !important; opacity: 0.6; }
        .peu-unloaded-icon { font-size: 11px; color: #aaa; display: block; margin-top: 2px; }
        .peu-diff-dot {
            display: block; width: 7px; height: 7px; border-radius: 50%;
            background: #2C6ED5; margin: 3px auto 0;
        }
        .peu-lock-badge {
            display: inline-block; font-size: 10px; font-weight: 700; border-radius: 3px;
            padding: 1px 4px; margin-left: 6px; vertical-align: middle;
        }
        .peu-lock-badge.sae  { background: #fdebd0; color: #e67e22; border: 1px solid #e67e22; }
        .peu-lock-badge.hard { background: #fadbd8; color: #c0392b; border: 1px solid #c0392b; }
        .peu-progress-wrap {
            padding: 10px 14px 8px; background: #f0f4fb;
            border-top: 1px solid #dde3ee; flex-shrink: 0;
        }
        .peu-progress-label {
            font-size: 11px; color: #2C6ED5; margin-bottom: 5px; text-align: center;
        }
        .peu-progress-bar-bg {
            background: #d0ddf5; border-radius: 6px; height: 8px; overflow: hidden;
        }
        .peu-progress-bar {
            background: #2C6ED5; height: 8px; width: 0%; border-radius: 6px;
            transition: width 0.3s ease;
        }
        .peu-btn-center {
            background: none; border: none; cursor: pointer; font-size: 13px;
            padding: 2px; border-radius: 4px; transition: background 0.1s;
        }
        .peu-btn-center:hover { background: #dce8fb; }
        .peu-input, .peu-textarea {
            width: 100%; box-sizing: border-box; height: 36px;
            font-size: 11.5px; font-family: inherit;
            border: 1px solid #c5d3e8; border-radius: 4px; padding: 3px 5px;
            background: #fff; color: #222; resize: none; transition: border 0.15s;
        }
        .peu-input:focus, .peu-textarea:focus {
            outline: none; border-color: #2C6ED5; background: #f0f6ff;
        }
        .peu-textarea { overflow-y: auto; line-height: 1.3; }
        .peu-checkbox { width: 14px; height: 14px; cursor: pointer; accent-color: #2C6ED5; }
        .peu-footer {
            padding: 6px 14px; background: #f0f4fb; border-top: 1px solid #dde3ee;
            font-size: 11px; color: #666; text-align: right; flex-shrink: 0;
        }
        .peu-footer-error {
            padding: 8px 14px; background: #fff5f5; border-top: 2px solid #e74c3c;
            font-size: 11px; flex-shrink: 0;
        }
        .peu-footer-error .peu-error-title {
            color: #c0392b; font-weight: 700; margin-bottom: 5px;
        }
        .peu-footer-error ul {
            margin: 4px 0 8px 16px; padding: 0; color: #555;
        }
        .peu-footer-error ul li { margin-bottom: 2px; }
        .peu-btn-retry {
            background: #e74c3c; color: #fff; border: none; border-radius: 5px;
            padding: 4px 12px; font-size: 11px; cursor: pointer; font-weight: 600;
        }
        .peu-btn-retry:hover { background: #c0392b; }
    `;

    function injectCSS() {
        if (document.getElementById('peu-style')) return;
        const style = document.createElement('style');
        style.id = 'peu-style';
        style.textContent = CSS;
        document.head.appendChild(style);
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
            saveOverlayGeom({
                left: box.offsetLeft, top: box.offsetTop,
                width: box.offsetWidth, height: box.offsetHeight
            });
        };

        // Géométrie initiale : restaure la taille/position mémorisées SI valides,
        // sinon largeur par défaut du CSS, ancrée en haut à droite (laisse voir la carte).
        const applyInitialGeom = () => {
            const saved = getOverlayGeom();
            const valid = saved && saved.width > 0 && saved.height > 0
                          && Number.isFinite(saved.left) && Number.isFinite(saved.top);
            const vw = window.innerWidth, vh = window.innerHeight;
            if (valid) {
                box.style.width  = Math.min(saved.width,  vw - 20) + 'px';
                box.style.height = Math.min(saved.height, vh - 20) + 'px';
            }
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
            // zoomLevel du permalink en le plafonnant à 17. La vue de l'utilisateur
            // est restaurée après le préchargement (savedCenter/savedZoom).
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
                    const wb = XLSX.read(new Uint8Array(ev.target.result), {type:'array'});
                    const all = [];
                    const warnings = [];

                    wb.SheetNames.filter(n => n !== 'Config').forEach(sheet => {
                        const sh = wb.Sheets[sheet];
                        if (!sh['A1']?.v || !sh['B1']?.v || !sh['C1']?.v) {
                            warnings.push(`[${sheet}] ${t('sheetHeaderErr')}`);
                            return;
                        }
                        const rows = XLSX.utils.sheet_to_json(sh, {header:['perm','name','desc'], range:1, defval:''});
                        const sheetPerms = new Map();
                        rows.forEach((r, idx) => {
                            const rowNum = idx + 2; // +2 car range:1 saute la ligne 1
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

                            all.push({event:sheet, perm:r.perm, name:r.name, desc:r.desc});
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

        // Sauvegarder la position/zoom actuelle pour y revenir après
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

        // Si annulé : nettoyer et sortir sans ouvrir l'overlay
        if (cancelRef.cancelled) {
            W.map.setCenter(savedCenter, savedZoom);
            loadingDiv.remove();
            return;
        }

        // Revenir à la position initiale
        W.map.setCenter(savedCenter, savedZoom);
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
                const isDiff = nameChanged || descChanged;
                tr.dataset.hasDiff = isDiff ? 'true' : 'false';
                // Ne pas écraser hard/sae
                if (lockStatus === 'ok') tr.classList.toggle('peu-row-diff', isDiff);
                diffDot.style.display = isDiff ? 'block' : 'none';
                // Barrer l'ancienne valeur uniquement si elle change réellement
                if (oldNameEl) oldNameEl.classList.toggle('changed', nameChanged);
                if (oldDescEl) oldDescEl.classList.toggle('changed', descChanged);
                // Cochage auto tant que l'utilisateur n'a pas décidé lui-même
                if (!cb.disabled && !tr._cbUserSet) cb.checked = isDiff;
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

            // ── Colonne case à cocher ──
            const tdCB = tr.insertCell(); tdCB.className = 'center';
            const cb = document.createElement('input'); cb.type = 'checkbox'; cb.className = 'peu-checkbox';
            cb.checked = false;
            if (lockStatus === 'hard' || !venueLoaded) cb.disabled = true;
            cb.addEventListener('change', () => { tr._cbUserSet = true; updateMasterCb(); });
            tdCB.appendChild(cb);

            tr._inputs = {vid, inpName, txtArea, cb};

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
                const {vid, inpName, txtArea, cb} = tr._inputs;
                if (!cb.checked) return;
                const poi = pois.find(p => getVenueIdFromPermalink(p.perm) === vid);
                if (poi) toApply.push({vid, perm: poi.perm, inpName, txtArea, name: inpName.value});
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

            const savedCenter = W.map.getCenter();
            const savedZoom   = W.map.getZoom();

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
                        W.model.actionManager.add(new UpdateObject(venue, {
                            id: venue.attributes.id,
                            name: inpName.value,
                            description: txtArea.value,
                            aliases: venue.attributes.aliases || []
                        }));
                        allResults.push({
                            oldName, newName: inpName.value,
                            oldDesc, newDesc: txtArea.value,
                            status: 'applied'
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

                W.map.setCenter(savedCenter, savedZoom);

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
