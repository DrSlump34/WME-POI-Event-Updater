// ==UserScript==
// @name         WME POI Event Updater
// @name:fr      WME POI Event Updater
// @namespace    http://tampermonkey.net/
// @version      0.54.00
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
// @homepageURL  https://github.com/DrSlump34/WME-POI-Event-Updater
// @supportURL   https://www.waze.com/discuss/t/script-wme-poi-event-updater/404593
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      update.greasyfork.org
// @icon         data:image/svg+xml;base64,PHN2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHdpZHRoPScyNCcgaGVpZ2h0PScyNCcgdmlld0JveD0nMCAwIDI0IDI0Jz48cGF0aCBkPSdNMTIgMS44QzcuOCAxLjggNC42IDUgNC42IDkuMWMwIDUuNiA3LjQgMTMuMSA3LjQgMTMuMXM3LjQtNy41IDcuNC0xMy4xQzE5LjQgNSAxNi4yIDEuOCAxMiAxLjh6JyBmaWxsPScjZmZiMzAwJyBzdHJva2U9JyM4ZDUzMDAnIHN0cm9rZS13aWR0aD0nMS40JyBzdHJva2UtbGluZWpvaW49J3JvdW5kJy8+PGNpcmNsZSBjeD0nMTInIGN5PSc5LjEnIHI9JzIuOScgZmlsbD0nI2ZmZicgc3Ryb2tlPScjOGQ1MzAwJyBzdHJva2Utd2lkdGg9JzEnLz48L3N2Zz4=
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js#sha256=cc015130aa8521e7f088f88898eba949ccdcbfb38df0bd129b44b7273c3a6f41
// @downloadURL  https://update.greasyfork.org/scripts/578776/WME%20POI%20Event%20Updater.user.js
// @updateURL    https://update.greasyfork.org/scripts/578776/WME%20POI%20Event%20Updater.meta.js
// ==/UserScript==

(function() {
    'use strict';

    /* ⚠️⚠️ LA PAGE SE LIT PAR `unsafeWindow`. La pastille de nouvelle version
       exige GM_xmlhttpRequest (la politique de sécurité de WME interdit d'appeler
       GreasyFork depuis la page) ; accorder une permission fait tourner le
       script dans un bac à sable, où `getWmeSdk` n'existe plus comme variable
       globale : il se lit par `pw`. banc-demarrage le vérifie en faisant
       tourner le script dans un faux bac à sable.
       ⚠️⚠️ 0.54.00 : PLUS AUCUN `W`, ni `OpenLayers`, ni `require`. Waze retire
          `W` de WME le 24/11/2026 (annonce Discuss 413664) : tout passe par le
          SDK. banc-demarrage fait tourner le script dans un WME SANS `W`. */
    const pw = (typeof unsafeWindow !== 'undefined' && unsafeWindow) ? unsafeWindow : window;

    const scriptId   = 'poi-event-updater';
    const URL_DISCUSS = 'https://www.waze.com/discuss/t/script-wme-poi-event-updater/404593';
    const URL_GF      = 'https://greasyfork.org/scripts/578776';
    const URL_GH      = 'https://github.com/DrSlump34/WME-POI-Event-Updater';
    const URL_INSTALLER = 'https://update.greasyfork.org/scripts/578776/WME%20POI%20Event%20Updater.user.js';
    const URL_META      = 'https://update.greasyfork.org/scripts/578776/WME%20POI%20Event%20Updater.meta.js';
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
    /* ⭐⭐ UNE SEULE ICÔNE, CELLE DE L'EN-TÊTE (`@icon`) : sur l'onglet, en tête du
       panneau, dans la fenêtre et sur le bouton de la carte (charte commune).
       Il y en avait trois — un carré sur GreasyFork, une épingle sur l'onglet et
       le bouton, un emoji dans les titres. check-architecture compare cette
       chaîne à celle de l'en-tête : les deux ne peuvent pas diverger. */
    const ICONE = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHdpZHRoPScyNCcgaGVpZ2h0PScyNCcgdmlld0JveD0nMCAwIDI0IDI0Jz48cGF0aCBkPSdNMTIgMS44QzcuOCAxLjggNC42IDUgNC42IDkuMWMwIDUuNiA3LjQgMTMuMSA3LjQgMTMuMXM3LjQtNy41IDcuNC0xMy4xQzE5LjQgNSAxNi4yIDEuOCAxMiAxLjh6JyBmaWxsPScjZmZiMzAwJyBzdHJva2U9JyM4ZDUzMDAnIHN0cm9rZS13aWR0aD0nMS40JyBzdHJva2UtbGluZWpvaW49J3JvdW5kJy8+PGNpcmNsZSBjeD0nMTInIGN5PSc5LjEnIHI9JzIuOScgZmlsbD0nI2ZmZicgc3Ryb2tlPScjOGQ1MzAwJyBzdHJva2Utd2lkdGg9JzEnLz48L3N2Zz4=';
    let poiData = [];
    /* ⚠️ LA POIGNEE VERS LE CHAMP DE FICHIER, pas une seconde lecture : le
       chemin qui lit le classeur valide vingt règles, tient un rapport
       d’anomalies et alimente l’historique. Le déplacer aurait été échanger
       une interface contre un risque. */
    let _peuFileInput = null;
    let _peuLang = 'en'; // initialisé dans initScript avant tout appel à t()

    // Détection langue
    /**
     * LA LANGUE DU SCRIPT : celle de WME, parmi les huit de la charte.
     * ⚠️ Le portugais se décide par le PAYS (pt-BR / pt-PT) : les deux diffèrent
     *    par le vocabulaire de l'éditeur, pas seulement par l'orthographe.
     */
    function detectLang(localeWme) {
        try {
            const l = String(localeWme
                || document.documentElement.lang || navigator.language || 'en').toLowerCase();
            if (l.startsWith('pt')) return l.indexOf('br') !== -1 ? 'pt-BR' : 'pt-PT';
            const c = l.slice(0, 2) === 'iw' ? 'he' : l.slice(0, 2);
            return ['fr', 'en', 'de', 'es', 'it', 'he'].indexOf(c) !== -1 ? c : 'en';
        } catch (e) { return 'en'; }
    }

    // Dictionnaire i18n construit une seule fois (mémoïsé) au 1er appel, puis réutilisé
    // — évite de reconstruire tout l'objet à chaque appel de t().
    let _strings = null;
    function t(key, ...args) {
        if (!_strings) _strings = {
            fr: {
                tabTitle:'WME POI Event Updater',
                panelTitle:'POI Event Updater', chooseFile:'📂 Choisir un fichier',
                chooseFileTitle:'Charger un classeur .xlsx depuis votre disque',
                selectSheet:'L’onglet du classeur à poser',
                fabTitle:'POI Event Updater — afficher la fenêtre',
                fabTitleOn:'POI Event Updater — masquer la fenêtre',
                btnApplyNone:'Rien de coché', btnApplyOne:'Appliquer 1 ligne',
                btnApplyN:(n)=>`Appliquer les ${n} lignes cochées`,
                btnApplyTitle:'Poser les valeurs des lignes cochées dans l’éditeur. Rien n’est enregistré : vous relirez dans WME.',
                btnExportTitle:'Enregistrer le rapport de cette pose',
                footerHelpVide:'Rien n’est écrit sur la carte tant que vous n’avez pas cliqué sur Appliquer.',
                guideFichier:'Choisissez le classeur de l’événement.',
                guideFichierSuite:'Ensuite vous choisirez l’onglet, puis vous relirez chaque ligne avant d’appliquer.',
                guideOnglet:'Choisissez l’onglet à poser.',
                guideOngletSuite:'Un onglet par événement. Celui « Hors Evenement » remet les lieux dans leur état ordinaire.',
                colSelect:'Poser', colEtat:'État',
                dropLigne1:'📄 Déposez un classeur ici',
                dropLigne2:'ou cliquez pour le choisir',
                dropTitre:'Déposez un fichier .xlsx n’importe où sur cette fenêtre, ou cliquez pour le choisir',
                dropRefus:(nom)=>`« ${nom} » n’est pas un classeur Excel : seuls les fichiers .xlsx et .xls se chargent ici.`,
                sbOuvrir:'Afficher la fenêtre', sbOuvrirTitre:'Ouvrir la fenêtre de travail — c’est là qu’on charge un classeur et qu’on relit avant d’appliquer',
                sbIntro:'Met à jour en lot le nom, la description et les champs des lieux, depuis un classeur : un onglet par événement.',
                sbAide:'Aide', historyVide:'Aucun classeur chargé pour l’instant.',
                aideClasseurT:'Le classeur',
                aideClasseur:'Un onglet par événement, une ligne par lieu. Les colonnes se lisent par leur en-tête : permalien, nom, description, puis les champs du lieu (téléphone, site, catégories, services, parking…).\nUne cellule vide d’un champ du lieu ne demande rien. Une description vide EFFACE celle du lieu : c’est ainsi qu’un onglet « Hors Evenement » remet les lieux à nu.',
                aideRelireT:'Relire et appliquer',
                aideRelire:'Chaque ligne montre l’avant (barré) et l’après. Ce qui change est coché d’office ; une retouche dans l’aperçu re-décide la case.\nAppliquer écrit dans l’éditeur sans enregistrer : relisez sur la carte, puis cliquez sur Enregistrer. Les lieux en échec restent cochés, pour réessayer.',
                aideEtatsT:'Les états d’une ligne',
                aideEtats:'Chiffre vert : champs à poser.\n-n orange : appliquer RETIRERAIT n valeurs — jamais coché d’office.\nSaE : lieu verrouillé au-dessus de votre niveau, la modification part en suggestion.\nL7 : verrou du staff, rien ne se pose.\n? : lieu introuvable.\n✔ : posé, en attente d’enregistrement.',
                sbEcrit:'✍️ Appliquer écrit dans l’éditeur ; le script n’enregistre jamais.',
                majDispo:(v)=>`Nouvelle version ${v} disponible.`, majInstaller:'Installer',
                cancelTitle:'Interrompre : ce qui est déjà lu est conservé',
                footerHelp:'Décochez ce que vous ne voulez pas poser. Les lignes orange RETIRENT des valeurs : elles ne sont jamais cochées d’office.',
                footerMasquees:(n)=>`${n} ligne(s) cochée(s) sont masquées par le filtre — elles seront posées aussi.`,
                bilanPartiel:(n)=>`⚠️ ${n} lieu(x) n’ont reçu qu’une partie des valeurs — voir le rapport.`,
                bilanEchec:(n)=>`⚠️ ${n} lieu(x) n’ont pas pu être traités — leurs lignes restent cochées :`,
                bilanSae:(n)=>`⚠️ ${n} lieu(x) verrouillé(s) au-dessus de votre niveau : la modification passe en suggestion (SaE).`,
                bilanNonEnregistre:(n)=>`${n} modification(s) ajoutée(s) à la pile de WME — RIEN N’EST ENREGISTRÉ : relisez, puis cliquez sur Enregistrer dans l’éditeur.`,
                bilanNonEnregistreSansCompte:'RIEN N’EST ENREGISTRÉ : relisez, puis cliquez sur Enregistrer dans l’éditeur.',
                bilanErreurGenerale:(m)=>`✖ La pose n’a pas pu commencer : ${m}`,
                echecLigne:(nom, statut, err)=>`${nom} — ${statut}${err ? ' : ' + err : ''}`,
                occupeDepot:'Un balayage ou une pose est en cours : attendez la fin avant de charger un autre classeur.',
                cbTitre:'Poser cette ligne dans l’éditeur',
                cbFige:'Cette ligne ne peut pas être posée — voir le badge d’état',
                colNameTitle:'Le nom à poser. Modifiable avant d’appliquer.',
                colDescTitle:'La description à poser. Une cellule vide EFFACE la description du lieu.',
                descEffacee:'⚠ la description du lieu sera effacée',
                nomEfface:'⚠ le nom du lieu serait effacé',
                triTitre:'Trier sur cette colonne',
                badgePerteTitle:(n)=>`Appliquer RETIRERAIT ${n} valeur(s) au lieu`,
                badgeDiffTitle:(n)=>`${n} champ(s) à poser`,
                badgeRienTitle:'Rien à poser : le lieu porte déjà ces valeurs',
                badgeOffTitle:'Lieu introuvable dans l’éditeur : rien ne peut être posé',
                badgePoseTitle:'Posé dans l’éditeur — en attente d’enregistrement',
                badgeSaeTitle:'Lieu verrouillé au-dessus de votre niveau : la modification passera en suggestion (SaE)',
                lockHardTitle:'Verrou niveau 7 (staff Waze) — édition impossible',
                noFile:'Aucun fichier choisi',
                historyTitle:'Fichiers récents', histLoaded:'📂 Chargé :', histApplied:'✔ Appliqué :',
                histNeverApplied:'Jamais appliqué', clearHistoryTitle:'Effacer l’historique',
                filterPlaceholder:'🔍 Filtrer par nom…',
                colName:'Nom', colDesc:'Description',
                btnReduce:'Réduire', btnRestore:'Restaurer', btnApply:'Appliquer', btnClose:'Fermer',
                btnDiffActive:'≠ Écarts', btnDiffAll:'≡ Tout',
                tooltipDiffOn:'N’afficher que les lieux à modifier', tooltipDiffOff:'Afficher tous les lieux',
                deplacerAide:'Déplacer la fenêtre : glisser, ou flèches du clavier (Maj : grand pas). Double-clic : position par défaut.',
                redimAide:'Redimensionner la fenêtre : glisser, ou flèches du clavier (Maj : grand pas)',
                loadingPois:(n,total)=>`Chargement des lieux… ${n} / ${total}`,
                applying:(n,total)=>`Application… ${n} / ${total}`,
                cancelBtn:'Annuler',
                noPoisLoaded:'✖ Aucun lieu valide chargé',
                anomalies:(n)=>`⚠️ ${n} anomalie${n>1?'s':''} dans le classeur`,
                btnRetry:(n)=>`🔄 Réessayer (${n} lieu${n>1?'x':''})`,
                successMsg:(n)=>`✔ ${n} lieu${n>1?'x':''} posé${n>1?'s':''}`,
                btnExport:'📥 Exporter le rapport',
                sheetHeaderErr:'En-têtes de colonnes absentes ou non reconnues — onglet ignoré',
                sheetHeaderFallback:'en-têtes non reconnues : colonnes lues par position (A = permalien, B = nom, C = description)',
                sheetHeaderConflit:'des en-têtes reconnues ne sont pas à leur place A/B/C — onglet ignoré : nommez les trois colonnes (permalien, nom, description)',
                colonneDoublon:(lib, cols)=>`deux colonnes pour « ${lib} » (${cols}) : seule la première est lue`,
                plusApplique:'✔ à poser :', plusMain:'✋ à poser à la main :', plusRefus:'⚠ non reconnu :',
                plusConforme:'· déjà conforme :', plusAvant:'remplace :', plusRetire:'RETIRE',
                urlInvalid:'URL invalide',
                urlBadHost:'URL non reconnue (doit être waze.com ou beta.waze.com/…/editor)',
                urlNoEnv:'paramètre env= manquant', urlBadLat:'lat= absent ou invalide',
                urlBadLon:'lon= absent ou invalide', urlBadZoom:'zoomLevel= absent ou invalide',
                urlNoVenues:'venues= absent', urlNoVid:'impossible de lire l’identifiant du lieu',
                urlAutreEnv:(env, courant)=>`permalien d’un autre serveur (env=${env}, l’éditeur est en ${courant}) : le lieu ne sera pas trouvé`,
                urlPlusieursLieux:'plusieurs lieux dans venues= : seul le premier est lu',
                nameEmpty:'nom vide', dupRow:(a,b)=>`Lignes ${a} et ${b} : permalien en double`,
                rowLabel:'ligne',
                reportHeaders:['Permalien','Nom avant','Nom après','Description avant','Description après','Statut','Champs non posés','Erreur'],
                statusApplied:'✔ Appliqué', statusPartial:'⚠ Partiel', statusTimeout:'✖ Introuvable', statusErreur:'✖ Erreur',
                masterCbTitle:'Tout cocher / décocher',
                poiCount:(n)=>`${n} lieu${n>1?'x':''}`,
                layerOffMsg:'⚠️ Le calque « Lieux » est éteint et n’a pas pu être allumé : activez-le (Calques > Lieux), puis rechoisissez l’onglet.',
                xlsxMissing:'⚠️ La bibliothèque Excel (SheetJS) n’a pas pu se charger. Vérifiez votre connexion ou autorisez cdn.sheetjs.com, puis rechargez la page (F5).',
                locateTitle:'Centrer la carte sur ce lieu',
                valOui:'oui', valNon:'non',
                chPerm:'Permalien', chName:'Nom', chDesc:'Description', chAliases:'Noms alternatifs',
                chPhone:'Téléphone', chUrl:'Site web', chServices:'Services', chCategories:'Catégories',
                chParkingType:'Type de parking', chHasTBR:'Type variable', chCostType:'Tarif',
                chPaymentType:'Modes de paiement', chParkingServices:'Services du parking', chLotType:'Situation',
                chSpots:'Nombre de places', chCanExit:'Sortie quand fermé',
                chHours:'Horaires', chAddress:'Adresse', chEntryPoints:'Points d’entrée', chOperator:'Opérateur de parking',
                chGoogleName:'Nom Google', chGoogleCategory:'Catégorie Google', chGooglePosition:'Position Google',
                moHours:'WME attend des créneaux, pas une phrase',
                moAddress:'WME attend un numéro et une rue de son propre modèle',
                moEntryPoints:'ce sont des points sur la carte, pas du texte',
                moOperator:'liste fermée chez WME, dont les clés ne sont pas relevées',
                moGoogleName:'ne relève pas de WME', moGoogleCategory:'ne relève pas de WME', moGooglePosition:'ne relève pas de WME',
                vaPublic:'Public', vaPrivate:'Privé', vaRestricted:'Restreint',
                vaFree:'Gratuit', vaLow:'Faible', vaModerate:'Modéré', vaExpensive:'Élevé',
                vaMultiLevel:'Plusieurs niveaux', vaStreetLevel:'Extérieur', vaStreetLevelCovered:'Extérieur couvert', vaUnderground:'Souterrain',
                vaCash:'Espèces', vaChecks:'Chèques', vaCredit:'Carte de crédit', vaDebitCard:'Carte bancaire',
                vaDigitalWallet:'Portefeuille numérique', vaElectronicPass:'Pass électronique', vaMembership:'Abonnement',
                vaParkingApp:'Application', vaPermit:'Laissez-passer', vaPrepaid:'Prépaiement', vaSmsCall:'SMS/Appel',
                vaAirConditioning:'Climatisation', vaCreditCards:'Accepte les cartes de crédit', vaCurbsidePickup:'Click & Collect',
                vaDeliveries:'Livraisons', vaDrivethrough:'Drive', vaOutsideSeating:'Terrasse extérieure',
                vaParkingForCustomers:'Parking client', vaReservations:'Réservations', vaRestrooms:'Toilettes',
                vaTakeAway:'À emporter', vaValletService:'Service de voiturier', vaWheelchairAccessible:'Accessible en fauteuil roulant',
                vaWiFi:'Wi-Fi', vaAirportShuttle:'Navette aéroport', vaCarpoolParking:'Places covoiturage', vaCarWash:'Lavage auto',
                vaCovered:'Couvert', vaDisabilityParking:'Places PMR', vaOnSiteAttendant:'Agent d’accueil', vaParkAndRide:'P+R',
                vaSecurity:'Surveillance', vaValet:'Voiturier', vaEvChargingStation:'Bornes de charge', vaUnknown:'Inconnu',
            },
            en: {
                tabTitle:'WME POI Event Updater',
                panelTitle:'POI Event Updater', chooseFile:'📂 Choose a file',
                chooseFileTitle:'Load an .xlsx workbook from your disk',
                selectSheet:'Which sheet to apply',
                fabTitle:'POI Event Updater — show the window',
                fabTitleOn:'POI Event Updater — hide the window',
                btnApplyNone:'Nothing ticked', btnApplyOne:'Apply 1 row',
                btnApplyN:(n)=>`Apply the ${n} ticked rows`,
                btnApplyTitle:'Write the ticked rows into the editor. Nothing is saved: you will review in WME.',
                btnExportTitle:'Save the report of this run',
                footerHelpVide:'Nothing is written to the map until you click Apply.',
                guideFichier:'Choose the event workbook.',
                guideFichierSuite:'Then pick the sheet, and review every row before applying.',
                guideOnglet:'Choose the sheet to apply.',
                guideOngletSuite:'One sheet per event. The off-event sheet puts places back to their ordinary state.',
                colSelect:'Apply', colEtat:'State',
                dropLigne1:'📄 Drop a workbook here',
                dropLigne2:'or click to pick one',
                dropTitre:'Drop an .xlsx file anywhere on this window, or click to pick one',
                dropRefus:(nom)=>`« ${nom} » is not an Excel workbook: only .xlsx and .xls files load here.`,
                sbOuvrir:'Show the window', sbOuvrirTitre:'Open the work window — that is where you load a workbook and review before applying',
                sbIntro:'Bulk-updates the name, description and fields of places from a workbook: one sheet per event.',
                sbAide:'Help', historyVide:'No workbook loaded yet.',
                aideClasseurT:'The workbook',
                aideClasseur:'One sheet per event, one row per place. Columns are read by their header: permalink, name, description, then the place fields (phone, website, categories, services, parking…).\nAn empty cell in a place field asks for nothing. An empty description ERASES the place’s one: that is how an off-event sheet puts places back to plain.',
                aideRelireT:'Review and apply',
                aideRelire:'Each row shows the before (struck through) and the after. What changes is ticked by default; editing a value in the preview decides the box again.\nApply writes into the editor without saving: review on the map, then click Save. Places that failed stay ticked, to retry.',
                aideEtatsT:'Row states',
                aideEtats:'Green number: fields to write.\nOrange -n: applying would REMOVE n values — never ticked by default.\nSaE: place locked above your level, the change goes as a suggestion.\nL7: staff lock, nothing is written.\n?: place not found.\n✔: written, waiting to be saved.',
                sbEcrit:'✍️ Apply writes into the editor; the script never saves.',
                majDispo:(v)=>`New version ${v} available.`, majInstaller:'Install',
                cancelTitle:'Stop: what is already loaded is kept',
                footerHelp:'Untick what you do not want to write. Orange rows REMOVE values: they are never ticked by default.',
                footerMasquees:(n)=>`${n} ticked row(s) are hidden by the filter — they will be written too.`,
                bilanPartiel:(n)=>`⚠️ ${n} place(s) only received part of the values — see the report.`,
                bilanEchec:(n)=>`⚠️ ${n} place(s) could not be processed — their rows stay ticked:`,
                bilanSae:(n)=>`⚠️ ${n} place(s) locked above your level: the change goes as a suggestion (SaE).`,
                bilanNonEnregistre:(n)=>`${n} change(s) added to the WME stack — NOTHING IS SAVED: review, then click Save in the editor.`,
                bilanNonEnregistreSansCompte:'NOTHING IS SAVED: review, then click Save in the editor.',
                bilanErreurGenerale:(m)=>`✖ Writing could not start: ${m}`,
                echecLigne:(nom, statut, err)=>`${nom} — ${statut}${err ? ': ' + err : ''}`,
                occupeDepot:'A scan or a write is in progress: wait for it to finish before loading another workbook.',
                cbTitre:'Write this row into the editor',
                cbFige:'This row cannot be written — see the state badge',
                colNameTitle:'The name to write. Editable before applying.',
                colDescTitle:'The description to write. An empty cell ERASES the place description.',
                descEffacee:'⚠ the place description will be erased',
                nomEfface:'⚠ the place name would be erased',
                triTitre:'Sort on this column',
                badgePerteTitle:(n)=>`Applying would REMOVE ${n} value(s) from the place`,
                badgeDiffTitle:(n)=>`${n} field(s) to write`,
                badgeRienTitle:'Nothing to write: the place already carries these values',
                badgeOffTitle:'Place not found in the editor: nothing can be written',
                badgePoseTitle:'Written into the editor — waiting to be saved',
                badgeSaeTitle:'Place locked above your level: the change will go as a suggestion (SaE)',
                lockHardTitle:'Level 7 lock (Waze staff) — editing not possible',
                noFile:'No file chosen',
                historyTitle:'Recent files', histLoaded:'📂 Loaded:', histApplied:'✔ Applied:',
                histNeverApplied:'Never applied', clearHistoryTitle:'Clear history',
                filterPlaceholder:'🔍 Filter by name…',
                colName:'Name', colDesc:'Description',
                btnReduce:'Minimize', btnRestore:'Restore', btnApply:'Apply', btnClose:'Close',
                btnDiffActive:'≠ Changes', btnDiffAll:'≡ All',
                tooltipDiffOn:'Show only the places to change', tooltipDiffOff:'Show all places',
                deplacerAide:'Move the window: drag, or keyboard arrows (Shift: big step). Double-click: default position.',
                redimAide:'Resize the window: drag, or keyboard arrows (Shift: big step)',
                loadingPois:(n,total)=>`Loading places… ${n} / ${total}`,
                applying:(n,total)=>`Applying… ${n} / ${total}`,
                cancelBtn:'Cancel',
                noPoisLoaded:'✖ No valid place loaded',
                anomalies:(n)=>`⚠️ ${n} issue${n>1?'s':''} in the workbook`,
                btnRetry:(n)=>`🔄 Retry (${n} place${n>1?'s':''})`,
                successMsg:(n)=>`✔ ${n} place${n>1?'s':''} written`,
                btnExport:'📥 Export report',
                sheetHeaderErr:'Column headers missing or unrecognised — sheet ignored',
                sheetHeaderFallback:'headers not recognised: columns read by position (A = permalink, B = name, C = description)',
                sheetHeaderConflit:'recognised headers are not in their A/B/C place — sheet ignored: name the three columns (permalink, name, description)',
                colonneDoublon:(lib, cols)=>`two columns for « ${lib} » (${cols}): only the first one is read`,
                plusApplique:'✔ to write:', plusMain:'✋ to set by hand:', plusRefus:'⚠ not recognised:',
                plusConforme:'· already correct:', plusAvant:'replaces:', plusRetire:'REMOVES',
                urlInvalid:'Invalid URL',
                urlBadHost:'Unrecognised URL (must be waze.com or beta.waze.com/…/editor)',
                urlNoEnv:'missing env= parameter', urlBadLat:'lat= missing or invalid',
                urlBadLon:'lon= missing or invalid', urlBadZoom:'zoomLevel= missing or invalid',
                urlNoVenues:'venues= missing', urlNoVid:'unable to read the place ID',
                urlAutreEnv:(env, courant)=>`permalink from another server (env=${env}, the editor is on ${courant}): the place will not be found`,
                urlPlusieursLieux:'several places in venues=: only the first one is read',
                nameEmpty:'empty name', dupRow:(a,b)=>`Rows ${a} and ${b}: duplicate permalink`,
                rowLabel:'row',
                reportHeaders:['Permalink','Name before','Name after','Description before','Description after','Status','Fields not written','Error'],
                statusApplied:'✔ Applied', statusPartial:'⚠ Partial', statusTimeout:'✖ Not found', statusErreur:'✖ Error',
                masterCbTitle:'Tick / untick all',
                poiCount:(n)=>`${n} place${n>1?'s':''}`,
                layerOffMsg:'⚠️ The « Places » layer is off and could not be switched on: enable it (Layers > Places), then pick the sheet again.',
                xlsxMissing:'⚠️ The Excel library (SheetJS) could not load. Check your connection or allow cdn.sheetjs.com, then reload the page (F5).',
                locateTitle:'Center the map on this place',
                valOui:'yes', valNon:'no',
                chPerm:'Permalink', chName:'Name', chDesc:'Description', chAliases:'Alternate names',
                chPhone:'Phone', chUrl:'Website', chServices:'Services', chCategories:'Categories',
                chParkingType:'Parking type', chHasTBR:'Type varies', chCostType:'Cost',
                chPaymentType:'Payment methods', chParkingServices:'Parking services', chLotType:'Lot type',
                chSpots:'Number of spots', chCanExit:'Exit while closed',
                chHours:'Opening hours', chAddress:'Address', chEntryPoints:'Entry points', chOperator:'Parking operator',
                chGoogleName:'Google name', chGoogleCategory:'Google category', chGooglePosition:'Google position',
                moHours:'WME expects time slots, not a sentence',
                moAddress:'WME expects a house number and a street from its own model',
                moEntryPoints:'these are points on the map, not text',
                moOperator:'closed list in WME, whose keys have not been collected',
                moGoogleName:'not a WME field', moGoogleCategory:'not a WME field', moGooglePosition:'not a WME field',
                vaPublic:'Public', vaPrivate:'Private', vaRestricted:'Restricted',
                vaFree:'Free', vaLow:'Low', vaModerate:'Moderate', vaExpensive:'Expensive',
                vaMultiLevel:'Multi-level', vaStreetLevel:'Street level', vaStreetLevelCovered:'Street level, covered', vaUnderground:'Underground',
                vaCash:'Cash', vaChecks:'Checks', vaCredit:'Credit card', vaDebitCard:'Debit card',
                vaDigitalWallet:'Digital wallet', vaElectronicPass:'Electronic pass', vaMembership:'Membership',
                vaParkingApp:'Parking app', vaPermit:'Permit', vaPrepaid:'Prepaid', vaSmsCall:'SMS/Call',
                vaAirConditioning:'Air conditioning', vaCreditCards:'Accepts credit cards', vaCurbsidePickup:'Curbside pickup',
                vaDeliveries:'Deliveries', vaDrivethrough:'Drive-through', vaOutsideSeating:'Outside seating',
                vaParkingForCustomers:'Parking for customers', vaReservations:'Reservations', vaRestrooms:'Restrooms',
                vaTakeAway:'Take away', vaValletService:'Valet service', vaWheelchairAccessible:'Wheelchair accessible',
                vaWiFi:'Wi-Fi', vaAirportShuttle:'Airport shuttle', vaCarpoolParking:'Carpool parking', vaCarWash:'Car wash',
                vaCovered:'Covered', vaDisabilityParking:'Disability parking', vaOnSiteAttendant:'On-site attendant', vaParkAndRide:'Park and ride',
                vaSecurity:'Security', vaValet:'Valet', vaEvChargingStation:'EV charging station', vaUnknown:'Unknown',
            },
            de: {
                tabTitle:'WME POI Event Updater',
                panelTitle:'POI Event Updater', chooseFile:'📂 Datei wählen',
                chooseFileTitle:'Eine .xlsx-Arbeitsmappe von Ihrem Rechner laden',
                selectSheet:'Das anzuwendende Tabellenblatt',
                fabTitle:'POI Event Updater — Fenster anzeigen',
                fabTitleOn:'POI Event Updater — Fenster ausblenden',
                btnApplyNone:'Nichts angehakt', btnApplyOne:'1 Zeile anwenden',
                btnApplyN:(n)=>`Die ${n} angehakten Zeilen anwenden`,
                btnApplyTitle:'Die angehakten Zeilen in den Editor schreiben. Nichts wird gespeichert: Sie prüfen danach in WME.',
                btnExportTitle:'Den Bericht dieses Durchlaufs speichern',
                footerHelpVide:'Auf der Karte wird nichts geschrieben, bevor Sie auf Anwenden klicken.',
                guideFichier:'Wählen Sie die Arbeitsmappe der Veranstaltung.',
                guideFichierSuite:'Danach wählen Sie das Tabellenblatt und prüfen jede Zeile vor dem Anwenden.',
                guideOnglet:'Wählen Sie das anzuwendende Tabellenblatt.',
                guideOngletSuite:'Ein Blatt pro Veranstaltung. Das Blatt „außerhalb der Veranstaltung“ setzt die Orte in ihren normalen Zustand zurück.',
                colSelect:'Anwenden', colEtat:'Status',
                dropLigne1:'📄 Arbeitsmappe hier ablegen',
                dropLigne2:'oder klicken, um sie zu wählen',
                dropTitre:'Eine .xlsx-Datei irgendwo auf dieses Fenster ziehen, oder klicken, um sie zu wählen',
                dropRefus:(nom)=>`„${nom}“ ist keine Excel-Arbeitsmappe: hier lassen sich nur .xlsx- und .xls-Dateien laden.`,
                sbOuvrir:'Fenster anzeigen', sbOuvrirTitre:'Das Arbeitsfenster öffnen — dort lädt man eine Arbeitsmappe und prüft vor dem Anwenden',
                sbIntro:'Aktualisiert Name, Beschreibung und Felder von Orten stapelweise aus einer Arbeitsmappe: ein Blatt pro Veranstaltung.',
                sbAide:'Hilfe', historyVide:'Noch keine Arbeitsmappe geladen.',
                aideClasseurT:'Die Arbeitsmappe',
                aideClasseur:'Ein Blatt pro Veranstaltung, eine Zeile pro Ort. Die Spalten werden über ihre Überschrift gelesen: Permalink, Name, Beschreibung, dann die Felder des Ortes (Telefon, Website, Kategorien, Dienste, Parkplatz…).\nEine leere Zelle in einem Feld des Ortes verlangt nichts. Eine leere Beschreibung LÖSCHT die des Ortes: so setzt ein Blatt „außerhalb der Veranstaltung“ die Orte zurück.',
                aideRelireT:'Prüfen und anwenden',
                aideRelire:'Jede Zeile zeigt das Vorher (durchgestrichen) und das Nachher. Was sich ändert, ist vorab angehakt; eine Änderung in der Vorschau entscheidet das Häkchen neu.\nAnwenden schreibt in den Editor, ohne zu speichern: auf der Karte prüfen, dann Speichern klicken. Fehlgeschlagene Orte bleiben angehakt, zum erneuten Versuch.',
                aideEtatsT:'Zeilenstatus',
                aideEtats:'Grüne Zahl: zu schreibende Felder.\nOrange -n: Anwenden würde n Werte ENTFERNEN — nie vorab angehakt.\nSaE: Ort über Ihrer Stufe gesperrt, die Änderung geht als Vorschlag.\nL7: Staff-Sperre, nichts wird geschrieben.\n?: Ort nicht gefunden.\n✔: geschrieben, wartet auf Speichern.',
                sbEcrit:'✍️ Anwenden schreibt in den Editor; das Skript speichert nie.',
                majDispo:(v)=>`Neue Version ${v} verfügbar.`, majInstaller:'Installieren',
                cancelTitle:'Abbrechen: bereits Geladenes bleibt erhalten',
                footerHelp:'Haken Sie ab, was Sie nicht schreiben wollen. Orange Zeilen ENTFERNEN Werte: sie sind nie vorab angehakt.',
                footerMasquees:(n)=>`${n} angehakte Zeile(n) sind durch den Filter ausgeblendet — sie werden trotzdem geschrieben.`,
                bilanPartiel:(n)=>`⚠️ ${n} Ort(e) haben nur einen Teil der Werte erhalten — siehe Bericht.`,
                bilanEchec:(n)=>`⚠️ ${n} Ort(e) konnten nicht verarbeitet werden — ihre Zeilen bleiben angehakt:`,
                bilanSae:(n)=>`⚠️ ${n} Ort(e) über Ihrer Stufe gesperrt: die Änderung geht als Vorschlag (SaE).`,
                bilanNonEnregistre:(n)=>`${n} Änderung(en) zum WME-Stapel hinzugefügt — NICHTS IST GESPEICHERT: prüfen, dann im Editor auf Speichern klicken.`,
                bilanNonEnregistreSansCompte:'NICHTS IST GESPEICHERT: prüfen, dann im Editor auf Speichern klicken.',
                bilanErreurGenerale:(m)=>`✖ Das Schreiben konnte nicht beginnen: ${m}`,
                echecLigne:(nom, statut, err)=>`${nom} — ${statut}${err ? ': ' + err : ''}`,
                occupeDepot:'Ein Durchlauf oder ein Schreibvorgang läuft: warten Sie das Ende ab, bevor Sie eine andere Arbeitsmappe laden.',
                cbTitre:'Diese Zeile in den Editor schreiben',
                cbFige:'Diese Zeile kann nicht geschrieben werden — siehe Statusabzeichen',
                colNameTitle:'Der zu schreibende Name. Vor dem Anwenden änderbar.',
                colDescTitle:'Die zu schreibende Beschreibung. Eine leere Zelle LÖSCHT die Beschreibung des Ortes.',
                descEffacee:'⚠ die Beschreibung des Ortes wird gelöscht',
                nomEfface:'⚠ der Name des Ortes würde gelöscht',
                triTitre:'Nach dieser Spalte sortieren',
                badgePerteTitle:(n)=>`Anwenden würde ${n} Wert(e) vom Ort ENTFERNEN`,
                badgeDiffTitle:(n)=>`${n} Feld(er) zu schreiben`,
                badgeRienTitle:'Nichts zu schreiben: der Ort trägt diese Werte bereits',
                badgeOffTitle:'Ort im Editor nicht gefunden: nichts kann geschrieben werden',
                badgePoseTitle:'In den Editor geschrieben — wartet auf Speichern',
                badgeSaeTitle:'Ort über Ihrer Stufe gesperrt: die Änderung geht als Vorschlag (SaE)',
                lockHardTitle:'Sperre Stufe 7 (Waze-Staff) — Bearbeiten nicht möglich',
                noFile:'Keine Datei gewählt',
                historyTitle:'Letzte Dateien', histLoaded:'📂 Geladen:', histApplied:'✔ Angewendet:',
                histNeverApplied:'Nie angewendet', clearHistoryTitle:'Verlauf löschen',
                filterPlaceholder:'🔍 Nach Namen filtern…',
                colName:'Name', colDesc:'Beschreibung',
                btnReduce:'Minimieren', btnRestore:'Wiederherstellen', btnApply:'Anwenden', btnClose:'Schließen',
                btnDiffActive:'≠ Änderungen', btnDiffAll:'≡ Alle',
                tooltipDiffOn:'Nur die zu ändernden Orte anzeigen', tooltipDiffOff:'Alle Orte anzeigen',
                deplacerAide:'Fenster verschieben: ziehen, oder Pfeiltasten (Umschalt: großer Schritt). Doppelklick: Standardposition.',
                redimAide:'Fenstergröße ändern: ziehen, oder Pfeiltasten (Umschalt: großer Schritt)',
                loadingPois:(n,total)=>`Orte werden geladen… ${n} / ${total}`,
                applying:(n,total)=>`Wird angewendet… ${n} / ${total}`,
                cancelBtn:'Abbrechen',
                noPoisLoaded:'✖ Kein gültiger Ort geladen',
                anomalies:(n)=>`⚠️ ${n} Auffälligkeit${n>1?'en':''} in der Arbeitsmappe`,
                btnRetry:(n)=>`🔄 Erneut versuchen (${n} Ort${n>1?'e':''})`,
                successMsg:(n)=>`✔ ${n} Ort${n>1?'e':''} geschrieben`,
                btnExport:'📥 Bericht exportieren',
                sheetHeaderErr:'Spaltenüberschriften fehlen oder sind unbekannt — Blatt ignoriert',
                sheetHeaderFallback:'Überschriften unbekannt: Spalten nach Position gelesen (A = Permalink, B = Name, C = Beschreibung)',
                sheetHeaderConflit:'erkannte Überschriften stehen nicht an ihrer Stelle A/B/C — Blatt ignoriert: benennen Sie die drei Spalten (Permalink, Name, Beschreibung)',
                colonneDoublon:(lib, cols)=>`zwei Spalten für „${lib}“ (${cols}): nur die erste wird gelesen`,
                plusApplique:'✔ zu schreiben:', plusMain:'✋ von Hand zu setzen:', plusRefus:'⚠ nicht erkannt:',
                plusConforme:'· bereits korrekt:', plusAvant:'ersetzt:', plusRetire:'ENTFERNT',
                urlInvalid:'Ungültige URL',
                urlBadHost:'Unbekannte URL (muss waze.com oder beta.waze.com/…/editor sein)',
                urlNoEnv:'Parameter env= fehlt', urlBadLat:'lat= fehlt oder ungültig',
                urlBadLon:'lon= fehlt oder ungültig', urlBadZoom:'zoomLevel= fehlt oder ungültig',
                urlNoVenues:'venues= fehlt', urlNoVid:'die Kennung des Ortes ist nicht lesbar',
                urlAutreEnv:(env, courant)=>`Permalink eines anderen Servers (env=${env}, der Editor ist auf ${courant}): der Ort wird nicht gefunden`,
                urlPlusieursLieux:'mehrere Orte in venues=: nur der erste wird gelesen',
                nameEmpty:'leerer Name', dupRow:(a,b)=>`Zeilen ${a} und ${b}: doppelter Permalink`,
                rowLabel:'Zeile',
                reportHeaders:['Permalink','Name vorher','Name nachher','Beschreibung vorher','Beschreibung nachher','Status','Nicht geschriebene Felder','Fehler'],
                statusApplied:'✔ Angewendet', statusPartial:'⚠ Teilweise', statusTimeout:'✖ Nicht gefunden', statusErreur:'✖ Fehler',
                masterCbTitle:'Alle an- / abhaken',
                poiCount:(n)=>`${n} Ort${n>1?'e':''}`,
                layerOffMsg:'⚠️ Die Ebene „Orte“ ist aus und ließ sich nicht einschalten: aktivieren Sie sie (Ebenen > Orte) und wählen Sie das Blatt erneut.',
                xlsxMissing:'⚠️ Die Excel-Bibliothek (SheetJS) konnte nicht geladen werden. Prüfen Sie die Verbindung oder erlauben Sie cdn.sheetjs.com, dann Seite neu laden (F5).',
                locateTitle:'Karte auf diesen Ort zentrieren',
                valOui:'ja', valNon:'nein',
                chPerm:'Permalink', chName:'Name', chDesc:'Beschreibung', chAliases:'Alternative Namen',
                chPhone:'Telefon', chUrl:'Website', chServices:'Dienste', chCategories:'Kategorien',
                chParkingType:'Parkplatztyp', chHasTBR:'Typ variiert', chCostType:'Kosten',
                chPaymentType:'Zahlungsarten', chParkingServices:'Parkplatzdienste', chLotType:'Lage',
                chSpots:'Anzahl der Stellplätze', chCanExit:'Ausfahrt bei Schließung',
                chHours:'Öffnungszeiten', chAddress:'Adresse', chEntryPoints:'Zufahrtspunkte', chOperator:'Parkplatzbetreiber',
                chGoogleName:'Google-Name', chGoogleCategory:'Google-Kategorie', chGooglePosition:'Google-Position',
                moHours:'WME erwartet Zeitfenster, keinen Satz',
                moAddress:'WME erwartet Hausnummer und Straße aus seinem eigenen Modell',
                moEntryPoints:'das sind Punkte auf der Karte, kein Text',
                moOperator:'geschlossene Liste in WME, deren Schlüssel nicht erfasst sind',
                moGoogleName:'kein WME-Feld', moGoogleCategory:'kein WME-Feld', moGooglePosition:'kein WME-Feld',
                vaPublic:'Öffentlich', vaPrivate:'Privat', vaRestricted:'Eingeschränkt',
                vaFree:'Kostenlos', vaLow:'Günstig', vaModerate:'Mittel', vaExpensive:'Teuer',
                vaMultiLevel:'Mehrgeschossig', vaStreetLevel:'Ebenerdig', vaStreetLevelCovered:'Ebenerdig, überdacht', vaUnderground:'Tiefgarage',
                vaCash:'Bargeld', vaChecks:'Schecks', vaCredit:'Kreditkarte', vaDebitCard:'Debitkarte',
                vaDigitalWallet:'Digitale Geldbörse', vaElectronicPass:'Elektronischer Pass', vaMembership:'Mitgliedschaft',
                vaParkingApp:'Park-App', vaPermit:'Genehmigung', vaPrepaid:'Vorauszahlung', vaSmsCall:'SMS/Anruf',
                vaAirConditioning:'Klimaanlage', vaCreditCards:'Kreditkarten akzeptiert', vaCurbsidePickup:'Abholung am Straßenrand',
                vaDeliveries:'Lieferungen', vaDrivethrough:'Drive-in', vaOutsideSeating:'Außenbereich',
                vaParkingForCustomers:'Kundenparkplatz', vaReservations:'Reservierungen', vaRestrooms:'Toiletten',
                vaTakeAway:'Zum Mitnehmen', vaValletService:'Parkservice', vaWheelchairAccessible:'Rollstuhlgerecht',
                vaWiFi:'WLAN', vaAirportShuttle:'Flughafen-Shuttle', vaCarpoolParking:'Fahrgemeinschaftsparkplätze', vaCarWash:'Autowäsche',
                vaCovered:'Überdacht', vaDisabilityParking:'Behindertenparkplätze', vaOnSiteAttendant:'Personal vor Ort', vaParkAndRide:'Park & Ride',
                vaSecurity:'Überwachung', vaValet:'Parkservice (Valet)', vaEvChargingStation:'Ladestation', vaUnknown:'Unbekannt',
            },
            es: {
                tabTitle:'WME POI Event Updater',
                panelTitle:'POI Event Updater', chooseFile:'📂 Elegir un archivo',
                chooseFileTitle:'Cargar un libro .xlsx desde su disco',
                selectSheet:'La hoja del libro que se aplicará',
                fabTitle:'POI Event Updater — mostrar la ventana',
                fabTitleOn:'POI Event Updater — ocultar la ventana',
                btnApplyNone:'Nada marcado', btnApplyOne:'Aplicar 1 fila',
                btnApplyN:(n)=>`Aplicar las ${n} filas marcadas`,
                btnApplyTitle:'Escribir las filas marcadas en el editor. No se guarda nada: revisará en WME.',
                btnExportTitle:'Guardar el informe de esta aplicación',
                footerHelpVide:'No se escribe nada en el mapa hasta que haga clic en Aplicar.',
                guideFichier:'Elija el libro del evento.',
                guideFichierSuite:'Después elegirá la hoja y revisará cada fila antes de aplicar.',
                guideOnglet:'Elija la hoja que se aplicará.',
                guideOngletSuite:'Una hoja por evento. La hoja «fuera de evento» devuelve los lugares a su estado habitual.',
                colSelect:'Aplicar', colEtat:'Estado',
                dropLigne1:'📄 Suelte un libro aquí',
                dropLigne2:'o haga clic para elegirlo',
                dropTitre:'Suelte un archivo .xlsx en cualquier parte de esta ventana, o haga clic para elegirlo',
                dropRefus:(nom)=>`«${nom}» no es un libro de Excel: aquí solo se cargan archivos .xlsx y .xls.`,
                sbOuvrir:'Mostrar la ventana', sbOuvrirTitre:'Abrir la ventana de trabajo — allí se carga un libro y se revisa antes de aplicar',
                sbIntro:'Actualiza por lotes el nombre, la descripción y los campos de los lugares desde un libro: una hoja por evento.',
                sbAide:'Ayuda', historyVide:'Todavía no se ha cargado ningún libro.',
                aideClasseurT:'El libro',
                aideClasseur:'Una hoja por evento, una fila por lugar. Las columnas se leen por su encabezado: enlace permanente, nombre, descripción y luego los campos del lugar (teléfono, sitio web, categorías, servicios, aparcamiento…).\nUna celda vacía en un campo del lugar no pide nada. Una descripción vacía BORRA la del lugar: así una hoja «fuera de evento» deja los lugares como estaban.',
                aideRelireT:'Revisar y aplicar',
                aideRelire:'Cada fila muestra el antes (tachado) y el después. Lo que cambia se marca de oficio; retocar un valor en la vista previa vuelve a decidir la casilla.\nAplicar escribe en el editor sin guardar: revise en el mapa y luego haga clic en Guardar. Los lugares con error siguen marcados, para reintentar.',
                aideEtatsT:'Estados de una fila',
                aideEtats:'Número verde: campos que se escribirán.\n-n naranja: aplicar QUITARÍA n valores — nunca se marca de oficio.\nSaE: lugar bloqueado por encima de su nivel, el cambio va como sugerencia.\nL7: bloqueo del staff, no se escribe nada.\n?: lugar no encontrado.\n✔: escrito, pendiente de guardar.',
                sbEcrit:'✍️ Aplicar escribe en el editor; el script nunca guarda.',
                majDispo:(v)=>`Nueva versión ${v} disponible.`, majInstaller:'Instalar',
                cancelTitle:'Interrumpir: lo ya cargado se conserva',
                footerHelp:'Desmarque lo que no quiera escribir. Las filas naranjas QUITAN valores: nunca se marcan de oficio.',
                footerMasquees:(n)=>`${n} fila(s) marcada(s) están ocultas por el filtro — también se escribirán.`,
                bilanPartiel:(n)=>`⚠️ ${n} lugar(es) solo recibieron parte de los valores — vea el informe.`,
                bilanEchec:(n)=>`⚠️ ${n} lugar(es) no se pudieron procesar — sus filas siguen marcadas:`,
                bilanSae:(n)=>`⚠️ ${n} lugar(es) bloqueado(s) por encima de su nivel: el cambio va como sugerencia (SaE).`,
                bilanNonEnregistre:(n)=>`${n} cambio(s) añadido(s) a la pila de WME — NO SE HA GUARDADO NADA: revise y luego haga clic en Guardar en el editor.`,
                bilanNonEnregistreSansCompte:'NO SE HA GUARDADO NADA: revise y luego haga clic en Guardar en el editor.',
                bilanErreurGenerale:(m)=>`✖ La escritura no pudo empezar: ${m}`,
                echecLigne:(nom, statut, err)=>`${nom} — ${statut}${err ? ': ' + err : ''}`,
                occupeDepot:'Hay una carga o una escritura en curso: espere a que termine antes de cargar otro libro.',
                cbTitre:'Escribir esta fila en el editor',
                cbFige:'Esta fila no se puede escribir — vea la insignia de estado',
                colNameTitle:'El nombre que se escribirá. Modificable antes de aplicar.',
                colDescTitle:'La descripción que se escribirá. Una celda vacía BORRA la descripción del lugar.',
                descEffacee:'⚠ se borrará la descripción del lugar',
                nomEfface:'⚠ se borraría el nombre del lugar',
                triTitre:'Ordenar por esta columna',
                badgePerteTitle:(n)=>`Aplicar QUITARÍA ${n} valor(es) al lugar`,
                badgeDiffTitle:(n)=>`${n} campo(s) que se escribirán`,
                badgeRienTitle:'Nada que escribir: el lugar ya tiene estos valores',
                badgeOffTitle:'Lugar no encontrado en el editor: no se puede escribir nada',
                badgePoseTitle:'Escrito en el editor — pendiente de guardar',
                badgeSaeTitle:'Lugar bloqueado por encima de su nivel: el cambio irá como sugerencia (SaE)',
                lockHardTitle:'Bloqueo de nivel 7 (staff de Waze) — no se puede editar',
                noFile:'Ningún archivo elegido',
                historyTitle:'Archivos recientes', histLoaded:'📂 Cargado:', histApplied:'✔ Aplicado:',
                histNeverApplied:'Nunca aplicado', clearHistoryTitle:'Borrar el historial',
                filterPlaceholder:'🔍 Filtrar por nombre…',
                colName:'Nombre', colDesc:'Descripción',
                btnReduce:'Minimizar', btnRestore:'Restaurar', btnApply:'Aplicar', btnClose:'Cerrar',
                btnDiffActive:'≠ Cambios', btnDiffAll:'≡ Todo',
                tooltipDiffOn:'Mostrar solo los lugares que cambian', tooltipDiffOff:'Mostrar todos los lugares',
                deplacerAide:'Mover la ventana: arrastrar, o flechas del teclado (Mayús: paso grande). Doble clic: posición por defecto.',
                redimAide:'Cambiar el tamaño de la ventana: arrastrar, o flechas del teclado (Mayús: paso grande)',
                loadingPois:(n,total)=>`Cargando lugares… ${n} / ${total}`,
                applying:(n,total)=>`Aplicando… ${n} / ${total}`,
                cancelBtn:'Cancelar',
                noPoisLoaded:'✖ No se cargó ningún lugar válido',
                anomalies:(n)=>`⚠️ ${n} anomalía${n>1?'s':''} en el libro`,
                btnRetry:(n)=>`🔄 Reintentar (${n} lugar${n>1?'es':''})`,
                successMsg:(n)=>`✔ ${n} lugar${n>1?'es':''} escrito${n>1?'s':''}`,
                btnExport:'📥 Exportar el informe',
                sheetHeaderErr:'Encabezados de columna ausentes o no reconocidos — hoja ignorada',
                sheetHeaderFallback:'encabezados no reconocidos: columnas leídas por posición (A = enlace, B = nombre, C = descripción)',
                sheetHeaderConflit:'hay encabezados reconocidos fuera de su lugar A/B/C — hoja ignorada: nombre las tres columnas (enlace, nombre, descripción)',
                colonneDoublon:(lib, cols)=>`dos columnas para «${lib}» (${cols}): solo se lee la primera`,
                plusApplique:'✔ a escribir:', plusMain:'✋ a poner a mano:', plusRefus:'⚠ no reconocido:',
                plusConforme:'· ya correcto:', plusAvant:'sustituye:', plusRetire:'QUITA',
                urlInvalid:'URL no válida',
                urlBadHost:'URL no reconocida (debe ser waze.com o beta.waze.com/…/editor)',
                urlNoEnv:'falta el parámetro env=', urlBadLat:'lat= ausente o no válido',
                urlBadLon:'lon= ausente o no válido', urlBadZoom:'zoomLevel= ausente o no válido',
                urlNoVenues:'falta venues=', urlNoVid:'no se puede leer el identificador del lugar',
                urlAutreEnv:(env, courant)=>`enlace de otro servidor (env=${env}, el editor está en ${courant}): no se encontrará el lugar`,
                urlPlusieursLieux:'varios lugares en venues=: solo se lee el primero',
                nameEmpty:'nombre vacío', dupRow:(a,b)=>`Filas ${a} y ${b}: enlace duplicado`,
                rowLabel:'fila',
                reportHeaders:['Enlace permanente','Nombre antes','Nombre después','Descripción antes','Descripción después','Estado','Campos no escritos','Error'],
                statusApplied:'✔ Aplicado', statusPartial:'⚠ Parcial', statusTimeout:'✖ No encontrado', statusErreur:'✖ Error',
                masterCbTitle:'Marcar / desmarcar todo',
                poiCount:(n)=>`${n} lugar${n>1?'es':''}`,
                layerOffMsg:'⚠️ La capa «Lugares» está apagada y no se pudo encender: actívela (Capas > Lugares) y vuelva a elegir la hoja.',
                xlsxMissing:'⚠️ No se pudo cargar la biblioteca de Excel (SheetJS). Compruebe la conexión o permita cdn.sheetjs.com, y recargue la página (F5).',
                locateTitle:'Centrar el mapa en este lugar',
                valOui:'sí', valNon:'no',
                chPerm:'Enlace permanente', chName:'Nombre', chDesc:'Descripción', chAliases:'Nombres alternativos',
                chPhone:'Teléfono', chUrl:'Sitio web', chServices:'Servicios', chCategories:'Categorías',
                chParkingType:'Tipo de aparcamiento', chHasTBR:'Tipo variable', chCostType:'Tarifa',
                chPaymentType:'Formas de pago', chParkingServices:'Servicios del aparcamiento', chLotType:'Ubicación',
                chSpots:'Número de plazas', chCanExit:'Salida estando cerrado',
                chHours:'Horario', chAddress:'Dirección', chEntryPoints:'Puntos de entrada', chOperator:'Operador del aparcamiento',
                chGoogleName:'Nombre en Google', chGoogleCategory:'Categoría en Google', chGooglePosition:'Posición en Google',
                moHours:'WME espera franjas horarias, no una frase',
                moAddress:'WME espera un número y una calle de su propio modelo',
                moEntryPoints:'son puntos en el mapa, no texto',
                moOperator:'lista cerrada en WME, cuyas claves no se han recopilado',
                moGoogleName:'no es un campo de WME', moGoogleCategory:'no es un campo de WME', moGooglePosition:'no es un campo de WME',
                vaPublic:'Público', vaPrivate:'Privado', vaRestricted:'Restringido',
                vaFree:'Gratuito', vaLow:'Bajo', vaModerate:'Moderado', vaExpensive:'Caro',
                vaMultiLevel:'Varias plantas', vaStreetLevel:'Exterior', vaStreetLevelCovered:'Exterior cubierto', vaUnderground:'Subterráneo',
                vaCash:'Efectivo', vaChecks:'Cheques', vaCredit:'Tarjeta de crédito', vaDebitCard:'Tarjeta de débito',
                vaDigitalWallet:'Monedero digital', vaElectronicPass:'Pase electrónico', vaMembership:'Abono',
                vaParkingApp:'Aplicación', vaPermit:'Permiso', vaPrepaid:'Prepago', vaSmsCall:'SMS/Llamada',
                vaAirConditioning:'Aire acondicionado', vaCreditCards:'Acepta tarjetas de crédito', vaCurbsidePickup:'Recogida en la acera',
                vaDeliveries:'Entregas a domicilio', vaDrivethrough:'Autoservicio en coche', vaOutsideSeating:'Terraza',
                vaParkingForCustomers:'Aparcamiento para clientes', vaReservations:'Reservas', vaRestrooms:'Aseos',
                vaTakeAway:'Para llevar', vaValletService:'Aparcacoches', vaWheelchairAccessible:'Accesible en silla de ruedas',
                vaWiFi:'Wi-Fi', vaAirportShuttle:'Lanzadera al aeropuerto', vaCarpoolParking:'Plazas para coche compartido', vaCarWash:'Lavado de coches',
                vaCovered:'Cubierto', vaDisabilityParking:'Plazas para movilidad reducida', vaOnSiteAttendant:'Personal in situ', vaParkAndRide:'Aparcamiento disuasorio',
                vaSecurity:'Vigilancia', vaValet:'Aparcacoches (valet)', vaEvChargingStation:'Puntos de recarga', vaUnknown:'Desconocido',
            },
            it: {
                tabTitle:'WME POI Event Updater',
                panelTitle:'POI Event Updater', chooseFile:'📂 Scegli un file',
                chooseFileTitle:'Carica una cartella di lavoro .xlsx dal tuo disco',
                selectSheet:'Il foglio da applicare',
                fabTitle:'POI Event Updater — mostra la finestra',
                fabTitleOn:'POI Event Updater — nascondi la finestra',
                btnApplyNone:'Nulla di selezionato', btnApplyOne:'Applica 1 riga',
                btnApplyN:(n)=>`Applica le ${n} righe selezionate`,
                btnApplyTitle:'Scrive le righe selezionate nell’editor. Nulla viene salvato: rileggerai in WME.',
                btnExportTitle:'Salva il resoconto di questa applicazione',
                footerHelpVide:'Nulla viene scritto sulla mappa finché non fai clic su Applica.',
                guideFichier:'Scegli la cartella di lavoro dell’evento.',
                guideFichierSuite:'Poi sceglierai il foglio e rileggerai ogni riga prima di applicare.',
                guideOnglet:'Scegli il foglio da applicare.',
                guideOngletSuite:'Un foglio per evento. Il foglio «fuori evento» riporta i luoghi al loro stato ordinario.',
                colSelect:'Applica', colEtat:'Stato',
                dropLigne1:'📄 Trascina qui una cartella di lavoro',
                dropLigne2:'oppure fai clic per sceglierla',
                dropTitre:'Trascina un file .xlsx in qualunque punto di questa finestra, oppure fai clic per sceglierlo',
                dropRefus:(nom)=>`«${nom}» non è una cartella di lavoro Excel: qui si caricano solo file .xlsx e .xls.`,
                sbOuvrir:'Mostra la finestra', sbOuvrirTitre:'Apri la finestra di lavoro — lì si carica una cartella e si rilegge prima di applicare',
                sbIntro:'Aggiorna in blocco nome, descrizione e campi dei luoghi da una cartella di lavoro: un foglio per evento.',
                sbAide:'Aiuto', historyVide:'Nessuna cartella di lavoro caricata finora.',
                aideClasseurT:'La cartella di lavoro',
                aideClasseur:'Un foglio per evento, una riga per luogo. Le colonne si leggono dalla loro intestazione: permalink, nome, descrizione, poi i campi del luogo (telefono, sito, categorie, servizi, parcheggio…).\nUna cella vuota in un campo del luogo non chiede nulla. Una descrizione vuota CANCELLA quella del luogo: è così che un foglio «fuori evento» riporta i luoghi allo stato base.',
                aideRelireT:'Rileggere e applicare',
                aideRelire:'Ogni riga mostra il prima (barrato) e il dopo. Ciò che cambia è selezionato d’ufficio; ritoccare un valore nell’anteprima ridecide la casella.\nApplica scrive nell’editor senza salvare: rileggi sulla mappa, poi fai clic su Salva. I luoghi in errore restano selezionati, per riprovare.',
                aideEtatsT:'Gli stati di una riga',
                aideEtats:'Numero verde: campi da scrivere.\n-n arancione: applicare RIMUOVEREBBE n valori — mai selezionata d’ufficio.\nSaE: luogo bloccato sopra il tuo livello, la modifica va come suggerimento.\nL7: blocco dello staff, non si scrive nulla.\n?: luogo non trovato.\n✔: scritto, in attesa di salvataggio.',
                sbEcrit:'✍️ Applica scrive nell’editor; lo script non salva mai.',
                majDispo:(v)=>`Nuova versione ${v} disponibile.`, majInstaller:'Installa',
                cancelTitle:'Interrompi: ciò che è già caricato resta',
                footerHelp:'Deseleziona ciò che non vuoi scrivere. Le righe arancioni RIMUOVONO valori: non sono mai selezionate d’ufficio.',
                footerMasquees:(n)=>`Righe selezionate nascoste dal filtro: ${n} — verranno scritte anch’esse.`,
                bilanPartiel:(n)=>`⚠️ Luoghi che hanno ricevuto solo una parte dei valori: ${n} — vedi il resoconto.`,
                bilanEchec:(n)=>`⚠️ Luoghi non elaborati: ${n} — le loro righe restano selezionate:`,
                bilanSae:(n)=>`⚠️ Luoghi bloccati sopra il tuo livello: ${n} — la modifica va come suggerimento (SaE).`,
                bilanNonEnregistre:(n)=>`Modifiche aggiunte alla pila di WME: ${n} — NULLA È SALVATO: rileggi, poi fai clic su Salva nell’editor.`,
                bilanNonEnregistreSansCompte:'NULLA È SALVATO: rileggi, poi fai clic su Salva nell’editor.',
                bilanErreurGenerale:(m)=>`✖ La scrittura non è potuta partire: ${m}`,
                echecLigne:(nom, statut, err)=>`${nom} — ${statut}${err ? ': ' + err : ''}`,
                occupeDepot:'È in corso un caricamento o una scrittura: attendi la fine prima di caricare un’altra cartella.',
                cbTitre:'Scrivi questa riga nell’editor',
                cbFige:'Questa riga non può essere scritta — vedi il badge di stato',
                colNameTitle:'Il nome da scrivere. Modificabile prima di applicare.',
                colDescTitle:'La descrizione da scrivere. Una cella vuota CANCELLA la descrizione del luogo.',
                descEffacee:'⚠ la descrizione del luogo sarà cancellata',
                nomEfface:'⚠ il nome del luogo verrebbe cancellato',
                triTitre:'Ordina per questa colonna',
                badgePerteTitle:(n)=>`Applicare RIMUOVEREBBE valori dal luogo: ${n}`,
                badgeDiffTitle:(n)=>`Campi da scrivere: ${n}`,
                badgeRienTitle:'Nulla da scrivere: il luogo ha già questi valori',
                badgeOffTitle:'Luogo non trovato nell’editor: non si può scrivere nulla',
                badgePoseTitle:'Scritto nell’editor — in attesa di salvataggio',
                badgeSaeTitle:'Luogo bloccato sopra il tuo livello: la modifica andrà come suggerimento (SaE)',
                lockHardTitle:'Blocco livello 7 (staff Waze) — modifica impossibile',
                noFile:'Nessun file scelto',
                historyTitle:'File recenti', histLoaded:'📂 Caricato:', histApplied:'✔ Applicato:',
                histNeverApplied:'Mai applicato', clearHistoryTitle:'Cancella la cronologia',
                filterPlaceholder:'🔍 Filtra per nome…',
                colName:'Nome', colDesc:'Descrizione',
                btnReduce:'Riduci', btnRestore:'Ripristina', btnApply:'Applica', btnClose:'Chiudi',
                btnDiffActive:'≠ Modifiche', btnDiffAll:'≡ Tutto',
                tooltipDiffOn:'Mostra solo i luoghi da modificare', tooltipDiffOff:'Mostra tutti i luoghi',
                deplacerAide:'Sposta la finestra: trascina, oppure frecce della tastiera (Maiusc: passo grande). Doppio clic: posizione predefinita.',
                redimAide:'Ridimensiona la finestra: trascina, oppure frecce della tastiera (Maiusc: passo grande)',
                loadingPois:(n,total)=>`Caricamento dei luoghi… ${n} / ${total}`,
                applying:(n,total)=>`Applicazione… ${n} / ${total}`,
                cancelBtn:'Annulla',
                noPoisLoaded:'✖ Nessun luogo valido caricato',
                anomalies:(n)=>`⚠️ ${n} anomali${n>1?'e':'a'} nella cartella di lavoro`,
                btnRetry:(n)=>`🔄 Riprova (${n} luog${n>1?'hi':'o'})`,
                successMsg:(n)=>`✔ ${n} luog${n>1?'hi scritti':'o scritto'}`,
                btnExport:'📥 Esporta il resoconto',
                sheetHeaderErr:'Intestazioni di colonna assenti o non riconosciute — foglio ignorato',
                sheetHeaderFallback:'intestazioni non riconosciute: colonne lette per posizione (A = permalink, B = nome, C = descrizione)',
                sheetHeaderConflit:'intestazioni riconosciute fuori dal loro posto A/B/C — foglio ignorato: dai un nome alle tre colonne (permalink, nome, descrizione)',
                colonneDoublon:(lib, cols)=>`due colonne per «${lib}» (${cols}): si legge solo la prima`,
                plusApplique:'✔ da scrivere:', plusMain:'✋ da impostare a mano:', plusRefus:'⚠ non riconosciuto:',
                plusConforme:'· già corretto:', plusAvant:'sostituisce:', plusRetire:'RIMUOVE',
                urlInvalid:'URL non valido',
                urlBadHost:'URL non riconosciuto (deve essere waze.com o beta.waze.com/…/editor)',
                urlNoEnv:'parametro env= mancante', urlBadLat:'lat= assente o non valido',
                urlBadLon:'lon= assente o non valido', urlBadZoom:'zoomLevel= assente o non valido',
                urlNoVenues:'venues= assente', urlNoVid:'impossibile leggere l’identificativo del luogo',
                urlAutreEnv:(env, courant)=>`permalink di un altro server (env=${env}, l’editor è su ${courant}): il luogo non sarà trovato`,
                urlPlusieursLieux:'più luoghi in venues=: si legge solo il primo',
                nameEmpty:'nome vuoto', dupRow:(a,b)=>`Righe ${a} e ${b}: permalink duplicato`,
                rowLabel:'riga',
                reportHeaders:['Permalink','Nome prima','Nome dopo','Descrizione prima','Descrizione dopo','Stato','Campi non scritti','Errore'],
                statusApplied:'✔ Applicato', statusPartial:'⚠ Parziale', statusTimeout:'✖ Non trovato', statusErreur:'✖ Errore',
                masterCbTitle:'Seleziona / deseleziona tutto',
                poiCount:(n)=>`${n} luog${n>1?'hi':'o'}`,
                layerOffMsg:'⚠️ Il livello «Luoghi» è spento e non è stato possibile accenderlo: attivalo (Livelli > Luoghi), poi scegli di nuovo il foglio.',
                xlsxMissing:'⚠️ Impossibile caricare la libreria Excel (SheetJS). Controlla la connessione o consenti cdn.sheetjs.com, poi ricarica la pagina (F5).',
                locateTitle:'Centra la mappa su questo luogo',
                valOui:'sì', valNon:'no',
                chPerm:'Permalink', chName:'Nome', chDesc:'Descrizione', chAliases:'Nomi alternativi',
                chPhone:'Telefono', chUrl:'Sito web', chServices:'Servizi', chCategories:'Categorie',
                chParkingType:'Tipo di parcheggio', chHasTBR:'Tipo variabile', chCostType:'Tariffa',
                chPaymentType:'Metodi di pagamento', chParkingServices:'Servizi del parcheggio', chLotType:'Collocazione',
                chSpots:'Numero di posti', chCanExit:'Uscita a parcheggio chiuso',
                chHours:'Orari', chAddress:'Indirizzo', chEntryPoints:'Punti di accesso', chOperator:'Gestore del parcheggio',
                chGoogleName:'Nome Google', chGoogleCategory:'Categoria Google', chGooglePosition:'Posizione Google',
                moHours:'WME si aspetta fasce orarie, non una frase',
                moAddress:'WME si aspetta un numero civico e una via del proprio modello',
                moEntryPoints:'sono punti sulla mappa, non testo',
                moOperator:'elenco chiuso in WME, le cui chiavi non sono state rilevate',
                moGoogleName:'non è un campo di WME', moGoogleCategory:'non è un campo di WME', moGooglePosition:'non è un campo di WME',
                vaPublic:'Pubblico', vaPrivate:'Privato', vaRestricted:'Riservato',
                vaFree:'Gratuito', vaLow:'Basso', vaModerate:'Moderato', vaExpensive:'Caro',
                vaMultiLevel:'Multipiano', vaStreetLevel:'A raso', vaStreetLevelCovered:'A raso, coperto', vaUnderground:'Sotterraneo',
                vaCash:'Contanti', vaChecks:'Assegni', vaCredit:'Carta di credito', vaDebitCard:'Carta di debito',
                vaDigitalWallet:'Portafoglio digitale', vaElectronicPass:'Pass elettronico', vaMembership:'Abbonamento',
                vaParkingApp:'App di parcheggio', vaPermit:'Permesso', vaPrepaid:'Prepagato', vaSmsCall:'SMS/Chiamata',
                vaAirConditioning:'Aria condizionata', vaCreditCards:'Accetta carte di credito', vaCurbsidePickup:'Ritiro al marciapiede',
                vaDeliveries:'Consegne', vaDrivethrough:'Drive-through', vaOutsideSeating:'Posti all’aperto',
                vaParkingForCustomers:'Parcheggio clienti', vaReservations:'Prenotazioni', vaRestrooms:'Servizi igienici',
                vaTakeAway:'Da asporto', vaValletService:'Servizio parcheggiatore', vaWheelchairAccessible:'Accessibile in sedia a rotelle',
                vaWiFi:'Wi-Fi', vaAirportShuttle:'Navetta aeroporto', vaCarpoolParking:'Posti car pooling', vaCarWash:'Autolavaggio',
                vaCovered:'Coperto', vaDisabilityParking:'Posti per disabili', vaOnSiteAttendant:'Personale sul posto', vaParkAndRide:'Parcheggio di scambio',
                vaSecurity:'Sorveglianza', vaValet:'Parcheggiatore (valet)', vaEvChargingStation:'Colonnine di ricarica', vaUnknown:'Sconosciuto',
            },
            'pt-BR': {
                tabTitle:'WME POI Event Updater',
                panelTitle:'POI Event Updater', chooseFile:'📂 Escolher um arquivo',
                chooseFileTitle:'Carregar uma planilha .xlsx do seu computador',
                selectSheet:'A aba da planilha a aplicar',
                fabTitle:'POI Event Updater — mostrar a janela',
                fabTitleOn:'POI Event Updater — ocultar a janela',
                btnApplyNone:'Nada marcado', btnApplyOne:'Aplicar 1 linha',
                btnApplyN:(n)=>`Aplicar as ${n} linhas marcadas`,
                btnApplyTitle:'Gravar as linhas marcadas no editor. Nada é salvo: você revisará no WME.',
                btnExportTitle:'Salvar o relatório desta aplicação',
                footerHelpVide:'Nada é gravado no mapa até você clicar em Aplicar.',
                guideFichier:'Escolha a planilha do evento.',
                guideFichierSuite:'Depois você escolherá a aba e revisará cada linha antes de aplicar.',
                guideOnglet:'Escolha a aba a aplicar.',
                guideOngletSuite:'Uma aba por evento. A aba «fora do evento» devolve os locais ao estado normal.',
                colSelect:'Aplicar', colEtat:'Estado',
                dropLigne1:'📄 Solte uma planilha aqui',
                dropLigne2:'ou clique para escolher',
                dropTitre:'Solte um arquivo .xlsx em qualquer lugar desta janela, ou clique para escolher',
                dropRefus:(nom)=>`«${nom}» não é uma planilha do Excel: aqui só se carregam arquivos .xlsx e .xls.`,
                sbOuvrir:'Mostrar a janela', sbOuvrirTitre:'Abrir a janela de trabalho — é lá que se carrega uma planilha e se revisa antes de aplicar',
                sbIntro:'Atualiza em lote o nome, a descrição e os campos dos locais a partir de uma planilha: uma aba por evento.',
                sbAide:'Ajuda', historyVide:'Nenhuma planilha carregada ainda.',
                aideClasseurT:'A planilha',
                aideClasseur:'Uma aba por evento, uma linha por local. As colunas são lidas pelo cabeçalho: link permanente, nome, descrição e depois os campos do local (telefone, site, categorias, serviços, estacionamento…).\nUma célula vazia num campo do local não pede nada. Uma descrição vazia APAGA a do local: é assim que uma aba «fora do evento» deixa os locais como eram.',
                aideRelireT:'Revisar e aplicar',
                aideRelire:'Cada linha mostra o antes (riscado) e o depois. O que muda vem marcado; editar um valor na prévia decide a caixa de novo.\nAplicar grava no editor sem salvar: revise no mapa e depois clique em Salvar. Os locais com falha continuam marcados, para tentar de novo.',
                aideEtatsT:'Estados de uma linha',
                aideEtats:'Número verde: campos a gravar.\n-n laranja: aplicar REMOVERIA n valores — nunca vem marcada.\nSaE: local bloqueado acima do seu nível, a alteração vai como sugestão.\nL7: bloqueio da equipe, nada é gravado.\n?: local não encontrado.\n✔: gravado, aguardando ser salvo.',
                sbEcrit:'✍️ Aplicar grava no editor; o script nunca salva.',
                majDispo:(v)=>`Nova versão ${v} disponível.`, majInstaller:'Instalar',
                cancelTitle:'Interromper: o que já foi carregado é mantido',
                footerHelp:'Desmarque o que não quiser gravar. As linhas laranja REMOVEM valores: nunca vêm marcadas.',
                footerMasquees:(n)=>`${n} linha(s) marcada(s) estão ocultas pelo filtro — também serão gravadas.`,
                bilanPartiel:(n)=>`⚠️ ${n} local(is) receberam só parte dos valores — veja o relatório.`,
                bilanEchec:(n)=>`⚠️ ${n} local(is) não puderam ser processados — suas linhas continuam marcadas:`,
                bilanSae:(n)=>`⚠️ ${n} local(is) bloqueado(s) acima do seu nível: a alteração vai como sugestão (SaE).`,
                bilanNonEnregistre:(n)=>`${n} alteração(ões) adicionada(s) à pilha do WME — NADA FOI SALVO: revise e depois clique em Salvar no editor.`,
                bilanNonEnregistreSansCompte:'NADA FOI SALVO: revise e depois clique em Salvar no editor.',
                bilanErreurGenerale:(m)=>`✖ A gravação não pôde começar: ${m}`,
                echecLigne:(nom, statut, err)=>`${nom} — ${statut}${err ? ': ' + err : ''}`,
                occupeDepot:'Um carregamento ou uma gravação está em andamento: espere terminar antes de carregar outra planilha.',
                cbTitre:'Gravar esta linha no editor',
                cbFige:'Esta linha não pode ser gravada — veja o selo de estado',
                colNameTitle:'O nome a gravar. Editável antes de aplicar.',
                colDescTitle:'A descrição a gravar. Uma célula vazia APAGA a descrição do local.',
                descEffacee:'⚠ a descrição do local será apagada',
                nomEfface:'⚠ o nome do local seria apagado',
                triTitre:'Ordenar por esta coluna',
                badgePerteTitle:(n)=>`Aplicar REMOVERIA ${n} valor(es) do local`,
                badgeDiffTitle:(n)=>`${n} campo(s) a gravar`,
                badgeRienTitle:'Nada a gravar: o local já tem estes valores',
                badgeOffTitle:'Local não encontrado no editor: nada pode ser gravado',
                badgePoseTitle:'Gravado no editor — aguardando ser salvo',
                badgeSaeTitle:'Local bloqueado acima do seu nível: a alteração irá como sugestão (SaE)',
                lockHardTitle:'Bloqueio nível 7 (equipe Waze) — edição impossível',
                noFile:'Nenhum arquivo escolhido',
                historyTitle:'Arquivos recentes', histLoaded:'📂 Carregado:', histApplied:'✔ Aplicado:',
                histNeverApplied:'Nunca aplicado', clearHistoryTitle:'Limpar o histórico',
                filterPlaceholder:'🔍 Filtrar por nome…',
                colName:'Nome', colDesc:'Descrição',
                btnReduce:'Minimizar', btnRestore:'Restaurar', btnApply:'Aplicar', btnClose:'Fechar',
                btnDiffActive:'≠ Alterações', btnDiffAll:'≡ Tudo',
                tooltipDiffOn:'Mostrar só os locais a alterar', tooltipDiffOff:'Mostrar todos os locais',
                deplacerAide:'Mover a janela: arrastar, ou setas do teclado (Shift: passo grande). Clique duplo: posição padrão.',
                redimAide:'Redimensionar a janela: arrastar, ou setas do teclado (Shift: passo grande)',
                loadingPois:(n,total)=>`Carregando locais… ${n} / ${total}`,
                applying:(n,total)=>`Aplicando… ${n} / ${total}`,
                cancelBtn:'Cancelar',
                noPoisLoaded:'✖ Nenhum local válido carregado',
                anomalies:(n)=>`⚠️ ${n} anomalia${n>1?'s':''} na planilha`,
                btnRetry:(n)=>`🔄 Tentar de novo (${n} loca${n>1?'is':'l'})`,
                successMsg:(n)=>`✔ ${n} loca${n>1?'is gravados':'l gravado'}`,
                btnExport:'📥 Exportar o relatório',
                sheetHeaderErr:'Cabeçalhos de coluna ausentes ou não reconhecidos — aba ignorada',
                sheetHeaderFallback:'cabeçalhos não reconhecidos: colunas lidas pela posição (A = link, B = nome, C = descrição)',
                sheetHeaderConflit:'há cabeçalhos reconhecidos fora do lugar A/B/C — aba ignorada: dê nome às três colunas (link, nome, descrição)',
                colonneDoublon:(lib, cols)=>`duas colunas para «${lib}» (${cols}): só a primeira é lida`,
                plusApplique:'✔ a gravar:', plusMain:'✋ a definir à mão:', plusRefus:'⚠ não reconhecido:',
                plusConforme:'· já correto:', plusAvant:'substitui:', plusRetire:'REMOVE',
                urlInvalid:'URL inválida',
                urlBadHost:'URL não reconhecida (deve ser waze.com ou beta.waze.com/…/editor)',
                urlNoEnv:'parâmetro env= ausente', urlBadLat:'lat= ausente ou inválido',
                urlBadLon:'lon= ausente ou inválido', urlBadZoom:'zoomLevel= ausente ou inválido',
                urlNoVenues:'venues= ausente', urlNoVid:'não foi possível ler o identificador do local',
                urlAutreEnv:(env, courant)=>`link de outro servidor (env=${env}, o editor está em ${courant}): o local não será encontrado`,
                urlPlusieursLieux:'vários locais em venues=: só o primeiro é lido',
                nameEmpty:'nome vazio', dupRow:(a,b)=>`Linhas ${a} e ${b}: link duplicado`,
                rowLabel:'linha',
                reportHeaders:['Link permanente','Nome antes','Nome depois','Descrição antes','Descrição depois','Estado','Campos não gravados','Erro'],
                statusApplied:'✔ Aplicado', statusPartial:'⚠ Parcial', statusTimeout:'✖ Não encontrado', statusErreur:'✖ Erro',
                masterCbTitle:'Marcar / desmarcar tudo',
                poiCount:(n)=>`${n} loca${n>1?'is':'l'}`,
                layerOffMsg:'⚠️ A camada «Locais» está desligada e não pôde ser ligada: ative-a (Camadas > Locais) e escolha a aba de novo.',
                xlsxMissing:'⚠️ Não foi possível carregar a biblioteca do Excel (SheetJS). Verifique a conexão ou permita cdn.sheetjs.com e recarregue a página (F5).',
                locateTitle:'Centralizar o mapa neste local',
                valOui:'sim', valNon:'não',
                chPerm:'Link permanente', chName:'Nome', chDesc:'Descrição', chAliases:'Nomes alternativos',
                chPhone:'Telefone', chUrl:'Site', chServices:'Serviços', chCategories:'Categorias',
                chParkingType:'Tipo de estacionamento', chHasTBR:'Tipo variável', chCostType:'Tarifa',
                chPaymentType:'Formas de pagamento', chParkingServices:'Serviços do estacionamento', chLotType:'Localização',
                chSpots:'Número de vagas', chCanExit:'Saída quando fechado',
                chHours:'Horário', chAddress:'Endereço', chEntryPoints:'Pontos de entrada', chOperator:'Operador do estacionamento',
                chGoogleName:'Nome no Google', chGoogleCategory:'Categoria no Google', chGooglePosition:'Posição no Google',
                moHours:'o WME espera faixas de horário, não uma frase',
                moAddress:'o WME espera um número e uma rua do seu próprio modelo',
                moEntryPoints:'são pontos no mapa, não texto',
                moOperator:'lista fechada no WME, cujas chaves não foram levantadas',
                moGoogleName:'não é um campo do WME', moGoogleCategory:'não é um campo do WME', moGooglePosition:'não é um campo do WME',
                vaPublic:'Público', vaPrivate:'Privado', vaRestricted:'Restrito',
                vaFree:'Gratuito', vaLow:'Baixo', vaModerate:'Moderado', vaExpensive:'Caro',
                vaMultiLevel:'Vários andares', vaStreetLevel:'Ao ar livre', vaStreetLevelCovered:'Ao ar livre, coberto', vaUnderground:'Subterrâneo',
                vaCash:'Dinheiro', vaChecks:'Cheques', vaCredit:'Cartão de crédito', vaDebitCard:'Cartão de débito',
                vaDigitalWallet:'Carteira digital', vaElectronicPass:'Passe eletrônico', vaMembership:'Mensalidade',
                vaParkingApp:'Aplicativo', vaPermit:'Autorização', vaPrepaid:'Pré-pago', vaSmsCall:'SMS/Ligação',
                vaAirConditioning:'Ar-condicionado', vaCreditCards:'Aceita cartão de crédito', vaCurbsidePickup:'Retirada na calçada',
                vaDeliveries:'Entregas', vaDrivethrough:'Drive-thru', vaOutsideSeating:'Mesas ao ar livre',
                vaParkingForCustomers:'Estacionamento para clientes', vaReservations:'Reservas', vaRestrooms:'Banheiros',
                vaTakeAway:'Para viagem', vaValletService:'Serviço de manobrista', vaWheelchairAccessible:'Acessível para cadeira de rodas',
                vaWiFi:'Wi-Fi', vaAirportShuttle:'Traslado para o aeroporto', vaCarpoolParking:'Vagas de carona', vaCarWash:'Lava-jato',
                vaCovered:'Coberto', vaDisabilityParking:'Vagas para deficientes', vaOnSiteAttendant:'Atendente no local', vaParkAndRide:'Estacionamento integrado',
                vaSecurity:'Segurança', vaValet:'Manobrista', vaEvChargingStation:'Carregador de veículos elétricos', vaUnknown:'Desconhecido',
            },
            'pt-PT': {
                tabTitle:'WME POI Event Updater',
                panelTitle:'POI Event Updater', chooseFile:'📂 Escolher um ficheiro',
                chooseFileTitle:'Carregar um livro .xlsx a partir do seu disco',
                selectSheet:'A folha do livro a aplicar',
                fabTitle:'POI Event Updater — mostrar a janela',
                fabTitleOn:'POI Event Updater — ocultar a janela',
                btnApplyNone:'Nada assinalado', btnApplyOne:'Aplicar 1 linha',
                btnApplyN:(n)=>`Aplicar as ${n} linhas assinaladas`,
                btnApplyTitle:'Escrever as linhas assinaladas no editor. Nada é guardado: irá rever no WME.',
                btnExportTitle:'Guardar o relatório desta aplicação',
                footerHelpVide:'Nada é escrito no mapa enquanto não clicar em Aplicar.',
                guideFichier:'Escolha o livro do evento.',
                guideFichierSuite:'Depois escolherá a folha e reverá cada linha antes de aplicar.',
                guideOnglet:'Escolha a folha a aplicar.',
                guideOngletSuite:'Uma folha por evento. A folha «fora do evento» repõe os locais no estado habitual.',
                colSelect:'Aplicar', colEtat:'Estado',
                dropLigne1:'📄 Largue aqui um livro',
                dropLigne2:'ou clique para o escolher',
                dropTitre:'Largue um ficheiro .xlsx em qualquer ponto desta janela, ou clique para o escolher',
                dropRefus:(nom)=>`«${nom}» não é um livro do Excel: aqui só se carregam ficheiros .xlsx e .xls.`,
                sbOuvrir:'Mostrar a janela', sbOuvrirTitre:'Abrir a janela de trabalho — é aí que se carrega um livro e se revê antes de aplicar',
                sbIntro:'Atualiza em lote o nome, a descrição e os campos dos locais a partir de um livro: uma folha por evento.',
                sbAide:'Ajuda', historyVide:'Ainda não foi carregado nenhum livro.',
                aideClasseurT:'O livro',
                aideClasseur:'Uma folha por evento, uma linha por local. As colunas leem-se pelo cabeçalho: ligação permanente, nome, descrição e depois os campos do local (telefone, site, categorias, serviços, parque…).\nUma célula vazia num campo do local não pede nada. Uma descrição vazia APAGA a do local: é assim que uma folha «fora do evento» repõe os locais.',
                aideRelireT:'Rever e aplicar',
                aideRelire:'Cada linha mostra o antes (riscado) e o depois. O que muda vem assinalado; alterar um valor na pré-visualização decide de novo a caixa.\nAplicar escreve no editor sem guardar: reveja no mapa e depois clique em Guardar. Os locais com falha ficam assinalados, para tentar de novo.',
                aideEtatsT:'Estados de uma linha',
                aideEtats:'Número verde: campos a escrever.\n-n laranja: aplicar REMOVERIA n valores — nunca vem assinalada.\nSaE: local bloqueado acima do seu nível, a alteração segue como sugestão.\nL7: bloqueio do staff, nada é escrito.\n?: local não encontrado.\n✔: escrito, à espera de ser guardado.',
                sbEcrit:'✍️ Aplicar escreve no editor; o script nunca guarda.',
                majDispo:(v)=>`Nova versão ${v} disponível.`, majInstaller:'Instalar',
                cancelTitle:'Interromper: o que já foi carregado mantém-se',
                footerHelp:'Retire o visto ao que não quiser escrever. As linhas laranja REMOVEM valores: nunca vêm assinaladas.',
                footerMasquees:(n)=>`${n} linha(s) assinalada(s) estão ocultas pelo filtro — também serão escritas.`,
                bilanPartiel:(n)=>`⚠️ ${n} local(ais) só receberam parte dos valores — veja o relatório.`,
                bilanEchec:(n)=>`⚠️ ${n} local(ais) não puderam ser tratados — as suas linhas ficam assinaladas:`,
                bilanSae:(n)=>`⚠️ ${n} local(ais) bloqueado(s) acima do seu nível: a alteração segue como sugestão (SaE).`,
                bilanNonEnregistre:(n)=>`${n} alteração(ões) acrescentada(s) à pilha do WME — NADA FOI GUARDADO: reveja e depois clique em Guardar no editor.`,
                bilanNonEnregistreSansCompte:'NADA FOI GUARDADO: reveja e depois clique em Guardar no editor.',
                bilanErreurGenerale:(m)=>`✖ A escrita não pôde começar: ${m}`,
                echecLigne:(nom, statut, err)=>`${nom} — ${statut}${err ? ': ' + err : ''}`,
                occupeDepot:'Está em curso um carregamento ou uma escrita: aguarde o fim antes de carregar outro livro.',
                cbTitre:'Escrever esta linha no editor',
                cbFige:'Esta linha não pode ser escrita — veja o selo de estado',
                colNameTitle:'O nome a escrever. Editável antes de aplicar.',
                colDescTitle:'A descrição a escrever. Uma célula vazia APAGA a descrição do local.',
                descEffacee:'⚠ a descrição do local será apagada',
                nomEfface:'⚠ o nome do local seria apagado',
                triTitre:'Ordenar por esta coluna',
                badgePerteTitle:(n)=>`Aplicar REMOVERIA ${n} valor(es) ao local`,
                badgeDiffTitle:(n)=>`${n} campo(s) a escrever`,
                badgeRienTitle:'Nada a escrever: o local já tem estes valores',
                badgeOffTitle:'Local não encontrado no editor: nada pode ser escrito',
                badgePoseTitle:'Escrito no editor — à espera de ser guardado',
                badgeSaeTitle:'Local bloqueado acima do seu nível: a alteração seguirá como sugestão (SaE)',
                lockHardTitle:'Bloqueio de nível 7 (staff Waze) — edição impossível',
                noFile:'Nenhum ficheiro escolhido',
                historyTitle:'Ficheiros recentes', histLoaded:'📂 Carregado:', histApplied:'✔ Aplicado:',
                histNeverApplied:'Nunca aplicado', clearHistoryTitle:'Limpar o histórico',
                filterPlaceholder:'🔍 Filtrar por nome…',
                colName:'Nome', colDesc:'Descrição',
                btnReduce:'Minimizar', btnRestore:'Restaurar', btnApply:'Aplicar', btnClose:'Fechar',
                btnDiffActive:'≠ Alterações', btnDiffAll:'≡ Tudo',
                tooltipDiffOn:'Mostrar só os locais a alterar', tooltipDiffOff:'Mostrar todos os locais',
                deplacerAide:'Mover a janela: arrastar, ou setas do teclado (Shift: passo grande). Duplo clique: posição predefinida.',
                redimAide:'Redimensionar a janela: arrastar, ou setas do teclado (Shift: passo grande)',
                loadingPois:(n,total)=>`A carregar locais… ${n} / ${total}`,
                applying:(n,total)=>`A aplicar… ${n} / ${total}`,
                cancelBtn:'Cancelar',
                noPoisLoaded:'✖ Nenhum local válido carregado',
                anomalies:(n)=>`⚠️ ${n} anomalia${n>1?'s':''} no livro`,
                btnRetry:(n)=>`🔄 Tentar de novo (${n} loca${n>1?'is':'l'})`,
                successMsg:(n)=>`✔ ${n} loca${n>1?'is escritos':'l escrito'}`,
                btnExport:'📥 Exportar o relatório',
                sheetHeaderErr:'Cabeçalhos de coluna ausentes ou não reconhecidos — folha ignorada',
                sheetHeaderFallback:'cabeçalhos não reconhecidos: colunas lidas pela posição (A = ligação, B = nome, C = descrição)',
                sheetHeaderConflit:'há cabeçalhos reconhecidos fora do lugar A/B/C — folha ignorada: dê nome às três colunas (ligação, nome, descrição)',
                colonneDoublon:(lib, cols)=>`duas colunas para «${lib}» (${cols}): só a primeira é lida`,
                plusApplique:'✔ a escrever:', plusMain:'✋ a definir à mão:', plusRefus:'⚠ não reconhecido:',
                plusConforme:'· já conforme:', plusAvant:'substitui:', plusRetire:'REMOVE',
                urlInvalid:'URL inválido',
                urlBadHost:'URL não reconhecido (deve ser waze.com ou beta.waze.com/…/editor)',
                urlNoEnv:'parâmetro env= em falta', urlBadLat:'lat= em falta ou inválido',
                urlBadLon:'lon= em falta ou inválido', urlBadZoom:'zoomLevel= em falta ou inválido',
                urlNoVenues:'venues= em falta', urlNoVid:'não foi possível ler o identificador do local',
                urlAutreEnv:(env, courant)=>`ligação de outro servidor (env=${env}, o editor está em ${courant}): o local não será encontrado`,
                urlPlusieursLieux:'vários locais em venues=: só o primeiro é lido',
                nameEmpty:'nome vazio', dupRow:(a,b)=>`Linhas ${a} e ${b}: ligação duplicada`,
                rowLabel:'linha',
                reportHeaders:['Ligação permanente','Nome antes','Nome depois','Descrição antes','Descrição depois','Estado','Campos não escritos','Erro'],
                statusApplied:'✔ Aplicado', statusPartial:'⚠ Parcial', statusTimeout:'✖ Não encontrado', statusErreur:'✖ Erro',
                masterCbTitle:'Assinalar / retirar tudo',
                poiCount:(n)=>`${n} loca${n>1?'is':'l'}`,
                layerOffMsg:'⚠️ A camada «Locais» está desligada e não foi possível ligá-la: ative-a (Camadas > Locais) e escolha de novo a folha.',
                xlsxMissing:'⚠️ Não foi possível carregar a biblioteca do Excel (SheetJS). Verifique a ligação ou autorize cdn.sheetjs.com e recarregue a página (F5).',
                locateTitle:'Centrar o mapa neste local',
                valOui:'sim', valNon:'não',
                chPerm:'Ligação permanente', chName:'Nome', chDesc:'Descrição', chAliases:'Nomes alternativos',
                chPhone:'Telefone', chUrl:'Site', chServices:'Serviços', chCategories:'Categorias',
                chParkingType:'Tipo de parque', chHasTBR:'Tipo variável', chCostType:'Tarifa',
                chPaymentType:'Formas de pagamento', chParkingServices:'Serviços do parque', chLotType:'Localização',
                chSpots:'Número de lugares', chCanExit:'Saída quando fechado',
                chHours:'Horário', chAddress:'Morada', chEntryPoints:'Pontos de entrada', chOperator:'Operador do parque',
                chGoogleName:'Nome no Google', chGoogleCategory:'Categoria no Google', chGooglePosition:'Posição no Google',
                moHours:'o WME espera intervalos horários, não uma frase',
                moAddress:'o WME espera um número de porta e uma rua do seu próprio modelo',
                moEntryPoints:'são pontos no mapa, não texto',
                moOperator:'lista fechada no WME, cujas chaves não foram recolhidas',
                moGoogleName:'não é um campo do WME', moGoogleCategory:'não é um campo do WME', moGooglePosition:'não é um campo do WME',
                vaPublic:'Público', vaPrivate:'Privado', vaRestricted:'Restrito',
                vaFree:'Gratuito', vaLow:'Baixo', vaModerate:'Moderado', vaExpensive:'Caro',
                vaMultiLevel:'Vários pisos', vaStreetLevel:'Ao ar livre', vaStreetLevelCovered:'Ao ar livre, coberto', vaUnderground:'Subterrâneo',
                vaCash:'Numerário', vaChecks:'Cheques', vaCredit:'Cartão de crédito', vaDebitCard:'Cartão de débito',
                vaDigitalWallet:'Carteira digital', vaElectronicPass:'Via Verde / passe eletrónico', vaMembership:'Assinatura',
                vaParkingApp:'Aplicação', vaPermit:'Autorização', vaPrepaid:'Pré-pagamento', vaSmsCall:'SMS/Chamada',
                vaAirConditioning:'Ar condicionado', vaCreditCards:'Aceita cartões de crédito', vaCurbsidePickup:'Levantamento à porta',
                vaDeliveries:'Entregas', vaDrivethrough:'Drive', vaOutsideSeating:'Esplanada',
                vaParkingForCustomers:'Parque para clientes', vaReservations:'Reservas', vaRestrooms:'Casas de banho',
                vaTakeAway:'Take-away', vaValletService:'Serviço de arrumador', vaWheelchairAccessible:'Acessível a cadeira de rodas',
                vaWiFi:'Wi-Fi', vaAirportShuttle:'Vaivém para o aeroporto', vaCarpoolParking:'Lugares de boleia partilhada', vaCarWash:'Lavagem auto',
                vaCovered:'Coberto', vaDisabilityParking:'Lugares para deficientes', vaOnSiteAttendant:'Funcionário no local', vaParkAndRide:'Parque dissuasor',
                vaSecurity:'Vigilância', vaValet:'Arrumador (valet)', vaEvChargingStation:'Postos de carregamento', vaUnknown:'Desconhecido',
            },
            he: {
                tabTitle:'WME POI Event Updater',
                panelTitle:'POI Event Updater', chooseFile:'📂 בחירת קובץ',
                chooseFileTitle:'טעינת חוברת ‎.xlsx מהמחשב',
                selectSheet:'הגיליון שיוחל',
                fabTitle:'POI Event Updater — הצגת החלון',
                fabTitleOn:'POI Event Updater — הסתרת החלון',
                btnApplyNone:'לא סומן דבר', btnApplyOne:'החלת שורה אחת',
                btnApplyN:(n)=>`החלת ${n} השורות המסומנות`,
                btnApplyTitle:'כתיבת השורות המסומנות לעורך. דבר אינו נשמר: תבדקו ב-WME.',
                btnExportTitle:'שמירת הדוח של החלה זו',
                footerHelpVide:'דבר אינו נכתב למפה עד שלוחצים על החלה.',
                guideFichier:'בחרו את חוברת האירוע.',
                guideFichierSuite:'לאחר מכן תבחרו גיליון ותבדקו כל שורה לפני ההחלה.',
                guideOnglet:'בחרו את הגיליון שיוחל.',
                guideOngletSuite:'גיליון לכל אירוע. גיליון „מחוץ לאירוע“ מחזיר את המקומות למצבם הרגיל.',
                colSelect:'החלה', colEtat:'מצב',
                dropLigne1:'📄 גררו לכאן חוברת',
                dropLigne2:'או לחצו כדי לבחור',
                dropTitre:'גררו קובץ ‎.xlsx לכל מקום בחלון, או לחצו כדי לבחור',
                dropRefus:(nom)=>`„${nom}“ אינו חוברת Excel: ניתן לטעון כאן רק קובצי ‎.xlsx ו-‎.xls.`,
                sbOuvrir:'הצגת החלון', sbOuvrirTitre:'פתיחת חלון העבודה — שם טוענים חוברת ובודקים לפני ההחלה',
                sbIntro:'עדכון מרוכז של שם, תיאור ושדות של מקומות מתוך חוברת: גיליון לכל אירוע.',
                sbAide:'עזרה', historyVide:'עדיין לא נטענה חוברת.',
                aideClasseurT:'החוברת',
                aideClasseur:'גיליון לכל אירוע, שורה לכל מקום. העמודות נקראות לפי הכותרת: קישור קבוע, שם, תיאור, ואחריהם שדות המקום (טלפון, אתר, קטגוריות, שירותים, חניה…).\nתא ריק בשדה של המקום אינו מבקש דבר. תיאור ריק מוחק את תיאור המקום: כך גיליון „מחוץ לאירוע“ מחזיר את המקומות למצבם.',
                aideRelireT:'בדיקה והחלה',
                aideRelire:'כל שורה מציגה את הלפני (מחוק) ואת האחרי. מה שמשתנה מסומן מראש; עריכת ערך בתצוגה המקדימה מחליטה מחדש על התיבה.\nהחלה כותבת לעורך בלי לשמור: בדקו במפה ואז לחצו על שמירה. מקומות שנכשלו נשארים מסומנים, לניסיון חוזר.',
                aideEtatsT:'מצבי שורה',
                aideEtats:'מספר ירוק: שדות לכתיבה.\n‎-n כתום: ההחלה תסיר n ערכים — לעולם לא מסומנת מראש.\nSaE: מקום נעול מעל הדרגה שלכם, השינוי יישלח כהצעה.\nL7: נעילת צוות, דבר אינו נכתב.\n?: המקום לא נמצא.\n✔: נכתב, ממתין לשמירה.',
                sbEcrit:'✍️ החלה כותבת לעורך; הסקריפט לעולם אינו שומר.',
                majDispo:(v)=>`גרסה חדשה ${v} זמינה.`, majInstaller:'התקנה',
                cancelTitle:'עצירה: מה שכבר נטען נשמר',
                footerHelp:'בטלו את הסימון של מה שאינכם רוצים לכתוב. שורות כתומות מסירות ערכים: הן לעולם אינן מסומנות מראש.',
                footerMasquees:(n)=>`${n} שורות מסומנות מוסתרות על ידי המסנן — גם הן ייכתבו.`,
                bilanPartiel:(n)=>`⚠️ ${n} מקומות קיבלו רק חלק מהערכים — ראו את הדוח.`,
                bilanEchec:(n)=>`⚠️ ${n} מקומות לא עובדו — השורות שלהם נשארות מסומנות:`,
                bilanSae:(n)=>`⚠️ ${n} מקומות נעולים מעל הדרגה שלכם: השינוי יישלח כהצעה (SaE).`,
                bilanNonEnregistre:(n)=>`${n} שינויים נוספו למחסנית של WME — דבר לא נשמר: בדקו, ואז לחצו על שמירה בעורך.`,
                bilanNonEnregistreSansCompte:'דבר לא נשמר: בדקו, ואז לחצו על שמירה בעורך.',
                bilanErreurGenerale:(m)=>`✖ הכתיבה לא יכלה להתחיל: ${m}`,
                echecLigne:(nom, statut, err)=>`${nom} — ${statut}${err ? ': ' + err : ''}`,
                occupeDepot:'טעינה או כתיבה מתבצעת כעת: המתינו לסיומה לפני טעינת חוברת אחרת.',
                cbTitre:'כתיבת שורה זו לעורך',
                cbFige:'לא ניתן לכתוב שורה זו — ראו את תג המצב',
                colNameTitle:'השם שייכתב. ניתן לערוך לפני ההחלה.',
                colDescTitle:'התיאור שייכתב. תא ריק מוחק את תיאור המקום.',
                descEffacee:'⚠ תיאור המקום יימחק',
                nomEfface:'⚠ שם המקום יימחק',
                triTitre:'מיון לפי עמודה זו',
                badgePerteTitle:(n)=>`ההחלה תסיר ${n} ערכים מהמקום`,
                badgeDiffTitle:(n)=>`${n} שדות לכתיבה`,
                badgeRienTitle:'אין מה לכתוב: למקום כבר יש ערכים אלה',
                badgeOffTitle:'המקום לא נמצא בעורך: לא ניתן לכתוב דבר',
                badgePoseTitle:'נכתב לעורך — ממתין לשמירה',
                badgeSaeTitle:'מקום נעול מעל הדרגה שלכם: השינוי יישלח כהצעה (SaE)',
                lockHardTitle:'נעילה בדרגה 7 (צוות Waze) — לא ניתן לערוך',
                noFile:'לא נבחר קובץ',
                historyTitle:'קבצים אחרונים', histLoaded:'📂 נטען:', histApplied:'✔ הוחל:',
                histNeverApplied:'לא הוחל מעולם', clearHistoryTitle:'ניקוי ההיסטוריה',
                filterPlaceholder:'🔍 סינון לפי שם…',
                colName:'שם', colDesc:'תיאור',
                btnReduce:'מזעור', btnRestore:'שחזור', btnApply:'החלה', btnClose:'סגירה',
                btnDiffActive:'≠ שינויים', btnDiffAll:'≡ הכול',
                tooltipDiffOn:'הצגת המקומות שישתנו בלבד', tooltipDiffOff:'הצגת כל המקומות',
                deplacerAide:'הזזת החלון: גרירה, או חיצי המקלדת (Shift: צעד גדול). לחיצה כפולה: מיקום ברירת המחדל.',
                redimAide:'שינוי גודל החלון: גרירה, או חיצי המקלדת (Shift: צעד גדול)',
                loadingPois:(n,total)=>`טעינת מקומות… ${n} / ${total}`,
                applying:(n,total)=>`מחיל… ${n} / ${total}`,
                cancelBtn:'ביטול',
                noPoisLoaded:'✖ לא נטען אף מקום תקין',
                anomalies:(n)=>`⚠️ ${n} חריגות בחוברת`,
                btnRetry:(n)=>`🔄 ניסיון חוזר (${n} מקומות)`,
                successMsg:(n)=>`✔ ${n} מקומות נכתבו`,
                btnExport:'📥 ייצוא הדוח',
                sheetHeaderErr:'כותרות העמודות חסרות או לא מזוהות — הגיליון נדלג',
                sheetHeaderFallback:'כותרות לא מזוהות: העמודות נקראות לפי מיקום (A = קישור, B = שם, C = תיאור)',
                sheetHeaderConflit:'כותרות מזוהות אינן במקומן A/B/C — הגיליון נדלג: תנו שם לשלוש העמודות (קישור, שם, תיאור)',
                colonneDoublon:(lib, cols)=>`שתי עמודות עבור „${lib}“ (${cols}): רק הראשונה נקראת`,
                plusApplique:'✔ לכתיבה:', plusMain:'✋ להזנה ידנית:', plusRefus:'⚠ לא מזוהה:',
                plusConforme:'· כבר תקין:', plusAvant:'מחליף:', plusRetire:'מסיר',
                urlInvalid:'כתובת לא תקינה',
                urlBadHost:'כתובת לא מזוהה (חייבת להיות waze.com או beta.waze.com/…/editor)',
                urlNoEnv:'הפרמטר env= חסר', urlBadLat:'lat= חסר או לא תקין',
                urlBadLon:'lon= חסר או לא תקין', urlBadZoom:'zoomLevel= חסר או לא תקין',
                urlNoVenues:'venues= חסר', urlNoVid:'לא ניתן לקרוא את מזהה המקום',
                urlAutreEnv:(env, courant)=>`קישור משרת אחר (env=${env}, העורך נמצא ב-${courant}): המקום לא יימצא`,
                urlPlusieursLieux:'כמה מקומות ב-venues=: רק הראשון נקרא',
                nameEmpty:'שם ריק', dupRow:(a,b)=>`שורות ${a} ו-${b}: קישור כפול`,
                rowLabel:'שורה',
                reportHeaders:['קישור קבוע','שם לפני','שם אחרי','תיאור לפני','תיאור אחרי','מצב','שדות שלא נכתבו','שגיאה'],
                statusApplied:'✔ הוחל', statusPartial:'⚠ חלקי', statusTimeout:'✖ לא נמצא', statusErreur:'✖ שגיאה',
                masterCbTitle:'סימון / ביטול הכול',
                poiCount:(n)=>`${n} מקומות`,
                layerOffMsg:'⚠️ שכבת „מקומות“ כבויה ולא ניתן היה להדליק אותה: הפעילו אותה (שכבות > מקומות) ובחרו שוב את הגיליון.',
                xlsxMissing:'⚠️ לא ניתן היה לטעון את ספריית Excel ‏(SheetJS). בדקו את החיבור או אפשרו את cdn.sheetjs.com, ואז טענו מחדש את הדף (F5).',
                locateTitle:'מרכוז המפה על מקום זה',
                valOui:'כן', valNon:'לא',
                chPerm:'קישור קבוע', chName:'שם', chDesc:'תיאור', chAliases:'שמות חלופיים',
                chPhone:'טלפון', chUrl:'אתר', chServices:'שירותים', chCategories:'קטגוריות',
                chParkingType:'סוג חניה', chHasTBR:'סוג משתנה', chCostType:'תעריף',
                chPaymentType:'אמצעי תשלום', chParkingServices:'שירותי החניה', chLotType:'מיקום',
                chSpots:'מספר מקומות חניה', chCanExit:'יציאה כשסגור',
                chHours:'שעות פתיחה', chAddress:'כתובת', chEntryPoints:'נקודות כניסה', chOperator:'מפעיל החניה',
                chGoogleName:'שם ב-Google', chGoogleCategory:'קטגוריה ב-Google', chGooglePosition:'מיקום ב-Google',
                moHours:'WME מצפה לחלונות זמן, לא למשפט',
                moAddress:'WME מצפה למספר בית ולרחוב מהמודל שלו',
                moEntryPoints:'אלה נקודות במפה, לא טקסט',
                moOperator:'רשימה סגורה ב-WME, שמפתחותיה לא נאספו',
                moGoogleName:'אינו שדה של WME', moGoogleCategory:'אינו שדה של WME', moGooglePosition:'אינו שדה של WME',
                vaPublic:'ציבורית', vaPrivate:'פרטית', vaRestricted:'מוגבלת',
                vaFree:'חינם', vaLow:'נמוך', vaModerate:'בינוני', vaExpensive:'יקר',
                vaMultiLevel:'רב-קומתי', vaStreetLevel:'במפלס הרחוב', vaStreetLevelCovered:'במפלס הרחוב, מקורה', vaUnderground:'תת-קרקעי',
                vaCash:'מזומן', vaChecks:'המחאות', vaCredit:'כרטיס אשראי', vaDebitCard:'כרטיס חיוב',
                vaDigitalWallet:'ארנק דיגיטלי', vaElectronicPass:'כרטיס אלקטרוני', vaMembership:'מנוי',
                vaParkingApp:'אפליקציית חניה', vaPermit:'היתר', vaPrepaid:'תשלום מראש', vaSmsCall:'SMS/שיחה',
                vaAirConditioning:'מיזוג אוויר', vaCreditCards:'מקבל כרטיסי אשראי', vaCurbsidePickup:'איסוף מהמדרכה',
                vaDeliveries:'משלוחים', vaDrivethrough:'דרייב-אין', vaOutsideSeating:'ישיבה בחוץ',
                vaParkingForCustomers:'חניה ללקוחות', vaReservations:'הזמנות', vaRestrooms:'שירותים',
                vaTakeAway:'טייק-אוויי', vaValletService:'שירות ואלה', vaWheelchairAccessible:'נגיש לכיסא גלגלים',
                vaWiFi:'Wi-Fi', vaAirportShuttle:'הסעה לשדה התעופה', vaCarpoolParking:'חניית קארפול', vaCarWash:'שטיפת רכב',
                vaCovered:'מקורה', vaDisabilityParking:'חניית נכים', vaOnSiteAttendant:'נציג במקום', vaParkAndRide:'חנה וסע',
                vaSecurity:'אבטחה', vaValet:'ואלה', vaEvChargingStation:'עמדת טעינה לרכב חשמלי', vaUnknown:'לא ידוע',
            },
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
    /** La locale des dates, des nombres et du tri : celle du SCRIPT, pas du navigateur. */
    function localeDuScript() {
        return { fr: 'fr-FR', en: 'en-GB', de: 'de-DE', es: 'es-ES', it: 'it-IT',
                 'pt-BR': 'pt-BR', 'pt-PT': 'pt-PT', he: 'he-IL' }[_peuLang] || 'en-GB';
    }
    function formatDateTime(iso) {
        if (!iso) return '—';
        const d = new Date(iso);
        const loc = localeDuScript();
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
    /* Posé par initScript : l'historique vit dans le panneau latéral, et la
       pose, qui a lieu dans la fenêtre, doit pouvoir le rafraîchir. */
    let _rafraichirHistorique = () => {};
    // ────────────────────────────────────────────────────────────────────────

    /**
     * ECHAPPE UNE DONNEE AVANT DE L'INSERER DANS DU HTML.
     *
     * ⚠️⚠️ LES CINQ CARACTERES, PAS TROIS. Oublier l'apostrophe et le chevron
     *    fermant suffit a faire sortir une valeur de son attribut : un nom de
     *    lieu venu d'un classeur est une donnee EXTERNE, et ce script en pose
     *    dans des title= a chaque ligne.
     */
    // ==== banc:esc ====
    // Extrait par les bancs qui rendent du HTML : une copie dans un banc
    // laisserait passer une régression de l'échappement (mesuré le 25/09/2026 :
    // `return String(v)` gardait banc-ligne et banc-coque au vert).
    function esc(v) {
        return String(v === undefined || v === null ? '' : v)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }
    // ==== /banc:esc ====

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
/* ⭐ CHARTE COMMUNE (25/09/2026) : #2196f3 pour les titres, les accents et le
   rail des interrupteurs ; #1976d2 pour ce qui est PLEIN avec du texte blanc
   (4,60:1 — le #2196f3 n'y donne que 3,12:1, sous le seuil AA). */
:root {
    --peu-blue:    #2196f3;
    --peu-blue-plein: #1976d2;
    --peu-blue-dk: #1565c0;
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
#peu-fab-btn.peu-fab-on { box-shadow: 0 0 0 2px var(--peu-blue, #2196f3), 0 2px 6px rgba(0,0,0,.3); }
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
    background: linear-gradient(135deg, var(--peu-blue-plein, #1976d2) 0%, #0d47a1 100%);
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
.peu-header-left .peu-icone { flex-shrink: 0; border-radius: 4px; }
/* ⚠️ Plus d'opacité .6 : 2,30:1 sur le dégradé, illisible (WCAG 1.4.3). */
.peu-header-version { font-size: .833em; color: #fff; font-weight: 400; flex-shrink: 0; }
/* ⚠️ Le contour bleu du focus sur l'en-tête bleu : 1,32:1, invisible (1.4.11). */
.peu-header :focus-visible, .peu-header:focus-visible { outline-color: #fff !important; }
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
/* La pilule de la charte : radius 50px, 3px 10px, 600 11px. */
.peu-btn {
    display: inline-flex; align-items: center; justify-content: center; gap: 5px;
    padding: 3px 10px; border: none; border-radius: 50px;
    font-family: inherit; font-size: .917em; font-weight: 600;
    height: auto; min-height: 0; line-height: 1.35;
    cursor: pointer; white-space: nowrap;
    transition: filter .15s, transform .1s, background .15s;
}
.peu-btn:active:not(:disabled) { transform: scale(.97); }
.peu-btn-primary { background: var(--peu-blue-plein, #1976d2);   color: #fff; }
.peu-btn-neutral { background: var(--peu-border, #dde3ea); color: var(--peu-text, #2d3748); }
.peu-btn-primary:hover:not(:disabled) { background: var(--peu-blue-dk, #1565c0); color: #fff; }
.peu-btn-neutral:hover:not(:disabled) { filter: brightness(.95); }
/* Un filtre actif se marque d'un contour : un seul bouton PLEIN par écran. */
.peu-btn-actif { box-shadow: inset 0 0 0 2px var(--peu-blue-plein, #1976d2); color: #0d47a1; }
.peu-btn:disabled { opacity: .45; cursor: not-allowed; }
.peu-btn-sm { font-size: .833em; }
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
    outline: none; border-color: var(--peu-blue, #2196f3);
    box-shadow: 0 0 0 3px rgba(33,150,243,.18);
}
.peu-input:disabled, .peu-textarea:disabled { background: #f1f3f6; color: var(--peu-grey, #9e9e9e); }
/* Un effacement se lit DANS le champ : le texte d'attente dit ce qui partira. */
.peu-efface { border-color: #e0a060; }
.peu-efface::placeholder { color: #b34700; font-style: italic; opacity: 1; }
.peu-select { width: auto; padding: .2em .4em; }

/* La case vit DANS un label : toute la zone devient cliquable. */
.peu-check { display: inline-flex; align-items: center; cursor: pointer; }
.peu-check input, .peu-checkbox {
    width: 15px; height: 15px; margin: 0;
    cursor: pointer; accent-color: var(--peu-blue-plein, #1976d2);
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
    background: var(--peu-bg, #f5f7f9); color: var(--peu-blue-dk, #1565c0);
    font-size: .833em; font-weight: 700; text-transform: uppercase; letter-spacing: .04em;
    padding: 7px 6px; border-bottom: 2px solid var(--peu-blue, #2196f3);
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
.peu-row-posee    > td:first-child { border-inline-start-color: var(--peu-blue, #2196f3); }
.peu-row-unloaded > td:first-child { border-inline-start-color: var(--peu-grey, #9e9e9e); }
.peu-row-unloaded > td { background: #f5f5f5; color: var(--peu-grey, #9e9e9e); }

/* ⚠️ #6b6b6b et non #8a8a8a : 3,45:1 sur blanc, sous le seuil AA (1.4.3). */
.peu-cell-old {
    color: #6b6b6b; font-size: .909em; line-height: 1.4; min-height: 14px;
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
/* ⚠️ Teintes foncées recalculées pour 4,5:1 sur leur fond (WCAG 1.4.3) :
   -n, SaE et = étaient à 3,46, 2,49 et 3,79 (audit du 25/09/2026). */
.peu-badge-diff  { background: #e8f5e9; color: #1b5e20; cursor: help; }
.peu-badge-perte { background: #fff3e0; color: #a33c00; cursor: help; }
.peu-badge-lock  { background: #ffebee; color: #c62828; cursor: help; }
.peu-badge-sae   { background: #fff8e1; color: #7a4f00; cursor: help; }
.peu-badge-off   { background: #eceff1; color: #37474f; cursor: help; }
.peu-badge-ok    { background: #eceff1; color: #455a64; cursor: help; }

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
/* Ce que la valeur remplace, en clair : une infobulle ne s'atteint ni au
   clavier, ni au doigt. */
.peu-pastille-avant { font-style: italic; opacity: .9; }
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
.peu-error-title { font-weight: 700; display: list-item; margin-bottom: 4px; cursor: pointer; }

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
    border-radius: 50%; background: var(--peu-blue-plein, #1976d2); color: #fff;
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
    border-color: var(--peu-blue, #2196f3); color: var(--peu-blue-dk, #1565c0);
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
.peu-progress-bar { display: block; height: 100%; width: 0; background: var(--peu-blue, #2196f3);
                    border-radius: 3px; transition: width .15s linear; }
.peu-prog-b { display: flex; align-items: center; gap: 6px; }
.peu-prog-d { flex: 1 1 auto; font-size: .909em; font-variant-numeric: tabular-nums; }

/* ----------------------------------------------------------------------
   LE PANNEAU LATERAL — les REGLAGES et l'HISTORIQUE, jamais le travail.
   ⚠️ Le panneau fait disparaitre son contenu des qu'on selectionne un objet
      sur la carte : on ne peut pas y travailler.
   ---------------------------------------------------------------------- */
/* Valeurs de la charte, relevées dans WME sur WCT, WJN et WRP le 25/09/2026.
   ⚠️ En px et non en em : le panneau ne suit pas la densité de la fenêtre. */
.peu-container {
    padding: 10px 12px; font-family: 'Rubik','Open Sans',sans-serif;
    font-size: 12px; color: var(--peu-text, #2d3748);
}
.peu-container h2 {
    display: flex; align-items: center; gap: 6px;
    font-size: 13px; font-weight: 700; color: var(--peu-blue, #2196f3); margin: 0 0 8px;
}
.peu-container h2 img { flex-shrink: 0; border-radius: 4px; }
.peu-container h2 span { font-size: 11px; font-weight: 400; color: #9e9e9e; }
.peu-hint { font-size: 11px; color: var(--peu-text2, #566372); line-height: 1.6; margin: 0 0 8px; }
.peu-sec {
    font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .05em;
    color: var(--peu-blue, #2196f3); margin: 14px 0 6px;
    display: flex; align-items: center; gap: 5px;
}
.peu-sb-maj { margin: 0 0 8px; padding: 5px 8px; border-radius: 8px; background: #ffebee; color: #c62828; font-size: 11px; font-weight: 600; }
.peu-sb-maj a { color: #c62828; }
.peu-help-section { border: 1px solid var(--peu-border, #dde3ea); border-radius: 8px; margin-bottom: 4px; overflow: hidden; }
.peu-help-hdr {
    display: flex; align-items: center; justify-content: space-between; width: 100%;
    height: auto; min-height: 0; margin: 0; border: none; font-family: inherit; text-align: start;
    padding: 5px 9px; font-size: 11px; font-weight: 700; cursor: pointer;
    background: var(--peu-bg, #f5f7f9); color: var(--peu-text, #2d3748); user-select: none;
}
.peu-help-hdr.on { color: var(--peu-blue-dk, #1565c0); background: #e3f2fd; }
.peu-help-hdr:hover { background: #eef4fb; }
.peu-help-body { padding: 7px 9px; font-size: 11px; line-height: 1.5; color: var(--peu-text, #2d3748); white-space: pre-line; }
.peu-sb-foot {
    margin: 12px 0 0; padding-top: 10px; border-top: 1px solid var(--peu-border, #dde3ea);
    font-size: 11px; line-height: 1.6; text-align: center;
}
.peu-sb-foot a { color: var(--peu-blue, #2196f3); }
.peu-sb-note { color: #9e9e9e; }
.peu-hist-row {
    margin-bottom: 5px; padding: 5px 7px; background: var(--peu-bg, #f5f7f9);
    border-radius: 4px; border-inline-start: 3px solid var(--peu-blue, #2196f3);
    font-size: 11px;
}
.peu-hist-name { font-weight: 700; color: var(--peu-blue-dk, #1565c0);
                 white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.peu-hist-meta { font-size: 11px; color: var(--peu-text2, #566372); margin-top: 1px; }

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
    outline: 2px solid var(--peu-blue-plein, #1976d2); outline-offset: 1px; border-radius: 3px;
}
/* Pendant une pose ou un balayage, le tableau est inerte, et cela se voit. */
#peu-overlay.peu-occupe .peu-table { opacity: .7; }

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

    /* ⚠️ CE QUE WME PEUT PORTER SANS QU'ON PUISSE LE DEMANDER. Un tarif jamais
       renseigné vaut `UNKNOWN` : l'aperçu l'affichait brut (« remplace :
       UNKNOWN », essai dans WME du 25/09/2026). On le TRADUIT à l'affichage, on
       ne l'accepte pas en entrée — demander « inconnu » n'a pas de sens. */
    const VALEURS_LUES_SEULEMENT = ['UNKNOWN'];

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
    /* En mode async, le SDK rend une promesse : deux fichiers ouverts coup sur
       coup partagent le même chargement au lieu d'interroger WME deux fois. */
    let _categoriesEnCours = null;
    async function chargerCategories(sdk) {
        const table = VALEURS_WME.categories;
        if (Object.keys(table).length) return Object.keys(table).length;
        if (!_categoriesEnCours) {
            _categoriesEnCours = (async () => {
                (await sdk.DataModel.Venues.getAllVenueCategories() || []).forEach(c => {
                    if (c && c.id) table[c.id] = [c.localizedName || c.id];
                });
                return Object.keys(table).length;
            })().finally(() => { _categoriesEnCours = null; });
        }
        return _categoriesEnCours;
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

        /* ⚠️⚠️ LE REPLI SE REFUSE QUAND LES EN-TÊTES DISENT AUTRE CHOSE QUE A/B/C.
           « Lien WME | Description | Nom » : une seule en-tête non reconnue, et le
           repli lisait la DESCRIPTION comme le NOM, sur tout l'onglet, chaque
           ligne cochée d'office (audit du 25/09/2026). Les deux en-têtes
           reconnues disaient pourtant où elles étaient.
           ⚠️ Cette garde ne joue QU'EN REPLI : un classeur lu par ses en-têtes
              n'est jamais concerné. */
        const repli = { perm: 0, name: 1, desc: 2 };
        const conflit = Object.keys(parNom).some(cle => parNom[cle] !== repli[cle]);
        const troisEnTetes = [0, 1, 2].every(i => normalizeHeader(ligne[i]) !== '');
        return { columns: repli, byPosition: true, usable: troisEnTetes && !conflit, conflit: conflit };
    }

    /**
     * LES CHAMPS QUE PLUSIEURS COLONNES REVENDIQUENT.
     *
     * ⚠️ La première colonne gagne (`mapChamps`), et c'était dit nulle part : un
     *    classeur qui porte « Site » et « Website » se lisait par la première,
     *    et l'autre était ignorée sans un mot.
     *
     * @return {{cle: string, colonnes: number[]}[]}
     */
    function colonnesEnDouble(entetes) {
        const ligne = Array.isArray(entetes) ? entetes : [];
        const doublons = [];
        CHAMPS.forEach(champ => {
            const colonnes = [];
            ligne.forEach((e, i) => { if (champ.entetes.indexOf(normalizeHeader(e)) !== -1) colonnes.push(i); });
            if (colonnes.length > 1) doublons.push({ cle: champ.cle, colonnes: colonnes });
        });
        return doublons;
    }

    /** La lettre d'une colonne de tableur : 0 → A, 26 → AA. */
    function lettreDeColonne(i) {
        let s = '';
        for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
        return s;
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

    /* Le nom, la description et les noms alternatifs s'écrivent À PART (`ecrireHerite`),
       comme depuis la 0.1 — le reste par `ecrireSdk`.
       ⚠️ 0.54.00 : `ecrireHerite` passait par `UpdateObject` sur `W.model`, que Waze
          retire le 24/11/2026 ; il passe désormais lui aussi par `updateVenue` du SDK,
          qui accepte ces trois champs. Deux appels restent distincts : un refus sur
          le nom ne doit pas emporter les champs complémentaires sans le dire. */
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

    /* ⭐⭐⭐ UNE LISTE DONT L'ORDRE COMPTE. La PREMIÈRE catégorie d'un lieu est sa
       catégorie principale — celle qui donne l'icône et le type. Comparées
       triées, `RESTAURANT, CAFE` et `CAFE, RESTAURANT` passaient pour égales :
       l'aperçu disait « rien à faire », la relecture disait « posé », et la
       catégorie principale changeait en silence dès qu'une autre colonne de la
       ligne la faisait cocher (audit du 25/09/2026). */
    const LISTES_ORDONNEES = ['categories'];

    function comparerAuLieu(attributs, aPoser) {
        const identiques = [], differents = [];
        const memeValeur = (a, b, ordonnee) => {
            if (Array.isArray(a) || Array.isArray(b)) {
                const x = (a || []).slice(), y = (b || []).slice();
                if (!ordonnee) { x.sort(); y.sort(); }
                return x.length === y.length && x.every((v, i) => v === y[i]);
            }
            return a === b;
        };

        /* ⚠️ Le nom et la description ont leur comparaison à eux (les deux
           colonnes de l'aperçu). Les NOMS ALTERNATIFS, eux, se comparent ici :
           exclus, un ajout arrivait « = », non coché, et n'était jamais posé. */
        Object.keys(aPoser).forEach(cible => {
            if (cible === 'name' || cible === 'description') return;
            (memeValeur(valeurDuLieu(attributs, cible), aPoser[cible], LISTES_ORDONNEES.includes(cible))
                ? identiques : differents).push(cible);
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
    function cocherDOffice(nomChange, descriptionChange, champsDifferents, pertes, nomVide) {
        const differe = nomChange || descriptionChange || champsDifferents > 0;

        /* ⚠️ UN NOM VIDE NE SE POSE PAS PAR DÉFAUT : il effacerait le nom du
           lieu. La ligne reste proposée — la cocher reste possible — mais c'est
           un geste, jamais un défaut. */
        return differe && !pertes && !nomVide;
    }
    // ==== /banc:pose ====

    // ==== banc:apercu ====
    // Rendu pur : il ne touche qu'au document qu'on lui donne, pour être
    // éprouvable hors de WME (tools/banc-apercu.html).

    const majusculeInit = (c) => c.charAt(0).toUpperCase() + c.slice(1);

    /**
     * Ce qu'une clé du dictionnaire donne, ou rien si elle n'y est pas.
     * ⚠️ `t` rend la clé elle-même quand elle manque : c'est ce qui se teste.
     */
    function traduitOuRien(traduire, cle) {
        if (!traduire) return null;
        const tr = traduire(cle);
        return tr && tr !== cle ? tr : null;
    }

    /**
     * LE LIBELLÉ D'UN CHAMP, DANS LA LANGUE DU SCRIPT.
     * ⚠️ `CHAMPS.libelle` reste la donnée, en français : les bancs s'y fient, et
     *    c'est le repli quand aucune traduction n'est fournie. L'écran, lui,
     *    passe par le dictionnaire (`ch` + clé du champ).
     */
    function libelleChamp(champ, traduire) {
        if (!champ) return '';
        return traduitOuRien(traduire, 'ch' + majusculeInit(champ.cle)) || champ.libelle;
    }

    /** Le libellé d'un champ posé, retrouvé par sa cible. */
    function libelleDeCible(cible, traduire) {
        const champ = CHAMPS.find(c => c.cible === cible && c.pose);
        return champ ? libelleChamp(champ, traduire) : cible;
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
    function libelleDeValeur(cle, traduire) {
        /* La langue du script d'abord (`va` + clé WME en casse chameau) ; les
           catégories, que WME rend déjà traduites, n'ont pas de clé et
           tombent sur la table. */
        const tr = traduitOuRien(traduire, 'va' + String(cle).toLowerCase().split('_').map(majusculeInit).join(''));
        if (tr) return tr;
        for (const table of Object.values(VALEURS_WME)) {
            if (table[cle] && table[cle].length) {
                return table[cle][0];
            }
        }
        return cle;
    }

    /**
     * Ce qu'une valeur posée donne à lire : les listes se NOMMENT, elles ne se comptent pas.
     * @param traduire `t` dans le script ; absent (banc), oui/non restent en français.
     */
    function valeurLisible(v, traduire) {
        if (v === true) return traduire ? traduire('valOui') : 'oui';
        if (v === false) return traduire ? traduire('valNon') : 'non';
        if (Array.isArray(v)) {
            const noms = v.map((x) => libelleDeValeur(x, traduire));
            const texte = noms.join(', ');
            /* Au-delà de trois, on nomme les deux premières et on compte le reste :
               une pastille qui déborde ne se lit plus. */
            return texte.length <= 44 ? texte : noms.slice(0, 2).join(', ') + ' +' + (noms.length - 2);
        }
        const texte = libelleDeValeur(String(v), traduire);
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
                const el = pastille(perdu ? 'perte' : 'pose', libelleDeCible(cible, traduire),
                    valeurLisible(valeurs.aPoser[cible], traduire)
                    + (perdu ? ' — ' + traduire('plusRetire') + ' ' + valeurLisible(perdu, traduire) : ''));
                /* ⚠️ CE QUE LA VALEUR REMPLACE, EN CLAIR DANS LA PASTILLE : un tableau
                   écrase tout son contenu dans WME. En infobulle seulement, on
                   écrasait sans les voir les horaires ou les services posés par un
                   autre éditeur — et l'infobulle ne s'atteint ni au clavier, ni au
                   doigt (audit du 25/09/2026). */
                const avant = attributs ? valeurDuLieu(attributs, cible) : undefined;
                if (avant !== undefined && avant !== null && String(avant) !== ''
                    && !(Array.isArray(avant) && !avant.length)) {
                    const rem = doc.createElement('span');
                    rem.className = 'peu-pastille-avant';
                    rem.textContent = ' (' + traduire('plusAvant') + ' ' + valeurLisible(avant, traduire) + ')';
                    el.appendChild(rem);
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
                const champ = CHAMPS.find(c => c.cle === m.cle);
                const el = pastille('montre', champ ? libelleChamp(champ, traduire) : m.libelle, valeurLisible(m.valeur, traduire));
                if (m.motif) el.title = traduitOuRien(traduire, 'mo' + majusculeInit(m.cle)) || m.motif;
                zone.appendChild(el);
            });
        }
        if (valeurs.refus.length) {
            zone.appendChild(titre(traduire('plusRefus')));
            valeurs.refus.forEach(r => {
                const champRefuse = CHAMPS.find(c => c.libelle === r.libelle);
                zone.appendChild(pastille('refus', champRefuse ? libelleChamp(champRefuse, traduire) : r.libelle, '« ' + r.valeurs.join(' », « ') + ' »'));
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
        if (typeof pw.getWmeSdk !== 'function') {
            throw new Error('SDK de WME indisponible');
        }
        _sdk = pw.getWmeSdk({ scriptId: 'poi-event-updater', scriptName: 'WME POI Event Updater', mode: 'async' });
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
    /* Le rang de l'éditeur connecté, relu par le SDK à chaque ouverture d'aperçu
       (`lireRang`) : getLockStatus reste synchrone. Même échelle que l'ancien
       `W.loginManager` (0 = L1) — mesuré identique le 04/10/2026. */
    let _rangEditeur = 0;
    async function lireRang() {
        try { _rangEditeur = (await obtenirSdk().State.getUserInfo())?.rank ?? 0; } catch (e) { /* on garde le dernier */ }
    }
    function getLockStatus(venue) {
        if (!venue) return 'ok';
        const lockRank = venue.attributes.lockRank ?? 0; // 0=L1 … 5=L6, 6=L7 staff
        const userRank = _rangEditeur;
        if (lockRank >= 6) return 'hard';  // niveau 7 staff
        if (lockRank > userRank) return 'sae';
        return 'ok';
    }

    /** La date du jour en AAAA-MM-JJ : un nom de fichier ne se lit pas dans une langue. */
    function dateIso(d) {
        const z = (n) => String(n).padStart(2, '0');
        return d.getFullYear() + '-' + z(d.getMonth() + 1) + '-' + z(d.getDate());
    }

    /**
     * LE RAPPORT DE LA POSE, EN CLASSEUR.
     *
     * ⭐⭐⭐ IL DIT QUOI REPRENDRE. Le bilan renvoyait « voir le rapport » sur une
     *    pose partielle, et le rapport n'avait que cinq colonnes : ni le
     *    permalien, ni les champs non posés, ni le message d'erreur — `partial`
     *    y sortait brut. On ne savait pas quel champ reprendre avant
     *    d'enregistrer (audit du 25/09/2026).
     */
    function exportReport(eventName, results) {
        const wb = XLSX.utils.book_new();
        const cles = { applied: 'statusApplied', partial: 'statusPartial', timeout: 'statusTimeout', erreur: 'statusErreur' };
        const statut = (r) => t(cles[r.status] || 'statusErreur') + (r.verrou === 'sae' ? ' (SaE)' : '');
        const rows = results.map(r => [r.perm || '', r.oldName, r.newName, r.oldDesc, r.newDesc, statut(r),
            (r.manques || []).map((c) => libelleDeCible(c, t)).join(', '), r.erreur || '']);
        const ws = XLSX.utils.aoa_to_sheet([t('reportHeaders'), ...rows]);
        ws['!cols'] = [40, 30, 30, 45, 45, 18, 30, 30].map(w => ({wch: w}));

        const jour = dateIso(new Date());
        /* ⚠️ Excel refuse \ / ? * [ ] : dans un nom d'onglet, et le plafonne à 31. */
        XLSX.utils.book_append_sheet(wb, ws, (jour + ' ' + eventName).replace(/[\\/?*[\]:]/g, '-').slice(0, 31));
        XLSX.writeFile(wb, 'POI_Report_' + eventName.replace(/[\\/:*?"<>|\s]+/g, '_') + '_' + jour + '.xlsx');
    }
    // ────────────────────────────────────────────────────────────────────────
    /**
     * UN LIEU LU AU SDK, SOUS LA FORME QUE TOUT LE SCRIPT CONNAÎT : `{attributes, boite}`.
     *
     * ⭐⭐⭐ C'EST UNE PHOTOGRAPHIE, PAS L'OBJET VIVANT du modèle de WME (`W.model`,
     *    retiré le 24/11/2026). Elle ne suit pas les écritures : après une pose, le
     *    lieu se RELIT (`appliquerLignes`), sans quoi l'aperçu montrerait l'avant.
     * ⚠️ Les champs de parking n'existent pas dans le lieu du SDK : ils se lisent
     *    un par un (`DataModel.Venues.ParkingLot`) et se rangent dans
     *    `categoryAttributes.PARKING_LOT`, là où l'ancien modèle les tenait.
     * ✅ Mesuré le 04/10/2026 à Avignon TGV : 32 lieux, dont 9 parkings — AUCUN
     *    écart avec `W.model` sur les champs que le script compare.
     * `boite` : l'emprise du lieu en degrés, pour cadrer la carte.
     */
    async function lireLieu(vid) {
        const sdk = obtenirSdk();
        let v = null;
        try { v = await sdk.DataModel.Venues.getById({ venueId: vid }); } catch (e) { return null; }
        if (!v) return null;
        const attributs = Object.assign({}, v);
        delete attributs.geometry;
        if ((v.categories || []).includes('PARKING_LOT')) {
            const P = sdk.DataModel.Venues.ParkingLot, q = { venueId: vid };
            const lire = async (f) => { try { return await f(); } catch (e) { return undefined; } };
            const parking = {
                parkingType: await lire(() => P.getParkingLotType(q)),
                hasTBR: await lire(() => P.isLotTypeDependentOnDayTime(q)),
                costType: await lire(() => P.getCostType(q)),
                paymentType: await lire(() => P.getPaymentMethods(q)),
                lotType: await lire(() => P.getLotTypes(q)),
                estimatedNumberOfSpots: await lire(() => P.getEstimatedNumberOfSpots(q)),
                canExitWhileClosed: await lire(() => P.canExitWhileClosed(q)),
            };
            /* Une valeur que le SDK n'a pas su lire est ABSENTE, comme dans l'ancien modèle. */
            Object.keys(parking).forEach((k) => { if (parking[k] === undefined) delete parking[k]; });
            attributs.categoryAttributes = { PARKING_LOT: parking };
        }
        return { attributes: attributs, boite: boiteDeGeometrie(v.geometry) };
    }

    /** L'emprise d'une géométrie GeoJSON (point ou polygone), en degrés. */
    function boiteDeGeometrie(g) {
        const pts = [];
        const parcourir = (c) => {
            if (!Array.isArray(c)) return;
            if (typeof c[0] === 'number') { pts.push(c); return; }
            c.forEach(parcourir);
        };
        parcourir(g && g.coordinates);
        if (!pts.length) return null;
        return {
            left: Math.min(...pts.map((p) => p[0])), bottom: Math.min(...pts.map((p) => p[1])),
            right: Math.max(...pts.map((p) => p[0])), top: Math.max(...pts.map((p) => p[1])),
        };
    }

    async function centerAndLoad(permalink, vid, timeoutMs = 4000) {
        {
            const coords = parseLatLon(permalink);
            if (!coords) return null;
            const { lat, lon } = coords;

            // Zoom de préchargement volontairement large (16-17) : le lat/lon d'un
            // permalink cadre souvent la CARTE, pas le POI (ex. Fresnes 1 : venue à
            // 361 m du point du permalink). À zoom 19, un POI décalé tombe hors des
            // tuiles chargées et n'est jamais trouvé → « non chargé ». On suit le
            // zoomLevel du permalink en le plafonnant à 17. Au terme du balayage,
            // la carte est cadrée sur le PÉRIMÈTRE des lieux (cadrerSurLesLieux),
            // et non ramenée là où elle était avant.
            let z = parseInt(new URL(permalink).searchParams.get('zoomLevel'), 10);
            if (isNaN(z)) z = 17;
            z = Math.max(16, Math.min(z, 17));

            try {
                await obtenirSdk().Map.setMapCenter({ lonLat: { lon: lon, lat: lat }, zoomLevel: z });
            } catch (e) { return null; }

            const t0 = Date.now();
            for (;;) {
                // « Chargé » = présent dans le modèle (`lieuPret`), avec ou sans nom.
                // On n'exige PAS d'être éditable : un POI verrouillé au-dessus du rang
                // de l'éditeur est bien chargé (cf. getLockStatus).
                const v = await lireLieu(vid);
                if (lieuPret(v)) return v;
                if (Date.now() - t0 > timeoutMs) return null;
                await new Promise((r) => setTimeout(r, 80));
            }
        }
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
            let venue = await lireLieu(vid);
            if (!lieuPret(venue)) {
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
    async function cadrerSurLesLieux(venues) {
        const sdk = obtenirSdk();
        const boites = Object.keys(venues || {}).map((cle) => (venues[cle] && venues[cle].boite) || null);

        const u = unionDesBoites(boites);
        if (!u) return false;
        const centre = { lon: (u.left + u.right) / 2, lat: (u.bottom + u.top) / 2 };

        if (boiteSansEtendue(u)) {
            const z = await sdk.Map.getZoomLevel();
            await sdk.Map.setMapCenter({ lonLat: centre, zoomLevel: Math.min(Math.max(z, 17), 19) });

            return true;
        }

        {
            await sdk.Map.zoomToExtent({ bbox: [u.left, u.bottom, u.right, u.top] });
            /* ⚠️⚠️ UN PLAFOND, ET PAS SEULEMENT UN PLANCHER. Cadrer sur UN SEUL lieu
               colle la carte au sol : l'echelle tombe a deux metres, on ne voit
               plus ni la rue, ni les lieux voisins, ni ou l'on est. Le plancher
               protege des lieux disperses, le plafond du lieu unique — et le
               second se rencontre bien plus souvent que le premier. */
            const z = await sdk.Map.getZoomLevel();
            if (z < 12) await sdk.Map.setMapCenter({ lonLat: centre, zoomLevel: 12 });
            else if (z > 19) await sdk.Map.setMapCenter({ lonLat: centre, zoomLevel: 19 });

            return true;
        }
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
            + '<div class="peu-header" id="peu-header" tabindex="0" title="' + esc(t('deplacerAide')) + '">'
            +   '<div class="peu-header-left">'
            +     '<img class="peu-icone" src="' + ICONE + '" alt="" width="18" height="18">'
            +     '<span id="peu-titre">' + esc(t('panelTitle')) + '</span>'
            +     '<span class="peu-header-version">v' + esc(version) + '</span>'
            +   '</div>'
            +   '<div class="peu-header-btns">'
            +     '<button type="button" class="peu-btn-icon" id="peu-btn-replier" title="' + esc(t('btnReduce')) + '"'
            +       ' aria-label="' + esc(t('btnReduce')) + '" aria-expanded="true">-</button>'
            +     '<button type="button" class="peu-btn-icon" id="peu-btn-fermer" title="' + esc(t('btnClose')) + '"'
            +       ' aria-label="' + esc(t('btnClose')) + '">X</button>'
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
            + '<div id="peu-resize" tabindex="0" role="button" title="' + esc(t('redimAide')) + '"'
            +   ' aria-label="' + esc(t('redimAide')) + '"></div>';
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
        /* ⚠️ LA PERTE AVANT LE SaE : l'ordre inverse masquait la perte d'une ligne
           verrouillée au-dessus du rang — badge « SaE », ni « -n » ni liseré
           orange (audit du 25/09/2026). Le badge de la perte dit alors les deux. */
        if (infos.pertes > 0)   return 'perte';
        if (infos.verrou === 'sae')  return 'sae';
        if (infos.posee)        return 'posee';
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
        if (etat === 'perte') {
            const sae = infos.verrou === 'sae';
            return { classe: 'peu-badge-perte', texte: '-' + infos.pertes + (sae ? ' SaE' : ''),
                titre: t('badgePerteTitle', infos.pertes) + (sae ? ' — ' + t('badgeSaeTitle') : '') };
        }
        if (etat === 'posee')    return { classe: 'peu-badge-diff',  texte: '✔', titre: t('badgePoseTitle') };
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
        /* ⚠️ UN EFFACEMENT SE LIT DANS LE CHAMP, sans décocher : une cellule vide
           EFFACE la description (c'est voulu — l'onglet ordinaire remet les lieux
           à nu), mais l'écran la montrait comme une modification ordinaire. */
        const efface = (oui, cle) => oui ? ' placeholder="' + esc(t(cle)) + '"' : '';
        const effaceDesc = vue.descChange && vue.desc === '';

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
            +   '<input type="text" class="peu-input' + (vue.nomVide ? ' peu-efface' : '') + '" data-nom dir="auto"'
            +     (fige ? ' disabled' : '') + efface(vue.nomVide, 'nomEfface')
            +     ' value="' + esc(vue.nom) + '" title="' + esc(t('colNameTitle')) + '"'
            +     ' aria-label="' + esc(t('colName')) + '"></td>'
            + '<td>' + vieux(vue.ancienDesc, vue.descChange)
            +   '<textarea class="peu-textarea' + (effaceDesc ? ' peu-efface' : '') + '" data-desc rows="1" dir="auto"'
            +     (fige ? ' disabled' : '') + efface(effaceDesc, 'descEffacee')
            +     ' title="' + esc(t('colDescTitle')) + '" aria-label="' + esc(t('colDesc')) + '">'
            +     esc(vue.desc) + '</textarea></td>'
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
            + '<th class="sortable" data-tri="nom" tabindex="0" aria-sort="none" title="' + esc(t('triTitre')) + '">'
            +   esc(t('colName')) + '<i class="peu-sort-icon" aria-hidden="true">&#9650;</i></th>'
            + '<th class="sortable" data-tri="desc" tabindex="0" aria-sort="none" title="' + esc(t('triTitre')) + '">'
            +   esc(t('colDesc')) + '<i class="peu-sort-icon" aria-hidden="true">&#9650;</i></th>'
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
        wrap.innerHTML = '<button type="button" id="peu-fab-btn" title="' + esc(t('fabTitle')) + '"'
            + ' aria-label="' + esc(t('fabTitle')) + '">'
            + '<img src="' + ICONE + '" alt="" width="22" height="22" style="display:block">'
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
        btn.setAttribute('aria-label', btn.title);
        btn.setAttribute('aria-expanded', String(ouvert));

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
            memoriserGeometrie(ov);
        });

        /* ⭐ AU CLAVIER AUSSI (WCAG 2.5.7) : l'en-tête prend le focus, les flèches
           déplacent la fenêtre — seulement quand c'est LUI qui a le focus. */
        poignee.addEventListener('keydown', (e) => {
            if (e.target === poignee) deplacerAuClavier(ov, e, false);
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
            memoriserGeometrie(ov);
        });

        poignee.addEventListener('keydown', (e) => deplacerAuClavier(ov, e, true));
    }

    /**
     * MÉMORISE LA GÉOMÉTRIE DE LA FENÊTRE.
     *
     * ⚠️ REPLIÉE, ELLE NE MESURE QUE SON EN-TÊTE : on gardait alors une hauteur
     *    de 120 px, et la fenêtre dépliée au chargement suivant n'avait plus de
     *    place pour son tableau. Repliée, on garde la hauteur d'avant.
     */
    function memoriserGeometrie(ov) {
        const r = ov.getBoundingClientRect();
        const memo = lireGeometrie();
        const h = ov.classList.contains('peu-replie')
            ? (memo ? memo.h : Math.round(bornesCarte().bas - r.top))
            : Math.round(r.height);
        ecrireGeometrie({ x: Math.round(r.left), y: Math.round(r.top), w: Math.round(r.width), h: h });
    }

    /**
     * DÉPLACER OU REDIMENSIONNER AU CLAVIER : flèches, Maj pour un grand pas.
     *
     * ⚠️ LA PROPAGATION S'ARRÊTE ICI, ET SEULEMENT POUR LES FLÈCHES : WME les
     *    écoute pour faire glisser la carte, et les deux bougeraient ensemble.
     *    Toute autre touche passe — Ctrl+S compris.
     */
    function deplacerAuClavier(ov, e, redim) {
        const pas = e.shiftKey ? 64 : 16;
        const d = { ArrowLeft: [-pas, 0], ArrowRight: [pas, 0], ArrowUp: [0, -pas], ArrowDown: [0, pas] }[e.key];
        if (!d || e.ctrlKey || e.altKey || e.metaKey) return;
        e.preventDefault();
        e.stopPropagation();
        const r = ov.getBoundingClientRect();
        const bornes = bornesCarte();
        const g = bornerFenetre(redim
            ? { x: r.left, y: r.top, w: r.width + d[0], h: r.height + d[1] }
            : { x: r.left + d[0], y: r.top + d[1], w: r.width, h: r.height }, bornes);
        ov.style.left = g.x + 'px';
        ov.style.top = g.y + 'px';
        ov.style.right = 'auto';
        if (redim) { ov.style.width = g.w + 'px'; ov.style.height = g.h + 'px'; }
        ov.style.maxHeight = Math.max(120, bornes.bas - g.y) + 'px';
        memoriserGeometrie(ov);
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
        /* Une fenêtre NON modale : la carte reste utilisable à côté. */
        ov.setAttribute('role', 'dialog');
        ov.setAttribute('aria-labelledby', 'peu-titre');
        if (_peuLang === 'he') ov.dir = 'rtl';
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
        ov.querySelector('#peu-btn-appliquer').addEventListener('click', () => { appliquerLignes().catch(signalerErreur); });

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
                ouvrirApercu(e.target.value).catch(signalerErreur);
            }
        });
        ov.querySelector('#peu-btn-replier').addEventListener('click', () => {
            const replie = ov.classList.toggle('peu-replie');
            const b = ov.querySelector('#peu-btn-replier');
            b.textContent = replie ? '+' : '-';
            b.title = t(replie ? 'btnRestore' : 'btnReduce');
            b.setAttribute('aria-label', b.title);
            b.setAttribute('aria-expanded', String(!replie));
        });

        window.addEventListener('resize', () => {
            if (ov.classList.contains('peu-open')) placerFenetre(ov);
        });

        return ov;
    }

    function ouvrirOverlay() {
        const ov = construireOverlay();
        const dejaOuverte = ov.classList.contains('peu-open');
        ov.classList.add('peu-open');
        placerFenetre(ov);
        /* ⭐ RIEN DE CHARGE ⇒ ON DIT PAR OU COMMENCER. Une fenetre vide avec un
           bouton grise n'explique rien. */
        if (!ov.querySelector('#peu-body').children.length) {
            montrerGuide(poiData.length ? 'guideOnglet' : 'guideFichier',
                poiData.length ? 'guideOngletSuite' : 'guideFichierSuite');
        }
        /* ⚠️ LE FOCUS ENTRE DANS LA FENÊTRE À L'OUVERTURE (WCAG 2.4.3) — et
           seulement là : la rouvrir alors qu'elle l'est ne le déplace pas. */
        if (!dejaOuverte) {
            const cible = ov.querySelector(poiData.length ? '#peu-select-onglet' : '#peu-btn-fichier');
            if (cible && !cible.hidden && !cible.disabled) cible.focus({ preventScroll: true });
        }
        majFab();
    }

    function fermerOverlay() {
        const ov = document.getElementById('peu-overlay');
        if (ov) ov.classList.remove('peu-open');
        /* Le focus revient d'où l'on vient : le bouton de la carte. */
        const fab = document.getElementById('peu-fab-btn');
        if (ov && fab && ov.contains(document.activeElement)) fab.focus({ preventScroll: true });
        majFab();
    }

    function basculerOverlay() {
        const ov = document.getElementById('peu-overlay');
        if (ov && ov.classList.contains('peu-open')) fermerOverlay(); else ouvrirOverlay();
    }

    async function initScript() {
        let localeWme = '';
        try { localeWme = (await obtenirSdk().Settings.getLocale())?.localeCode || ''; } catch (e) { /* langue du navigateur */ }
        _peuLang = detectLang(localeWme);
        injectCSS();
        const { tabLabel, tabPane } = await obtenirSdk().Sidebar.registerScriptTab();
        // Icône (pin) à la place du nom, nom conservé en infobulle.
        tabLabel.textContent = '';
        tabLabel.style.display = 'flex';
        tabLabel.style.alignItems = 'center';
        tabLabel.style.justifyContent = 'center';
        tabLabel.style.height = '100%';
        const tabIcon = document.createElement('img');
        tabIcon.src = ICONE;
        tabIcon.alt = t('tabTitle');
        tabIcon.width = 18;
        tabIcon.height = 18;
        tabIcon.style.display = 'block';
        tabLabel.appendChild(tabIcon);
        tabLabel.title = t('tabTitle');
        /* L'onglet du SDK s'insère dans le panneau un peu après : on attend qu'il y soit (2 s au plus). */
        for (let i = 0; i < 40 && !tabPane.isConnected; i++) await new Promise((r) => setTimeout(r, 50));
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
        if (_peuLang === 'he') container.dir = 'rtl';

        /* ⭐ LA CHARTE COMMUNE (WCT, WJN, WRP, WDA) : l'icône du script, son nom et
           sa version en tête ; la pastille de nouvelle version ; une phrase qui
           dit à quoi il sert. Valeurs relevées dans WME le 25/09/2026. */
        const title = document.createElement('h2');
        title.innerHTML = '<img src="' + ICONE + '" alt="" width="18" height="18">'
            + esc(t('panelTitle')) + ' <span>v' + esc(PEU_VERSION) + '</span>';
        container.appendChild(title);

        const maj = document.createElement('p');
        maj.className = 'peu-sb-maj';
        maj.id = 'peu-sb-maj';
        maj.hidden = true;
        maj.innerHTML = '<span></span> <a href="#" data-maj>' + esc(t('majInstaller')) + '</a>';
        maj.querySelector('[data-maj]').addEventListener('click', (e) => {
            e.preventDefault();
            window.open(URL_INSTALLER, '_blank', 'noopener');
        });
        container.appendChild(maj);

        const intro = document.createElement('p');
        intro.className = 'peu-hint';
        intro.textContent = t('sbIntro');
        container.appendChild(intro);

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

        /* ⚠️ UN SEUL TITRE « FICHIERS RÉCENTS » : il s'affichait deux fois, le
           second portant la corbeille. La corbeille vit désormais dans le titre,
           et n'apparaît que s'il y a quelque chose à effacer. */
        const titreHist = document.createElement('div');
        titreHist.className = 'peu-sec';
        titreHist.innerHTML = '<span aria-hidden="true">🕘</span><span style="flex:1">' + esc(t('historyTitle')) + '</span>'
            + '<button type="button" class="peu-btn-center" data-raz hidden title="' + esc(t('clearHistoryTitle')) + '"'
            + ' aria-label="' + esc(t('clearHistoryTitle')) + '">🗑</button>';
        container.appendChild(titreHist);
        const btnRaz = titreHist.querySelector('[data-raz]');
        btnRaz.addEventListener('click', () => { saveHistory([]); renderHistory(); });

        const historyDiv = document.createElement('div');
        container.appendChild(historyDiv);

        /* ⭐ L'AIDE SE REPLIE (charte) : elle est là quand on la cherche, et ne
           pousse pas le reste hors de vue quand on ne la cherche pas. */
        const titreAide = document.createElement('div');
        titreAide.className = 'peu-sec';
        titreAide.innerHTML = '<span aria-hidden="true">❓</span>' + esc(t('sbAide'));
        container.appendChild(titreAide);
        [['aideClasseurT', 'aideClasseur'], ['aideRelireT', 'aideRelire'], ['aideEtatsT', 'aideEtats']].forEach(([cleT, cleC], i) => {
            const bloc = document.createElement('div');
            bloc.className = 'peu-help-section';
            bloc.innerHTML = '<button type="button" class="peu-help-hdr" aria-expanded="false" aria-controls="peu-aide-' + i + '">'
                + '<span>' + esc(t(cleT)) + '</span><span aria-hidden="true">▶</span></button>'
                + '<div class="peu-help-body" id="peu-aide-' + i + '" hidden></div>';
            bloc.querySelector('.peu-help-body').textContent = t(cleC);
            const hdr = bloc.querySelector('.peu-help-hdr');
            hdr.addEventListener('click', () => {
                const ouvert = hdr.getAttribute('aria-expanded') !== 'true';
                hdr.setAttribute('aria-expanded', String(ouvert));
                hdr.classList.toggle('on', ouvert);
                hdr.lastElementChild.textContent = ouvert ? '▼' : '▶';
                bloc.querySelector('.peu-help-body').hidden = !ouvert;
            });
            container.appendChild(bloc);
        });

        /* LE PIED : où parler du script, où l'installer, où lire son code — puis
           ce qu'il fait de la carte. ⚠️ PAS « ne modifie jamais la carte » comme
           les scripts voisins : celui-ci ÉCRIT, et il doit le dire. */
        const pied = document.createElement('p');
        pied.className = 'peu-sb-foot';
        pied.innerHTML = '💬 <a href="' + URL_DISCUSS + '" target="_blank" rel="noopener">Discuss</a>'
            + ' &nbsp;·&nbsp; 🔗 <a href="' + URL_GF + '" target="_blank" rel="noopener">GreasyFork</a>'
            + ' &nbsp;·&nbsp; <a href="' + URL_GH + '" target="_blank" rel="noopener">GitHub</a>'
            + '<br><span class="peu-sb-note"></span>';
        pied.querySelector('.peu-sb-note').textContent = t('sbEcrit');
        container.appendChild(pied);

        function renderHistory() {
            historyDiv.innerHTML = '';
            const history = getHistory();
            btnRaz.hidden = !history.length;
            if (!history.length) {
                const vide = document.createElement('div');
                vide.className = 'peu-hist-meta';
                vide.textContent = t('historyVide');
                historyDiv.appendChild(vide);
                return;
            }

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
        verifierMaj();

        _rafraichirHistorique = renderHistory;

        fileInput.addEventListener('change', e => {
            const file = e.target.files[0];
            /* ⚠️ ON REPART DE ZERO A CHAQUE FICHIER : laisser le bandeau de
               l'ancien classeur pendant qu'on en lit un autre, c'est afficher
               deux verites a la fois. L'aperçu et ses anomalies partent aussi. */
            poiData = [];
            _apercu = null;
            _anomalies = [];
            _fichierCourant = null;
            majStrip(null, [], 0);
            montrerGuide('guideFichier', 'guideFichierSuite');
            if (!file) return;
            const reader = new FileReader();
            reader.onload = async ev => {
                /* Le référentiel des catégories vient de l'éditeur, dans SA langue.
                   Un échec laisse la table vide : les catégories seront alors
                   refusées et signalées, jamais posées à l'aveugle. */
                try { await chargerCategories(obtenirSdk()); } catch (e) { /* signalé à la ligne */ }
                try {
                    lireClasseur(file.name, ev.target.result);
                } catch (err) {
                    /* ⚠️ UNE LECTURE QUI ECHOUE SE DIT, ET SE DIT LA OU L'ON REGARDE.
                       Un message pose dans un panneau que la carte fait disparaitre
                       ne serait lu par personne. */
                    ouvrirOverlay();
                    showValidationReport(['✖ ' + err.message]);
                } finally {
                    /* ⚠️ TOUJOURS, refus compris : sinon rechoisir le MÊME fichier,
                       corrigé, ne déclenche pas `change`, et rien n'est relu. */
                    fileInput.value = '';
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    /* ======================================================================
       LA PASTILLE DE NOUVELLE VERSION — au plus une vérification par 24 h
       ====================================================================== */
    const VER_RE = /^\d+(\.\d+)*$/;
    const MAJ_KEY = 'peu_maj', MAJ_DELAI = 864e5;

    /* Segment par segment, en nombres : en chaînes, « 0.9 » passerait pour plus
       récent que « 0.53.00 ». Un segment absent vaut zéro (0.52 = 0.52.00). */
    function majCmp(a, b) {
        const x = a.split('.').map(Number), y = b.split('.').map(Number);
        for (let i = 0; i < Math.max(x.length, y.length); i++) {
            const d = (x[i] || 0) - (y[i] || 0);
            if (d) return d;
        }
        return 0;
    }

    function montrerMaj(v) {
        const p = document.getElementById('peu-sb-maj');
        if (!p) return;
        p.querySelector('span').textContent = t('majDispo', v);
        p.hidden = false;
    }

    /**
     * ⚠️ PAR GM_xmlhttpRequest, PAS PAR fetch : la politique de sécurité de WME
     *    (connect-src) n'autorise pas GreasyFork depuis la page. Sans la
     *    permission (script injecté à la main), la vérification se tait.
     * ⚠️ Un 404 arrive AUSSI par onload : la page d'erreur ne se lit pas comme
     *    un en-tête de script.
     */
    function verifierMaj() {
        if (!VER_RE.test(PEU_VERSION) || typeof GM_xmlhttpRequest !== 'function') return;
        let memo = null;
        try { memo = JSON.parse(localStorage.getItem(MAJ_KEY) || 'null'); } catch (e) { /* rien de lisible */ }
        if (memo && Date.now() - memo.t < MAJ_DELAI) {
            if (memo.v && VER_RE.test(memo.v) && majCmp(PEU_VERSION, memo.v) < 0) montrerMaj(memo.v);
            return;
        }
        const retenir = (v) => { try { localStorage.setItem(MAJ_KEY, JSON.stringify({ t: Date.now(), v: v })); } catch (e) { /* tant pis */ } };
        GM_xmlhttpRequest({
            method: 'GET', url: URL_META, timeout: 10000, nocache: true,
            onload: (r) => {
                if (r.status < 200 || r.status >= 300) { retenir(null); return; }
                const m = (r.responseText || '').match(/^\/\/\s*@version\s+(\S+)/m);
                const v = m && VER_RE.test(m[1]) ? m[1] : null;
                retenir(v);
                if (v && majCmp(PEU_VERSION, v) < 0) montrerMaj(v);
            },
            onerror: () => {}, ontimeout: () => {},
        });
    }

    /** Le fichier dont l'aperçu est à l'écran — pour l'historique « appliqué ». */
    let _fichierCourant = null;

    /** L'environnement de la page (row, usa, il), s'il est dans l'adresse. */
    function envDeLaPage() {
        try { return new URL(location.href).searchParams.get('env'); } catch (e) { return null; }
    }

    /** Le libellé d'un champ, par sa clé. */
    function libelleDuChamp(cle) {
        const champ = CHAMPS.find(c => c.cle === cle);
        return champ ? libelleChamp(champ, t) : cle;
    }

    /**
     * LIT UN CLASSEUR : chaque onglet, chaque ligne, vingt règles.
     *
     * ⚠️ Toute ligne écartée se DIT dans le rapport d'anomalies : ni le tableau,
     *    qui ne montre que ce qui est retenu, ni la carte ne la montreraient.
     */
    function lireClasseur(nomFichier, donnees) {
        /* Les catégories sont chargées AVANT l'appel (`reader.onload`) : le SDK async rend une promesse. */
        const wb = XLSX.read(new Uint8Array(donnees), {type:'array'});
        const all = [];
        const warnings = [];
        const envCourant = envDeLaPage();

        wb.SheetNames.filter(n => n !== 'Config').forEach(sheet => {
            const sh = wb.Sheets[sheet];
            // Les colonnes se lisent par leur EN-TÊTE, avec repli sur les
            // positions A/B/C : l'ordre des colonnes cesse d'être un contrat
            // tacite, et une colonne ajoutée à droite ne décale plus rien.
            const rows = XLSX.utils.sheet_to_json(sh, {header:1, defval:''});
            const plan = mapColumns(rows[0]);
            if (!plan.usable) {
                warnings.push(`[${sheet}] ${t(plan.conflit ? 'sheetHeaderConflit' : 'sheetHeaderErr')}`);
                return;
            }
            if (plan.byPosition) warnings.push(`[${sheet}] ${t('sheetHeaderFallback')}`);
            colonnesEnDouble(rows[0]).forEach(d => warnings.push(`[${sheet}] `
                + t('colonneDoublon', libelleDuChamp(d.cle), d.colonnes.map(lettreDeColonne).join(', '))));
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
                /* ⚠️ EN TEXTE, SANS ESPACE AUTOUR : un nom lu brut gardait son
                   espace final — posé tel quel — et un nombre ne se comparait
                   jamais égal au texte du lieu. */
                const r = {
                    perm: String(valeur(plan.columns.perm)).trim(),
                    name: String(valeur(plan.columns.name)).trim(),
                    desc: String(valeur(plan.columns.desc)).trim()
                };
                if (!r.perm) return; // ligne vide ignorée silencieusement

                // 1. URL syntaxiquement valide ?
                let parsedUrl;
                try { parsedUrl = new URL(r.perm); } catch {
                    warnings.push(`[${sheet}] ${t('urlInvalid')} (${t('rowLabel')} ${rowNum})`);
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
                /* Gardée, mais dite : un lieu d'un autre serveur ne sera pas
                   trouvé, et « introuvable » n'en donnerait pas la cause. */
                if (envCourant && params.get('env') !== envCourant) {
                    warnings.push(`[${sheet}] ${t('urlAutreEnv', params.get('env'), envCourant)} (${t('rowLabel')} ${rowNum})`);
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
                if (venues.indexOf(',') !== -1) {
                    warnings.push(`[${sheet}] ${t('urlPlusieursLieux')} (${t('rowLabel')} ${rowNum})`);
                }

                // 4. Extraire l'ID du venue
                const vid = getVenueIdFromPermalink(r.perm);
                if (!vid) {
                    warnings.push(`[${sheet}] ${t('urlNoVid')} (${t('rowLabel')} ${rowNum})`);
                    return;
                }

                // 5. Nom vide ?
                if (!r.name) {
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
            ouvrirOverlay();
            montrerGuide('guideFichier', 'guideFichierSuite');
            showValidationReport(warnings.length ? warnings : [t('noPoisLoaded')]);
            return;
        }
        poiData = all;
        _fichierCourant = nomFichier;
        recordFileLoaded(nomFichier);
        _rafraichirHistorique();
        /* ⭐ LE BANDEAU DE LA FENETRE PORTE TOUT : le nom du classeur, la
           liste des onglets et le nombre de POI. Un seul endroit le dit,
           donc il n'y a plus deux comptes a garder d'accord.
           ⚠️ Les anomalies d'abord : le balayage qui suit les affiche. */
        ouvrirOverlay();
        showValidationReport(warnings);
        majStrip(nomFichier, Array.from(new Set(poiData.map((p) => p.event))), all.length);
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
        if (aDesOnglets) ouvrirApercu(onglets[0]).catch(signalerErreur);
    }

    /** Le corps de la fenetre, vide. */
    function corpsFenetre() {
        const ov = construireOverlay();
        const corps = ov.querySelector('#peu-body');
        corps.innerHTML = '';

        return corps;
    }

    /** Les anomalies du dernier classeur lu. Elles survivent au rendu du tableau. */
    let _anomalies = [];

    /**
     * LE RAPPORT D'ANOMALIES — dans la FENETRE, pas dans le panneau.
     *
     * ⭐ IL DIT CE QUI N'EST PAS ENTRE. Une ligne ecartee du classeur ne se
     *    voit nulle part ailleurs : ni dans le tableau, qui ne montre que ce
     *    qui est retenu, ni sur la carte. Sans ce rapport, le fichier parait
     *    complet et il ne l'est pas.
     *
     * ⚠️⚠️ IL VIVAIT DANS initScript, hors de portée du dépôt de fichier : lâcher
     *    un fichier qui n'est pas un classeur levait une ReferenceError, et rien
     *    ne s'affichait. Et le tableau l'effaçait en se rendant (audit du
     *    25/09/2026) : il est désormais redessiné à chaque rendu.
     */
    function showValidationReport(warnings) {
        _anomalies = (warnings || []).slice();
        const ov = document.getElementById('peu-overlay');
        const corps = ov ? ov.querySelector('#peu-body') : null;
        if (corps) rendreAnomalies(corps);
    }

    /** Le compte reste visible ; la liste se replie quand elle est longue. */
    function rendreAnomalies(corps) {
        const vieux = corps.querySelector('[data-rapport]');
        if (vieux) vieux.remove();
        if (!_anomalies.length) return;
        const bloc = document.createElement('details');
        bloc.className = 'peu-alert peu-alert-warn';
        bloc.setAttribute('data-rapport', '');
        bloc.open = _anomalies.length <= 3;
        const resume = document.createElement('summary');
        resume.className = 'peu-error-title';
        resume.textContent = t('anomalies', _anomalies.length);
        bloc.appendChild(resume);
        const ul = document.createElement('ul');
        _anomalies.forEach(w => {
            const li = document.createElement('li');
            li.textContent = w;
            ul.appendChild(li);
        });
        bloc.appendChild(ul);
        corps.prepend(bloc);
    }

    /** Un message qui passe, en tête du corps : un refus, une consigne. */
    function messagePassager(texte) {
        const ov = document.getElementById('peu-overlay');
        const corps = ov ? ov.querySelector('#peu-body') : null;
        if (!corps) return;
        const div = document.createElement('div');
        div.className = 'peu-alert peu-alert-warn';
        div.setAttribute('role', 'status');
        div.textContent = texte;
        corps.prepend(div);
        setTimeout(() => div.remove(), 6000);
    }

    /** Une erreur imprévue se DIT, dans la fenêtre — jamais seulement en console. */
    function signalerErreur(e) {
        console.error('[WPEU]', e);
        occuper(false);
        ouvrirOverlay();
        messagePassager('✖ ' + ((e && e.message) || String(e)));
    }

    /**
     * UN PRÉCHARGEMENT OU UNE POSE EST EN COURS.
     *
     * ⭐⭐⭐ RIEN NE L'EMPÊCHAIT DE RECOMMENCER. Après la pose, le bouton se
     *    réactivait sur les mêmes cases, toujours cochées : un clic de plus
     *    posait tout une seconde fois. Et pendant la pose, le menu d'onglet, le
     *    bouton de fichier et le dépôt restaient actifs — le bilan se rattachait
     *    alors au mauvais événement (audit du 25/09/2026).
     */
    let _occupe = false;
    function occuper(oui) {
        _occupe = oui;
        const ov = document.getElementById('peu-overlay');
        if (!ov) return;
        ov.classList.toggle('peu-occupe', oui);
        ['#peu-select-onglet', '#peu-btn-fichier'].forEach((sel) => {
            const el = ov.querySelector(sel);
            if (el) el.disabled = oui;
        });
        const table = ov.querySelector('#peu-body .peu-table');
        if (table) table.inert = oui;
        majPied();
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
        if (zone) zone.addEventListener('click', () => { if (_peuFileInput && !_occupe) _peuFileInput.click(); });
        rendreAnomalies(corps);
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

            /* ⚠️ PAS DE NOUVEAU CLASSEUR PENDANT UN BALAYAGE OU UNE POSE : le bilan
               se rattacherait à un aperçu qui n'est plus à l'écran. */
            if (_occupe) { messagePassager(t('occupeDepot')); return; }

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
        const cases = [...ov.querySelectorAll('#peu-body .peu-ligne .peu-checkbox:checked')];
        const nb = cases.length;
        /* ⚠️ UNE LIGNE COCHÉE PUIS MASQUÉE PAR LE FILTRE EST POSÉE QUAND MÊME :
           le compte du bouton est juste, mais il doit dire qu'une partie ne se
           voit plus. */
        const masquees = cases.filter((cb) => {
            const tr = cb.closest('.peu-ligne');
            return tr && tr.style.display === 'none';
        }).length;

        btn.textContent = libelleAppliquer(nb);
        btn.disabled = nb === 0 || _occupe;
        if (exp) exp.hidden = !_apercu || !_apercu.resultats;
        if (aide) aide.textContent = masquees ? t('footerMasquees', masquees) : t(_apercu ? 'footerHelp' : 'footerHelpVide');
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
            const idx = Number(tr.dataset.idx);
            const p = _apercu.pois[idx];
            const vue = _apercu.vues ? _apercu.vues[idx] : null;

            return {
                idx: idx,
                vid: getVenueIdFromPermalink(p.perm),
                perm: p.perm,
                nom: tr.querySelector('[data-nom]').value.trim(),
                desc: tr.querySelector('[data-desc]').value.trim(),
                valeurs: p.valeurs,
                verrou: vue ? vue.verrou : 'ok',
            };
        });
    }

    // ==== banc:poser ====
    // Extrait par tools/banc-application.mjs. Tout ce qui touche WME passe par
    // `env` : le banc y met des bouchons qui REFUSENT ce qu'ils ne connaissent
    // pas, l'éditeur y met W.model et le SDK (`environnementDePose`).

    /**
     * POSE UN LIEU.
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
    /**
     * UN LIEU EST-IL CHARGÉ ?
     *
     * ⚠️ PAS « A-T-IL UN NOM ». Le test exigeait un nom : un lieu SANS nom était
     *    déclaré introuvable, sa ligne figée, et quatre secondes perdues — alors
     *    que donner un nom à un lieu qui n'en a pas est justement l'usage de la
     *    colonne Nom (audit du 25/09/2026).
     */
    function lieuPret(v) {
        return !!(v && v.attributes);
    }

    async function poserUnLieu(item, env) {
        /* ⭐ La carte ne bouge que si le lieu n'est pas déjà en mémoire : le
           préchargement vient presque toujours de le charger. */
        let venue = await env.lieuCharge(item.vid);
        if (!lieuPret(venue)) venue = await env.charger(item.perm, item.vid);
        if (!lieuPret(venue)) {
            return { echec: true, resultat: { oldName: '', newName: item.nom, oldDesc: '', newDesc: item.desc,
                status: 'timeout', poses: [], manques: [] } };
        }

        const oldName = venue.attributes.name || '';
        const oldDesc = venue.attributes.description || '';
        /* ⚠️ Les noms alternatifs viennent du classeur s'il en porte, sinon on
           REPASSE ceux du lieu tels quels : ne jamais les perdre au passage. */
        const aPoser = (item.valeurs && item.valeurs.aPoser) || {};
        /* `await` : en mode async, un refus du SDK est une promesse rejetée — qui remonte
           alors à poserLesLignes et classe la ligne en erreur, au lieu de se perdre. */
        await env.ecrireHerite(venue, {
            id: venue.attributes.id,
            name: item.nom,
            description: item.desc,
            aliases: aPoser.aliases || venue.attributes.aliases || [],
        });

        const { maj, ignores } = construireMaj(aPoser);
        let erreurSdk = null;
        if (Object.keys(maj).length) {
            /* `await` : en mode async, un refus du SDK est une promesse rejetée, que le `try` ne verrait pas sans lui. */
            try { await env.ecrireSdk(item.vid, maj); } catch (e) { erreurSdk = e.message; }
        }

        /* ⭐⭐⭐ LA RELECTURE FAIT FOI, POUR TOUT — le nom et la description compris.
           Ils n'étaient pas relus : « appliqué » voulait dire « UpdateObject n'a
           pas levé », y compris sur un lieu verrouillé au-dessus du rang, dont
           rien ne dit encore ce que WME fait de la modification.
           ⚠️ Un champ à poser jamais envoyé (hors liste blanche) reste un MANQUE :
              la relecture le trouve absent, et `ignores` le rappelle. */
        const relu = (await env.relire(item.vid)) || {};
        const manques = [];
        if ((relu.name || '') !== item.nom) manques.push('name');
        if ((relu.description || '') !== item.desc) manques.push('description');
        const ecarts = comparerAuLieu(relu, aPoser);
        ecarts.differents.concat(ignores).forEach((c) => { if (!manques.includes(c)) manques.push(c); });

        return {
            echec: false,
            resultat: {
                oldName: oldName, newName: item.nom,
                oldDesc: oldDesc, newDesc: item.desc,
                status: manques.length ? 'partial' : 'applied',
                poses: ecarts.identiques,
                manques: manques,
                erreur: erreurSdk,
            },
        };
    }

    /**
     * POSE TOUTES LES LIGNES — UNE EXCEPTION N'EN ARRÊTE AUCUNE.
     *
     * ⚠️⚠️ SANS CE try, UNE EXCEPTION AU LIEU k ARRÊTAIT LA BOUCLE : k-1
     *    modifications restaient dans la pile de WME, la barre restait figée, et
     *    aucun bilan ne s'affichait. Cause possible, non mesurée : un lieu
     *    verrouillé au-dessus du rang que WME refuse (audit du 25/09/2026).
     *
     * @return {{resultats: Object[], aReprendre: Object[]}} `aReprendre` : les
     *    lignes à relancer — introuvables ou en erreur.
     */
    async function poserLesLignes(items, env, surAvance) {
        const resultats = [], aReprendre = [];
        for (let i = 0; i < items.length; i++) {
            const it = items[i];
            let r;
            try {
                r = await poserUnLieu(it, env);
            } catch (e) {
                r = { echec: true, resultat: { oldName: '', newName: it.nom, oldDesc: '', newDesc: it.desc,
                    status: 'erreur', erreur: (e && e.message) || String(e), poses: [], manques: [] } };
            }
            r.resultat.vid = it.vid;
            r.resultat.perm = it.perm;
            r.resultat.verrou = it.verrou || 'ok';
            resultats.push(r.resultat);
            if (r.echec) aReprendre.push(it);
            if (surAvance) surAvance(i + 1, items.length);
        }
        return { resultats: resultats, aReprendre: aReprendre };
    }

    /**
     * LA CASE D'UNE LIGNE APRÈS UNE RETOUCHE DANS L'APERÇU.
     *
     * ⭐⭐⭐⭐ UNE RETOUCHE ÉTAIT JETÉE EN SILENCE. La case n'était décidée qu'au
     *    premier rendu : on corrigeait le nom d'une ligne « = », on cliquait
     *    Appliquer, la correction n'était pas posée — et l'écran continuait de
     *    l'afficher. La 0.47 re-cochait ; la 0.52 l'avait perdu à l'extraction.
     *
     * ⚠️ UNE LIGNE QUI RETIRE N'EST JAMAIS TOUCHÉE : ni cochée seule, ni
     *    décochée si l'éditeur l'a cochée — c'est son geste, pas le nôtre.
     */
    function caseApresRetouche(coche, vue) {
        if (vue.pertes) return coche;
        return cocherDOffice(vue.nomChange, vue.descChange, vue.champsDiff, 0, vue.nomVide);
    }
    // ==== /banc:poser ====

    /** Ce que la pose demande a WME, dans l'editeur — tout par le SDK (0.54.00). */
    function environnementDePose() {
        return {
            lieuCharge: (vid) => lireLieu(vid),
            charger: (perm, vid) => centerAndLoad(perm, vid, 3000),
            /* ⚠️ La description VIDE part aussi : c'est ainsi qu'un onglet « hors
               événement » remet un lieu à nu (251 descriptions vides sur 342 dans ACO). */
            ecrireHerite: (venue, champs) => obtenirSdk().DataModel.Venues.updateVenue({
                venueId: venue.attributes.id,
                name: champs.name, description: champs.description, aliases: champs.aliases,
            }),
            ecrireSdk: (vid, maj) => obtenirSdk().DataModel.Venues.updateVenue(Object.assign({ venueId: vid }, maj)),
            relire: async (vid) => { const v = await lireLieu(vid); return v ? v.attributes : null; },
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
        div.setAttribute('role', 'status');
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
        if (_occupe) return;
        const pois = poiData.filter((p) => p.event === eventName);
        if (!pois.length) { montrerGuide('guideOnglet', 'guideOngletSuite'); return; }

        /* ⚠️⚠️ LE CALQUE « LIEUX » DOIT ETRE ALLUME, SANS QUOI RIEN N'EXISTE. WME ne
           charge pas les lieux d'un calque eteint : le prechargement ne trouverait
           AUCUN POI et l'apercu annoncerait que tout est introuvable — un diagnostic
           faux, sur un fichier juste. On l'allume donc, on attend que WME serve
           les lieux, et l'on RELIT : un clic qui n'a rien allumé ne vaut pas un
           calque allumé. */
        /* `null` = on ne sait pas lire : on ne bloque pas sur une incertitude. */
        const lieuxVisibles = async () => {
            try { return await obtenirSdk().Map.isLayerVisible({ layerName: 'venues' }); } catch (e) { return null; }
        };
        if ((await lieuxVisibles()) === false) {
            const bascule = document.querySelector('#layer-switcher-group_places');
            if (bascule) {
                bascule.click();
                await new Promise((r) => setTimeout(r, 1500));
            }
            if (!bascule || (await lieuxVisibles()) === false) {
                ouvrirOverlay();
                montrerGuide('guideOnglet', 'guideOngletSuite');
                messagePassager(t('layerOffMsg'));
                return;
            }
        }

        ouvrirOverlay();
        occuper(true);
        await lireRang();
        let venueMap = null;
        const annule = { cancelled: false };
        try {
            const corps = corpsFenetre();
            rendreAnomalies(corps);
            const prog = poserProgression(corps, t('loadingPois', 0, pois.length), pois.length,
                () => { annule.cancelled = true; });
            venueMap = await preloadVenues(pois, (n, total) => {
                prog.avance(n, total);
                prog.libelle(t('loadingPois', n, total));
            }, annule);
            prog.retirer();
        } finally {
            occuper(false);
        }
        if (annule.cancelled) { montrerGuide('guideOnglet', 'guideOngletSuite'); return; }

        // ⭐ ON RESTE SUR LE PERIMETRE qu'on vient de parcourir.
        try { await cadrerSurLesLieux(venueMap); } catch (e) { /* la carte reste où elle est */ }

        _apercu = { eventName: eventName, pois: pois, venueMap: venueMap, resultats: null, vues: null };
        /* ⚠️ LE MENU DIT CE QUI EST OUVERT. Ouvert par un autre chemin que lui,
           il afficherait autre chose que ce que le tableau montre. */
        const menu = document.getElementById('peu-select-onglet');
        if (menu && menu.value !== eventName) menu.value = eventName;
        rendreTableau();
    }

    /**
     * Ce qu'il y a a dire d'une ligne, avant de la rendre.
     *
     * @param saisie `{nom, desc}` retouchés dans l'aperçu ; sinon, ceux du classeur.
     *    ⚠️ Le lieu est relu VIVANT : après une pose, il porte déjà les valeurs
     *    posées, et une nouvelle retouche se compare à elles.
     */
    function vueDeLaLigne(p, idx, venueMap, saisie) {
        const vid = getVenueIdFromPermalink(p.perm);
        const venue = venueMap[vid];
        const attributs = venue ? venue.attributes : null;
        const aPoser = (p.valeurs && p.valeurs.aPoser) || null;
        const verrou = venue ? getLockStatus(venue) : 'ok';
        const nom = saisie ? saisie.nom : p.name;
        const desc = saisie ? saisie.desc : p.desc;
        const nomChange = !!attributs && nom !== (attributs.name || '');

        return {
            idx: idx,
            nom: nom, desc: desc,
            ancienNom: attributs ? (attributs.name || '') : '',
            ancienDesc: attributs ? (attributs.description || '') : '',
            nomChange: nomChange,
            nomVide: nomChange && nom === '',
            descChange: !!attributs && desc !== (attributs.description || ''),
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
        _apercu.vues = vues;

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
        [...tbody.querySelectorAll('.peu-ligne:not(.peu-comp)')].forEach((tr) => {
            const v = vues[Number(tr.dataset.idx)];
            const cb = tr.querySelector('.peu-checkbox');
            if (cb && !cb.disabled) {
                cb.checked = cocherDOffice(v.nomChange, v.descChange, v.champsDiff, v.pertes, v.nomVide);
            }
        });

        brancherTableau(table, barre, vues);
        rendreAnomalies(corps);
        majPied();
    }

    /** Une ligne redessinée sur place après une retouche ou une pose — sans rien reconstruire. */
    function majLigne(tr, v) {
        const etat = etatDeLaLigne(v);
        const b = badgeDeLigne(etat, v);
        tr.className = 'peu-ligne peu-row-' + etat;
        const badge = tr.querySelector('.peu-badge');
        if (badge) { badge.className = 'peu-badge ' + b.classe; badge.textContent = b.texte; badge.title = b.titre; }
        const comp = tr.nextElementSibling;
        if (comp && comp.classList.contains('peu-comp')) comp.className = 'peu-ligne peu-comp peu-row-' + etat;
        const vieux = tr.querySelectorAll('.peu-cell-old');
        if (vieux[0]) vieux[0].classList.toggle('changed', v.nomChange);
        if (vieux[1]) vieux[1].classList.toggle('changed', v.descChange);
        const efface = (el, oui, cle) => {
            if (!el) return;
            el.classList.toggle('peu-efface', oui);
            el.placeholder = oui ? t(cle) : '';
        };
        efface(tr.querySelector('[data-nom]'), v.nomVide, 'nomEfface');
        efface(tr.querySelector('[data-desc]'), v.descChange && v.desc === '', 'descEffacee');
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

        /* ⭐⭐⭐⭐ UNE RETOUCHE RE-DÉCIDE LA CASE, LE BADGE ET LE FILTRE — voir
           `caseApresRetouche`. Sans cet écouteur, corriger un nom ne changeait
           rien à ce qui serait posé, et l'écran continuait de l'afficher. */
        table.addEventListener('input', (e) => {
            if (!e.target.matches('[data-nom], [data-desc]')) return;
            const tr = e.target.closest('.peu-ligne');
            const idx = Number(tr.dataset.idx);
            const v = vueDeLaLigne(_apercu.pois[idx], idx, _apercu.venueMap, {
                nom: tr.querySelector('[data-nom]').value.trim(),
                desc: tr.querySelector('[data-desc]').value.trim(),
            });
            vues[idx] = v;
            majLigne(tr, v);
            const cb = tr.querySelector('.peu-checkbox');
            if (cb && !cb.disabled) cb.checked = caseApresRetouche(cb.checked, v);
            majPied();
        });

        /* LE TRI PROMIS PAR L'EN-TÊTE. Chaque ligne voyage avec la ligne de ses
           champs du lot D2 : détachées, on lirait les champs d'un lieu sous le
           nom d'un autre. */
        let tri = null;
        const collation = new Intl.Collator(localeDuScript(), { sensitivity: 'base', numeric: true });
        table.querySelectorAll('th[data-tri]').forEach((th) => {
            const trier = () => {
                const col = th.dataset.tri;
                const sens = tri && tri.col === col && tri.sens === 1 ? -1 : 1;
                tri = { col: col, sens: sens };
                const tbody = table.querySelector('tbody');
                const paires = [...tbody.querySelectorAll('.peu-ligne:not(.peu-comp)')].map((tr) => {
                    const suivante = tr.nextElementSibling;
                    return {
                        tr: tr,
                        comp: suivante && suivante.classList.contains('peu-comp') ? suivante : null,
                        cle: tr.querySelector(col === 'nom' ? '[data-nom]' : '[data-desc]').value,
                    };
                });
                paires.sort((a, b) => sens * collation.compare(a.cle, b.cle));
                paires.forEach((p) => { tbody.appendChild(p.tr); if (p.comp) tbody.appendChild(p.comp); });
                table.querySelectorAll('th[data-tri]').forEach((h) => {
                    h.classList.remove('sort-asc', 'sort-desc');
                    h.setAttribute('aria-sort', 'none');
                });
                th.classList.add(sens === 1 ? 'sort-asc' : 'sort-desc');
                th.setAttribute('aria-sort', sens === 1 ? 'ascending' : 'descending');
                th.querySelector('.peu-sort-icon').textContent = sens === 1 ? '▲' : '▼';
            };
            th.addEventListener('click', trier);
            th.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); trier(); }
            });
        });

        const filtre = barre.querySelector('[data-filtre]');
        const btnEcarts = barre.querySelector('[data-ecarts]');
        let ecartsSeuls = false;

        const appliquerFiltres = () => {
            const mot = filtre.value.trim().toLowerCase();
            let vus = 0;
            table.querySelectorAll('tbody .peu-ligne:not(.peu-comp)').forEach((tr) => {
                const v = vues[Number(tr.dataset.idx)];
                const texte = (tr.querySelector('[data-nom]').value + ' ' + tr.querySelector('[data-desc]').value).toLowerCase();
                const garde = (!mot || texte.includes(mot))
                    && (!ecartsSeuls || v.champsDiff > 0 || v.nomChange || v.descChange || v.pertes > 0);
                tr.style.display = garde ? '' : 'none';
                const comp = tr.nextElementSibling;
                if (comp && comp.classList.contains('peu-comp')) comp.style.display = garde ? '' : 'none';
                if (garde) vus++;
            });
            barre.querySelector('.peu-search-count').textContent = t('poiCount', vus);
            majPied();
        };

        filtre.addEventListener('input', appliquerFiltres);
        btnEcarts.setAttribute('aria-pressed', 'false');
        btnEcarts.addEventListener('click', () => {
            ecartsSeuls = !ecartsSeuls;
            /* ⚠️ UN SEUL BOUTON PLEIN PAR ÉCRAN (charte) : Appliquer. Le filtre
               actif se marque d'un contour, il ne devient pas plein. */
            btnEcarts.classList.toggle('peu-btn-actif', ecartsSeuls);
            btnEcarts.setAttribute('aria-pressed', String(ecartsSeuls));
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
        if (_occupe || !_apercu) return;
        /* ⚠️ L'APERÇU EST FIGÉ AU CLIC : le bilan et le rapport se rattachent à
           celui-ci, même si l'écran en montrait un autre à la fin. */
        const apercu = _apercu;
        const items = lignesCochees();
        if (!items.length) return;

        const ov = document.getElementById('peu-overlay');
        const corps = ov.querySelector('#peu-body');
        const avant = await nbModifsEnAttente();
        occuper(true);
        /* ⚠️ EN TETE DU CORPS : la liste peut etre longue, et une barre posee en
           bas d une zone defilante travaille hors de vue. */
        const prog = poserProgression(corps, t('applying', 0, items.length), items.length, null, true);

        let bilan = { resultats: [], aReprendre: [] };
        let erreurGenerale = null;
        try {
            bilan = await poserLesLignes(items, environnementDePose(), (n, total) => {
                prog.avance(n, total);
                prog.libelle(t('applying', n, total));
            });
        } catch (e) {
            /* Ce qui lève AVANT la boucle (le module d'écriture de WME absent) :
               rien n'a été posé, et cela se dit. */
            erreurGenerale = (e && e.message) || String(e);
        } finally {
            prog.retirer();
            occuper(false);
        }

        try { await cadrerSurLesLieux(apercu.venueMap); } catch (e) { /* la carte reste où elle est */ }
        apercu.resultats = bilan.resultats;

        /* ⭐ UNE LIGNE POSÉE SE DÉCOCHE, ET LE DIT : recliquer ne repose que ce qui
           reste coché — les lieux en échec, pour « Réessayer ». */
        /* ⚠️ La ligne se relit sur le lieu VIVANT, qui porte désormais les
           valeurs posées : une pose PARTIELLE montre alors ce qui reste à poser,
           au lieu d'un ✔ qui mentirait. */
        /* ⚠️ LE LIEU EST UNE PHOTOGRAPHIE (`lireLieu`) : on la reprend après la pose,
           sans quoi la ligne se comparerait à l'état d'AVANT. */
        for (let i = 0; i < bilan.resultats.length; i++) {
            const r = bilan.resultats[i];
            if (r.status !== 'applied' && r.status !== 'partial') continue;
            try { const frais = await lireLieu(items[i].vid); if (frais) apercu.venueMap[items[i].vid] = frais; } catch (e) { /* on garde l'ancienne */ }
        }
        bilan.resultats.forEach((r, i) => {
            if (r.status !== 'applied' && r.status !== 'partial') return;
            const idx = items[i].idx;
            const tr = ov.querySelector('#peu-body .peu-ligne[data-idx="' + idx + '"]');
            if (!tr || _apercu !== apercu) return;
            const cb = tr.querySelector('.peu-checkbox');
            if (cb) cb.checked = false;
            const v = vueDeLaLigne(apercu.pois[idx], idx, apercu.venueMap, { nom: items[i].nom, desc: items[i].desc });
            v.posee = r.status === 'applied';
            if (apercu.vues) apercu.vues[idx] = v;
            majLigne(tr, v);
        });
        if (bilan.resultats.some((r) => r.status === 'applied' || r.status === 'partial') && _fichierCourant) {
            recordFileApplied(_fichierCourant);
            _rafraichirHistorique();
        }
        montrerBilan(corps, bilan, erreurGenerale, avant, await nbModifsEnAttente());
        majPied();
    }

    /**
     * LE BILAN — il dit ce qui est pose, ce qui manque, et que RIEN N'EST
     * ENREGISTRE.
     *
     * ⚠️ « Applique » ne veut pas dire « enregistre » : la confusion coute une
     *    session de travail perdue, et elle ne se voit qu'au rechargement.
     */
    function montrerBilan(corps, bilan, erreurGenerale, avant, apres) {
        const r = bilan.resultats;
        const compte = (s) => r.filter((x) => x.status === s).length;
        const poses = compte('applied'), partiels = compte('partial');
        const echecs = r.filter((x) => x.status === 'timeout' || x.status === 'erreur');
        const sae = r.filter((x) => x.verrou === 'sae' && (x.status === 'applied' || x.status === 'partial')).length;
        /* ⚠️ LE CHIFFRE DIT CE QUE CETTE POSE A AJOUTÉ, et rien quand on ne le sait
           pas : « 0 en attente » sur une erreur de lecture rassurait à tort, et
           toute la pile de WME comptait comme si le script l'avait remplie. */
        const ajoutees = avant !== null && apres !== null ? apres - avant : null;

        const vieux = corps.querySelector('[data-bilan]');
        if (vieux) vieux.remove();
        const div = document.createElement('div');
        div.className = 'peu-alert ' + (echecs.length || partiels || erreurGenerale ? 'peu-alert-warn' : 'peu-alert-ok');
        div.setAttribute('data-bilan', '');
        div.setAttribute('role', 'status');
        div.tabIndex = -1;
        const ligne = (texte, gras) => {
            const el = document.createElement('div');
            if (gras) { const b = document.createElement('b'); b.textContent = texte; el.appendChild(b); } else el.textContent = texte;
            div.appendChild(el);
        };

        if (erreurGenerale) ligne(t('bilanErreurGenerale', erreurGenerale), true);
        ligne(t('successMsg', poses), true);
        if (partiels) ligne(t('bilanPartiel', partiels));
        if (sae) ligne(t('bilanSae', sae));
        if (echecs.length) {
            ligne(t('bilanEchec', echecs.length));
            /* ⭐ LES LIEUX EN ÉCHEC SE NOMMENT, et se reprennent d'un clic : leurs
               lignes sont restées cochées. */
            const ul = document.createElement('ul');
            echecs.forEach((x) => {
                const li = document.createElement('li');
                li.textContent = t('echecLigne', x.oldName || x.newName || x.vid,
                    t(x.status === 'erreur' ? 'statusErreur' : 'statusTimeout'), x.erreur || '');
                ul.appendChild(li);
            });
            div.appendChild(ul);
            const reessayer = document.createElement('button');
            reessayer.type = 'button';
            reessayer.className = 'peu-btn peu-btn-neutral peu-btn-sm';
            reessayer.textContent = t('btnRetry', echecs.length);
            reessayer.addEventListener('click', () => { appliquerLignes().catch(signalerErreur); });
            div.appendChild(reessayer);
        }
        ligne(ajoutees !== null ? t('bilanNonEnregistre', ajoutees) : t('bilanNonEnregistreSansCompte'), true);
        corps.prepend(div);
        div.focus({ preventScroll: true });
    }

    /** Le nombre de modifications non enregistrées, ou `null` si on ne peut pas le lire. */
    async function nbModifsEnAttente() {
        try { return await obtenirSdk().Editing.getUnsavedChangesCount(); } catch (e) { return null; }
    }

    let _peuInited = false;
    function _peuInit() {
        if (_peuInited) return;
        _peuInited = true;
        initScript().catch((e) => console.error('[WPEU] démarrage', e));
    }

    // Démarrage dès que le SDK est prêt, à TOUS les zooms (comme WNA et WZM) : « wme-ready » n'arrive qu'à un zoom
    // éditable (≥ 12), et le script — son bouton de carte compris — restait absent tant qu'on regardait la carte de
    // loin (demande de l'auteur, 04/10/2026). Garde : wme-initialized et wme-ready peuvent arriver tous les deux.
    (() => {
        let lance = false;
        const go = () => { if (lance) return; lance = true; clearInterval(minuterie); Promise.resolve(pw.SDK_INITIALIZED).then(_peuInit); };
        const pret = () => !!pw.SDK_INITIALIZED;
        const minuterie = setInterval(() => { if (pret()) go(); }, 300);
        if (pret()) go();
        document.addEventListener('wme-initialized', go, { once: true });
        document.addEventListener('wme-ready', go, { once: true });
    })();
})();
