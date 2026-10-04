# WME POI Event Updater

A userscript for the [Waze Map Editor](https://www.waze.com/editor) that bulk-updates the names,
descriptions and fields of places (POIs) from an Excel file — one row per place, one sheet per event.

Built for editors who re-label the same set of places for every edition of a recurring event:
car parks, entrances, shuttle stops, campsites. Instead of opening each place by hand, you fill
in a spreadsheet once and let the script apply it.

[![Install from GreasyFork](https://img.shields.io/badge/install-GreasyFork-red)](https://greasyfork.org/scripts/578776)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

![WME POI Event Updater](Capture%200.53.03.jpg)

## How it works

Each row of the spreadsheet holds a **WME permalink** to a place, its **new name** and its **new
description**. The script reads the file, shows you a preview, and — when you press **Apply** —
writes the ticked rows into the editor. **It never saves**: you review on the map, then click
Save yourself.

A blank template is provided: **[`WME_POI_Event_Updater_Template.xlsx`](WME_POI_Event_Updater_Template.xlsx)**.

| Column | Contents |
| --- | --- |
| `POI Permalink` | WME permalink pointing at the place (with `env`, `lat`, `lon`, `zoomLevel` and `venues`) |
| `POI Name` | Name to set |
| `POI Description` | Description to set — **an empty cell erases the place's description** |

Columns are matched **by header name**, not by position (case, spacing and accents are ignored;
French headers `Permalien` / `Nom` / `Description` are accepted too). You may therefore reorder
them or add columns of your own — extra columns are ignored.

If the headers are not recognised, the script falls back to reading columns A, B and C, and says
so. It **refuses** that fallback when a recognised header sits elsewhere (say `Description` in
column B): reading by position would then put the description in the name, for the whole sheet.

⚠️ **An empty description is written as empty.** That is on purpose — an "off-event" sheet puts
places back to plain — and the preview says so in the field itself ("the place description will
be erased"). An empty *name* is never ticked by default.

### More fields

Beyond name and description, the file may carry these columns. **For these fields, a blank cell
asks for nothing and erases nothing.**

| Column | Contents |
| --- | --- |
| `Categories` | Place categories — as shown in *your* editor's language. **The first one is the main category**: order matters |
| `Alternative Names` · `Phone` · `Website` | as in WME |
| `Services` | Place services (Wi-Fi, Restrooms…) |
| `Parking Type` · `Parking Cost` · `Parking Spots` | Public/Private/Restricted · Free…Expensive · 1-10…>600 |
| `Parking Payment` · `Parking Services` · `Parking Situation` | several values, separated by `;` |
| `Parking Exit When Closed` · `Parking Type Varies` | `Yes` / `No` (or `Oui` / `Non`) |

Values are accepted as **WME's French labels** (the ones the EVIDRA export prints) or as WME's
**internal keys** (`FREE`, `CASH`, `WI_FI`…). Categories are the exception: they are read in the
language of your editor. **Anything else is rejected and shown**, never silently applied.

Some columns are **read and displayed, but never written**: `Opening Hours`, `Address`,
`Entry Points`, `Parking Operator` and the Google fields. Writing them would mean guessing — a
misread opening-hours sentence would land on the map with nothing to signal it. They appear in the
preview so you can set them by hand.

⚠️ **A list replaces the whole field in WME.** If a row lists fewer values than the place already
has, the preview shows it in **orange**, says what it **removes**, and leaves the row **unticked**.
The value a field *replaces* is written in clear next to the new one.

## The window

Everything you do lives in a floating window, opened by the map button. Drop a workbook
anywhere on it, or click to pick one. The window stays to the left of the map buttons; it can be
moved by its header and resized from the corner — with the mouse, or with the keyboard arrows once
the header or the corner has the focus. A double-click on the header puts it back.

- Each row carries its state in **three** signs — a left border, a background and a badge.
- The tick box is the **first** column. What changes is ticked by default; **editing a name or a
  description in the preview decides the box again**.
- **Apply** says how many rows it will write, including the ones hidden by the filter.
- After Apply, written rows are unticked and marked ✔; places that failed are **named** and stay
  ticked, so **Retry** writes only them. The report (Excel) lists, for each place, the permalink,
  the before and after, the status and **the fields that were not written**.

The interface is available in **French, English, German, Spanish, Italian, Portuguese (Brazil and
Portugal) and Hebrew**, following the editor's language.

## Installation

1. Install a userscript manager — [Tampermonkey](https://www.tampermonkey.net/) is recommended.
2. Install the script from **[GreasyFork](https://greasyfork.org/scripts/578776)**.
3. Open the Waze Map Editor. A **POI Event Updater** button appears in the map button
   column, on the right: it opens the work window. The script's tab in the **Scripts** panel
   keeps the recent files, the help, and the links — and says when a new version is out.

Updates are delivered automatically through GreasyFork. The script asks for one permission,
`GM_xmlhttpRequest` on `update.greasyfork.org`, used only to check for a new version (at most once
a day). The Excel library is [SheetJS](https://sheetjs.com/) 0.20.3, loaded from its own CDN with
an integrity hash.

## Support

- Questions and discussion: **[Waze forum thread](https://www.waze.com/discuss/t/script-wme-poi-event-updater/404593)**
- Bug reports: **[open an issue](../../issues)**

When reporting a bug, please include your script version, your browser, and the steps to
reproduce it. Do not attach spreadsheets containing real event data.

## License

[MIT](LICENSE) © DrSlump34
