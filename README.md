# WME POI Event Updater

A userscript for the [Waze Map Editor](https://www.waze.com/editor) that bulk-updates the names
and descriptions of places (POIs) from an Excel file — one row per place, one file per event.

Built for editors who re-label the same set of places for every edition of a recurring event:
car parks, entrances, shuttle stops, campsites. Instead of opening each place by hand, you fill
in a spreadsheet once and let the script apply it.

[![Install from GreasyFork](https://img.shields.io/badge/install-GreasyFork-red)](https://greasyfork.org/scripts/578776)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

![WME POI Event Updater](Capture%200.52.jpg)

## How it works

Each row of the spreadsheet holds a **WME permalink** to a place, its **new name** and its **new
description**. The script reads the file, walks the rows, and applies the changes in the editor.

A blank template is provided: **[`WME_POI_Event_Updater_Template.xlsx`](WME_POI_Event_Updater_Template.xlsx)**.

| Column | Contents |
| --- | --- |
| `POI Permalink` | WME permalink pointing at the place |
| `POI Name` | Name to set |
| `POI Description` | Description to set |

Columns are matched **by header name**, not by position (case, spacing and accents are ignored;
French headers `Permalien` / `Nom` / `Description` are accepted too). You may therefore reorder
them or add columns of your own — extra columns are ignored. If none of the three headers is
recognised, the script falls back to reading columns A, B and C as before, and says so.

### More fields (0.50)

Beyond name and description, the file may carry these columns. **A blank cell asks for nothing and
erases nothing.**

| Column | Contents |
| --- | --- |
| `Categories` | Place categories — as shown in *your* editor's language |
| `Alternative Names` · `Phone` · `Website` | as in WME |
| `Services` | Place services (Wi-Fi, Restrooms…) |
| `Parking Type` · `Parking Cost` · `Parking Spots` | Public/Private/Restricted · Free…Expensive · 1-10…>600 |
| `Parking Payment` · `Parking Services` · `Parking Situation` | several values, separated by `;` |
| `Parking Exit When Closed` · `Parking Type Varies` | `Yes` / `No` |

Values are accepted **as WME's own labels** (in your language) or as its internal keys. **Anything
else is rejected and shown**, never silently applied.

Some columns are **read and displayed, but never written**: `Opening Hours`, `Address`,
`Entry Points`, `Parking Operator` and the Google fields. Writing them would mean guessing — a
misread opening-hours sentence would land on the map with nothing to signal it. They appear in the
preview so you can set them by hand.

⚠️ **A list replaces the whole field in WME.** If a row lists fewer values than the place already
has, the preview shows it in **orange**, says what it **removes**, and leaves the row **unticked**.

## The window (0.52)

Everything you do lives in a floating window, opened by the map button. Drop a workbook
anywhere on it, or click to pick one. The window can be dragged by its header, resized from
the corner, and collapsed; a double-click on the header puts it back.

Each row carries its state in **three** signs — a left border, a background and a badge — so
it stays readable under the hover. The tick box is the **first** column, and **Apply** is a
named button that says how many rows it will write. Nothing is written to the map until you
press it, and nothing is ever saved: you review in WME, then click Save yourself.

## Installation

1. Install a userscript manager — [Tampermonkey](https://www.tampermonkey.net/) is recommended.
2. Install the script from **[GreasyFork](https://greasyfork.org/scripts/578776)**.
3. Open the Waze Map Editor. A **POI Event Updater** button appears in the map button
   column, on the right: it opens the work window. The **POI Events** sidebar tab keeps
   the history and the settings.

Updates are delivered automatically through GreasyFork.

## Support

- Questions and discussion: **[Waze forum thread](https://www.waze.com/discuss/t/script-wme-poi-event-updater/404593)**
- Bug reports: **[open an issue](../../issues)**

When reporting a bug, please include your script version, your browser, and the steps to
reproduce it. Do not attach spreadsheets containing real event data.

## License

[MIT](LICENSE) © DrSlump34
