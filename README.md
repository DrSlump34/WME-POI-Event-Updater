# WME POI Event Updater

A userscript for the [Waze Map Editor](https://www.waze.com/editor) that bulk-updates the names
and descriptions of places (POIs) from an Excel file — one row per place, one file per event.

Built for editors who re-label the same set of places for every edition of a recurring event:
car parks, entrances, shuttle stops, campsites. Instead of opening each place by hand, you fill
in a spreadsheet once and let the script apply it.

[![Install from GreasyFork](https://img.shields.io/badge/install-GreasyFork-red)](https://greasyfork.org/scripts/578776)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

![WME POI Event Updater](Capture%200.48.png)

## How it works

Each row of the spreadsheet holds a **WME permalink** to a place, its **new name** and its **new
description**. The script reads the file, walks the rows, and applies the changes in the editor.

A blank template is provided: **[`WME_POI_Event_Updater_Template.xlsx`](WME_POI_Event_Updater_Template.xlsx)**.

| Column | Contents |
| --- | --- |
| `POI Permalink` | WME permalink pointing at the place |
| `POI Name` | Name to set |
| `POI Description` | Description to set |

## Installation

1. Install a userscript manager — [Tampermonkey](https://www.tampermonkey.net/) is recommended.
2. Install the script from **[GreasyFork](https://greasyfork.org/scripts/578776)**.
3. Open the Waze Map Editor. The updater appears in the sidebar.

Updates are delivered automatically through GreasyFork.

## Support

- Questions and discussion: **[Waze forum thread](https://www.waze.com/discuss/t/script-wme-poi-event-updater/404593)**
- Bug reports: **[open an issue](../../issues)**

When reporting a bug, please include your script version, your browser, and the steps to
reproduce it. Do not attach spreadsheets containing real event data.

## License

[MIT](LICENSE) © DrSlump34
