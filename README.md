# POE2 Scout Price Scraper for Excel

An Excel VBA add-in that fetches real-time Path of Exile 2 item prices from [poe2scout.com](https://poe2scout.com).

## Features

- Real-time price lookup via poe2scout.com API
- In-memory cache (5 minutes) to reduce API calls
- Non-volatile functions — recalculates only on manual refresh (better performance with many formulas)
- Filter by category (currency, fragments, runes, talismans, essences, accessories, armour, weapon)
- Sorted item listings by price
- Manual and automatic refresh macros
- Fallback JSON parser if MSScriptControl is unavailable

## Installation

1. Open Excel and press `Alt + F11` to open the VBA editor.
2. In the menu, go to **Insert > Module**.
3. Copy the contents of `POE2ScoutScraper.bas` and paste into the new module.
4. Save the workbook as `.xlsm` (macro-enabled).

## Functions

### Price Lookup

```vba
=getPOE2Price(itemName, [category], [league])
```
Returns the current price (in Exalted Orbs) for the given item.

**Example:**
```
=getPOE2Price("Exalted Orb")
=getPOE2Price("Chaos Orb", "currency")
=getPOE2Price("Chaos Orb", "currency", "Fate of the Vaal")
```

### Item Details

```vba
=getPOE2ItemDetails(itemName, [category], [league])
```
Returns an array: `[Name, Price, Quantity, Category]`. Use as an array formula (`Ctrl+Shift+Enter`).

### Item Listings

```vba
=getPOE2Items([category], [league])
=getPOE2Currency([league])
=getPOE2Fragments([league])
=getPOE2Runes([league])
=getPOE2Talismans([league])
=getPOE2Essences([league])
=getPOE2Accessories([league])
=getPOE2Armour([league])
=getPOE2Weapons([league])
```
Returns a table of items sorted by price descending. Use as an array formula.

### Category List

```vba
=getPOE2Categories([league])
```
Returns all available categories and their item counts.

## Macros

| Macro | Description |
|---|---|
| `AtualizarPrecos` | Clears cache and recalculates all sheets |
| `AtualizarPlanilhaAtiva` | Clears cache and recalculates the active sheet |
| `AtualizarSelecao` | Clears cache and recalculates the current selection |
| `ForcarAtualizacao` | Full rebuild of all formulas |
| `ConfigurarAtualizacaoAutomatica` | Schedule automatic refresh at a chosen interval |
| `ClearCache` | Clears the in-memory cache |
| `TestAPIConnection` | Tests connectivity and prints results to Immediate Window |

## Keyboard Shortcut

You can assign `AtualizarPrecos` to `Ctrl+Shift+A` via **Tools > Macros > Options** in the VBA editor.

## Requirements

- Microsoft Excel (Windows) with macro support enabled
- Internet access to reach `poe2scout.com`
- `MSXML2.ServerXMLHTTP.6.0` (standard on Windows)
- `MSScriptControl.ScriptControl` (preferred JSON parser; falls back to built-in parser if unavailable)

## Default League

The default league is **Fate of the Vaal**. Pass the `league` parameter to any function to target a different league.

## License

MIT
