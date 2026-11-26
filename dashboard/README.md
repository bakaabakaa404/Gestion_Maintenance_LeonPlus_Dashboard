# Maintenance dashboard: LM leon anaesthesia workstations

Open `index.html` in a browser (double-click). There is nothing to install or build, and no server is needed. Excel files are read locally in the browser, and no data leaves the computer. The libraries are in `vendor/`, so it also works offline.

## Usage
The sidebar has one page per task:
| Page | Question it answers |
|---|---|
| Vue d'ensemble | What state is the fleet in, and what should I deal with first? |
| Équipements | Status of each machine, and its details (manufacturer due dates, breakdowns, parts) |
| Pannes | What breaks, why, and for how long |
| Maintenance préventive | Which tasks are overdue, coming up, or done |
| Ordres de travail | Which work orders are open, late or urgent |
| Pièces et composants | Which parts were replaced, their references and cost |
| Coûts / Équipe | Where the budget goes, and who does what |
| Guide constructeur | Leon plus intervals, consumables, glossary |
| Données et import | Import, blank template, export, list of recognised columns, targets |

- The period, department and model filters in the top bar apply to every page.
- You can import a file with the **Importer Excel** button or by dropping it anywhere on the page.
- Every page has its own URL (`#/pannes`, `#/equipements/LP-01`…), so you can bookmark a page or share its link.

## Accepted files
- The SAMQ template *Dashboard Maintenance Industrielle* is read as-is.
- Any workbook with the same column headers also works. Column order and sheet names don't matter, because sheets are recognised by their headers.
- The **Données et import** page lists which sheets and columns were recognised.

Optional columns that add leon plus features:
| Sheet | Column | Used for |
|---|---|---|
| Equipment | `Modèle`, `N° de série` | Model filter, list of equipment |
| Equipment | `Date dernière maintenance annuelle` | Next annual service + CST date |
| Breakdowns | `Code alarme`, `Référence composant` | Breakdown log |
| Interventions | `Référence pièce` | Components / spare parts table |

A completed preventive task whose name contains "annuelle" or "CST" also counts as the last annual service.

## Logo
Save the host company's logo as `logo.png` next to `index.html`. It appears at the top of the sidebar, in a white tile. Until then, a generic icon is shown.

## Files
- `core.js`: parsing, KPI calculations, manufacturer due dates, demo data, Excel export. It has no DOM code and can be tested in Node.
- `app.js` + `styles.css`: interface (navigation, pages, detail panel, sortable tables, charts).
- `vendor/`: SheetJS, Chart.js, Lucide icons, Inter font (local copies).
- Manufacturer reference values come from the leon plus user manual, rev. 3.11.18, chapters 13 and 16.
- The `DEMO-REF-…` part numbers are fictional. Replace them with the references from the *leon plus, leon and leon mri List of accessories and replacement parts*.
