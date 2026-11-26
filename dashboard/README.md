# Maintenance dashboard: LM leon anaesthesia workstations

Open `index.html` in a browser (double-click). There is nothing to install or build, and no server is needed. Excel files are read locally in the browser, and no data leaves the computer. The libraries are in `vendor/`, so it also works offline.

## Usage
- **Importer un fichier Excel** (or drag and drop a file onto the page) to analyse a workbook.
- **Modèle Excel vierge** downloads the empty template to fill in.
- **Exporter les données** saves the current data back into that template.
- The filters cover period, equipment, model, department/zone, criticality and maintenance type.
- Each chart has a **Tableau** button that shows its values as a table.

## Accepted files
- The SAMQ template *Dashboard Maintenance Industrielle* is read as-is.
- Any workbook with the same column headers also works. Column order and sheet names don't matter, because sheets are recognised by their headers.
- The *Diagnostic d'import* section at the bottom of the page lists which sheets and columns were recognised.

Optional columns that add leon plus features:
| Sheet | Column | Used for |
|---|---|---|
| Equipment | `Modèle`, `N° de série` | Model filter, list of equipment |
| Equipment | `Date dernière maintenance annuelle` | Next annual service + CST date |
| Breakdowns | `Code alarme`, `Référence composant` | Breakdown log |
| Interventions | `Référence pièce` | Components / spare parts table |

A completed preventive task whose name contains "annuelle" or "CST" also counts as the last annual service.

## Logo
Save the host company's logo as `logo.png` next to `index.html`. Until then, a "LOGO" placeholder is shown.

## Files
- `core.js`: parsing, KPI calculations, manufacturer due dates, demo data, Excel export. It has no DOM code and can be tested in Node.
- `app.js`: page interface (filters, charts, tables).
- Manufacturer reference values come from the leon plus user manual, rev. 3.11.18, chapters 13 and 16.
- The `DEMO-REF-…` part numbers are fictional. Replace them with the references from the *leon plus, leon and leon mri List of accessories and replacement parts*.
