# GMAO leon: maintenance of LM leon anaesthesia workstations

Open `index.html` in a browser (double-click). There is nothing to install or build, and no server is needed. The libraries are in `vendor/`, so it also works offline.

## Usage
The sidebar has one page per task. Pages marked ✎ let you add, edit and delete records (click a row, or use the button at the top of the table).
| Page | Purpose |
|---|---|
| Vue d'ensemble | State of the fleet and what to deal with first |
| Équipements ✎ | One card per machine. The detail panel has buttons to report a breakdown, schedule maintenance and edit the machine. |
| Pannes ✎ | Breakdown register: reporting, diagnosis and repair, following the specification's diagnostic tree (§5). |
| Maintenance préventive ✎ | Maintenance plan. The **Valider** button opens the protocol checklist (§4). |
| Ordres de travail | Work orders. A corrective work order is created automatically for each breakdown. |
| Services ✎ | Hospital services and the machines installed in each |
| Équipe ✎ | Team members, Löwenstein authorisation (with expiry date) and workload |
| Pièces et stock ✎ | Parts catalogue, stock levels, parts to order |
| Coûts | Breakdown of the maintenance budget |
| Guide technique | Preventive protocol, diagnostic tree, electrical safety limits, calibration gases, glossary |
| Données et import | Excel export and import, settings (currency, hourly rate), demo, empty database |

What happens automatically:
- **Reporting a breakdown** creates the corrective work order. The machine is shown as "En réparation" until the breakdown is closed.
- **Closing a breakdown** requires the root cause, the downtime, the recalibration and a full autotest (specification §5).
- **A replaced part** lowers the stock and is costed at its unit price. Labour is costed at the hourly rate (settings).
- **Validating a maintenance task** does the following:
  - records the checklist;
  - uses up the kit;
  - updates the last annual service and the hour counter;
  - closes the preventive work order;
  - schedules the next task.

  A "Non conforme" result offers to report a breakdown straight away.
- **Renaming a service, a person, a part or a machine code** also renames it everywhere in the history.
- **Deleting** a service is refused while machines are still attached to it. Deleting a machine is refused if it has a history: set its status to "Réformé" instead.
- Assigning a technician without a valid Löwenstein authorisation to a leon shows a warning.

## Where the data is stored
- Everything entered is saved automatically in the browser, on this computer (`localStorage`).
- For a backup, or to move the data to another computer: **Données et import → Exporter en Excel**, then import that file on the other computer.
- An import replaces the current database (a confirmation is shown first).
- To work on real data: **Nouvelle base vide**. Then add the services, the team, the parts and the machines.

## Accepted files
- The SAMQ template *Dashboard Maintenance Industrielle* is read as-is. Services, team and parts are rebuilt from the history.
- Any workbook with the same column headers also works. Column order and sheet names don't matter.
- The export (or **Modèle Excel vierge**) contains every sheet the app manages:
  - registers: `04_Equipements`, `02_Registre_Pannes`, `03_Historique_Interventions`, `05_Plan_Preventif`, `06_Ordres_Travail`;
  - managed lists: `09_Services`, `10_Equipe`, `11_Pieces`;
  - settings and help: `07_Parametres`, `08_Guide`.

## Preventive maintenance: sources
- The leon plus **Service Manual** is not public: Löwenstein Medical provides it only to authorised partners.
- The plan and the checklist are therefore built from two sources:
  - the specification (*cahier des charges*, §4): annual kit, calibration, final quality control;
  - chapter 13 of the user manual (rev. 3.11.18), "Maintenance by authorised service technicians":
    - service intervals: 12 months, 3 years / 10,000 h, 6 years / 20,000 h;
    - safety technical check (STC, DIN EN 62353);
    - high-pressure regulator;
    - gas measurement calibration.
- Only the tasks done by technicians are tracked. The user's own checks (water trap, filters, soda lime…) are left out.

## Logo
Save the host company's logo as `logo.png` next to `index.html`. It appears at the top of the sidebar.

## Files
- `core.js`: reading Excel files, KPI calculations, manufacturer due dates, protocols, demo data, Excel export.
- `store.js`: adding, editing and deleting records (renaming across the history, stock, linked work orders and costs).

  `core.js` and `store.js` contain no DOM code and can be tested in Node.
- `forms.js`: input forms (native dialog, validation messages, confirmation).
- `app.js` + `styles.css`: interface.
- `vendor/`: SheetJS, Chart.js, Lucide icons, Inter font.
- The `DEMO-REF-…` part numbers and the demo data are fictional.
