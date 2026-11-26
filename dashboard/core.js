/* Maintenance dashboard core: workbook parsing, KPI computation, demo data.
 * Pure logic, no DOM. Works in the browser (window.GMCore) and in Node (require).
 * Dates are handled as Excel serial day numbers (days since 1899-12-30), like the source workbook.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.GMCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------- text & value helpers ----------
  const norm = (s) =>
    (s == null ? '' : String(s))
      .replace(/[œŒ]/g, 'oe')
      .normalize('NFD')
      .replace(/[̀-ͯ]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .trim();

  const str = (v) => (v == null ? '' : String(v).trim());

  const num = (v) => {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    const s = String(v).replace(/\s/g, '').replace(',', '.').replace(/[^0-9.\-]/g, '');
    if (!s) return null;
    const n = parseFloat(s);
    return isFinite(n) ? n : null;
  };

  const DAY_MS = 86400000;
  const serialFromUTC = (ms) => Math.floor(ms / DAY_MS) + 25569;
  const utcFromSerial = (s) => (s - 25569) * DAY_MS;

  const toSerial = (v) => {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v > 0 ? Math.floor(v) : null;
    if (v instanceof Date && !isNaN(v)) return serialFromUTC(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/);
    if (m) {
      let y = +m[3];
      if (y < 100) y += 2000;
      return serialFromUTC(Date.UTC(y, +m[2] - 1, +m[1]));
    }
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return serialFromUTC(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    return null;
  };

  const toTime = (v) => {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v - Math.floor(v);
    if (v instanceof Date && !isNaN(v)) return (v.getHours() * 60 + v.getMinutes()) / 1440;
    const m = String(v).trim().match(/^(\d{1,2})[:hH](\d{2})/);
    return m ? (+m[1] * 60 + +m[2]) / 1440 : null;
  };

  const todaySerial = () => {
    const d = new Date();
    return serialFromUTC(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  };

  const addMonths = (serial, n) => {
    const d = new Date(utcFromSerial(serial));
    const day = d.getUTCDate();
    const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + n, 1));
    const last = new Date(Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + 1, 0)).getUTCDate();
    t.setUTCDate(Math.min(day, last));
    return serialFromUTC(t.getTime());
  };

  const fmtDate = (s) => {
    if (s == null) return '';
    const d = new Date(utcFromSerial(s));
    const p = (x) => String(x).padStart(2, '0');
    return `${p(d.getUTCDate())}/${p(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;
  };

  const monthKey = (s) => {
    const d = new Date(utcFromSerial(s));
    return d.getUTCFullYear() * 12 + d.getUTCMonth();
  };
  const MONTHS_FR = ['janv.', 'févr.', 'mars', 'avr.', 'mai', 'juin', 'juil.', 'août', 'sept.', 'oct.', 'nov.', 'déc.'];
  const monthLabel = (k) => `${MONTHS_FR[k % 12]} ${String(Math.floor(k / 12)).slice(2)}`;
  const monthStart = (k) => serialFromUTC(Date.UTC(Math.floor(k / 12), k % 12, 1));

  // ---------- schema (French headers of the SAMQ template + Leon extensions) ----------
  // [key, export label, aliases]. Aliases are matched on normalized headers: exact first,
  // then "starts with" for multi-word aliases. Column order in the file does not matter.
  const SCHEMA = {
    equipements: {
      sheet: '04_Equipements',
      must: ['code', 'name'],
      fields: [
        ['code', 'Code équipement', ['code equipement', 'code']],
        ['name', 'Désignation', ['designation', 'nom de l equipement', 'nom']],
        ['model', 'Modèle', ['modele', 'type d appareil']],
        ['serial', 'N° de série', ['n de serie', 'no de serie', 'numero de serie', 'serie']],
        ['family', 'Famille technique', ['famille technique']],
        ['zone', 'Zone / service', ['zone', 'zone service', 'zone ou atelier', 'service']],
        ['line', 'Ligne / salle', ['ligne de production', 'ligne salle', 'salle', 'ligne']],
        ['maker', 'Fabricant', ['fabricant']],
        ['ref', 'Référence', ['reference']],
        ['commissioned', 'Date de mise en service', ['date de mise en service']],
        ['crit', 'Criticité', ['criticite']],
        ['hours', 'Heures de fonctionnement (compteur)', ['heures de fonctionnement']],
        ['resp', 'Responsable', ['responsable']],
        ['status', "Statut de l'équipement", ['statut de l equipement', 'statut']],
        ['hpd', "Temps d'ouverture planifié (h/jour)", ['temps d ouverture planifie', 'temps d ouverture']],
        ['lastService', 'Date dernière maintenance annuelle', ['date derniere maintenance annuelle', 'derniere maintenance annuelle']],
      ],
    },
    pannes: {
      sheet: '02_Registre_Pannes',
      must: ['date', 'code', 'cause'],
      fields: [
        ['id', 'ID panne', ['id panne', 'id']],
        ['date', 'Date de panne', ['date de panne', 'date']],
        ['start', 'Heure de début', ['heure de debut']],
        ['end', 'Heure de remise en service', ['heure de remise en service']],
        ['code', 'Code équipement', ['code equipement', 'code']],
        ['name', "Nom de l'équipement", ['nom de l equipement', 'designation']],
        ['zone', 'Zone ou service', ['zone ou atelier', 'zone ou service', 'zone', 'service']],
        ['line', 'Ligne / salle', ['ligne de production', 'ligne salle', 'salle']],
        ['desc', 'Description de la panne', ['description de la panne', 'description']],
        ['cause', 'Cause de panne', ['cause de panne', 'cause']],
        ['family', 'Famille de panne (sous-ensemble)', ['famille de panne', 'sous ensemble']],
        ['type', "Type d'intervention", ['type d intervention']],
        ['downtime', "Temps d'arrêt (h)", ['temps d arret']],
        ['repair', 'Temps de réparation (h)', ['temps de reparation']],
        ['resp', "Responsable de l'intervention", ['responsable de l intervention', 'responsable', 'technicien']],
        ['comment', 'Commentaire', ['commentaire', 'commentaires']],
        ['restoreDate', 'Date de remise en service (si arrêt ≥ 24 h)', ['date de remise en service']],
        ['include', 'Inclure dans les KPI', ['inclure dans les kpi']],
        ['alarm', 'Code alarme', ['code alarme', 'alarme', 'code erreur']],
        ['part', 'Référence composant', ['reference composant', 'composant', 'reference piece']],
      ],
    },
    interventions: {
      sheet: '03_Historique_Interventions',
      must: ['date', 'code', 'type'],
      fields: [
        ['id', 'ID intervention', ['id intervention', 'id']],
        ['date', 'Date', ['date', 'date d intervention']],
        ['code', 'Code équipement', ['code equipement', 'code']],
        ['type', "Type d'intervention", ['type d intervention', 'type de maintenance']],
        ['desc', 'Description des travaux', ['description des travaux', 'description']],
        ['cause', 'Cause identifiée', ['cause identifiee', 'cause']],
        ['action', 'Action réalisée', ['action realisee']],
        ['tech', 'Technicien', ['technicien', 'responsable']],
        ['dur', 'Durée de travail (h)', ['duree de travail']],
        ['downtime', "Temps d'arrêt associé (h)", ['temps d arret associe', 'temps d arret']],
        ['parts', 'Pièces de rechange utilisées', ['pieces de rechange utilisees', 'pieces de rechange', 'pieces']],
        ['partRef', 'Référence pièce', ['reference piece', 'reference composant', 'ref piece']],
        ['costMO', "Coût de main-d'œuvre", ['cout de main d oeuvre']],
        ['costParts', 'Coût des pièces', ['cout des pieces']],
        ['costTotal', 'Coût total', ['cout total']],
        ['status', 'Statut', ['statut']],
        ['comment', 'Commentaires', ['commentaires', 'commentaire']],
        ['costOther', 'Autres coûts (sous-traitance…)', ['autres couts']],
        ['panneRef', 'Réf. panne (optionnel)', ['ref panne']],
        ['name', 'Désignation équipement', ['designation equipement', 'designation']],
        ['zone', 'Zone', ['zone', 'service']],
      ],
    },
    preventif: {
      sheet: '05_Plan_Preventif',
      must: ['code', 'due'],
      fields: [
        ['id', 'ID tâche', ['id tache', 'id']],
        ['code', 'Code équipement', ['code equipement', 'code']],
        ['name', 'Désignation', ['designation']],
        ['op', 'Opération de maintenance', ['operation de maintenance', 'operation']],
        ['freq', 'Fréquence', ['frequence']],
        ['last', 'Dernière réalisation', ['derniere realisation']],
        ['due', 'Prochaine échéance', ['prochaine echeance', 'echeance']],
        ['dur', 'Durée prévue (h)', ['duree prevue']],
        ['resp', 'Responsable', ['responsable']],
        ['done', 'Date de réalisation effective', ['date de realisation effective', 'date de realisation']],
        ['status', 'Statut', ['statut']],
        ['comment', 'Commentaire', ['commentaire', 'commentaires']],
        ['zone', 'Zone', ['zone', 'service']],
      ],
    },
    ot: {
      sheet: '06_Ordres_Travail',
      must: ['id', 'planned'],
      fields: [
        ['id', 'Numéro OT', ['numero ot', 'n ot', 'ot']],
        ['created', 'Date de création', ['date de creation']],
        ['code', 'Code équipement', ['code equipement', 'code']],
        ['desc', 'Description du travail', ['description du travail', 'description']],
        ['prio', 'Priorité', ['priorite']],
        ['type', 'Type de maintenance', ['type de maintenance']],
        ['resp', 'Responsable', ['responsable']],
        ['planned', 'Date prévue', ['date prevue']],
        ['closed', 'Date de clôture', ['date de cloture']],
        ['estH', 'Temps estimé (h)', ['temps estime']],
        ['realH', 'Temps réel (h)', ['temps reel']],
        ['estCost', 'Coût estimé', ['cout estime']],
        ['realCost', 'Coût réel', ['cout reel']],
        ['status', 'Statut', ['statut']],
        ['comment', 'Commentaires', ['commentaires', 'commentaire']],
        ['late', 'En retard ?', ['en retard']],
        ['zone', 'Zone', ['zone', 'service']],
      ],
    },
  };

  const KPI_DEFS = [
    { key: 'pannes', label: 'Nombre de pannes', match: 'nombre total de pannes', unit: 'pannes', higher: false, obj: 110, thr: 130 },
    { key: 'dispo', label: 'Disponibilité technique', match: 'disponibilite technique', unit: '%', higher: true, obj: 0.995, thr: 0.99 },
    { key: 'prev', label: 'Taux de réalisation du préventif', match: 'taux de realisation du preventif', unit: '%', higher: true, obj: 0.95, thr: 0.85 },
    { key: 'mtbf', label: 'MTBF', match: 'mtbf', unit: 'h', higher: true, obj: 550, thr: 450 },
    { key: 'mttr', label: 'MTTR', match: 'mttr', unit: 'h', higher: false, obj: 2, thr: 2.5 },
    { key: 'arret', label: "Temps d'arrêt total", match: 'temps d arret total', unit: 'h', higher: false, obj: 450, thr: 600 },
    { key: 'cout', label: 'Coût total de maintenance', match: 'cout total de maintenance', unit: 'devise', higher: false, obj: 450000, thr: 550000 },
    { key: 'otRetard', label: 'OT en retard', match: 'ot en retard', unit: 'OT', higher: false, obj: 0, thr: 3 },
  ];

  // ---------- workbook parsing ----------
  function matchHeaders(headerRow, fields) {
    const hs = headerRow.map(norm);
    const used = new Set();
    const map = {};
    for (const pass of [0, 1]) {
      for (const [key, , aliases] of fields) {
        if (key in map) continue;
        for (const a of aliases) {
          const na = norm(a);
          const idx = hs.findIndex(
            (h, i) => h && !used.has(i) && (pass === 0 ? h === na : na.includes(' ') && h.startsWith(na))
          );
          if (idx >= 0) {
            map[key] = idx;
            used.add(idx);
            break;
          }
        }
      }
    }
    return map;
  }

  function sheetRows(XLSX, ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  }

  function parseWorkbook(XLSX, wb) {
    const report = { tables: {}, warnings: [], sheets: wb.SheetNames.slice() };
    const rowsBySheet = {};
    for (const name of wb.SheetNames) rowsBySheet[name] = sheetRows(XLSX, wb.Sheets[name]);

    const found = {};
    for (const [tkey, def] of Object.entries(SCHEMA)) {
      let best = null;
      for (const name of wb.SheetNames) {
        const rows = rowsBySheet[name];
        for (let r = 0; r < Math.min(rows.length, 20); r++) {
          const row = rows[r] || [];
          const map = matchHeaders(row, def.fields);
          if (!def.must.every((k) => k in map)) continue;
          const score = Object.keys(map).length;
          if (!best || score > best.score) best = { sheet: name, headerRow: r, map, score };
        }
      }
      if (best) found[tkey] = best;
    }
    // A sheet can host only one table: keep the best-scoring claim.
    const bySheet = {};
    for (const [t, f] of Object.entries(found)) {
      if (!bySheet[f.sheet] || bySheet[f.sheet].score < f.score) bySheet[f.sheet] = { t, score: f.score };
    }
    for (const [t, f] of Object.entries(found)) if (bySheet[f.sheet].t !== t) delete found[t];

    const raw = {};
    for (const tkey of Object.keys(SCHEMA)) {
      const f = found[tkey];
      raw[tkey] = [];
      if (!f) {
        report.tables[tkey] = null;
        continue;
      }
      const rows = rowsBySheet[f.sheet];
      for (let r = f.headerRow + 1; r < rows.length; r++) {
        const row = rows[r] || [];
        const o = {};
        let any = false;
        for (const [k, idx] of Object.entries(f.map)) {
          const v = row[idx];
          o[k] = v === '' ? null : v;
          if (o[k] != null) any = true;
        }
        if (any) raw[tkey].push(o);
      }
      const labels = SCHEMA[tkey].fields.filter(([k]) => k in f.map).map(([, l]) => l);
      const missing = SCHEMA[tkey].fields.filter(([k]) => !(k in f.map)).map(([, l]) => l);
      report.tables[tkey] = { sheet: f.sheet, headerRow: f.headerRow + 1, columns: labels, missing };
    }

    const assigned = new Set(Object.values(found).map((f) => f.sheet));
    const params = parseParams(wb.SheetNames.filter((n) => !assigned.has(n)).map((n) => rowsBySheet[n]));
    const data = normalizeData(raw, params);
    for (const t of Object.keys(SCHEMA)) {
      if (report.tables[t]) report.tables[t].rows = data[t].length;
    }
    if (!report.tables.pannes) report.warnings.push('Aucun registre des pannes reconnu : les KPI de pannes seront vides.');
    if (!report.tables.equipements)
      report.warnings.push('Aucune liste des équipements reconnue : MTBF et disponibilité ne peuvent pas être calculés.');
    return { data, report };
  }

  function parseParams(sheets) {
    const p = { joursSem: 6, devise: 'MAD', refDate: null, inclPrev: false, targets: {} };
    const rightOf = (row, i, pred) => {
      for (let j = i + 1; j < row.length; j++) if (row[j] != null && row[j] !== '' && pred(row[j])) return row[j];
      return null;
    };
    for (const rows of sheets) {
      for (const row of rows) {
        if (!row) continue;
        for (let i = 0; i < row.length; i++) {
          const n = norm(row[i]);
          if (!n) continue;
          if (n.startsWith('jours d exploitation par semaine')) {
            const v = num(rightOf(row, i, (x) => num(x) != null));
            if (v) p.joursSem = v;
          } else if (n === 'devise') {
            const v = rightOf(row, i, (x) => typeof x === 'string' && x.length <= 6);
            if (v) p.devise = str(v);
          } else if (n === 'date de reference') {
            const v = rightOf(row, i, (x) => toSerial(x) != null && (typeof x !== 'number' || x > 30000));
            if (v != null) p.refDate = toSerial(v);
          } else if (n.startsWith('arrets preventifs inclus')) {
            const v = rightOf(row, i, (x) => /^(oui|non)$/i.test(str(x)));
            if (v) p.inclPrev = /oui/i.test(v);
          } else {
            const def = KPI_DEFS.find((d) => n === d.match || n === norm(d.label));
            if (def && !p.targets[def.key]) {
              const nums = row.slice(i + 1).map(num).filter((x) => x != null);
              const sens = row.slice(i + 1).map(norm).find((x) => x.startsWith('plus haut') || x.startsWith('plus bas'));
              if (nums.length >= 2) {
                p.targets[def.key] = { obj: nums[0], thr: nums[1], higher: sens ? sens.startsWith('plus haut') : def.higher };
              }
            }
          }
        }
      }
    }
    return p;
  }

  function normalizeData(raw, params) {
    const equipements = raw.equipements
      .map((r) => ({
        code: str(r.code),
        name: str(r.name),
        model: str(r.model),
        serial: str(r.serial),
        family: str(r.family),
        zone: str(r.zone),
        line: str(r.line),
        maker: str(r.maker),
        ref: str(r.ref),
        commissioned: toSerial(r.commissioned),
        crit: str(r.crit),
        hours: num(r.hours),
        resp: str(r.resp),
        status: str(r.status),
        hpd: num(r.hpd),
        lastService: toSerial(r.lastService),
      }))
      .filter((e) => e.code);
    const eqBy = Object.fromEntries(equipements.map((e) => [e.code, e]));

    const pannes = raw.pannes
      .map((r) => {
        const date = toSerial(r.date);
        const start = toTime(r.start);
        const end = toTime(r.end);
        let downtime = num(r.downtime);
        if (downtime == null && date != null && start != null && end != null) {
          const restore = toSerial(r.restoreDate);
          const h = restore != null ? (restore + end - (date + start)) * 24 : (((end - start) % 1) + 1) % 1 * 24;
          downtime = Math.round(Math.max(0, h) * 100) / 100;
        }
        const code = str(r.code);
        const eq = eqBy[code];
        return {
          id: str(r.id),
          date,
          start,
          end,
          code,
          name: str(r.name) || (eq ? eq.name : ''),
          zone: str(r.zone) || (eq ? eq.zone : ''),
          line: str(r.line),
          desc: str(r.desc),
          cause: str(r.cause),
          family: str(r.family),
          type: str(r.type),
          downtime,
          repair: num(r.repair),
          resp: str(r.resp),
          comment: str(r.comment),
          include: norm(r.include) !== 'non',
          alarm: str(r.alarm),
          part: str(r.part),
        };
      })
      .filter((p) => p.date != null && p.code);

    const interventions = raw.interventions
      .map((r) => {
        const code = str(r.code);
        const eq = eqBy[code];
        const costMO = num(r.costMO);
        const costParts = num(r.costParts);
        const costOther = num(r.costOther);
        let costTotal = num(r.costTotal);
        if (costTotal == null && (costMO != null || costParts != null || costOther != null))
          costTotal = (costMO || 0) + (costParts || 0) + (costOther || 0);
        return {
          id: str(r.id),
          date: toSerial(r.date),
          code,
          type: str(r.type),
          desc: str(r.desc),
          cause: str(r.cause),
          action: str(r.action),
          tech: str(r.tech),
          dur: num(r.dur),
          downtime: num(r.downtime),
          parts: str(r.parts),
          partRef: str(r.partRef),
          costMO,
          costParts,
          costOther,
          costTotal,
          status: str(r.status),
          panneRef: str(r.panneRef),
          name: str(r.name) || (eq ? eq.name : ''),
          zone: str(r.zone) || (eq ? eq.zone : ''),
        };
      })
      .filter((i) => i.date != null && (i.code || i.type));

    const preventif = raw.preventif
      .map((r) => {
        const code = str(r.code);
        const eq = eqBy[code];
        return {
          id: str(r.id),
          code,
          name: str(r.name) || (eq ? eq.name : ''),
          op: str(r.op),
          freq: str(r.freq),
          last: toSerial(r.last),
          due: toSerial(r.due),
          dur: num(r.dur),
          resp: str(r.resp),
          done: toSerial(r.done),
          status: str(r.status),
          zone: str(r.zone) || (eq ? eq.zone : ''),
        };
      })
      .filter((t) => t.code && (t.due != null || t.op));

    const ot = raw.ot
      .map((r) => {
        const code = str(r.code);
        const eq = eqBy[code];
        return {
          id: str(r.id),
          created: toSerial(r.created),
          code,
          desc: str(r.desc),
          prio: str(r.prio),
          type: str(r.type),
          resp: str(r.resp),
          planned: toSerial(r.planned),
          closed: toSerial(r.closed),
          estH: num(r.estH),
          realH: num(r.realH),
          estCost: num(r.estCost),
          realCost: num(r.realCost),
          status: str(r.status),
          zone: str(r.zone) || (eq ? eq.zone : ''),
        };
      })
      .filter((o) => o.id && (o.created != null || o.code));

    return { equipements, pannes, interventions, preventif, ot, params };
  }

  // ---------- Leon plus manufacturer service schedule (user manual rev. 3.11.18, ch. 13 & 16) ----------
  const isLeon = (e) => /\bleon\b/i.test(`${e.model} ${e.name} ${e.family}`);
  const ANNUAL_RE = /(annuel|annuelle|cst|stc|controle de securite|securite technique)/;

  function serviceInfo(e, data, ref) {
    if (!isLeon(e)) return null;
    let last = e.lastService;
    for (const t of data.preventif)
      if (t.code === e.code && t.done != null && ANNUAL_RE.test(norm(t.op)) && (last == null || t.done > last)) last = t.done;
    for (const i of data.interventions)
      if (i.code === e.code && ANNUAL_RE.test(norm(`${i.desc} ${i.action}`)) && norm(i.status) !== 'annulee' && (last == null || i.date > last))
        last = i.date;
    const nextAnnual = last != null ? addMonths(last, 12) : null;
    let annualState = 'unknown';
    if (nextAnnual != null) annualState = nextAnnual < ref ? 'late' : nextAnnual <= ref + 30 ? 'soon' : 'ok';

    let overhaul = null;
    if (e.commissioned != null) {
      let k = 1;
      while (addMonths(e.commissioned, 36 * k) < ref) k++;
      const date = addMonths(e.commissioned, 36 * k);
      const hoursNext = e.hours != null ? Math.floor(e.hours / 10000 + 1) * 10000 : null;
      overhaul = {
        date,
        kind: k % 2 === 0 ? '6 ans' : '3 ans',
        hoursNext,
        hoursKind: hoursNext != null ? (hoursNext % 20000 === 0 ? '6 ans / 20 000 h' : '3 ans / 10 000 h') : null,
        hoursLeft: hoursNext != null ? hoursNext - e.hours : null,
      };
    }
    return { lastAnnual: last, nextAnnual, annualState, overhaul };
  }

  const LEON_REFERENCE = {
    source: 'Manuel utilisateur leon plus, rév. 3.11.18 (Löwenstein Medical), chap. 13 et 16',
    intervals: [
      ['Tous les 12 mois', 'CST (contrôle de sécurité technique) avant et après travaux, maintenance annuelle, ajustement / étalonnage du système', 'Technicien agréé'],
      ['Tous les 3 ans ou 10 000 h de fonctionnement', 'Révision générale : maintenance annuelle + maintenance 3 ans', 'Technicien agréé'],
      ['Tous les 6 ans ou 20 000 h de fonctionnement', 'Révision générale : maintenance annuelle + 3 ans + 6 ans', 'Technicien agréé'],
      ['Tous les 12 mois', 'Détendeur haute pression : inspection ; révision générale tous les 6 ans', 'Personnel qualifié'],
      ['Annuellement', 'Étalonnage de la mesure latérale des gaz (gaz d’étalonnage CO2 6 %, N2O 45 %, O2 45 %, desflurane 4 %)', 'Technicien agréé'],
    ],
    consumables: [
      ['Piège à eau et ligne d’échantillonnage', '1 mois, ou si endommagé (sinon fin de garantie de l’analyseur)'],
      ['Filtre d’aspiration bronchique', '2 mois, ou si souillé, endommagé ou baisse d’aspiration'],
      ['Chaux sodée', 'Changement de couleur ou CO2 inspiré en hausse'],
      ['Cellule O2 (mesure latérale)', '10 000 h à 100 % O2, ou si endommagée'],
      ['Cellule FiO2', '20 000 h à 100 % O2, ou si endommagée'],
      ['Membrane de valve PEEP', 'Maintenance annuelle, ou si non étanche / endommagée'],
      ['Membranes des valves insp./exp.', 'Maintenance annuelle, ou si endommagées'],
      ['Filtre de ventilateur', 'Maintenance annuelle, ou si encrassé / endommagé'],
      ['LM Flowsensor, absorbeur CO2 réutilisable, hublots', 'Garantie 1 an ou 52 cycles de nettoyage ; remplacer si encrassement persistant ou dommage'],
    ],
  };

  // ---------- computation ----------
  const TYPE_ALL = 'Tous';
  const isCorrectif = (t) => norm(t).startsWith('correctif');
  const isPlanned = (t) => {
    const n = norm(t);
    return n.startsWith('preventif') || n.startsWith('conditionnel');
  };
  const typeGroup = (t) => {
    const n = norm(t);
    if (n.startsWith('correctif')) return 'Correctif';
    if (n.startsWith('preventif')) return 'Préventif';
    if (n.startsWith('conditionnel') || n.startsWith('predictif')) return 'Conditionnel / Prédictif';
    return 'Autre';
  };
  const CRIT_ORDER = ['Critique', 'Élevée', 'Moyenne', 'Faible'];

  function kpiStatus(value, t) {
    if (value == null || !isFinite(value)) return 'na';
    if (t.higher) return value >= t.obj ? 'good' : value >= t.thr ? 'warn' : 'bad';
    return value <= t.obj ? 'good' : value <= t.thr ? 'warn' : 'bad';
  }

  function dataRange(data) {
    let min = null, max = null;
    const see = (s) => {
      if (s == null) return;
      if (min == null || s < min) min = s;
      if (max == null || s > max) max = s;
    };
    data.pannes.forEach((p) => see(p.date));
    data.interventions.forEach((i) => see(i.date));
    return { min, max };
  }

  function compute(data, f) {
    const P = data.params;
    const ref = f.ref != null ? f.ref : P.refDate != null ? P.refDate : todaySerial();
    const from = f.from;
    const to = f.to;
    const toEff = Math.min(to, ref);
    const days = Math.max(0, toEff - from + 1);
    const js = P.joursSem || 6;
    const type = f.type || TYPE_ALL;

    const eqBy = Object.fromEntries(data.equipements.map((e) => [e.code, e]));
    const eqOK = (e) =>
      (!f.eq || e.code === f.eq) &&
      (!f.zone || e.zone === f.zone) &&
      (!f.crit || e.crit === f.crit) &&
      (!f.model || (e.model || e.name) === f.model);
    const recOK = (code, zone) => {
      if (f.eq && code !== f.eq) return false;
      const e = eqBy[code];
      if (f.zone && (zone || (e && e.zone)) !== f.zone) return false;
      if (f.crit && (!e || e.crit !== f.crit)) return false;
      if (f.model && (!e || (e.model || e.name) !== f.model)) return false;
      return true;
    };
    const typeOK = (t) => type === TYPE_ALL || typeGroup(t) === type || t === type;
    const inP = (s) => s != null && s >= from && s <= to;

    const scopeEq = data.equipements.filter(eqOK);
    const activeEq = scopeEq.filter((e) => norm(e.status) !== 'reforme');
    const hpdSum = activeEq.reduce((a, e) => a + (e.hpd || 0), 0);

    const pannesOn = type === TYPE_ALL || type === 'Correctif';
    const pannesAll = data.pannes.filter((p) => pannesOn && inP(p.date) && recOK(p.code, p.zone));
    const pannes = pannesAll.filter((p) => p.include);
    const excluded = pannesAll.length - pannes.length;
    const interv = data.interventions.filter((i) => inP(i.date) && recOK(i.code, i.zone) && typeOK(i.type));
    const intervLive = interv.filter((i) => norm(i.status) !== 'annulee');
    const plannedStops = intervLive.filter((i) => isPlanned(i.type));

    const downP = pannes.reduce((a, p) => a + (p.downtime || 0), 0);
    const downPlan = plannedStops.reduce((a, i) => a + (i.downtime || 0), 0);
    const repairs = pannes.filter((p) => p.repair != null && p.repair >= 0);
    const mttr = repairs.length ? repairs.reduce((a, p) => a + p.repair, 0) / repairs.length : null;
    const open = hpdSum * days * js / 7;
    const run = Math.max(0, open - downP - downPlan);
    const hasEq = data.equipements.length > 0 && hpdSum > 0;
    const mtbf = hasEq && pannes.length ? run / pannes.length : null;
    const dispo = hasEq && run + downP > 0 ? (P.inclPrev ? run / (run + downP + downPlan) : run / (run + downP)) : null;

    const prevOn = type === TYPE_ALL || type === 'Préventif' || type === 'Conditionnel / Prédictif';
    const dueTasks = data.preventif.filter((t) => prevOn && t.due != null && t.due >= from && t.due <= toEff && recOK(t.code, t.zone));
    const doneTasks = dueTasks.filter((t) => t.done != null);
    const prevRate = dueTasks.length ? doneTasks.length / dueTasks.length : null;

    const cost = intervLive.reduce((a, i) => a + (i.costTotal || 0), 0);
    const otScope = data.ot.filter((o) => recOK(o.code, o.zone) && typeOK(o.type));
    const otOpen = otScope.filter((o) => !['cloture', 'annule'].includes(norm(o.status)));
    const otLate = otOpen.filter((o) => o.planned != null && o.planned < ref);

    const kv = {
      pannes: pannes.length,
      dispo,
      prev: prevRate,
      mtbf,
      mttr,
      arret: downP + downPlan,
      cout: cost,
      otRetard: otLate.length,
    };
    const kpis = KPI_DEFS.map((d) => {
      const t = Object.assign({ obj: d.obj, thr: d.thr, higher: d.higher }, P.targets[d.key] || {});
      return Object.assign({}, d, { value: kv[d.key], target: t, status: kpiStatus(kv[d.key], t) });
    });
    const details = {
      pannes: excluded ? `${excluded} exclue(s) des KPI` : `Sur ${days} jour(s) analysés`,
      dispo: hasEq ? `Arrêts sur pannes : ${round(downP, 1)} h` : 'Référentiel équipements requis',
      prev: `${doneTasks.length} réalisée(s) / ${dueTasks.length} due(s)`,
      mtbf: hasEq ? `Fonctionnement estimé : ${Math.round(run).toLocaleString('fr-FR')} h` : 'Référentiel équipements requis',
      mttr: `Sur ${repairs.length} réparation(s)`,
      arret: `Pannes ${round(downP, 1)} h · Planifié ${round(downPlan, 1)} h`,
      cout: 'Hors interventions annulées',
      otRetard: `Sur ${otOpen.length} OT ouvert(s)`,
    };
    kpis.forEach((k) => (k.detail = details[k.key]));

    // monthly series
    const months = [];
    if (days > 0) for (let k = monthKey(from); k <= monthKey(toEff); k++) months.push(k);
    const monthly = months.map((k) => {
      const ms = Math.max(from, monthStart(k));
      const me = Math.min(toEff, monthStart(k + 1) - 1);
      const d = Math.max(0, me - ms + 1);
      const mp = pannes.filter((p) => monthKey(p.date) === k);
      const down = mp.reduce((a, p) => a + (p.downtime || 0), 0);
      const plan = plannedStops.filter((i) => monthKey(i.date) === k).reduce((a, i) => a + (i.downtime || 0), 0);
      const o = hpdSum * d * js / 7;
      const r = Math.max(0, o - down - plan);
      const rep = mp.filter((p) => p.repair != null);
      return {
        key: k,
        label: monthLabel(k),
        pannes: mp.length,
        downtime: down,
        dispo: hasEq && r + down > 0 ? r / (r + down) : null,
        mttr: rep.length ? rep.reduce((a, p) => a + p.repair, 0) / rep.length : null,
      };
    });

    const groupSum = (rows, keyFn, valFn) => {
      const m = new Map();
      for (const r of rows) {
        const k = keyFn(r) || 'Non renseigné';
        m.set(k, (m.get(k) || 0) + valFn(r));
      }
      return [...m.entries()].map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value);
    };

    const correctivePannes = pannes;
    const causes = groupSum(correctivePannes, (p) => p.cause, () => 1);
    const totalCauses = causes.reduce((a, c) => a + c.value, 0);
    let cum = 0;
    causes.forEach((c) => {
      cum += c.value;
      c.cumPct = totalCauses ? cum / totalCauses : 0;
      c.downtime = pannes.filter((p) => (p.cause || 'Non renseigné') === c.label).reduce((a, p) => a + (p.downtime || 0), 0);
    });

    const families = groupSum(pannes, (p) => p.family, () => 1);
    const eqLabel = (code) => {
      const e = eqBy[code];
      return e && (e.model || e.name) ? `${code} · ${e.model || e.name}` : code;
    };
    const downByEq = groupSum(pannes, (p) => eqLabel(p.code), (p) => p.downtime || 0);

    const critRows = CRIT_ORDER.concat(['Non renseignée']).map((c) => {
      const ps = pannes.filter((p) => {
        const e = eqBy[p.code];
        const v = e && e.crit ? e.crit : 'Non renseignée';
        return v === c;
      });
      return { label: c, pannes: ps.length, downtime: ps.reduce((a, p) => a + (p.downtime || 0), 0) };
    });
    const critData = critRows.filter((r, i) => i < 4 || r.pannes > 0);

    // cost by equipment (stacked)
    const costMap = new Map();
    for (const i of intervLive) {
      const k = eqLabel(i.code);
      const c = costMap.get(k) || { label: k, mo: 0, parts: 0, other: 0, total: 0 };
      c.mo += i.costMO || 0;
      c.parts += i.costParts || 0;
      c.other += i.costOther || 0;
      c.total += i.costTotal || 0;
      costMap.set(k, c);
    }
    const costByEq = [...costMap.values()].sort((a, b) => b.total - a.total);

    // workload by responsible (hours, stacked by maintenance type)
    const respMap = new Map();
    const TYPES = ['Correctif', 'Préventif', 'Conditionnel / Prédictif', 'Autre'];
    const addResp = (name, t, h, isInt) => {
      const k = name || 'Non renseigné';
      const r = respMap.get(k) || { label: k, hours: Object.fromEntries(TYPES.map((x) => [x, 0])), count: 0, total: 0 };
      r.hours[t] += h || 0;
      r.total += h || 0;
      if (isInt) r.count++;
      respMap.set(k, r);
    };
    if (intervLive.length) intervLive.forEach((i) => addResp(i.tech, typeGroup(i.type), i.dur, true));
    else pannes.forEach((p) => addResp(p.resp, 'Correctif', p.repair, true));
    const byResp = [...respMap.values()].sort((a, b) => b.total - a.total);

    // components / spare parts
    const partMap = new Map();
    for (const i of intervLive) {
      const label = i.parts || '';
      const refc = i.partRef || '';
      if (!label && !refc) continue;
      const k = `${refc}||${label}`;
      const c = partMap.get(k) || { ref: refc, label, qty: 0, cost: 0, eqs: new Set(), last: null };
      c.qty++;
      c.cost += i.costParts || 0;
      c.eqs.add(i.code);
      if (c.last == null || i.date > c.last) c.last = i.date;
      partMap.set(k, c);
    }
    if (!partMap.size) {
      for (const p of pannes) {
        if (!p.part) continue;
        const c = partMap.get(p.part) || { ref: p.part, label: '', qty: 0, cost: 0, eqs: new Set(), last: null };
        c.qty++;
        c.eqs.add(p.code);
        if (c.last == null || p.date > c.last) c.last = p.date;
        partMap.set(p.part, c);
      }
    }
    const parts = [...partMap.values()]
      .map((c) => Object.assign(c, { eqs: [...c.eqs].sort() }))
      .sort((a, b) => b.qty - a.qty || b.cost - a.cost);

    // equipment fleet table
    const fleet = scopeEq.map((e) => {
      const ps = pannes.filter((p) => p.code === e.code);
      const down = ps.reduce((a, p) => a + (p.downtime || 0), 0);
      const plan = plannedStops.filter((i) => i.code === e.code).reduce((a, i) => a + (i.downtime || 0), 0);
      const o = norm(e.status) === 'reforme' ? 0 : (e.hpd || 0) * days * js / 7;
      const r = Math.max(0, o - down - plan);
      const c = intervLive.filter((i) => i.code === e.code).reduce((a, i) => a + (i.costTotal || 0), 0);
      const last = ps.reduce((m, p) => (m == null || p.date > m ? p.date : m), null);
      return Object.assign({}, e, {
        pannes: ps.length,
        downtime: down,
        dispo: r + down > 0 ? r / (r + down) : null,
        mtbf: ps.length && o ? r / ps.length : null,
        cost: c,
        lastPanne: last,
        service: serviceInfo(e, data, ref),
      });
    });

    // preventive: overdue & upcoming (next 30 days), independent of the period filter
    const prevScope = data.preventif.filter((t) => recOK(t.code, t.zone));
    const prevLate = prevScope.filter((t) => t.done == null && t.due != null && t.due < ref).sort((a, b) => a.due - b.due);
    const prevSoon = prevScope.filter((t) => t.done == null && t.due != null && t.due >= ref && t.due <= ref + 30).sort((a, b) => a.due - b.due);

    const recentPannes = pannesAll.slice().sort((a, b) => b.date - a.date || (b.start || 0) - (a.start || 0));

    return {
      ref, from, to, toEff, days,
      kpis,
      monthly,
      causes,
      families,
      downByEq,
      critData,
      costByEq,
      byResp,
      respTypes: TYPES,
      parts,
      fleet,
      prevLate,
      prevSoon,
      otOpen: otOpen.slice().sort((a, b) => (a.planned || 0) - (b.planned || 0)),
      recentPannes,
      devise: P.devise,
      hasLeon: data.equipements.some(isLeon),
    };
  }

  const round = (v, d) => (v == null ? '—' : (Math.round(v * 10 ** d) / 10 ** d).toLocaleString('fr-FR', { maximumFractionDigits: d }));

  function filterOptions(data) {
    const uniq = (arr) => [...new Set(arr.filter(Boolean))].sort((a, b) => a.localeCompare(b, 'fr'));
    const crits = uniq(data.equipements.map((e) => e.crit));
    return {
      equipements: data.equipements.map((e) => ({ value: e.code, label: e.name ? `${e.code} · ${e.name}` : e.code })),
      models: uniq(data.equipements.map((e) => e.model || e.name)),
      zones: uniq(data.equipements.map((e) => e.zone).concat(data.pannes.map((p) => p.zone))),
      crits: CRIT_ORDER.filter((c) => crits.includes(c)).concat(crits.filter((c) => !CRIT_ORDER.includes(c))),
      types: ['Correctif', 'Préventif', 'Conditionnel / Prédictif', 'Amélioratif', 'Autre'],
    };
  }

  // ---------- demo data: fictitious fleet of LM leon anaesthesia workstations ----------
  function mulberry32(a) {
    return function () {
      a |= 0;
      a = (a + 0x6d2b79f5) | 0;
      let t = Math.imul(a ^ (a >>> 15), 1 | a);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  function demoData(refDate) {
    const ref = refDate != null ? refDate : todaySerial();
    const rnd = mulberry32(20261126);
    const pick = (arr) => arr[Math.floor(rnd() * arr.length)];
    const between = (a, b) => a + rnd() * (b - a);
    const r2 = (v) => Math.round(v * 100) / 100;
    const q = (v) => Math.round(v * 4) / 4;
    const s = (y, m, d) => serialFromUTC(Date.UTC(y, m - 1, d));

    const techs = ['H. Mansour', 'S. Gharbi', 'Y. Jlassi', 'K. Ayari'];
    const LM = 'Technicien agréé LM';
    const eqDefs = [
      ['LP-01', 'leon plus', 'Bloc opératoire A', 'Salle 1', 'Critique', s(2019, 3, 12), 15800, 12, 'H. Mansour'],
      ['LP-02', 'leon plus', 'Bloc opératoire A', 'Salle 2', 'Critique', s(2019, 3, 12), 15150, 12, 'H. Mansour'],
      ['LP-03', 'leon plus', 'Bloc opératoire B', 'Salle 1', 'Élevée', s(2021, 6, 2), 9420, 10, 'S. Gharbi'],
      ['LP-04', 'leon plus', 'Bloc opératoire B', 'Salle 2', 'Élevée', s(2021, 6, 2), 9180, 10, 'S. Gharbi'],
      ['LP-05', 'leon plus', 'Maternité', 'Salle de césarienne', 'Critique', s(2022, 9, 20), 8240, 8, 'Y. Jlassi'],
      ['LN-01', 'leon', 'Bloc ambulatoire', 'Salle 1', 'Moyenne', s(2016, 1, 18), 21900, 8, 'K. Ayari'],
      ['LN-02', 'leon', 'Bloc ambulatoire', 'Salle 2', 'Moyenne', s(2016, 1, 18), 19650, 8, 'K. Ayari'],
      ['LM-01', 'leon mri', 'Imagerie (IRM)', 'Salle IRM', 'Élevée', s(2020, 11, 5), 6100, 6, 'Y. Jlassi'],
    ];
    const equipements = eqDefs.map(([code, model, zone, line, crit, commissioned, hours, hpd, resp], i) => ({
      code,
      name: `Station d'anesthésie ${model}`,
      model,
      serial: `DEMO-SN-${String(i + 1).padStart(4, '0')}`,
      family: "Station d'anesthésie",
      zone,
      line,
      maker: 'Löwenstein Medical',
      ref: '',
      commissioned,
      crit,
      hours,
      resp,
      status: 'En service',
      hpd,
      lastService: null,
    }));
    const weight = { 'LP-01': 1.3, 'LP-02': 1.2, 'LP-03': 1, 'LP-04': 0.9, 'LP-05': 0.8, 'LN-01': 1.4, 'LN-02': 1.2, 'LM-01': 0.6 };

    // [family, cause, description, alarm, part, part ref, repair h range, extra wait h range, part cost, weight]
    const catalog = [
      ['Mesure des gaz', "Piège à eau / ligne d'échantillonnage saturés", "Mesure des gaz instable, alarme ligne d'échantillonnage", '', "Piège à eau DRYLINE + ligne d'échantillonnage", 'DEMO-REF-WT', [0.25, 0.75], [0, 0.5], 85, 14],
      ['Cellule O2 / FiO2', 'Cellule O2 en fin de vie', 'Mesure O2 hors tolérance, étalonnage refusé', '135 – Étalonner la cellule O2', 'Cellule O2 (mesure latérale)', 'DEMO-REF-O2', [0.75, 1.5], [0, 2], 450, 7],
      ['Cellule O2 / FiO2', 'Cellule FiO2 non étalonnée / usée', 'Alarme étalonnage FiO2 au démarrage', '140 – Étalonner la cellule FiO2', 'Cellule FiO2', 'DEMO-REF-FIO2', [0.75, 1.5], [0, 2], 600, 4],
      ['LM Flowsensor', 'Capteur de débit encrassé ou endommagé', 'Test système échoué au bloc LM Flowsensors', '', 'LM Flowsensor', 'DEMO-REF-FS', [0.5, 1.25], [0, 1], 350, 10],
      ['Respirateur', 'Membranes des valves insp./exp. usées', "Test d'étanchéité non passé (fuite respirateur)", '', 'Membranes de valves insp./exp.', 'DEMO-REF-VD', [1, 2.5], [0, 2], 120, 7],
      ['Respirateur', 'Membrane de valve PEEP non étanche', 'PEEP instable pendant la ventilation', '', 'Membrane de valve PEEP', 'DEMO-REF-PEEP', [0.75, 1.5], [0, 1], 90, 5],
      ['Système patient', 'Fuite système patient (joint, dôme, soufflet)', 'Fuite importante au test du système complet', '', 'Joint de dôme / soufflet', 'DEMO-REF-SEAL', [0.5, 1.5], [0, 1], 60, 8],
      ['Absorbeur CO2', 'Chaux sodée épuisée / absorbeur endommagé', 'CO2 inspiré élevé', '', 'Absorbeur CO2 réutilisable', 'DEMO-REF-CO2', [0.5, 1], [0, 0.5], 700, 5],
      ['Alimentation électrique / batterie', 'Batterie défectueuse', 'Autonomie batterie insuffisante au test', '1 – Batterie défectueuse', 'Batterie', 'DEMO-REF-BAT', [1, 2], [2, 24], 900, 4],
      ["Ventilation de l'appareil", 'Ventilateur défaillant / filtre encrassé', 'Alarme ventilateur, échauffement', '5 – Défaut ventilateur', 'Filtre de ventilateur', 'DEMO-REF-FAN', [0.5, 1.5], [0, 4], 40, 4],
      ['Alimentation en gaz', 'Pression réseau insuffisante (cause externe)', 'Alarme pression d’alimentation O2 / AIR', '', '', '', [0.25, 1], [0, 2], 0, 4],
      ['Écran tactile / IHM', 'Défaut écran tactile / logiciel', "L'écran tactile ne répond plus", '', '', '', [1, 3], [0, 4], 0, 2],
      ['Mesure des gaz', 'Analyseur multigaz défectueux', "Mesure des gaz indisponible, alarme module d'analyse", '82 – Module de mesure des gaz défectueux', 'Module de mesure des gaz', 'DEMO-REF-GAS', [2, 4], [24, 72], 6500, 1.5],
      ['Utilisation', 'Mauvaise manipulation (évaporateur, raccords)', "Évaporateur mal verrouillé / mauvais raccordement", '', '', '', [0.25, 0.75], [0, 0.25], 0, 6],
    ];
    const totalW = catalog.reduce((a, c) => a + c[9], 0);
    const pickFailure = () => {
      let x = rnd() * totalW;
      for (const c of catalog) if ((x -= c[9]) <= 0) return c;
      return catalog[0];
    };
    const RATE = 45; // TND / h (fictitious)

    const start = ref - 365;
    const pannes = [];
    const interventions = [];
    let pid = 0, iid = 0;
    for (const e of equipements) {
      let d = start + Math.floor(between(2, 25));
      while (d <= ref) {
        const c = pickFailure();
        const repair = q(between(c[6][0], c[6][1]));
        const wait = rnd() < 0.35 ? between(c[7][0], c[7][1]) : 0;
        const downtime = r2(repair + wait + 0.25);
        const startT = (7 * 60 + Math.floor(between(0, 11 * 60))) / 1440;
        const endT = ((startT * 24 + downtime) % 24) / 24;
        const tech = rnd() < 0.15 && c[8] > 400 ? LM : rnd() < 0.7 ? e.resp : pick(techs);
        pid++;
        const id = `PAN-${String(pid).padStart(4, '0')}`;
        pannes.push({
          id, date: d, start: startT, end: endT, code: e.code, name: e.name, zone: e.zone, line: e.line,
          desc: c[2], cause: c[1], family: c[0], type: wait > 4 ? 'Correctif palliatif' : 'Correctif curatif',
          downtime, repair, resp: tech, comment: wait > 4 ? 'Attente pièce' : '', include: true, alarm: c[3], part: c[5],
        });
        iid++;
        const mo = Math.round(repair * RATE);
        const partCost = c[8] ? Math.round(c[8] * between(0.9, 1.1)) : 0;
        const other = tech === LM ? 350 : 0;
        interventions.push({
          id: `INT-${String(iid).padStart(4, '0')}`, date: d, code: e.code, type: 'Correctif', desc: c[2], cause: c[1],
          action: c[4] ? `Remplacement : ${c[4]}, test système` : 'Diagnostic, correction et test système',
          tech, dur: repair, downtime, parts: c[4], partRef: c[5], costMO: mo, costParts: partCost, costOther: other,
          costTotal: mo + partCost + other, status: 'Terminée', panneRef: id, name: e.name, zone: e.zone,
        });
        d += Math.max(3, Math.round(-Math.log(1 - rnd()) * (40 / weight[e.code])));
      }
    }

    // preventive plan
    const tasks = [
      ["Remplacement piège à eau et ligne d'échantillonnage", 'Mensuelle', 1, 0.25, "Piège à eau DRYLINE + ligne d'échantillonnage", 'DEMO-REF-WT', 85, 0],
      ["Remplacement filtre d'aspiration bronchique", 'Bimestrielle', 2, 0.25, "Filtre d'aspiration bronchique", 'DEMO-REF-BAF', 30, 0],
      ['Nettoyage et contrôle LM Flowsensors, absorbeur CO2', 'Trimestrielle', 3, 0.75, '', '', 0, 0],
      ['Maintenance annuelle + CST (contrôle de sécurité technique) et étalonnage', 'Annuelle', 12, 4, 'Kit maintenance annuelle (membranes, filtre ventilateur)', 'DEMO-REF-KIT12', 1400, 4],
    ];
    const preventif = [];
    let tid = 0;
    equipements.forEach((e, ei) => {
      tasks.forEach(([op, freq, months, dur, part, partRef, partCost, stop], ti) => {
        tid++;
        const id = `PM-${String(tid).padStart(2, '0')}`;
        let due = ti === 3 ? addMonths(start, Math.floor((ei * 1.4) % 12)) : start + Math.floor(between(0, 25));
        let last = addMonths(due, -months);
        while (due <= ref + 40) {
          const late = due < ref && ref - due < 20 && rnd() < 0.5;
          const skipped = due < ref - 20 && due > ref - 75 && rnd() < 0.08;
          const done = due > ref || late || skipped ? null : due + Math.round(between(-3, 6));
          const doneEff = done != null && done > ref ? null : done;
          const resp = ti === 3 ? LM : e.resp;
          preventif.push({ id, code: e.code, name: e.name, op, freq, last, due, dur, resp, done: doneEff, status: '', zone: e.zone });
          if (doneEff != null) {
            iid++;
            const mo = Math.round(dur * RATE);
            const pc = partCost ? Math.round(partCost * between(0.95, 1.05)) : 0;
            const other = ti === 3 ? 600 : 0;
            interventions.push({
              id: `INT-${String(iid).padStart(4, '0')}`, date: doneEff, code: e.code, type: ti === 2 ? 'Conditionnel / Prédictif' : 'Préventif',
              desc: op, cause: `Plan préventif ${id}`, action: `Opération réalisée selon gamme ${id}`, tech: resp, dur, downtime: stop,
              parts: part, partRef, costMO: mo, costParts: pc, costOther: other, costTotal: mo + pc + other, status: 'Terminée',
              panneRef: '', name: e.name, zone: e.zone,
            });
          }
          last = doneEff != null ? doneEff : last;
          due = addMonths(due, months);
        }
      });
    });
    for (const t of preventif) t.status = t.done != null ? 'Réalisée' : t.due < ref ? 'En retard' : t.due === ref ? "Aujourd'hui" : 'À venir';
    interventions.sort((a, b) => a.date - b.date);
    interventions.forEach((x, i) => (x.id = `INT-${String(i + 1).padStart(4, '0')}`));

    // work orders for the last ~3 months
    const ot = [];
    let oid = 0;
    const addOT = (o) => {
      oid++;
      ot.push(Object.assign({ id: `OT-${String(oid).padStart(4, '0')}` }, o));
    };
    for (const p of pannes.filter((p) => p.date >= ref - 90)) {
      const open = p.date >= ref - 6 && rnd() < 0.6;
      addOT({
        created: p.date, code: p.code, desc: p.desc, prio: equipements.find((e) => e.code === p.code).crit === 'Critique' ? 'Urgente' : 'Haute',
        type: 'Correctif', resp: p.resp, planned: p.date + (open ? Math.round(between(0, 4)) : 0), closed: open ? null : p.date + (p.downtime > 24 ? 2 : 0),
        estH: p.repair, realH: open ? null : p.repair, estCost: null, realCost: null, status: open ? pick(['En cours', 'En attente de pièces']) : 'Clôturé', zone: p.zone,
      });
    }
    for (const t of preventif.filter((t) => t.due >= ref - 60 && t.due <= ref + 30 && /annuelle|flowsensor/i.test(t.op))) {
      addOT({
        created: t.due - 10, code: t.code, desc: t.op, prio: 'Normale', type: 'Préventif', resp: t.resp, planned: t.due, closed: t.done,
        estH: t.dur, realH: t.done != null ? t.dur : null, estCost: null, realCost: null, status: t.done != null ? 'Clôturé' : 'Planifié', zone: t.zone,
      });
    }
    ot.sort((a, b) => a.created - b.created);
    ot.forEach((o, i) => (o.id = `OT-${String(i + 1).padStart(4, '0')}`));

    const params = {
      joursSem: 6,
      devise: 'TND',
      refDate: null,
      inclPrev: false,
      targets: {
        pannes: { obj: 80, thr: 100, higher: false },
        dispo: { obj: 0.99, thr: 0.98, higher: true },
        prev: { obj: 0.95, thr: 0.85, higher: true },
        mtbf: { obj: 350, thr: 250, higher: true },
        mttr: { obj: 1.5, thr: 2.5, higher: false },
        arret: { obj: 250, thr: 400, higher: false },
        cout: { obj: 70000, thr: 90000, higher: false },
        otRetard: { obj: 0, thr: 3, higher: false },
      },
    };
    return { equipements, pannes, interventions, preventif, ot, params };
  }

  // ---------- export to the Excel template (same sheet names/headers the parser reads) ----------
  function toAOA(tkey, rows) {
    const fields = SCHEMA[tkey].fields;
    const DATE_KEYS = new Set(['date', 'commissioned', 'lastService', 'last', 'due', 'done', 'created', 'planned', 'closed', 'restoreDate']);
    const TIME_KEYS = new Set(['start', 'end']);
    const cell = (k, v) => {
      if (v == null || v === '') return null;
      if (DATE_KEYS.has(k) && typeof v === 'number') return new Date(utcFromSerial(v));
      if (TIME_KEYS.has(k) && typeof v === 'number') {
        const m = Math.round(v * 1440);
        return `${String(Math.floor(m / 60) % 24).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;
      }
      if (k === 'include') return v ? 'Oui' : 'Non';
      return v;
    };
    return [fields.map(([, l]) => l)].concat(rows.map((r) => fields.map(([k]) => cell(k, r[k]))));
  }

  function buildWorkbook(XLSX, data) {
    const wb = XLSX.utils.book_new();
    const order = ['equipements', 'pannes', 'interventions', 'preventif', 'ot'];
    for (const t of order) {
      const ws = XLSX.utils.aoa_to_sheet(toAOA(t, data ? data[t] : []), { cellDates: true, dateNF: 'dd/mm/yyyy' });
      ws['!cols'] = SCHEMA[t].fields.map(([, l]) => ({ wch: Math.max(12, Math.min(40, l.length + 2)) }));
      XLSX.utils.book_append_sheet(wb, ws, SCHEMA[t].sheet);
    }
    const P = (data && data.params) || { joursSem: 6, devise: 'TND', targets: {} };
    const params = [
      ['PARAMÈTRES GÉNÉRAUX'],
      ['Date de référence', null, 'Vide = date du jour'],
      ['Devise', P.devise || 'TND'],
      ["Jours d'exploitation par semaine", P.joursSem || 6],
      ['Arrêts préventifs inclus dans la disponibilité ?', P.inclPrev ? 'Oui' : 'Non'],
      [],
      ['OBJECTIFS ET SEUILS'],
      ['Indicateur', 'Sens', 'Objectif', "Seuil d'alerte"],
    ];
    for (const d of KPI_DEFS) {
      const t = Object.assign({ obj: d.obj, thr: d.thr, higher: d.higher }, P.targets[d.key] || {});
      const label = { pannes: 'Nombre total de pannes', prev: 'Taux de réalisation du préventif', cout: 'Coût total de maintenance' }[d.key] || d.label;
      params.push([label, t.higher ? 'Plus haut = meilleur' : 'Plus bas = meilleur', t.obj, t.thr]);
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(params), '07_Parametres');
    const guide = [
      ['MODE D’EMPLOI'],
      ['Une ligne par enregistrement, en-têtes sur la première ligne. L’ordre des colonnes est libre : le tableau de bord reconnaît les en-têtes par leur nom.'],
      ['Code équipement : identique dans toutes les feuilles (ex. LP-01).'],
      ['Dates au format jj/mm/aaaa ; heures au format hh:mm. Le temps d’arrêt est recalculé si la cellule est vide.'],
      ['Temps d’ouverture planifié (h/jour) est nécessaire pour le MTBF et la disponibilité.'],
      ['Date dernière maintenance annuelle (ou une tâche préventive « Maintenance annuelle / CST » réalisée) permet de calculer la prochaine échéance constructeur.'],
      ['Les fichiers du modèle SAMQ « Dashboard Maintenance Industrielle » sont également lus tels quels.'],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(guide), '08_Guide');
    return wb;
  }

  return {
    norm, num, toSerial, toTime, todaySerial, addMonths, fmtDate, monthLabel,
    SCHEMA, KPI_DEFS, LEON_REFERENCE, CRIT_ORDER,
    parseWorkbook, compute, filterOptions, dataRange, demoData, buildWorkbook, isLeon, typeGroup,
  };
});
