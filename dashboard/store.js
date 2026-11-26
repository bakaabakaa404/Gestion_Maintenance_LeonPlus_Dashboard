/* GMAO leon: data mutations (create / update / delete) on the in-memory database.
 * Pure logic, no DOM: usable in the browser (window.GMStore) and in Node (require).
 * Every function takes the database object returned by GMCore and mutates it in place.
 * Errors meant for the user are thrown as Error with a French message. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core.js'));
  else root.GMStore = factory(root.GMCore);
})(typeof self !== 'undefined' ? self : this, function (C) {
  'use strict';
  const norm = C.norm;
  const TABLES = ['equipements', 'pannes', 'interventions', 'preventif', 'ot', 'clients', 'services', 'equipe', 'pieces'];
  const DEFAULT_CLIENT = 'Client principal';

  const PANNE_STATUS = ['Ouverte', 'En cours de diagnostic', 'En attente de pièce', 'Résolue'];
  const OT_STATUS_FOR = { Ouverte: 'Ouvert', 'En cours de diagnostic': 'En cours', 'En attente de pièce': 'En attente de pièces', 'Résolue': 'Clôturé' };
  const isOpenPanne = (p) => {
    const n = norm(p.status);
    return !!n && !['resolue', 'cloturee', 'cloture', 'annulee', 'terminee'].includes(n);
  };

  // ---------- identity ----------
  function newUid(data) {
    data.meta.seq = (data.meta.seq || 0) + 1;
    return `r${data.meta.seq}`;
  }
  function nextId(list, prefix, width = 4) {
    let max = 0;
    for (const r of list) {
      const m = String(r.id || '').match(new RegExp(`^${prefix}(\\d+)$`));
      if (m) max = Math.max(max, +m[1]);
    }
    return `${prefix}${String(max + 1).padStart(width, '0')}`;
  }
  const find = (data, table, uid) => data[table].find((r) => r.uid === uid) || null;

  // ---------- registries derived from history (imported files have no Services / Équipe / Pièces sheet) ----------
  function ensure(data) {
    if (!data.meta) data.meta = { seq: 0 };
    for (const t of TABLES) if (!Array.isArray(data[t])) data[t] = [];
    for (const t of TABLES) for (const r of data[t]) if (!r.uid) r.uid = newUid(data);
    if (!data.meta.registries) deriveRegistries(data);
    assignClients(data);
    return data;
  }

  // Every service belongs to a client; files without a Clients sheet get one default client.
  function assignClients(data) {
    const known = new Set(data.clients.map((c) => c.name));
    const orphans = data.services.filter((x) => !x.client || !known.has(x.client));
    if (!orphans.length) return;
    for (const o of orphans) {
      if (o.client && !known.has(o.client)) {
        data.clients.push({ uid: newUid(data), name: o.client, city: '', contact: '', phone: '', email: '', contract: '', contractEnd: null, notes: '' });
        known.add(o.client);
      }
    }
    const rest = data.services.filter((x) => !x.client);
    if (!rest.length) return;
    let def = data.clients.length === 1 ? data.clients[0] : data.clients.find((c) => c.name === DEFAULT_CLIENT);
    if (!def) {
      def = { uid: newUid(data), name: DEFAULT_CLIENT, city: '', contact: '', phone: '', email: '', contract: '', contractEnd: null, notes: '' };
      data.clients.push(def);
    }
    rest.forEach((x) => (x.client = def.name));
  }

  function deriveRegistries(data) {

    const addOnce = (table, key, val, mk) => {
      if (!val || data[table].some((r) => norm(r[key]) === norm(val))) return;
      data[table].push(Object.assign({ uid: newUid(data) }, mk(val)));
    };
    const zones = data.equipements.map((e) => e.zone).concat(data.pannes.map((p) => p.zone));
    zones.forEach((z) => addOnce('services', 'name', z, (name) => ({ name, site: '', chef: '', phone: '', email: '', notes: '' })));

    const people = [].concat(
      data.equipements.map((e) => e.resp), data.pannes.map((p) => p.resp), data.interventions.map((i) => i.tech),
      data.preventif.map((t) => t.resp), data.ot.map((o) => o.resp));
    people.forEach((n) => addOnce('equipe', 'name', n, (name) => {
      const agreed = /(agree|lowenstein|loewenstein|\blm\b)/.test(norm(name));
      return { name, role: agreed ? 'Prestataire agréé Löwenstein' : 'Technicien biomédical', habil: agreed, habilEnd: null, phone: '', email: '', active: true };
    }));

    // parts: one entry per reference (or per name when there is no reference), unit price = average cost
    const seen = new Map();
    for (const i of data.interventions) {
      const key = i.partRef || i.parts;
      if (!key) continue;
      const s = seen.get(norm(key)) || { ref: i.partRef || '', name: i.parts || '', n: 0, cost: 0 };
      s.n++;
      s.cost += i.costParts || 0;
      if (!s.name && i.parts) s.name = i.parts;
      seen.set(norm(key), s);
    }
    for (const p of data.pannes) if (p.part && !seen.has(norm(p.part))) seen.set(norm(p.part), { ref: p.part, name: '', n: 0, cost: 0 });
    for (const s of seen.values()) {
      const ref = s.ref || s.name;
      addOnce('pieces', 'ref', ref, () => ({
        ref, name: s.name || s.ref, category: '', models: '', price: s.n && s.cost ? Math.round(s.cost / s.n) : null,
        stock: null, stockMin: null, supplier: '',
      }));
    }
    data.meta.registries = true;
  }

  function hourlyRate(data) {
    if (data.params.rate != null) return data.params.rate;
    let h = 0, c = 0;
    for (const i of data.interventions) if (i.dur > 0 && i.costMO > 0) { h += i.dur; c += i.costMO; }
    return h ? Math.round(c / h) : 0;
  }

  const partBy = (data, ref) => (ref ? data.pieces.find((p) => norm(p.ref) === norm(ref)) || null : null);
  function adjustStock(data, ref, delta) {
    const p = partBy(data, ref);
    if (p && p.stock != null && delta) p.stock = Math.max(0, p.stock + delta);
  }

  const unique = (list, key, val, uid, label) => {
    if (!String(val || '').trim()) throw new Error(`${label} est obligatoire.`);
    if (list.some((r) => r.uid !== uid && norm(r[key]) === norm(val))) throw new Error(`${label} « ${val} » existe déjà.`);
  };
  function cascade(data, pairs, oldV, newV) {
    if (!oldV || oldV === newV) return;
    for (const [table, key] of pairs) for (const r of data[table]) if (r[key] === oldV) r[key] = newV;
  }

  // ---------- pannes (each one carries its corrective work order and its intervention / cost line) ----------
  const PRIO_FOR = { Critique: 'Urgente', 'Élevée': 'Haute' };

  function savePanne(data, input, uid) {
    const eq = data.equipements.find((e) => e.code === input.code);
    if (!eq) throw new Error('Choisissez un équipement existant.');
    if (input.date == null) throw new Error('La date de la panne est obligatoire.');
    let p = uid ? find(data, 'pannes', uid) : null;
    const isNew = !p;
    const before = p ? { part: p.part, qty: p.partQty || (p.part ? 1 : 0), repair: p.repair, date: p.date, code: p.code } : null;
    if (!p) {
      p = { uid: newUid(data), id: nextId(data.pannes, 'PAN-'), include: true, type: 'Correctif curatif', comment: '' };
      data.pannes.push(p);
    }
    Object.assign(p, input, { name: eq.name, zone: eq.zone, line: eq.line });
    if (!p.part) p.partQty = null;
    else if (!p.partQty) p.partQty = 1;
    const open = isOpenPanne(p);

    // stock: give back what the previous version consumed, then consume the new quantity
    if (before && before.part) adjustStock(data, before.part, before.qty);
    if (p.part) adjustStock(data, p.part, -p.partQty);

    // intervention line (cost). Imported files may not link it: fall back to same equipment + date.
    let it = data.interventions.find((i) => i.panneRef && i.panneRef === p.id)
      || (!isNew ? data.interventions.find((i) => !i.panneRef && i.code === before.code && i.date === before.date && C.typeGroup(i.type) === 'Correctif') : null);
    const partChanged = !before || before.part !== p.part || before.qty !== (p.partQty || 0);
    const repairChanged = !before || before.repair !== p.repair;
    let created = false;
    if (!it && (isNew || partChanged || repairChanged)) {
      created = true;
      it = { uid: newUid(data), id: nextId(data.interventions, 'INT-'), costMO: 0, costParts: 0, costOther: 0 };
      data.interventions.push(it);
    }
    if (it) {
      const piece = partBy(data, p.part);
      if (created || repairChanged) it.costMO = Math.round((p.repair || 0) * hourlyRate(data));
      if (created || partChanged) it.costParts = piece && piece.price != null ? Math.round(piece.price * (p.partQty || 0)) : p.part ? it.costParts || 0 : 0;
      Object.assign(it, {
        panneRef: p.id, date: p.date, code: p.code, type: 'Correctif', desc: p.desc, cause: p.cause,
        action: p.part ? `Remplacement : ${piece ? piece.name : p.part}${p.autotest ? ', autotest complet OK' : ''}` : `Diagnostic et réparation${p.autotest ? ', autotest complet OK' : ''}`,
        tech: p.resp, dur: p.repair, downtime: p.downtime, parts: p.part ? (piece ? piece.name : p.part) : '', partRef: p.part || '',
        status: open ? 'En cours' : 'Terminée', name: eq.name, zone: eq.zone,
      });
      it.costTotal = (it.costMO || 0) + (it.costParts || 0) + (it.costOther || 0);
    }

    // corrective work order
    let o = data.ot.find((x) => x.panneRef && x.panneRef === p.id)
      || (!isNew ? data.ot.find((x) => !x.panneRef && x.code === before.code && x.created === before.date && C.typeGroup(x.type) === 'Correctif') : null);
    if (!o && (isNew || open)) {
      o = { uid: newUid(data), id: nextId(data.ot, 'OT-'), created: p.date, type: 'Correctif', prio: PRIO_FOR[eq.crit] || 'Normale', estH: p.repair };
      data.ot.push(o);
    }
    if (o) {
      Object.assign(o, {
        panneRef: p.id, code: p.code, desc: p.desc, resp: p.resp, zone: eq.zone,
        planned: o.planned != null ? o.planned : p.date,
        status: OT_STATUS_FOR[p.status] || (open ? 'En cours' : 'Clôturé'),
        closed: open ? null : p.restoreDate != null ? p.restoreDate : p.date,
        realH: open ? null : p.repair,
      });
    }
    return p;
  }

  function deletePanne(data, uid) {
    const p = find(data, 'pannes', uid);
    if (!p) return;
    if (p.part) adjustStock(data, p.part, p.partQty || 1);
    data.pannes = data.pannes.filter((x) => x !== p);
    if (p.id) {
      data.interventions = data.interventions.filter((i) => i.panneRef !== p.id);
      data.ot = data.ot.filter((o) => o.panneRef !== p.id);
    }
  }

  // ---------- equipment ----------
  function saveEquipment(data, input, uid) {
    unique(data.equipements, 'code', input.code, uid, 'Le code équipement');
    let e = uid ? find(data, 'equipements', uid) : null;
    const oldCode = e ? e.code : null;
    if (!e) {
      e = { uid: newUid(data), family: "Station d'anesthésie", maker: 'Löwenstein Medical', ref: '', status: 'En service', lastService: null };
      data.equipements.push(e);
    }
    Object.assign(e, input);
    if (!e.name) e.name = e.model ? `Station d'anesthésie ${e.model}` : e.code;
    cascade(data, [['pannes', 'code'], ['interventions', 'code'], ['preventif', 'code'], ['ot', 'code']], oldCode, e.code);
    return e;
  }
  function deleteEquipment(data, uid) {
    const e = find(data, 'equipements', uid);
    if (!e) return;
    const n = ['pannes', 'interventions', 'preventif', 'ot'].reduce((a, t) => a + data[t].filter((r) => r.code === e.code).length, 0);
    if (n) throw new Error(`${e.code} a un historique (${n} enregistrements). Pour le retirer du parc sans perdre cet historique, passez son statut à « Réformé ».`);
    data.equipements = data.equipements.filter((x) => x !== e);
  }

  // ---------- clients ----------
  function saveClient(data, input, uid) {
    unique(data.clients, 'name', input.name, uid, 'Le nom du client');
    let c = uid ? find(data, 'clients', uid) : null;
    const old = c ? c.name : null;
    if (!c) { c = { uid: newUid(data) }; data.clients.push(c); }
    Object.assign(c, input);
    cascade(data, [['services', 'client']], old, c.name);
    return c;
  }
  function deleteClient(data, uid) {
    const c = find(data, 'clients', uid);
    if (!c) return;
    const n = data.services.filter((x) => x.client === c.name).length;
    if (n) throw new Error(`${n} service${n > 1 ? 's sont rattachés' : ' est rattaché'} à « ${c.name} ». Supprimez-${n > 1 ? 'les' : 'le'} ou rattachez-${n > 1 ? 'les' : 'le'} à un autre client d’abord.`);
    data.clients = data.clients.filter((x) => x !== c);
  }

  // ---------- services ----------
  function saveService(data, input, uid) {
    unique(data.services, 'name', input.name, uid, 'Le nom du service');
    if (!input.client || !data.clients.some((c) => c.name === input.client)) throw new Error('Choisissez le client auquel appartient ce service.');
    let s = uid ? find(data, 'services', uid) : null;
    const old = s ? s.name : null;
    if (!s) { s = { uid: newUid(data) }; data.services.push(s); }
    Object.assign(s, input);
    cascade(data, [['equipements', 'zone'], ['pannes', 'zone'], ['interventions', 'zone'], ['preventif', 'zone'], ['ot', 'zone']], old, s.name);
    return s;
  }
  function deleteService(data, uid) {
    const s = find(data, 'services', uid);
    if (!s) return;
    const n = data.equipements.filter((e) => e.zone === s.name).length;
    if (n) throw new Error(`${n} appareil${n > 1 ? 's sont rattachés' : ' est rattaché'} au service « ${s.name} ». Réaffectez-${n > 1 ? 'les' : 'le'} à un autre service avant de le supprimer.`);
    data.services = data.services.filter((x) => x !== s);
  }

  // ---------- team ----------
  function saveMember(data, input, uid) {
    unique(data.equipe, 'name', input.name, uid, 'Le nom');
    let m = uid ? find(data, 'equipe', uid) : null;
    const old = m ? m.name : null;
    if (!m) { m = { uid: newUid(data), active: true }; data.equipe.push(m); }
    Object.assign(m, input);
    cascade(data, [['equipements', 'resp'], ['pannes', 'resp'], ['interventions', 'tech'], ['preventif', 'resp'], ['ot', 'resp']], old, m.name);
    return m;
  }
  function deleteMember(data, uid) {
    data.equipe = data.equipe.filter((m) => m.uid !== uid);
  }

  // ---------- parts ----------
  function savePart(data, input, uid) {
    unique(data.pieces, 'ref', input.ref, uid, 'La référence');
    let p = uid ? find(data, 'pieces', uid) : null;
    const old = p ? p.ref : null;
    if (!p) { p = { uid: newUid(data) }; data.pieces.push(p); }
    Object.assign(p, input);
    cascade(data, [['pannes', 'part'], ['interventions', 'partRef']], old, p.ref);
    return p;
  }
  function deletePart(data, uid) {
    data.pieces = data.pieces.filter((p) => p.uid !== uid);
  }

  // ---------- preventive plan ----------
  function freqMonths(f) {
    const n = norm(f);
    if (!n) return null;
    let m = n.match(/(\d+)\s*mois/);
    if (m) return +m[1];
    m = n.match(/(\d+)\s*ans?\b/);
    if (m) return +m[1] * 12;
    if (n.startsWith('mensuel')) return 1;
    if (n.startsWith('bimestriel')) return 2;
    if (n.startsWith('trimestriel')) return 3;
    if (n.startsWith('semestriel')) return 6;
    if (n.startsWith('annuel')) return 12;
    return null;
  }

  function saveTask(data, input, uid) {
    const eq = data.equipements.find((e) => e.code === input.code);
    if (!eq) throw new Error('Choisissez un équipement existant.');
    if (input.due == null) throw new Error('La date d’échéance est obligatoire.');
    let t = uid ? find(data, 'preventif', uid) : null;
    if (!t) {
      t = { uid: newUid(data), id: nextId(data.preventif, 'PM-', 2), done: null, last: null, status: '' };
      data.preventif.push(t);
    }
    Object.assign(t, input, { name: eq.name, zone: eq.zone });
    if (t.done == null) t.status = 'À venir';
    return t;
  }
  function deleteTask(data, uid) {
    data.preventif = data.preventif.filter((t) => t.uid !== uid);
  }

  const ANNUAL_RE = /(annuel|annuelle|cst|stc|controle de securite|securite technique|revision generale)/;

  /* Record that a planned task was carried out: closes the task, adds the intervention (cost),
   * consumes the kit, updates the equipment (last annual service, hour counter), closes the
   * matching preventive work order and schedules the next occurrence. */
  function recordTask(data, uid, rec) {
    const t = find(data, 'preventif', uid);
    if (!t) throw new Error('Tâche introuvable.');
    if (rec.done == null) throw new Error('La date de réalisation est obligatoire.');
    const eq = data.equipements.find((e) => e.code === t.code);
    const piece = partBy(data, rec.part);
    const qty = rec.part ? rec.partQty || 1 : 0;
    Object.assign(t, { done: rec.done, resp: rec.resp, dur: rec.dur != null ? rec.dur : t.dur, status: 'Réalisée', comment: rec.comment || '' });
    if (rec.part) adjustStock(data, rec.part, -qty);
    const mo = Math.round((t.dur || 0) * hourlyRate(data));
    const cp = piece && piece.price != null ? Math.round(piece.price * qty) : 0;
    const summary = rec.total ? `Protocole : ${rec.checked}/${rec.total} points validés · ${rec.result}` : rec.result || 'Réalisée';
    data.interventions.push({
      uid: newUid(data), id: nextId(data.interventions, 'INT-'), date: rec.done, code: t.code, type: 'Préventif',
      desc: t.op, cause: `Plan préventif ${t.id}`, action: summary, tech: rec.resp, dur: t.dur, downtime: rec.downtime || 0,
      parts: rec.part ? (piece ? piece.name : rec.part) : '', partRef: rec.part || '', costMO: mo, costParts: cp, costOther: 0,
      costTotal: mo + cp, status: 'Terminée', panneRef: '', comment: [rec.comment, rec.failed && rec.failed.length ? `Non conforme : ${rec.failed.join(' ; ')}` : ''].filter(Boolean).join(' · '),
      name: eq ? eq.name : t.name, zone: eq ? eq.zone : t.zone,
    });
    if (eq) {
      if (ANNUAL_RE.test(norm(t.op)) && (eq.lastService == null || rec.done > eq.lastService)) eq.lastService = rec.done;
      if (rec.hours != null && (eq.hours == null || rec.hours >= eq.hours)) eq.hours = rec.hours;
    }
    for (const o of data.ot) {
      if (o.code === t.code && C.typeGroup(o.type) === 'Préventif' && o.planned === t.due && !['cloture', 'annule'].includes(norm(o.status))) {
        Object.assign(o, { status: 'Clôturé', closed: rec.done, realH: t.dur });
      }
    }
    const months = freqMonths(t.freq);
    let next = null;
    if (months) {
      const due = C.addMonths(rec.done, months);
      const exists = data.preventif.some((x) => x !== t && x.id === t.id && x.code === t.code && x.done == null && x.due != null && x.due > t.due);
      if (!exists) {
        next = Object.assign({}, t, { uid: newUid(data), last: rec.done, due, done: null, status: 'À venir', comment: '' });
        data.preventif.push(next);
      }
    }
    return { task: t, next };
  }

  function saveParams(data, input) {
    Object.assign(data.params, input);
  }

  function emptyData() {
    return ensure({
      equipements: [], pannes: [], interventions: [], preventif: [], ot: [], clients: [], services: [], equipe: [], pieces: [],
      params: { joursSem: 6, devise: 'TND', refDate: null, inclPrev: false, rate: null, targets: {} },
      meta: { seq: 0, registries: true },
    });
  }

  return {
    TABLES, PANNE_STATUS, isOpenPanne, ensure, hourlyRate, freqMonths, nextId, emptyData,
    saveClient, deleteClient, savePanne, deletePanne, saveEquipment, deleteEquipment, saveService, deleteService,
    saveMember, deleteMember, savePart, deletePart, saveTask, deleteTask, recordTask, saveParams,
  };
});
